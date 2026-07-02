from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel
import uvicorn
import tempfile, os, re, difflib, io, json
import textwrap
from datetime import datetime
from typing import Optional, List
from rapidfuzz import fuzz

# Document parsers
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# PDF report generation
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)

app = FastAPI(title="Publishing QA Validation API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── Helpers ────────────────────────────────────────────────────────────────

def extract_docx(path: str) -> dict:
    doc = Document(path)
    paragraphs = []
    headings   = []
    tables     = []
    current_page = 1

    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph as DocxParagraph
    from docx.table import Table as DocxTable

    ti = 0
    pi = 0
    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            para = DocxParagraph(child, doc)
            text  = para.text.strip()
            style = para.style.name if para.style else "Normal"

            # Detect page break in runs
            for run in para.runs:
                if run.text and '\f' in run.text:
                    current_page += 1

            # Detect explicit XML page break
            for br in para._element.findall('.//' + qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    current_page += 1

            # Detect section break (new section = new page in our sample doc)
            sectPr = para._element.find(qn('w:pPr') + '/' + qn('w:sectPr'))
            if sectPr is not None:
                current_page += 1

            if not text:
                continue

            entry = {
                "index"    : pi,
                "text"     : text,
                "style"    : style,
                "word_page": current_page,
            }
            paragraphs.append(entry)
            if "Heading" in style:
                headings.append(entry)
            pi += 1

        elif child.tag.endswith('tbl'):
            table = DocxTable(child, doc)
            rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            ti += 1
            tables.append({
                "table_index": ti,
                "rows": rows,
                "word_page": current_page,
            })

    full_text = "\n".join(p["text"] for p in paragraphs)
    return {
        "paragraphs": paragraphs,
        "headings"  : headings,
        "tables"    : tables,
        "full_text" : full_text,
    }


def extract_pdf(path: str) -> dict:
    doc   = fitz.open(path)
    pages = []
    for i, page in enumerate(doc):
        text   = page.get_text("text").strip()
        blocks = page.get_text("blocks")
        pages.append({
            "page_num" : i + 1,
            "text"     : text,
            "blocks"   : [b[4] for b in blocks if b[4].strip()],
        })
    full_text = "\n".join(p["text"] for p in pages)
    return {"pages": pages, "full_text": full_text, "total_pages": len(pages)}


# ─── Utility: find which PDF page contains a text snippet ───────────────────

def find_pdf_page(text: str, pdf_data: dict, min_chars: int = 15, threshold: int = 90) -> int | None:
    """Find the PDF page whose text best matches the given snippet, using
    RapidFuzz's partial_ratio (good for 'is this text contained somewhere in
    this page' style matching). Returns the highest-scoring page if its
    score clears `threshold`, else None."""
    norm_text = normalize(text)
    if len(norm_text) < min_chars:
        return None

    # partial_ratio works best on a reasonably sized snippet, not the whole
    # paragraph — long snippets dilute the score against short pages.
    snippet = norm_text[:200]

    best_page  = None
    best_score = 0
    for page in pdf_data["pages"]:
        score = fuzz.partial_ratio(snippet, normalize(page["text"]))
        if score > best_score:
            best_score = score
            best_page  = page["page_num"]

    return best_page if best_score >= threshold else None


def closest_pdf_page(text: str, pdf_data: dict, threshold: int = 0) -> int | None:
    """Return the PDF page most similar to the given text using RapidFuzz's
    token_set_ratio (robust to word order / extra or missing words — good
    for 'roughly which page is this near' when there's no exact match).
    URLs/DOIs are stripped before scoring since PDF text extraction often
    mangles them with stray spaces, which would otherwise drag the score
    down on the page the text actually belongs to."""
    norm_text = _strip_url_noise(text)[:500]
    if not norm_text:
        norm_text = normalize(text)[:500]
    best_page  = None
    best_score = 0
    for page in pdf_data["pages"]:
        page_text = _strip_url_noise(page["text"])[:5000]
        score = fuzz.token_set_ratio(norm_text, page_text)
        if score > best_score:
            best_score = score
            best_page  = page["page_num"]
    return best_page if best_score > threshold else None


def _best_page_by_partial(snippet: str, pages: list, page_indices: Optional[list] = None) -> tuple[Optional[int], float]:
    """Return (page_num, score) for the best partial match of snippet."""
    if not snippet:
        return None, 0.0
    snippet = normalize(snippet)[:220]
    if not snippet:
        return None, 0.0

    best_page = None
    best_score = 0.0
    indices = page_indices if page_indices is not None else list(range(len(pages)))
    for i in indices:
        score = fuzz.partial_ratio(snippet, normalize(pages[i]["text"]))
        if score > best_score:
            best_score = score
            best_page = pages[i]["page_num"]
    return best_page, float(best_score)


def build_word_to_pdf_page_map(word_data: dict, pdf_data: dict) -> dict:
    """Create a stable mapping from Word page number -> PDF page number.

    We anchor the map with headings that can be confidently located in the PDF,
    then interpolate pages between anchors. This avoids random global fuzzy jumps.
    """
    pages = pdf_data["pages"]
    n_pdf = max(1, len(pages))

    anchors = []
    for h in word_data.get("headings", []):
        text = normalize(h.get("text", ""))
        if len(text) < 8:
            continue
        if is_toc_line(text):
            continue

        page_num = find_pdf_page(text, pdf_data, min_chars=8, threshold=88)
        if page_num is None:
            page_num = closest_pdf_page(text, pdf_data, threshold=70)
        if page_num is not None:
            anchors.append((int(h.get("word_page", 1)), int(page_num)))

    if not anchors:
        # Fallback: identity mapping when no anchors are available.
        return {w: min(max(1, w), n_pdf) for w in range(1, n_pdf + 1)}

    # Keep monotonic anchor progression.
    anchors.sort(key=lambda x: x[0])
    monotonic = []
    last_pdf = 0
    for w, p in anchors:
        if p >= last_pdf:
            monotonic.append((w, p))
            last_pdf = p
    anchors = monotonic or anchors[:1]

    max_word_page = max([p.get("word_page", 1) for p in word_data.get("paragraphs", [])] + [1])
    page_map = {}

    first_w, first_p = anchors[0]
    for w in range(1, min(first_w, max_word_page + 1)):
        page_map[w] = max(1, first_p - (first_w - w))

    for i, (w1, p1) in enumerate(anchors):
        page_map[w1] = p1
        if i == len(anchors) - 1:
            for w in range(w1 + 1, max_word_page + 1):
                page_map[w] = min(n_pdf, p1 + (w - w1))
            break

        w2, p2 = anchors[i + 1]
        span = max(1, w2 - w1)
        for w in range(w1 + 1, w2):
            t = (w - w1) / span
            mapped = round(p1 + t * (p2 - p1))
            page_map[w] = min(n_pdf, max(1, mapped))

    return page_map


def resolve_exact_pdf_page(
    text: str,
    word_page: int,
    pdf_data: dict,
    page_map: dict,
    preferred_page: Optional[int] = None,
    local_only: bool = False,
) -> int:
    """Resolve the most likely exact PDF page for an error text.

    Search expected page first, then nearby pages, then full document.
    """
    pages = pdf_data["pages"]
    n = len(pages)
    if n == 0:
        return 1

    snippet = _strip_url_noise(text)[:220] or normalize(text)[:220]
    mapped_page = page_map.get(word_page) or min(max(1, word_page), n)
    expected_page = min(max(1, preferred_page or mapped_page), n)
    expected_idx = min(max(0, expected_page - 1), n - 1)

    # Near search window around expected page (best for keeping page accuracy)
    window_indices = list(range(max(0, expected_idx - 2), min(n, expected_idx + 3)))
    near_page, near_score = _best_page_by_partial(snippet, pages, window_indices)
    if near_page is not None and near_score >= 70:
        return int(near_page)

    if local_only:
        return int(expected_page)

    # Global exact-ish match
    global_page, global_score = _best_page_by_partial(snippet, pages)
    if global_page is not None and global_score >= 85:
        return int(global_page)

    # Semantic fallback restricted around expected page, then global
    near_text = "\n".join(pages[i]["text"] for i in window_indices)
    if fuzz.token_set_ratio(snippet, _strip_url_noise(near_text)[:5000]) >= 65:
        return int(expected_page)

    cp = closest_pdf_page(snippet, pdf_data, threshold=0)
    if cp is not None:
        return int(cp)

    return int(expected_page)


def is_toc_line(text: str) -> bool:
    """Return True if this line looks like a Table of Contents entry."""
    if re.search(r'\.{2,}', text):
        return True
    if re.match(r'^[\d\.]+\s+.{3,80}\s+\d{1,4}$', text.strip()):
        return True
    return False


def normalize(text: str) -> str:
    """Normalize text for comparison: collapse spaces, strip."""
    return re.sub(r'\s+', ' ', text).strip()


def _strip_url_noise(text: str) -> str:
    """Strip URLs/DOIs before fuzzy matching (NOT for display).

    PyMuPDF frequently inserts stray spaces inside long URLs/DOIs when they
    wrap across a line in the PDF (e.g. 'https://doi.org/...' becomes
    'https:// doi . org/...'). That noise drags down the similarity score
    against the page the text actually lives on, sometimes letting an
    unrelated, shorter page win instead. Citation/reference text is still
    identifiable from the author/title portion alone, so we drop URLs
    entirely for the purpose of scoring page matches."""
    t = re.sub(r'https?\s*:?\s*/?\s*/\s*\S*', ' ', text)
    t = re.sub(r'\bdoi\.org\S*', ' ', t, flags=re.I)
    t = re.sub(r'\bwww\.\S*', ' ', t, flags=re.I)
    return normalize(t)


# ─── Individual check functions ─────────────────────────────────────────────

def _detect_printed_folio(text: str) -> int | None:
    """Try to find the printed folio (page number) on a page, preferring
    a standalone number on the first or last non-empty line (where folios
    usually sit), and falling back to any small number in the text."""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    candidates = []
    for l in (lines[:2] + lines[-2:]):
        if re.fullmatch(r'\d{1,4}', l):
            candidates.append(int(l))
    if candidates:
        return candidates[0]

    numbers = re.findall(r'(?<!\d)(\d{1,4})(?!\d)', text)
    nums    = [int(n) for n in numbers if 1 <= int(n) <= 9999]
    return nums[0] if nums else None


def check_page_sequence(pdf_data):
    """Check that printed folios increase in order from one page to the next.
    Does NOT assume numbering starts at 1 — front matter, prelims, etc. can
    use any starting number or scheme; we only care that the sequence is
    consistent and increasing."""
    errors   = []
    detected = []  # list of (pdf_page_num, printed_number)

    for page in pdf_data["pages"]:
        folio = _detect_printed_folio(page["text"])
        if folio is not None:
            detected.append((page["page_num"], folio))

    for i in range(1, len(detected)):
        prev_pdf_page, prev_num = detected[i - 1]
        curr_pdf_page, curr_num = detected[i]

        if curr_num <= prev_num:
            errors.append({
                "check"        : "Page Number Sequence & Folio",
                "page"         : str(curr_pdf_page),
                "search_string": str(curr_num),
                "location": (
                    f"The printed page number on Page {curr_pdf_page} is {curr_num}, "
                    f"but the previous page was numbered {prev_num}. "
                    f"Page numbers should increase in order — Page {curr_pdf_page} may be "
                    f"out of sequence, duplicated, or misnumbered."
                ),
            })

    return errors


def check_running_heads(pdf_data):
    errors = []
    for page in pdf_data["pages"]:
        lines = [l.strip() for l in page["text"].split("\n") if l.strip()]
        if not lines:
            continue
        head = lines[0]
        if len(head) > 120:
            errors.append({
                "check"        : "Running Head Style & Position",
                "page"         : str(page["page_num"]),
                "search_string": head[:50],
                "location": (
                    f"The header text at the top of Page {page['page_num']} is too long and may be "
                    f"cut off or placed in the wrong position. "
                    f"Found: '{head[:60]}…'"
                ),
            })
    return errors


def check_slug_line(pdf_data):
    errors = []
    slug_pattern   = re.compile(r'\b(Ch|Chapter|Sec|Section|Part|Unit|Module)\b', re.I)
    first_lines = []
    for page in pdf_data["pages"]:
        lines = [l.strip() for l in page["text"].split("\n") if l.strip()]
        first_lines.append(lines[0] if lines else "")

    from collections import Counter
    line_counts = Counter(l for l in first_lines if l)
    repeating_slugs = {line for line, cnt in line_counts.items() if cnt >= 3}

    for page in pdf_data["pages"]:
        lines = [l.strip() for l in page["text"].split("\n") if l.strip()]
        if not lines:
            continue
        first_line = lines[0]

        if re.match(r'^[\-\—\–\d\s]+$', first_line):
            continue

        if not slug_pattern.search(first_line):
            if first_line in repeating_slugs:
                errors.append({
                    "check"        : "Running Head & File Name",
                    "page"         : str(page["page_num"]),
                    "search_string": first_line[:50],
                    "location": (
                        f"The slug line on Page {page['page_num']} reads '{first_line[:80]}' — "
                        f"this appears to be just the book title with no chapter or section identifier. "
                        f"The slug should include the chapter name or page range (e.g. 'Chapter 2 | {first_line[:40]}')."
                    ),
                })
            elif len(first_line) < 5 and not re.search(r'[a-zA-Z]', first_line):
                errors.append({
                    "check"        : "Running Head & File Name",
                    "page"         : str(page["page_num"]),
                    "search_string": first_line,
                    "location": (
                        f"The Running Head on Page {page['page_num']} appears to be missing or corrupt. "
                        f"Found: '{first_line[:60]}'. "
                        f"Expected a chapter name or section title here."
                    ),
                })
    return errors


import os as _os
DEBUG_MATCHING = _os.environ.get("QC_DEBUG_MATCHING", "0") == "1"


def check_word_comparison(word_data, pdf_data, page_map=None):
    """Compare Word paragraphs against the PDF text using RapidFuzz.

    For each paragraph we score it against every PDF page with
    fuzz.partial_ratio (catches the paragraph appearing as a substring/near-
    substring of the page, robust to spacing/punctuation differences) and
    pick the best-scoring page:
      - score >= MATCH_THRESHOLD        → present, no error
      - SOFT_THRESHOLD <= score < MATCH → present but altered (spacing/wording)
      - score < SOFT_THRESHOLD          → treated as missing; closest_pdf_page
                                           (token_set_ratio) gives a best-guess page

    Page lookup is locality-aware: we search near the last matched page
    first, since Word and PDF content follow the same reading order. This
    avoids snapping to a distant, unrelated page that happens to share a
    short phrase (e.g. a repeated running head).
    """
    MATCH_THRESHOLD = 90   # confident match — text is present as-is
    SOFT_THRESHOLD  = 75   # present but altered — flag as a spacing/wording issue

    errors = []
    matched_items = []

    def find_normalized_offset(page_text: str, snippet: str) -> int:
        norm_page = normalize(page_text)
        norm_snip = normalize(snippet)
        idx = norm_page.find(norm_snip)
        if idx != -1:
            return idx
        for length in (120, 90, 60, 40, 20, 10):
            if len(norm_snip) > length:
                sub = norm_snip[:length]
                idx = norm_page.find(sub)
                if idx != -1:
                    return idx
        return 0

    def get_lis_indices(arr: list[int]) -> set[int]:
        if not arr:
            return set()
        n = len(arr)
        dp = [1] * n
        parent = [-1] * n
        for i in range(1, n):
            for j in range(i):
                if arr[j] < arr[i] and dp[j] + 1 > dp[i]:
                    dp[i] = dp[j] + 1
                    parent[i] = j
        
        max_len_idx = 0
        for i in range(1, n):
            if dp[i] > dp[max_len_idx]:
                max_len_idx = i
                
        lis_indices = set()
        curr = max_len_idx
        while curr != -1:
            lis_indices.add(curr)
            curr = parent[curr]
        return lis_indices

    word_paras    = word_data["paragraphs"]
    heading_texts = set(h["text"].strip() for h in word_data["headings"])
    pages   = pdf_data["pages"]
    n_pages = len(pages)
    last_idx = 0  # 0-based index of the last page we matched on, for locality
    last_matched_context = None  # tail of the last paragraph confirmed present in the PDF; anchors MISSING boxes near the real gap
    last_matched_para_text = None
    page_map = page_map or {}

    def search_order(start_idx):
        return list(range(start_idx, n_pages)) + list(range(start_idx - 1, -1, -1))

    for para_idx, para_obj in enumerate(word_paras):
        para      = para_obj["text"]
        word_page = para_obj.get("word_page", 1)

        # Explicit marker/tag mismatch detection from Word source
        # (examples: <NL>, <H1>, <B>, etc.). If present in Word but absent on
        # the mapped PDF page, report it directly.
        marker_tokens = re.findall(r'<[A-Za-z0-9/]{1,20}>', para)
        if marker_tokens:
            expected_page = resolve_exact_pdf_page(
                para,
                word_page,
                pdf_data,
                page_map,
                preferred_page=(pages[last_idx]["page_num"] if 0 <= last_idx < n_pages else None),
                local_only=True,
            )
            page_text = pages[max(0, min(n_pages - 1, expected_page - 1))]["text"]
            for tok in marker_tokens:
                if tok not in page_text:
                    errors.append({
                        "check": "Word-to-Word Comparison",
                        "page": str(expected_page),
                        "search_string": tok,
                        "location": (
                            f"Page {expected_page}: marker '{tok}' is in Word but missing in PDF."
                        ),
                    })

        if len(para) < 20:
            continue
        if is_toc_line(para):
            continue
        if para.strip() in heading_texts:
            continue
        if re.match(r'^(\d+[\.\d]*)\s+\w', para) and len(para) < 60:
            continue

        norm_para = normalize(para)
        match_para = _strip_url_noise(para) or norm_para
        snippet   = match_para[:200]
        order     = search_order(last_idx)

        best_idx   = None
        best_score = 0
        for idx in order:
            norm_page = _strip_url_noise(pages[idx]["text"]) or normalize(pages[idx]["text"])
            score = fuzz.partial_ratio(snippet, norm_page)
            if score > best_score:
                best_score = score
                best_idx   = idx
            # early exit once we've found a confident match near our locality —
            # no need to keep scanning the whole book
            if score >= MATCH_THRESHOLD:
                break

        if best_score >= MATCH_THRESHOLD:
            last_idx = best_idx  # advance reading position
            # Remember the tail of this paragraph — if the *next* paragraph
            # turns out to be missing from the PDF, this is the real text
            # it should immediately follow, so we can anchor the box there
            # instead of guessing a direction (top/bottom/margin).
            last_matched_context = norm_para[-100:]
            last_matched_para_text = para
            
            offset = find_normalized_offset(pages[best_idx]["text"], snippet)
            matched_items.append({
                "word_idx": para_idx,
                "pdf_page": pages[best_idx]["page_num"],
                "pdf_pos": best_idx * 1000000 + offset,
                "text": para
            })
            continue  # Found — no error

        if best_score >= SOFT_THRESHOLD:
            last_idx = best_idx
            last_matched_context = norm_para[-100:]
            last_matched_para_text = para
            preferred = pages[best_idx]["page_num"] if best_idx is not None else pages[last_idx]["page_num"]
            exact_page = resolve_exact_pdf_page(
                norm_para,
                word_page,
                pdf_data,
                page_map,
                preferred_page=preferred,
                local_only=False,
            )
            errors.append({
                "check"   : "Word-to-Word Comparison",
                "page"    : str(exact_page),
                "search_string": norm_para[:120],
                "location": (
                    f"Page {exact_page}: text is present but changed slightly (spacing/punctuation/wording). "
                    f"Check this text: '{norm_para[:80]}…'"
                ),
            })
            matched_page_idx = exact_page - 1
            offset = find_normalized_offset(pages[max(0, min(n_pages - 1, matched_page_idx))]["text"], snippet)
            matched_items.append({
                "word_idx": para_idx,
                "pdf_page": exact_page,
                "pdf_pos": matched_page_idx * 1000000 + offset,
                "text": para
            })
            continue

        # For missing text, keep page assignment local to reading order.
        # This avoids jumping to unrelated pages that happen to share one word.
        local_hint = pages[last_idx]["page_num"] if 0 <= last_idx < n_pages else None
        exact_page = resolve_exact_pdf_page(
            norm_para,
            word_page,
            pdf_data,
            page_map,
            preferred_page=local_hint,
            local_only=True,
        )
        pdf_page = str(exact_page)

        if DEBUG_MATCHING:
            match_text = _strip_url_noise(norm_para)[:500] or norm_para[:500]
            closest_page = resolve_exact_pdf_page(norm_para, word_page, pdf_data, page_map)
            print("\n" + "=" * 80)
            print(f"[QC_DEBUG] Word paragraph reported as MISSING / wrong page")
            print(f"[QC_DEBUG] Word para text : {norm_para[:120]!r}")
            print(f"[QC_DEBUG] partial_ratio best score: {best_score} (page {pages[best_idx]['page_num'] if best_idx is not None else '—'})")
            print(f"[QC_DEBUG] closest_pdf_page result : page {closest_page}")
            print(f"[QC_DEBUG] Per-page token_set_ratio scores:")
            scored = []
            for p in pages:
                page_text = _strip_url_noise(p["text"])[:5000] or normalize(p["text"])[:5000]
                tscore = fuzz.token_set_ratio(match_text, page_text)
                scored.append((p["page_num"], tscore))
            for pnum, tscore in sorted(scored, key=lambda x: -x[1])[:5]:
                print(f"[QC_DEBUG]   page {pnum:>3}  token_set_ratio={tscore}")
            print("=" * 80 + "\n")

        # The truncation often happens mid-paragraph (e.g. a sentence gets
        # cut off partway through), not cleanly between two paragraphs. In
        # that case the useful anchor is the tail of THIS paragraph's own
        # matched portion — whatever part of it genuinely is on the page —
        # rather than the previous paragraph entirely.
        #
        # We want the matching run closest to the actual cutoff, not just
        # the single LONGEST matching run in the whole paragraph — an
        # earlier stretch of plain prose can easily be longer than the
        # stretch right before the cut (e.g. a smart-quote mismatch inside
        # a blockquote can break the match early, well before the real
        # gap), which would anchor the box to the wrong, earlier spot.
        # get_matching_blocks() gives every matching run; we take the one
        # that reaches furthest into the paragraph (largest end position),
        # among those substantial enough to be a real match rather than
        # coincidental short overlap.
        own_context = None
        candidate_idx = best_idx if best_idx is not None else last_idx
        if 0 <= candidate_idx < n_pages:
            candidate_page_text = (
                _strip_url_noise(pages[candidate_idx]["text"])
                or normalize(pages[candidate_idx]["text"])
            )
            if candidate_page_text:
                sm = difflib.SequenceMatcher(None, match_para, candidate_page_text, autojunk=False)
                good_blocks = [b for b in sm.get_matching_blocks() if b.size >= 15]
                if good_blocks:
                    last_block = max(good_blocks, key=lambda b: b.a + b.size)
                    own_context = match_para[:last_block.a + last_block.size][-100:]

        before_text = last_matched_para_text
        
        after_text = None
        for j in range(para_idx + 1, len(word_paras)):
            p = word_paras[j]["text"]
            if len(p) >= 20 and not is_toc_line(p) and p.strip() not in heading_texts:
                norm_p = normalize(p)
                match_p = _strip_url_noise(p) or norm_p
                snippet_p = match_p[:200]
                found_match = False
                for p_idx in search_order(last_idx):
                    norm_page = _strip_url_noise(pages[p_idx]["text"]) or normalize(pages[p_idx]["text"])
                    if fuzz.partial_ratio(snippet_p, norm_page) >= MATCH_THRESHOLD:
                        found_match = True
                        break
                if found_match:
                    after_text = p
                    break

        errors.append({
            "check"   : "Word-to-Word Comparison",
            "page"    : pdf_page,
            "search_string": norm_para[:120],
            "location": (
                f"Page {pdf_page}: this text is missing in PDF. "
                f"Check this text: '{norm_para[:80]}…'"
            ),
            "anchor_context": own_context or last_matched_context,
            "before_text": before_text,
            "after_text": after_text,
        })
        # Advance minimally to keep subsequent page guesses monotonic.
        last_idx = min(n_pages - 1, max(last_idx, exact_page - 1))

    # ─── Order Mismatch Detection ───
    if matched_items:
        # Sort matched items by their actual order of appearance in the PDF
        sorted_by_pdf = sorted(matched_items, key=lambda x: x["pdf_pos"])
        arr = [x["word_idx"] for x in sorted_by_pdf]
        lis_indices = get_lis_indices(arr)
        
        for i, item in enumerate(sorted_by_pdf):
            if i not in lis_indices:
                # This item is out of order compared to the general flow of the document!
                # We can report it as an order mismatch.
                pdf_page = item["pdf_page"]
                errors.append({
                    "check"   : "Word-to-Word Comparison",
                    "page"    : str(pdf_page),
                    "search_string": item["text"][:120],
                    "location": (
                        f"Page {pdf_page}: Paragraph order mismatch. "
                        f"This paragraph appears out of order in the PDF compared to the Word document: "
                        f"'{item['text'][:80]}...'"
                    ),
                })

    return errors

def check_typos(word_data, pdf_data):
    errors = []
    common_typos = {
        "teh"       : "the",
        "recieve"   : "receive",
        "occured"   : "occurred",
        "seperete"  : "separate",
        "definately": "definitely",
        "accomodate": "accommodate",
        "untill"    : "until",
        "publsihed" : "published",
        "refernce"  : "reference",
        "Figuure"   : "Figure",
        "Tablel"    : "Table",
    }
    for src, correction in common_typos.items():
        pattern = re.compile(r'\b' + re.escape(src) + r'\b', re.I)
        for page in pdf_data["pages"]:
            if pattern.search(page["text"]):
                errors.append({
                    "check"   : "Typos",
                    "page"    : str(page["page_num"]),
                    "search_string": src,
                    "location": (
                        f"The word '{src}' on Page {page['page_num']} appears to be a spelling mistake. "
                        f"It should be '{correction}'."
                    ),
                })

    # Catch accidental repeated adjacent words like "are Are".
    # This is common in manual edits and should be reported even if the
    # paragraph otherwise matches the Word source.
    repeated_word_re = re.compile(r"\b([A-Za-z]{2,})\b\s+\b\1\b", re.I)
    for page in pdf_data["pages"]:
        text = page["text"]
        for m in repeated_word_re.finditer(text):
            repeated = m.group(0)
            word = m.group(1)
            errors.append({
                "check": "Typos",
                "page": str(page["page_num"]),
                "search_string": repeated,
                "location": (
                    f"The phrase '{repeated}' on Page {page['page_num']} repeats the same word twice. "
                    f"Please keep only one '{word}'."
                ),
            })

    return errors


def check_missing_content(word_data, pdf_data, page_map=None):
    errors = []
    page_map = page_map or {}

    for heading in word_data["headings"]:
        h = heading["text"].strip()
        if len(h) < 4:
            continue
        if is_toc_line(h):
            continue

        norm_h = normalize(h)
        word_page = heading.get("word_page", 1)

        found_page = None
        for page in pdf_data["pages"]:
            norm_page = normalize(page["text"])
            if norm_h.lower() in norm_page.lower():
                found_page = page["page_num"]
                break
            elif len(norm_h) > 20 and norm_h[:20].lower() in norm_page.lower():
                found_page = page["page_num"]
                break

        if found_page is not None:
            continue

        exact_page = resolve_exact_pdf_page(norm_h, word_page, pdf_data, page_map)
        display_page = str(exact_page)

        errors.append({
            "check"   : "Missing Content",
            "page"    : display_page,
            "search_string": norm_h,
            "location": (
                f"Page {display_page}: heading is in Word but missing in PDF. "
                f"Missing heading: '{h}'"
            ),
        })

    return errors[:6]


def check_content_order(word_data, pdf_data):
    errors = []
    positions = []
    for heading in word_data["headings"]:
        h   = normalize(heading["text"])
        if is_toc_line(h):
            continue
        found_page = find_pdf_page(h, pdf_data)
        positions.append((heading["text"], found_page))

    for i in range(1, len(positions)):
        p_curr = positions[i][1]
        p_prev = positions[i-1][1]
        if p_curr is not None and p_prev is not None:
            if p_curr < p_prev:
                errors.append({
                    "check"   : "Content Order",
                    "page"    : str(p_curr),
                    "location": (
                        f"The section '{positions[i][0][:60]}' appears on Page {p_curr}, "
                        f"which is before '{positions[i-1][0][:60]}' on Page {p_prev}. "
                        f"The order of sections may be wrong."
                    ),
                })
    return errors


def check_headings(word_data, pdf_data, page_map=None):
    errors = []
    page_map = page_map or {}
    num_re = re.compile(r'^(\d+[\.\d]*)\s')
    for h in word_data["headings"]:
        m = num_re.match(h["text"])
        if not m:
            continue
        num = m.group(1)
        found_page = None
        for page in pdf_data["pages"]:
            if num in page["text"]:
                found_page = page["page_num"]
                break
        if found_page is None:
            word_page_raw = int(h.get("word_page", 1))
            exact_page = resolve_exact_pdf_page(h["text"], word_page_raw, pdf_data, page_map)
            word_page = str(exact_page)
            errors.append({
                "check"   : "Heading Levels & Numbering",
                "page"    : word_page,
                "search_string": num,
                "location": (
                    f"Page {word_page}: heading number '{num}' is missing or changed in PDF."
                ),
            })
    return errors[:5]


def check_equations(word_data, pdf_data):
    errors = []
    eq_re = re.compile(r'[A-Za-z]\s*=\s*[\w\(\)\+\-\*\/\^]+')
    for page in pdf_data["pages"]:
        matches = eq_re.findall(page["text"])
        for m in matches:
            if "?" in m or "□" in m or "■" in m:
                errors.append({
                    "check"   : "Equations",
                    "page"    : str(page["page_num"]),
                    "location": (
                        f"An equation on Page {page['page_num']} did not render correctly and contains "
                        f"broken or missing characters. Check: '{m[:60]}'"
                    ),
                })
    return errors


def check_special_chars(pdf_data):
    """Flag only characters that indicate a genuine rendering/font problem —
    missing-glyph boxes and the Unicode replacement character. Normal,
    legitimate typographic symbols (©, ®, ™, …, §, ¶, †, ‡, bullets, etc.)
    are NOT flagged just because they occur — they are expected content,
    not errors."""
    errors = []
    bad = re.compile(r'[□■▢◻\ufffd]')

    SYMBOL_NAMES = {
        '□': 'empty box / missing glyph (□)',
        '■': 'filled box / missing glyph (■)',
        '▢': 'empty box / missing glyph (▢)',
        '◻': 'empty box / missing glyph (◻)',
        '\ufffd': 'unknown-character replacement symbol ()',
    }

    for page in pdf_data["pages"]:
        found = bad.findall(page["text"])
        if found:
            unique = list(set(found))
            named  = [SYMBOL_NAMES.get(c, f"'{c}'") for c in unique[:8]]
            errors.append({
                "check"   : "Special Characters & Symbols",
                "page"    : str(page["page_num"]),
                "location": (
                    f"Broken or missing glyph(s) found on Page {page['page_num']} — "
                    f"{', '.join(named)}. "
                    f"This usually means a special font is not properly embedded in the PDF. "
                    f"Check the source file for these characters."
                ),
            })
    return errors[:4]


def check_footnotes(word_data, pdf_data):
    errors = []
    fn_re = re.compile(r'\[\d+\]|\(\d+\)|(?<!\d)\d{1,2}(?!\d)\s*$', re.M)
    for page in pdf_data["pages"]:
        refs = fn_re.findall(page["text"])
        for r in refs:
            num = re.search(r'\d+', r)
            if num and int(num.group()) > 200:
                errors.append({
                    "check"   : "Footnote Citation & Placement",
                    "page"    : str(page["page_num"]),
                    "location": (
                        f"The number {num.group()} on Page {page['page_num']} is being detected as a "
                        f"footnote reference, but it is too large to be a real footnote number. "
                        f"It may be a year (like 2023) being mistakenly flagged."
                    ),
                })
    return errors[:4]


def check_unwanted_chars(pdf_data):
    errors = []
    patterns = {
        r'\$\$\$'           : "Dollar signs ('$$$') used as a placeholder were found. Please replace with the actual content.",
        r'\bxxx\b'          : "'xxx' placeholder text was found. This must be replaced with real content before publishing.",
        r'\bXXX\b'          : "'XXX' placeholder text was found. This must be replaced with real content before publishing.",
        r'\bTBD\b'          : "'TBD' (To Be Decided) was found. This content was never finalized — please update it before publishing.",
        r'\bLorem\s+Ipsum\b': "'Lorem Ipsum' dummy filler text was found. This must be replaced with the actual content.",
        r'\bFPO\b'          : "'FPO' (For Position Only) marker was found. This means a placeholder image was never replaced with the real image.",
        r'\bPLACEHOLDER\b'  : "'PLACEHOLDER' text was found. This filler was never replaced with actual content.",
    }
    for pattern, explanation in patterns.items():
        rx = re.compile(pattern, re.I)
        for page in pdf_data["pages"]:
            if rx.search(page["text"]):
                errors.append({
                    "check"   : "Unwanted Characters",
                    "page"    : str(page["page_num"]),
                    "search_string": pattern.replace('\\b', '').replace('\\', ''),
                    "location": f"On Page {page['page_num']}: {explanation}",
                })
    return errors


def check_citations(pdf_data):
    errors = []
    cit_re = re.compile(r'\(([A-Z][a-z]+(?:,?\s(?:et al\.?|&|and)\s[A-Z][a-z]+)?,?\s*\d{4})\)')
    for page in pdf_data["pages"]:
        citations = cit_re.findall(page["text"])
        for c in citations:
            ctx   = page["text"]
            idx   = ctx.find(c)
            if idx > 0:
                before = ctx[max(0, idx-5):idx]
                if re.search(r'[A-Z]$', before.strip()):
                    errors.append({
                        "check"   : "Citations & Placement",
                        "page"    : str(page["page_num"]),
                        "location": (
                            f"The citation '({c})' on Page {page['page_num']} appears to be placed in the "
                            f"middle of a sentence. Citations should normally appear at the "
                            f"end of a sentence."
                        ),
                    })
    return errors[:4]


def check_lists(pdf_data):
    errors = []
    list_re  = re.compile(r'^(\s*[\•\-\*]\s|\s*\d+\.\s)', re.M)
    blank_re = re.compile(r'\n{3,}')
    for page in pdf_data["pages"]:
        if list_re.search(page["text"]) and blank_re.search(page["text"]):
            errors.append({
                "check"   : "List Spacing",
                "page"    : str(page["page_num"]),
                "location": (
                    f"There are too many blank lines near the list items on Page {page['page_num']}. "
                    f"The spacing between list items should be consistent and not have extra gaps."
                ),
            })
    return errors[:3]


def check_font_consistency(pdf_data):
    errors = []
    allcaps = re.compile(r'\b[A-Z]{4,}\b')
    for page in pdf_data["pages"]:
        caps_count = len(allcaps.findall(page["text"]))
        if caps_count > 10:
            errors.append({
                "check"   : "Font Consistency",
                "page"    : str(page["page_num"]),
                "location": (
                    f"Page {page['page_num']} has {caps_count} words written in ALL CAPS, which is unusually high. "
                    f"This may indicate an incorrect font style or a formatting error."
                ),
            })
    return errors[:3]


def check_quotations(word_data, pdf_data, page_map=None):
    """Flag quoted text from the Word file that is genuinely missing from the
    PDF. Quote *style* differences (straight " vs curly “ ”) are normal and
    expected after layout — they are NOT errors and are not flagged here."""
    errors = []
    quote_re    = re.compile(r'["\u201c\u201d]([^""\u201c\u201d]{10,300})["\u201c\u201d]')
    
    word_quotes = []
    for para_obj in word_data["paragraphs"]:
        text = para_obj["text"]
        word_page = para_obj.get("word_page", 1)
        for q in quote_re.findall(text):
            word_quotes.append((q, word_page))

    page_map = page_map or {}

    for q, word_page in word_quotes[:20]:
        norm_q = normalize(q)

        # Try matching regardless of straight vs curly quote style / punctuation
        found_page = find_pdf_page(norm_q, pdf_data, min_chars=20)
        if found_page is not None:
            continue  # genuinely present — not an error

        # Also check with quote-style normalized away, in case the only
        # difference is straight vs curly quotes around similar wording
        loose_q = re.sub(r'[\u2018\u2019\u201c\u201d"\']', '', norm_q)
        loose_found = False
        for page in pdf_data["pages"]:
            loose_page = re.sub(r'[\u2018\u2019\u201c\u201d"\']', '', normalize(page["text"]))
            if loose_q[:30] in loose_page:
                loose_found = True
                break
        if loose_found:
            continue  # present, just a quote-style difference — not an error

        exact_page = resolve_exact_pdf_page(norm_q, word_page, pdf_data, page_map)
        display_page = str(exact_page)
        errors.append({
            "check"   : "Quotations",
            "page"    : display_page,
            "search_string": norm_q[:120],
            "location": (
                f"Page {display_page}: quoted text is missing or changed in PDF: '\"{norm_q[:60]}\"'"
            ),
        })
    return errors[:4]


def check_fpo(pdf_data):
    errors = []
    fpo_re = re.compile(r'\bFPO\b|\bFor\s+Position\s+Only\b|\bPLACEHOLDER\s+IMAGE\b', re.I)
    for page in pdf_data["pages"]:
        if fpo_re.search(page["text"]):
            errors.append({
                "check"   : "FPO / Placeholder Images",
                "page"    : str(page["page_num"]),
                "search_string": "FPO",
                "location": (
                    f"A placeholder image marker (FPO / For Position Only) was found on Page {page['page_num']}. "
                    f"This means the real image was never placed. Please replace it before publishing."
                ),
            })
    return errors


def check_tables(word_data, pdf_data, page_map=None):
    errors = []
    page_map = page_map or {}
    for ti, table in enumerate(word_data["tables"]):
        if not table["rows"]:
            continue
        header_cell = table["rows"][0][0] if table["rows"][0] else ""
        if not header_cell:
            continue
        norm_header = normalize(header_cell)
        found_page  = find_pdf_page(norm_header, pdf_data, min_chars=10)

        if found_page is None:
            word_page = table.get("word_page", 1)
            exact_page = resolve_exact_pdf_page(norm_header, word_page, pdf_data, page_map)
            display_page = str(exact_page)
            errors.append({
                "check"   : "Table Style Consistency",
                "page"    : display_page,
                "search_string": norm_header[:120],
                "location": (
                    f"Page {display_page}: table {ti+1} is missing or moved in PDF. "
                    f"Table header: '{norm_header[:40]}'"
                ),
            })
    return errors[:4]


# ─── Report generation models ───────────────────────────────────────────────

class ErrorItem(BaseModel):
    check: str
    page: str
    location: str
    search_string: Optional[str] = None


class ReportRequest(BaseModel):
    errors: List[ErrorItem]
    total_errors: int
    total_pages: int
    affected_pages: List[int] = []
    checks_run: int = 0
    format: str = "docx"          # "docx" | "pdf"
    book_title: Optional[str] = None


# ─── Highlighted PDF generation ─────────────────────────────────────────────

# Use red for all highlighted PDF annotations and boxes.
RED_HIGHLIGHT_COLOR = (1.0, 0.0, 0.0)
CHECK_COLORS = {}
DEFAULT_HIGHLIGHT_COLOR = RED_HIGHLIGHT_COLOR


def _extract_snippet(location: str) -> Optional[str]:
    """Pull the most likely searchable text snippet out of an error's
    'location' message — these messages consistently wrap the relevant
    quoted text in single quotes, e.g. "...: 'some excerpt here…'". Picks
    the longest quoted span found, since shorter ones are often labels."""
    candidates = re.findall(r"'([^']{2,300})'", location)
    if not candidates:
        return None
    snippet = max(candidates, key=len)
    # Trim a trailing ellipsis (added by our own truncation) and any
    # trailing partial word it may have cut off mid-way.
    snippet = snippet.rstrip(" …").strip()
    return snippet or None


def _find_text_rects(page, snippet: str, allow_single_word_fallback: bool = True):
    """Progressively shorten the snippet until PyMuPDF finds it on the page.
    Returns (rects, matched_snippet) or ([], None) if nothing was found."""
    if not snippet:
        return [], None

    lengths_to_try = [len(snippet), 120, 90, 70, 45, 25, len(snippet)]
    seen_lengths = set()
    for length in lengths_to_try:
        length = min(length, len(snippet))
        if length < 3 or length in seen_lengths:
            continue
        seen_lengths.add(length)
        candidate = snippet[:length].strip()
        if len(candidate) < 3:
            continue
        rects = page.search_for(candidate, quads=False)
        if rects:
            return rects, candidate

    if allow_single_word_fallback:
        # Exact-word fallback: try longest meaningful words from snippet.
        words = [w for w in re.findall(r"[A-Za-z0-9][A-Za-z0-9_\-]{2,}", snippet) if len(w) >= 3]
        words.sort(key=len, reverse=True)
        for w in words[:8]:
            rects = page.search_for(w, quads=False)
            if rects:
                return [rects[0]], w

    return [], None


def _find_para_edge_rect(page, text: str, search_from_end: bool):
    """Locate the bounding box of the start or end of a paragraph text on page."""
    if not text:
        return None
    norm = normalize(text)
    words = [w for w in re.findall(r"[A-Za-z0-9][A-Za-z0-9_\-]{2,}", norm) if len(w) >= 4]
    if not words:
        return None

    # First try exact end/start phrases
    for n in (8, 6, 4, 3):
        if len(words) < n:
            continue
        if search_from_end:
            phrase = " ".join(words[-n:])
        else:
            phrase = " ".join(words[:n])

        rects = page.search_for(phrase, quads=False)
        if rects:
            if search_from_end:
                return max(rects, key=lambda r: r.y1)
            else:
                return min(rects, key=lambda r: r.y0)

    # Fallback for split paragraphs: slide window of single words and score context
    word_sequence = reversed(words) if search_from_end else words
    for i, w in enumerate(word_sequence):
        if i > 30:
            break
        rects = page.search_for(w, quads=False)
        if not rects:
            continue
            
        best_rect = None
        best_score = 0
        for r in rects:
            # Context window around the word
            probe = fitz.Rect(r.x0 - 50, r.y0 - 40, page.rect.x1, min(page.rect.y1, r.y1 + 40))
            probe_text = normalize(page.get_text("text", clip=probe))
            score = fuzz.partial_ratio(norm[:200] if not search_from_end else norm[-200:], probe_text)
            if score > best_score:
                best_score = score
                best_rect = r
                
        if best_rect is not None and best_score >= 60:
            return best_rect

    return None


def _find_anchor_rect_for_missing(page, snippet, prefer_top: bool = False):
    """Find a nearby anchor rect for missing text placement.

    Searches progressively shorter phrases from the missing snippet. Since
    common words (e.g. "restoration") can appear multiple times on a page,
    every candidate match is verified by fuzzy-comparing the text around it
    against the snippet — the arrow should point at the passage that
    actually resembles the missing content, not just the first coincidental
    word match PyMuPDF happens to return.

    If NONE of the snippet's words appear anywhere on the page at all,
    we fall back to a real text block on the page so the arrow still
    points at content rather than disappearing entirely. When
    `prefer_top` is true, we bias that fallback toward the uppermost text
    block so running head/file-name errors land at the top of the page
    instead of in the body text.
    """
    if not snippet:
        return None

    norm_snippet = normalize(snippet)
    words = [w for w in re.findall(r"[A-Za-z0-9][A-Za-z0-9_\-]{2,}", snippet) if len(w) >= 3]

    def score_candidate(rect):
        # Look at a chunk of text starting at this rect and extending a
        # couple of lines down/right — enough to judge if this is really
        # the surrounding context of the missing passage.
        probe = fitz.Rect(rect.x0, rect.y0, page.rect.x1, min(page.rect.y1, rect.y1 + 40))
        probe_text = normalize(page.get_text("text", clip=probe))
        return fuzz.partial_ratio(norm_snippet[:150], probe_text[:250])

    best_rect  = None
    best_score = -1  # track best candidate regardless of threshold

    if words:
        # Prefer longer, more specific phrases first — less likely to have
        # false-positive matches elsewhere on the page.
        for n in (5, 4, 3, 2):
            if len(words) < n:
                continue
            phrase = " ".join(words[:n])
            for r in page.search_for(phrase, quads=False):
                s = score_candidate(r)
                if s > best_score:
                    best_score, best_rect = s, r
            if best_rect is not None and best_score >= 60:
                return best_rect  # confident match — stop early

        # Single-word fallback — accept the best-scoring occurrence even if
        # it's a weak match; a rough anchor is better than none.
        for w in sorted(words[:10], key=len, reverse=True):
            for r in page.search_for(w, quads=False):
                s = score_candidate(r)
                if s > best_score:
                    best_score, best_rect = s, r

    if best_rect is not None:
        return best_rect

    # None of the snippet's words exist on this page at all — this usually
    # means the error text was assigned to the wrong page. Previously this
    # fell back to pointing at "the largest text block on the page", but
    # that's an arbitrary guess (often a body paragraph in the middle of
    # the page) and made boxes land in seemingly random spots like page
    # center. Better to admit we have no real anchor and let the caller
    # use the reliable top-left/bottom fallback placement instead.
    return None



# FIX: scans the page's real text/image geometry and returns a rect that
# does NOT overlap any existing content or any previously-placed fallback
# box on this page. Previously every fallback box was pinned to a fixed
# top-left coordinate, so it routinely landed on top of real paragraphs or
# images (see screenshots: box covering "Restorative justice is..." body
# text, and a box fully covering a photo). This searches a grid of
# candidate positions — right margin (preferred), left margin, then a
# denser full-page scan — and picks the first one that's genuinely empty,
# preferring a spot near the anchor when one is available.
def _find_free_rect(page, box_w: float, box_h: float, avoid_rects: list, anchor: "fitz.Rect" = None) -> "fitz.Rect":
    page_rect = page.rect
    margin = 6.0

    occupied = avoid_rects[:]
    for b in page.get_text("blocks"):
        occupied.append(fitz.Rect(b[0], b[1], b[2], b[3]))
    for img in page.get_image_info():
        bbox = img.get("bbox")
        if bbox:
            occupied.append(fitz.Rect(bbox))

    def fits(rect: "fitz.Rect") -> bool:
        return all(not rect.intersects(o) for o in occupied)

    step_y = 16.0
    x_candidates = [
        page_rect.x1 - box_w - margin,   # right margin column (preferred)
        page_rect.x0 + margin,           # left margin column
    ]

    # Prefer a spot near the anchor first, expanding downward from it —
    # a box near the actual gap reads better than one parked at the top
    # of the page.
    if anchor is not None:
        y = max(page_rect.y0 + margin, anchor.y1 - step_y)
        while y + box_h <= page_rect.y1 - margin:
            for x in x_candidates:
                rect = fitz.Rect(x, y, x + box_w, y + box_h)
                if fits(rect):
                    return rect
            y += step_y

    # Full-page fallback — scan bottom-up instead of top-down, so an
    # unanchored box lands toward the lower half of the page rather than
    # the top.
    y = page_rect.y1 - margin - box_h
    while y >= page_rect.y0 + margin:
        for x in x_candidates:
            rect = fitz.Rect(x, y, x + box_w, y + box_h)
            if fits(rect):
                return rect
        y -= step_y

    # Last resort: dense page, no empty spot found — place bottom-right
    # instead of top-left.
    return fitz.Rect(page_rect.x1 - margin - box_w, page_rect.y1 - margin - box_h,
                      page_rect.x1 - margin, page_rect.y1 - margin)


def _find_context_anchor(page, context_text: str):
    """Locate the tail-end of the last real paragraph that WAS found on the
    page, so a missing-text box can be anchored at the exact spot the
    missing content should appear (right after this text) instead of a
    generic direction like top-left or bottom."""
    if not context_text:
        return None
    norm = normalize(context_text)
    if len(norm) < 12:
        return None

    for n_words in (12, 8, 5):
        words = norm.split()
        if len(words) < n_words:
            continue
        phrase = " ".join(words[-n_words:])
        rects = page.search_for(phrase, quads=False)
        if rects:
            # last occurrence on the page, in case the phrase repeats
            return max(rects, key=lambda r: r.y0)
    return None


def _find_corner_rect(page, box_w: float, box_h: float, avoid_rects: list, corner: str = "top-left") -> "fitz.Rect":
    """Find a free rect anchored to a page corner/edge instead of near a
    text anchor. Used to give the FIRST missing-content box on a page a
    fixed spot in the top-left corner, and every subsequent box on that
    page a spot along the bottom of the page (packed left-to-right, then
    stacking upward if the bottom row fills up)."""
    page_rect = page.rect
    margin = 6.0

    occupied = avoid_rects[:]
    for b in page.get_text("blocks"):
        occupied.append(fitz.Rect(b[0], b[1], b[2], b[3]))
    for img in page.get_image_info():
        bbox = img.get("bbox")
        if bbox:
            occupied.append(fitz.Rect(bbox))

    def fits(rect: "fitz.Rect") -> bool:
        return all(not rect.intersects(o) for o in occupied)

    step = 16.0

    if corner == "top-left":
        y = page_rect.y0 + margin
        while y + box_h <= page_rect.y1 - margin:
            rect = fitz.Rect(page_rect.x0 + margin, y, page_rect.x0 + margin + box_w, y + box_h)
            if fits(rect):
                return rect
            y += step
        # Dense page — no free spot found scanning down the left column;
        # just pin it to the corner anyway.
        return fitz.Rect(page_rect.x0 + margin, page_rect.y0 + margin,
                          page_rect.x0 + margin + box_w, page_rect.y0 + margin + box_h)

    # corner == "bottom": pack left-to-right along the bottom edge, and if
    # a row fills up, start a new row stacked above it.
    y = page_rect.y1 - margin - box_h
    while y >= page_rect.y0 + margin:
        x = page_rect.x0 + margin
        while x + box_w <= page_rect.x1 - margin:
            rect = fitz.Rect(x, y, x + box_w, y + box_h)
            if fits(rect):
                return rect
            x += box_w + 8.0
        y -= (box_h + 8.0)

    # Dense page fallback — bottom-right corner.
    return fitz.Rect(page_rect.x1 - margin - box_w, page_rect.y1 - margin - box_h,
                      page_rect.x1 - margin, page_rect.y1 - margin)


def _page_free_bands(page):
    """Return (col_left, col_right, top_band, bottom_band).

    top_band / bottom_band are fitz.Rect areas of genuine whitespace above
    the first line of body text and below the last line, spanning the full
    width of the text column. This mirrors how an editor drops margin notes
    into the existing gutter space at the top/bottom of a page instead of
    scattering a separate box out in the margin next to every single line —
    which is what made earlier versions of this feature feel disconnected
    from the page. If a page has no real top/bottom whitespace (dense page),
    both bands come back None and callers fall back to the old per-spot
    free-rect search.
    """
    blocks = [b for b in page.get_text("blocks") if b[4].strip()]
    if not blocks:
        return page.rect.x0 + 40, page.rect.x1 - 40, None, None

    col_left  = min(b[0] for b in blocks)
    col_right = max(b[2] for b in blocks)
    top_y     = min(b[1] for b in blocks)
    bottom_y  = max(b[3] for b in blocks)

    margin = 20.0
    min_gap = 30.0  # don't bother with a band that's too thin to hold a box

    top_band = None
    if top_y - page.rect.y0 - margin > min_gap:
        top_band = fitz.Rect(col_left, page.rect.y0 + margin, col_right, top_y - 6)

    bottom_band = None
    if page.rect.y1 - margin - bottom_y > min_gap:
        bottom_band = fitz.Rect(col_left, bottom_y + 6, col_right, page.rect.y1 - margin)

    return col_left, col_right, top_band, bottom_band


def _flow_place_boxes(band: "fitz.Rect", items: list, box_h: float, gap: float = 8.0):
    """Pack `items` left-to-right, top-to-bottom inside `band`.

    Uses two boxes per row when the band is wide enough for that (matches
    the paired side-by-side layout editors use for short margin notes),
    otherwise one full-width box per row. Returns (placed, leftover) where
    placed is a list of (item, rect) and leftover is whatever didn't fit.
    """
    if band is None or band.width <= 0 or band.height < box_h:
        return [], items

    two_up = band.width >= 320
    n_cols = 2 if two_up else 1
    box_w  = (band.width - gap) / n_cols

    placed = []
    x, y, col = band.x0, band.y0, 0
    for item in items:
        if y + box_h > band.y1:
            break
        rect = fitz.Rect(x, y, x + box_w, y + box_h)
        placed.append((item, rect))
        col += 1
        if col >= n_cols:
            col = 0
            x = band.x0
            y += box_h + gap
        else:
            x += box_w + gap

    leftover = items[len(placed):]
    return placed, leftover


def _draw_pointer_arrow(page, box_rect: "fitz.Rect", target_rect: "fitz.Rect", color) -> None:
    """Draw a small arrow from the fallback box to the target area (in between paragraphs)."""
    if not box_rect or not target_rect:
        return

    # Check if box is to the right of the target (e.g., in the right margin)
    if box_rect.x0 > target_rect.x1:
        start = fitz.Point(box_rect.x0, (box_rect.y0 + box_rect.y1) / 2)
        end   = fitz.Point(target_rect.x1 - 10, (target_rect.y0 + target_rect.y1) / 2)
        page.draw_line(start, end, color=color, width=1.2)
        page.draw_line(end, fitz.Point(end.x + 8, end.y - 4), color=color, width=1.2)
        page.draw_line(end, fitz.Point(end.x + 8, end.y + 4), color=color, width=1.2)
    else:
        start = fitz.Point(box_rect.x1, (box_rect.y0 + box_rect.y1) / 2)
        end   = fitz.Point(target_rect.x0 + 10, (target_rect.y0 + target_rect.y1) / 2)
        page.draw_line(start, end, color=color, width=1.2)
        page.draw_line(end, fitz.Point(end.x - 8, end.y - 4), color=color, width=1.2)
        page.draw_line(end, fitz.Point(end.x - 8, end.y + 4), color=color, width=1.2)


def _insert_wrapped_box_text(page, rect: "fitz.Rect", text: str, color) -> None:
    """Insert wrapped visible text into a box line by line.

    Wrap width scales with the box's actual pixel width (~3.7pt per
    character at 7pt Helvetica) instead of a fixed character count, since
    boxes are no longer a fixed narrow size — a full-width box at the top/
    bottom of the page should use that width instead of wrapping as if it
    were still a narrow margin box.
    """
    inner = fitz.Rect(rect.x0 + 4, rect.y0 + 4, rect.x1 - 4, rect.y1 - 4)
    wrap_chars = max(20, int(inner.width / 3.7))
    lines = []
    for raw_line in (text.splitlines() or [text]):
        wrapped = textwrap.wrap(raw_line, width=wrap_chars) or [""]
        lines.extend(wrapped)

    font_size = 7.0
    line_height = 8.2
    y = inner.y0 + 1
    max_lines = max(3, int(inner.height / line_height))
    for line in lines[:max_lines]:
        if y + line_height > inner.y1:
            break
        page.insert_text(
            fitz.Point(inner.x0, y),
            line,
            fontsize=font_size,
            fontname="helv",
            color=color,
        )
        y += line_height


def generate_highlighted_pdf(pdf_bytes: bytes, errors: List[dict]) -> bytes:
    """Return a copy of the PDF with each error's location highlighted.

    For each error, the relevant page is opened and the extracted text
    snippet is searched for. If found, the EXACT matching line/word gets a
    real highlight annotation (color-coded by check type) with a popup note
    containing the full location message — nothing else on the page is
    touched.

    If the snippet cannot be located on the page at all (the content is
    genuinely absent — e.g. a missing paragraph), a visible text box is
    drawn in empty space on that page, showing a preview of the actual
    missing text so the error can be identified by reading the page alone,
    without needing to open the popup note.
    """
    with open("debug.log", "w", encoding="utf-8") as f_log:
        f_log.write("=== PDF Highlight Debug Log ===\n")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n_pages = doc.page_count

    # Track rects already used by fallback boxes on each page, so multiple
    # missing-content errors on one page stack without overlapping.
    placed_rects: dict[int, list] = {}

    # PASS 1: resolve every error to either an immediate highlight (order
    # doesn't matter for these) or a pending "missing" box. We don't place
    # missing boxes yet — we first need to know where every anchor on a
    # given page falls, so boxes can be stacked in the same top-to-bottom
    # order as their target text, instead of the order errors happen to
    # appear in the error list.
    pending_missing = []  # list of dicts: page_num, check_name, color, box_text, anchor, note_text

    for original_idx, err in enumerate(errors):
        check_name = err.get("check", "")
        location   = err.get("location", "")
        page_str   = str(err.get("page", "")).strip()

        if not page_str.isdigit():
            continue
        page_num = int(page_str)
        if page_num < 1 or page_num > n_pages:
            continue

        page  = doc[page_num - 1]
        color = CHECK_COLORS.get(check_name, DEFAULT_HIGHLIGHT_COLOR)

        snippet = err.get("search_string")
        if not snippet:
            snippet = _extract_snippet(location)

        is_missing_word_compare = (
            check_name == "Word-to-Word Comparison" and "missing in PDF" in location
        )
        rects, matched = _find_text_rects(
            page,
            snippet,
            allow_single_word_fallback=not is_missing_word_compare,
        ) if snippet else ([], None)

        note_text = f"[{check_name}]\n{location}"

        if rects:
            # Highlight ONLY the exact matched line/word found on the page
            # — no surrounding text is touched, so the highlight points
            # precisely at the error. Highlights don't need ordering, so
            # these are applied immediately.
            annot = page.add_highlight_annot(rects)
            annot.set_colors(stroke=color)
            annot.set_info(title=check_name, content=note_text)
            annot.update()
            continue

        # The text genuinely isn't on this page. Build a short preview
        # of what's missing (instead of a generic label) so the box
        # alone identifies the error.
        preview = (snippet or location)[:90].strip()
        if len(snippet or "") > 90:
            preview += "…"
        is_link_case = bool(re.search(r'https?://|www\.|doi\.|\blink\b', preview, re.I) or "link" in location.lower())
        if is_link_case:
            box_text = f"MISSING ? {check_name}\nLink above image\n\"{preview}\""
        else:
            box_text = f"MISSING ? {check_name}\n\"{preview}\""

        is_running_head = any(key in check_name.lower() for key in ("running head", "file name", "slug"))

        gap_rect = None
        if is_missing_word_compare:
            before_text = err.get("before_text")
            after_text = err.get("after_text")
            before_rect = _find_para_edge_rect(page, before_text, search_from_end=True) if before_text else None
            after_rect = _find_para_edge_rect(page, after_text, search_from_end=False) if after_text else None

            if before_rect is not None and after_rect is not None:
                if before_rect.y1 < after_rect.y0:
                    x_left = min(before_rect.x0, after_rect.x0)
                    x_right = max(before_rect.x1, after_rect.x1)
                    gap_rect = fitz.Rect(x_left, before_rect.y1, x_right, after_rect.y0)
            elif before_rect is not None:
                gap_rect = fitz.Rect(before_rect.x0, before_rect.y1, before_rect.x1, before_rect.y1 + 10)
            elif after_rect is not None:
                gap_rect = fitz.Rect(after_rect.x0, after_rect.y0 - 10, after_rect.x1, after_rect.y0)

        if gap_rect is not None:
            anchor = gap_rect
        else:
            context_anchor = None
            if is_missing_word_compare:
                context_anchor = _find_context_anchor(page, err.get("anchor_context"))

            if context_anchor is not None:
                # We found the exact text this content should follow — anchor
                # the box right there instead of guessing a direction.
                anchor = context_anchor
            elif is_missing_word_compare:
                # No exact position could be found at all — this is the only
                # case that falls back to a fixed bottom-of-page spot.
                anchor = None
            else:
                anchor = _find_anchor_rect_for_missing(page, snippet or location, prefer_top=is_running_head)

        if is_missing_word_compare:
            with open("debug.log", "a", encoding="utf-8") as f_log:
                f_log.write(f"\nMissing paragraph error: {preview}\n")
                f_log.write(f"  before_text: {before_text}\n")
                f_log.write(f"  after_text: {after_text}\n")
                f_log.write(f"  before_rect: {before_rect}\n")
                f_log.write(f"  after_rect: {after_rect}\n")
                f_log.write(f"  gap_rect: {gap_rect}\n")
                f_log.write(f"  anchor: {anchor}\n")

        pending_missing.append({
            "original_idx"  : original_idx,
            "page_num"      : page_num,
            "check_name"    : check_name,
            "color"         : color,
            "box_text"      : box_text,
            "note_text"     : note_text,
            "anchor"        : anchor,
            "is_running_head": is_running_head,
            "force_bottom"  : is_missing_word_compare and anchor is None,
            "is_missing_para": is_missing_word_compare,
        })

    # PASS 2: for each page, place its pending boxes in the same
    # top-to-bottom order as their anchors appear on the page. Items whose
    # anchor couldn't be resolved fall back to the order they were reported
    # in, placed after the anchored ones.
    by_page: dict[int, list] = {}
    for item in pending_missing:
        by_page.setdefault(item["page_num"], []).append(item)

    for page_num, items in by_page.items():
        page = doc[page_num - 1]
        existing = placed_rects.setdefault(page_num, [])

        items.sort(key=lambda it: (
            0 if it["is_running_head"] else (1 if it["force_bottom"] else (2 if it["anchor"] is not None else 3)),
            it["anchor"].y0 if it["anchor"] is not None else 0,
            it["original_idx"],
        ))

        box_w = min(205.0, max(150.0, page.rect.width * 0.18))
        box_h = 34.0

        top_left_taken = False
        for item in items:
            anchor          = item["anchor"]
            color           = item["color"]
            box_text        = item["box_text"]
            note_text       = item["note_text"]
            check_name      = item["check_name"]
            is_running_head = item["is_running_head"]
            force_bottom    = item["force_bottom"]
            is_missing_para = item.get("is_missing_para", False)

            if is_running_head:
                # Running Head / File Name errors always get the top-left
                # corner, regardless of whether an approximate anchor line
                # was found — this is a fixed, predictable spot for this
                # check type rather than one that moves around with content.
                rect = _find_corner_rect(page, box_w, box_h, existing, corner="top-left")
                top_left_taken = True
            elif force_bottom:
                # Word-to-Word Comparison "missing in PDF" errors: the text
                # legitimately isn't anywhere on this page, so there's no
                # honest line to sit beside. These always go below the last
                # real content on the page instead of floating next to a
                # fuzzy, potentially-misleading anchor line.
                rect = _find_corner_rect(page, box_w, box_h, existing, corner="bottom")
            elif anchor is not None:
                # Place the box directly in the text column at the anchor's vertical location.
                col_left, col_right, _, _ = _page_free_bands(page)
                # Determine horizontal bounds: if the anchor spans a column-like width (like gap_rect),
                # use it. Otherwise, span the entire column from col_left to col_right.
                if anchor.width > 200:
                    x0, x1 = anchor.x0, anchor.x1
                else:
                    x0, x1 = col_left, col_right

                anchor_mid = (anchor.y0 + anchor.y1) / 2
                target_y   = anchor_mid - box_h / 2
                target_y   = max(page.rect.y0 + 6, min(page.rect.y1 - box_h - 6, target_y))

                rect = fitz.Rect(x0, target_y, x1, target_y + box_h)
                
                # Check for collision with existing fallback boxes on the page and nudge vertically if needed.
                for step in range(1, 40):
                    if not any(rect.intersects(e) for e in existing):
                        break
                    y_try = target_y + (step * 6 if step % 2 == 0 else -step * 6)
                    y_try = max(page.rect.y0 + 6, min(page.rect.y1 - box_h - 6, y_try))
                    rect = fitz.Rect(x0, y_try, x1, y_try + box_h)
            else:
                # No anchor at all could be resolved on this page — the
                # error genuinely can't be tied to any location on it.
                # The first such box takes the top-left corner (if a
                # Running Head box hasn't already claimed it), any further
                # ones pack along the bottom.
                corner = "top-left" if not top_left_taken else "bottom"
                rect = _find_corner_rect(page, box_w, box_h, existing, corner=corner)
                if corner == "top-left":
                    top_left_taken = True
            existing.append(rect)

            # solid white background + colored border so the box is legible
            # against any underlying page content
            page.draw_rect(rect, color=color, fill=(1, 1, 1), width=1.2, fill_opacity=0.95)

            # Render the text line-by-line so it doesn't disappear when the
            # content is a little too tall for PyMuPDF's textbox fitting.
            _insert_wrapped_box_text(page, rect, box_text, color)

            # keep the full location detail available on click, without
            # cluttering the visible box itself
            annot = page.add_text_annot(fitz.Point(rect.x1 + 2, rect.y0), note_text)
            annot.set_info(title=check_name)
            annot.update()

            # NOTE: We intentionally do NOT draw pointer arrows anymore as requested.

            # NOTE: we intentionally do NOT draw a pointer arrow from the box
            # to its anchor. Drawing a visible line/arrowhead onto the page
            # lands the arrowhead in the middle of running text, which is
            # illegible. Precise row-alignment (above) plus the box border
            # is enough to show which line the box belongs to.

            # NOTE: we intentionally do NOT draw a pointer arrow from the box
            # to its anchor. The anchor is only used to decide reading-order
            # (top-to-bottom) among boxes grouped in the same whitespace
            # band; drawing a visible line/arrowhead onto the page lands the
            # arrowhead in the middle of running text, which is illegible.
            # The box + its border + click-to-read popup note is enough to
            # identify the error without marking up the underlying text.

    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    buf.seek(0)
    return buf.read()


# ─── Highlighted PDF download endpoint ──────────────────────────────────────

@app.post("/highlight-pdf")
async def highlight_pdf(
    pdf_file: UploadFile = File(...),
    errors  : str        = Form(...),
):
    try:
        error_list = json.loads(errors)
    except (json.JSONDecodeError, TypeError):
        raise HTTPException(status_code=400, detail="`errors` must be a valid JSON array")

    pdf_bytes = await pdf_file.read()

    try:
        highlighted = generate_highlighted_pdf(pdf_bytes, error_list)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to generate highlighted PDF: {str(e)}")

    base_name = re.sub(r'\.pdf$', '', pdf_file.filename or "document", flags=re.I)
    safe_name = re.sub(r'[^A-Za-z0-9_-]+', '_', base_name).strip('_') or "document"
    filename  = f"{safe_name}_highlighted.pdf"

    return StreamingResponse(
        io.BytesIO(highlighted),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ─── Word (.docx) report ────────────────────────────────────────────────────

def generate_word_report(data: ReportRequest) -> bytes:
    doc = Document()

    title = doc.add_heading("Publishing QA Validation Report", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if data.book_title:
        sub = doc.add_paragraph(data.book_title)
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub.runs[0].font.size = Pt(14)
        sub.runs[0].font.bold = True

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_run = meta.add_run(f"Generated on {datetime.now().strftime('%d %B %Y, %I:%M %p')}")
    meta_run.font.size = Pt(10)
    meta_run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()

    # Summary table
    doc.add_heading("Summary", level=1)
    summary_table = doc.add_table(rows=4, cols=2)
    summary_table.style = "Light Grid Accent 1"
    rows = [
        ("Total Errors Found", str(data.total_errors)),
        ("Total Pages in PDF", str(data.total_pages)),
        ("Pages with Issues", str(len(data.affected_pages))),
        ("Checks Run", str(data.checks_run)),
    ]
    for i, (label, value) in enumerate(rows):
        summary_table.cell(i, 0).text = label
        summary_table.cell(i, 1).text = value
        summary_table.cell(i, 0).paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()

    # Group errors by check type
    grouped: dict[str, list[ErrorItem]] = {}
    for err in data.errors:
        grouped.setdefault(err.check, []).append(err)

    doc.add_heading("Detailed Findings", level=1)

    if not data.errors:
        p = doc.add_paragraph("No issues were found. The document passed all selected checks.")
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(0x1a, 0x7a, 0x1a)
    else:
        for check_name, items in grouped.items():
            doc.add_heading(f"{check_name} ({len(items)})", level=2)
            for err in items:
                p = doc.add_paragraph(style="List Bullet")
                page_run = p.add_run(f"Page {err.page}: ")
                page_run.font.bold = True
                page_run.font.color.rgb = RGBColor(0xB0, 0x00, 0x00)
                p.add_run(err.location)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─── PDF report ──────────────────────────────────────────────────────────────

def generate_pdf_report(data: ReportRequest) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch,
        leftMargin=0.75 * inch, rightMargin=0.75 * inch,
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "ReportTitle", parent=styles["Title"], fontSize=20, spaceAfter=6,
    )
    meta_style = ParagraphStyle(
        "Meta", parent=styles["Normal"], fontSize=9, textColor=colors.grey,
        alignment=1, spaceAfter=20,
    )
    section_style = ParagraphStyle(
        "Section", parent=styles["Heading1"], fontSize=14,
        spaceBefore=16, spaceAfter=8, textColor=colors.HexColor("#1a1a1a"),
    )
    check_style = ParagraphStyle(
        "CheckName", parent=styles["Heading2"], fontSize=12,
        spaceBefore=10, spaceAfter=6, textColor=colors.HexColor("#333333"),
    )
    body_style = ParagraphStyle(
        "Body", parent=styles["Normal"], fontSize=10, leading=14, spaceAfter=6,
    )
    page_label_style = ParagraphStyle(
        "PageLabel", parent=styles["Normal"], fontSize=10,
        textColor=colors.HexColor("#B00000"), spaceAfter=2,
    )

    story = []
    story.append(Paragraph("Publishing QA Validation Report", title_style))
    if data.book_title:
        story.append(Paragraph(data.book_title, ParagraphStyle(
            "BookTitle", parent=styles["Normal"], fontSize=13, alignment=1,
            spaceAfter=4,
        )))
    story.append(Paragraph(
        f"Generated on {datetime.now().strftime('%d %B %Y, %I:%M %p')}", meta_style
    ))

    # Summary table
    story.append(Paragraph("Summary", section_style))
    summary_data = [
        ["Total Errors Found", str(data.total_errors)],
        ["Total Pages in PDF", str(data.total_pages)],
        ["Pages with Issues", str(len(data.affected_pages))],
        ["Checks Run", str(data.checks_run)],
    ]
    summary_table = Table(summary_data, colWidths=[3 * inch, 2.5 * inch])
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0f0f0")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 16))

    # Detailed findings, grouped by check type
    story.append(Paragraph("Detailed Findings", section_style))

    if not data.errors:
        story.append(Paragraph(
            "No issues were found. The document passed all selected checks.",
            ParagraphStyle("Pass", parent=body_style, textColor=colors.HexColor("#1a7a1a"),
                           fontName="Helvetica-Bold"),
        ))
    else:
        grouped: dict[str, list[ErrorItem]] = {}
        for err in data.errors:
            grouped.setdefault(err.check, []).append(err)

        for check_name, items in grouped.items():
            story.append(Paragraph(f"{check_name} ({len(items)})", check_style))
            for err in items:
                story.append(Paragraph(f"Page {err.page}", page_label_style))
                story.append(Paragraph(err.location.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"), body_style))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ─── Report download endpoint ───────────────────────────────────────────────

@app.post("/generate-report")
async def generate_report(req: ReportRequest):
    safe_title = re.sub(r'[^A-Za-z0-9_-]+', '_', req.book_title or "QA_Report").strip('_') or "QA_Report"
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")

    if req.format == "pdf":
        content      = generate_pdf_report(req)
        media_type   = "application/pdf"
        filename     = f"{safe_title}_QA_Report_{timestamp}.pdf"
    elif req.format == "docx":
        content      = generate_word_report(req)
        media_type   = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        filename     = f"{safe_title}_QA_Report_{timestamp}.docx"
    else:
        raise HTTPException(status_code=400, detail="format must be 'docx' or 'pdf'")

    return StreamingResponse(
        io.BytesIO(content),
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )




@app.post("/validate")
async def validate(
    word_file: UploadFile = File(...),
    pdf_file : UploadFile = File(...),
    checks   : str        = "",
):
    selected = [c.strip() for c in checks.split(",") if c.strip()] if checks else []

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as wf:
        wf.write(await word_file.read())
        word_path = wf.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as pf:
        pf.write(await pdf_file.read())
        pdf_path = pf.name

    try:
        word_data = extract_docx(word_path)
        pdf_data  = extract_pdf(pdf_path)
        page_map  = build_word_to_pdf_page_map(word_data, pdf_data)
    except Exception as e:
        os.unlink(word_path)
        os.unlink(pdf_path)
        raise HTTPException(status_code=400, detail=f"Failed to parse files: {str(e)}")

    os.unlink(word_path)
    os.unlink(pdf_path)

    all_errors = []

    def run_if_selected(check_name, fn):
        if not selected or check_name in selected:
            try:
                result = fn()
                all_errors.extend(result)
            except Exception:
                pass

    run_if_selected("Page Number Sequence & Folio",    lambda: check_page_sequence(pdf_data))
    run_if_selected("Running Head Style & Position",   lambda: check_running_heads(pdf_data))
    run_if_selected("Running Head & File Name",        lambda: check_slug_line(pdf_data))
    run_if_selected("Word-to-Word Comparison",         lambda: check_word_comparison(word_data, pdf_data, page_map))
    # Typos check intentionally disabled per current user requirement.
    run_if_selected("Missing Content",                 lambda: check_missing_content(word_data, pdf_data, page_map))
    run_if_selected("Content Order",                   lambda: check_content_order(word_data, pdf_data))
    run_if_selected("Heading Levels & Numbering",      lambda: check_headings(word_data, pdf_data, page_map))
    run_if_selected("Equations",                       lambda: check_equations(word_data, pdf_data))
    run_if_selected("Special Characters & Symbols",    lambda: check_special_chars(pdf_data))
    run_if_selected("Footnote Citation & Placement",   lambda: check_footnotes(word_data, pdf_data))
    run_if_selected("List Spacing",                    lambda: check_lists(pdf_data))
    run_if_selected("Font Consistency",                lambda: check_font_consistency(pdf_data))
    run_if_selected("Quotations",                      lambda: check_quotations(word_data, pdf_data, page_map))
    run_if_selected("Citations & Placement",           lambda: check_citations(pdf_data))
    run_if_selected("Unwanted Characters",             lambda: check_unwanted_chars(pdf_data))
    run_if_selected("FPO / Placeholder Images",        lambda: check_fpo(pdf_data))
    run_if_selected("Table Style Consistency",         lambda: check_tables(word_data, pdf_data, page_map))

    affected_pages = sorted(set(
        int(e["page"]) for e in all_errors
        if e.get("page", "").isdigit()
    ))

    return {
        "errors"        : all_errors,
        "total_errors"  : len(all_errors),
        "total_pages"   : pdf_data["total_pages"],
        "affected_pages": affected_pages,
        "checks_run"    : len(selected) if selected else 17,
    }


@app.get("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)