import React, { useState } from 'react'

const CHECKS = [
  "Page Number Sequence and Folio Placement",
  "Running Head Style and Position",
  "Page Range and File Name in Slug Line",
  "Word-to-Word Comparison",
  "No Typos",
  "No Content Missing",
  "Content Order",
  "Heading Levels Style Numbering and Turnover Indent",
  "Mini TOC in Chapter Opener",
  "Equation Entry",
  "Equation Special Character Symbol",
  "Footnote Citation and Placement",
  "Above Below Space for Lists",
  "Double Digit Alignment",
  "Facing Page Alignment",
  "Global Instructions",
  "Quotation Check",
  "Citation and Placement",
  "Image Figure Cutoffs",
  "Image Size",
  "Line Art Readability",
  "Credit Lines",
  "Heading Style Consistency",
  "Box Style Consistency",
  "Table Style Consistency",
  "Font Consistency",
  "Keyterm Page Numbers",
  "Figure Photo FPO Check",
  "Unwanted Characters"
]

export { CHECKS }

export default function CheckSelector({ selected, onChange }) {
  const allSelected = selected.length === CHECKS.length

  function toggle(check) {
    if (selected.includes(check)) onChange(selected.filter(c => c !== check))
    else onChange([...selected, check])
  }

  function toggleAll() {
    onChange(allSelected ? [] : [...CHECKS])
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-2">
        <p className="text-xs font-semibold text-gray-500 uppercase tracking-wider">
          Validation Checks ({selected.length}/{CHECKS.length})
        </p>
        <button
          onClick={toggleAll}
          className="text-xs text-brand underline font-semibold"
        >
          {allSelected ? 'Deselect All' : 'Select All'}
        </button>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-3 gap-1.5">
        {CHECKS.map(c => (
          <label key={c} className="check-pill">
            <input
              type="checkbox"
              checked={selected.includes(c)}
              onChange={() => toggle(c)}
            />
            <span className="text-xs leading-tight">{c}</span>
          </label>
        ))}
      </div>
    </div>
  )
}
