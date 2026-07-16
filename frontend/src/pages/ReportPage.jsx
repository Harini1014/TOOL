import React, { useMemo } from 'react'
import { useLocation, useNavigate } from 'react-router-dom'
import { AgGridReact } from 'ag-grid-react'
import 'ag-grid-community/styles/ag-grid.css'
import 'ag-grid-community/styles/ag-theme-alpine.css'

const API_BASE = "https://tool-2-3w1t.onrender.com/"

function BadgeCell({ value, color }) {
  const colors = {
    purple: 'bg-purple-950/40 text-purple-300 border border-purple-800/40',
    blue  : 'bg-blue-950/40 text-blue-300 border border-blue-800/40',
    red   : 'bg-red-950/40 text-red-300 border border-red-800/40',
    green : 'bg-green-950/40 text-green-300 border border-green-800/40',
    orange: 'bg-orange-950/40 text-orange-300 border border-orange-800/40',
    yellow: 'bg-yellow-950/40 text-yellow-300 border border-yellow-800/40'
  }
  return (
    <span className={`inline-block px-2.5 py-0.5 rounded-md text-xs font-semibold border ${colors[color] || colors.purple}`}>
      {value}
    </span>
  )
}

export default function ReportPage() {
  const location = useLocation()
  const navigate = useNavigate()
  const result = location.state?.result
  const checks = location.state?.checks || []
  const pdfFile = window._qaValidationPdfFile || null

  if (!result) {
    navigate('/')
    return null
  }

  const { errors = [], total_errors, total_pages, affected_pages, checks_run } = result

  const rowData = useMemo(() =>
    errors.map((e, i) => ({
      no         : i + 1,
      check      : e.check,
      type       : e.type || 'Error',
      color      : e.color || 'purple',
      page       : e.page || '?',
      line       : e.line || '—',
      expected   : e.expected || '',
      actual     : e.actual || '',
      description: e.description || e.location || '',
    })),
    [errors]
  )

  const colDefs = useMemo(() => [
    { field: 'no',          headerName: '#',                 width: 50,  sortable: true },
    { field: 'check',       headerName: 'Check Item',        width: 180, sortable: true, filter: true },
    { field: 'type',        headerName: 'Error Type',        width: 140, sortable: true, filter: true,
      cellRenderer: (p) => <BadgeCell value={p.value} color={p.data.color} /> },
    { field: 'page',        headerName: 'Page',              width: 80,  sortable: true,
      cellRenderer: (p) => <BadgeCell value={`Pg ${p.value}`} color="blue" /> },
    { field: 'expected',    headerName: 'Expected (Word)',   width: 160, sortable: true, filter: true,
      cellStyle: { fontFamily: 'monospace', color: '#10b981' } },
    { field: 'actual',      headerName: 'Actual (PDF)',      width: 160, sortable: true, filter: true,
      cellStyle: { fontFamily: 'monospace', color: '#ef4444' } },
    { field: 'description', headerName: 'Description',       flex: 2,    sortable: true, filter: true,
      wrapText: true, autoHeight: true,
      cellStyle: { lineHeight: '1.5', paddingTop: '8px', paddingBottom: '8px' } },
  ], [])

  const defaultColDef = useMemo(() => ({
    resizable: true,
    suppressMovable: false,
  }), [])

  // ── Download: Word or PDF report ─────────────────────────────────────────
  async function downloadReport(format) {
    try {
      const form = new FormData()
      form.append('format', format)
      form.append('errors', JSON.stringify(errors))

      const res = await fetch(`${API_BASE}/download-report`, {
        method: 'POST',
        body: form
      })
      if (!res.ok) {
        const text = await res.text()
        console.error('Download report backend error:', text)
        throw new Error(`Server error: ${text.slice(0, 100)}`)
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = format === 'docx' ? 'qa_report.docx' : 'qa_report.pdf'
      a.click()
      URL.revokeObjectURL(url)
    } catch (e) {
      console.error('Download error:', e)
      alert('Failed to download report. Check browser console for details.')
    }
  }


  // ── Download: Highlighted PDF ────────────────────────────────────────────
  async function downloadHighlightedPdf() {
    if (!pdfFile) {
      alert('PDF file not available. Please re-run validation.')
      return
    }
    try {
      const form = new FormData()
      form.append('pdf_file', pdfFile)
      form.append('errors', JSON.stringify(errors))

      const res = await fetch(`${API_BASE}/highlighted-pdf`, { method: 'POST', body: form })
      if (!res.ok) throw new Error('Server error')
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'highlighted_report.pdf'
      a.click()
      URL.revokeObjectURL(url)
    } catch (e) {
      console.error('Highlighted PDF error:', e)
      alert('Failed to generate highlighted PDF. Check browser console for details.')
    }
  }

  return (
    <div className="min-h-screen py-12 px-4 relative">
      {/* Background ambient glows */}
      <div className="absolute top-1/4 left-1/4 w-96 h-96 bg-indigo-500/5 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute bottom-1/4 right-1/4 w-96 h-96 bg-emerald-500/5 rounded-full blur-3xl pointer-events-none" />

      <div className="max-w-7xl mx-auto relative z-10">

        {/* Header */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-8">
          <div className="flex items-center gap-4">
            <button
              onClick={() => navigate('/')}
              className="text-sm font-semibold text-slate-400 hover:text-slate-200 transition-colors bg-slate-900 border border-slate-800 rounded-lg px-3.5 py-2 flex items-center gap-1.5"
            >
              <span>←</span> Back
            </button>
            <div>
              <h1 className="text-2xl font-bold bg-gradient-to-r from-slate-100 to-slate-300 bg-clip-text text-transparent">
                Validation Report
              </h1>
              <p className="text-xs text-slate-500 mt-0.5">Publishing QA Verification Output</p>
            </div>
          </div>
        </div>

        {/* Summary cards */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
          {[
            { label: 'Total Errors', value: total_errors, color: total_errors > 0 ? 'text-red-400' : 'text-emerald-400', glow: total_errors > 0 ? 'shadow-red-500/5' : 'shadow-emerald-500/5' },
            { label: 'Checks Run', value: checks_run, color: 'text-slate-200', glow: 'shadow-slate-500/5' },
            { label: 'Total Pages', value: total_pages, color: 'text-slate-200', glow: 'shadow-slate-500/5' },
            { label: 'Affected Pages', value: affected_pages.length, color: affected_pages.length > 0 ? 'text-amber-400' : 'text-emerald-400', glow: affected_pages.length > 0 ? 'shadow-amber-500/5' : 'shadow-emerald-500/5' },
          ].map(s => (
            <div key={s.label} className={`glass-panel rounded-2xl p-5 shadow-lg ${s.glow} hover:border-slate-700 transition-colors`}>
              <p className="text-xs text-slate-500 font-medium tracking-wider uppercase mb-1">{s.label}</p>
              <p className={`text-3xl font-extrabold ${s.color}`}>{s.value}</p>
            </div>
          ))}
        </div>

        {/* Affected pages info block */}
        {affected_pages.length > 0 && (
          <div className="bg-amber-950/20 border border-amber-800/40 rounded-2xl px-5 py-4 mb-6 text-sm text-amber-300 flex items-center gap-2">
            <span className="text-base">📍</span>
            <span>
              <strong>Affected pages checklist:</strong> {affected_pages.map(p => `Page ${p}`).join(', ')}
            </span>
          </div>
        )}

        {/* No errors state */}
        {total_errors === 0 && (
          <div className="glass-panel rounded-3xl p-16 text-center border-emerald-500/20 shadow-emerald-500/5 shadow-2xl">
            <p className="text-5xl mb-4 animate-bounce">🎉</p>
            <p className="text-xl font-bold text-emerald-400">All Checks Passed Successfully!</p>
            <p className="text-sm text-slate-400 mt-2 max-w-md mx-auto leading-relaxed">
              No errors were found between the Word source document and the typeset PDF. All {checks_run} verification constraints are satisfied.
            </p>
          </div>
        )}

        {/* AG Grid table */}
        {total_errors > 0 && (
          <div className="glass-panel rounded-3xl overflow-hidden shadow-2xl relative">
            <div className="absolute top-0 left-0 right-0 h-[1px] bg-gradient-to-r from-transparent via-slate-800 to-transparent" />

            {/* Toolbar */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 px-6 py-5 border-b border-slate-800 bg-slate-900/30">
              <div>
                <p className="text-sm font-semibold text-slate-300">
                  Detected Discrepancies ({total_errors} item{total_errors !== 1 ? 's' : ''})
                </p>
                <p className="text-xs text-slate-500 mt-0.5">Click any row to view cell details.</p>
              </div>
              <div className="flex gap-2 flex-wrap">
                <button
                  onClick={() => downloadReport('pdf')}
                  className="text-xs bg-slate-950 hover:bg-slate-900 border border-slate-800 text-slate-300 px-4 py-2.5 rounded-xl font-semibold flex items-center gap-1.5 transition-all active:scale-[0.98]"
                >
                  📕 Download PDF Report
                </button>
                <button
                  onClick={downloadHighlightedPdf}
                  className="text-xs bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-400 hover:to-orange-400 text-slate-950 px-4 py-2.5 rounded-xl font-black flex items-center gap-1.5 shadow-lg shadow-amber-500/10 transition-all active:scale-[0.98]"
                >
                  🔆 Download Annotated PDF
                </button>
              </div>
            </div>

            {/* Grid */}
            <div className="ag-theme-alpine w-full" style={{ height: Math.min(80 + rowData.length * 56, 560) }}>
              <AgGridReact
                rowData={rowData}
                columnDefs={colDefs}
                defaultColDef={defaultColDef}
                animateRows={true}
                rowSelection="single"
                suppressCellFocus={true}
                pagination={rowData.length > 100}
                paginationPageSize={100}
              />
            </div>
          </div>
        )}

        <p className="text-xs text-slate-600 mt-6 text-center">
          Publishing QA Tool — Powered by PyMuPDF + python-docx
        </p>
      </div>
    </div>
  )
}
