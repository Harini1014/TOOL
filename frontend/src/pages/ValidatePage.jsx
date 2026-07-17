import React, { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import axios from 'axios'
import UploadZone from '../components/UploadZone'
import CheckSelector, { CHECKS } from '../components/CheckSelector'

export default function ValidatePage() {
  const navigate = useNavigate()
  const [wordFile, setWordFile] = useState(null)
  const [pdfFile, setPdfFile] = useState(null)
  const [selected, setSelected] = useState([...CHECKS])
  const [loading, setLoading] = useState(false)
  const [progress, setProgress] = useState(0)
  const [error, setError] = useState('')

  const canRun = wordFile && pdfFile && selected.length > 0

  async function runValidation() {
    if (!canRun) return
    setError('')
    setLoading(true)
    setProgress(0)

    const timer = setInterval(() => {
      setProgress(p => Math.min(p + Math.random() * 12 + 3, 92))
    }, 300)

    try {
      const form = new FormData()
      form.append('word_file', wordFile)
      form.append('pdf_file', pdfFile)
      form.append('checks', selected.join(','))

      const API_BASE = "http://localhost:8000"

      const { data } = await axios.post(
        `${API_BASE}/validate`,
        form,
        { headers: { 'Content-Type': 'multipart/form-data' } }
      )

      clearInterval(timer)
      setProgress(100)

      setTimeout(() => {
        setLoading(false)
        window._qaValidationPdfFile = pdfFile
        navigate('/report', { state: { result: data, checks: selected } })
      }, 400)
    } catch (e) {
      clearInterval(timer)
      setLoading(false)
      setError(e?.response?.data?.detail || 'Validation failed. Is the backend running?')
    }
  }

  return (
    <div className="min-h-screen py-12 px-4 relative flex items-center justify-center">
      {/* Background ambient glows */}
      <div className="absolute top-1/4 left-1/4 w-96 h-96 bg-emerald-500/10 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute bottom-1/4 right-1/4 w-96 h-96 bg-indigo-500/10 rounded-full blur-3xl pointer-events-none" />

      <div className="max-w-4xl w-full mx-auto relative z-10">

        {/* Header */}
        <div className="text-center mb-10">
          <div className="inline-flex items-center justify-center p-3 bg-slate-900/80 rounded-2xl border border-slate-800 mb-4 shadow-xl shadow-emerald-500/5">
            <span className="text-3xl">🔍</span>
          </div>
          <h1 className="text-4xl font-extrabold tracking-tight bg-gradient-to-r from-emerald-400 via-teal-300 to-indigo-400 bg-clip-text text-transparent">
            Publishing QA Comparison Tool
          </h1>
          <p className="text-slate-400 text-sm mt-3 max-w-xl mx-auto leading-relaxed">
            Upload your source Word document and final typeset PDF. Our strict comparison engine validates formatting, text accuracy, and layout constraints at character level.
          </p>
        </div>

        <div className="glass-panel rounded-3xl p-8 shadow-2xl space-y-8 relative overflow-hidden">
          {/* Top subtle highlight border */}
          <div className="absolute top-0 left-0 right-0 h-[1px] bg-gradient-to-r from-transparent via-emerald-500/50 to-transparent" />

          {/* Upload Section */}
          <div>
            <p className="text-xs font-semibold text-slate-400 uppercase tracking-widest mb-4">
              Step 1: Upload Documents
            </p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <UploadZone
                label="Word Source Document"
                hint="Microsoft Word (.docx)"
                accept=".docx"
                icon="📄"
                onFile={setWordFile}
              />
              <UploadZone
                label="Typeset PDF Output"
                hint="Adobe PDF (.pdf)"
                accept=".pdf"
                icon="📕"
                onFile={setPdfFile}
              />
            </div>
          </div>

          {/* Checks Selection Section */}
          <div className="border-t border-slate-800/80 pt-6">
            <CheckSelector selected={selected} onChange={setSelected} />
          </div>

          {/* Error Message */}
          {error && (
            <div className="bg-red-950/40 border border-red-800/80 text-red-300 rounded-xl px-4 py-3 text-sm flex items-center gap-2 shadow-inner">
              <span className="text-base">⚠️</span>
              <span>{error}</span>
            </div>
          )}

          {/* Progress Section */}
          {loading && (
            <div className="space-y-2">
              <div className="flex justify-between text-xs font-medium text-slate-400">
                <span className="flex items-center gap-1.5">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-ping" />
                  Aligning characters & verifying pages…
                </span>
                <span>{Math.round(progress)}%</span>
              </div>
              <div className="h-2 bg-slate-950 rounded-full overflow-hidden border border-slate-800/60 p-[1px]">
                <div
                  className="h-full bg-gradient-to-r from-emerald-500 to-teal-400 rounded-full transition-all duration-300 shadow-[0_0_8px_rgba(16,185,129,0.4)]"
                  style={{ width: `${progress}%` }}
                />
              </div>
            </div>
          )}

          {/* Actions */}
          <button
            onClick={runValidation}
            disabled={!canRun || loading}
            className={`w-full py-4 rounded-xl font-bold text-sm flex items-center justify-center gap-2 transition-all duration-300 transform active:scale-[0.98] shadow-lg
              ${canRun && !loading
                ? 'bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-400 hover:to-teal-400 text-slate-950 shadow-emerald-500/10 hover:shadow-emerald-500/25 cursor-pointer font-black'
                : 'bg-slate-900 border border-slate-800 text-slate-500 cursor-not-allowed'}`}
          >
            {loading ? (
              <>
                <span className="animate-spin text-base">⏳</span>
                <span>Running Strict Verification...</span>
              </>
            ) : (
              <>
                <span>▶</span>
                <span>Run QA Verification</span>
              </>
            )}
          </button>

        </div>

        {/* Footer */}
        <p className="text-center text-xs text-slate-600 mt-8">
          Powered by PyMuPDF + python-docx • Localhost mode
        </p>
      </div>
    </div>
  )
}
