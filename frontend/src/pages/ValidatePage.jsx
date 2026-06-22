import React, { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import axios from 'axios'
import UploadZone    from '../components/UploadZone'
import CheckSelector, { CHECKS } from '../components/CheckSelector'

export default function ValidatePage() {
  const navigate = useNavigate()
  const [wordFile,  setWordFile]  = useState(null)
  const [pdfFile,   setPdfFile]   = useState(null)
  const [selected,  setSelected]  = useState([...CHECKS])
  const [loading,   setLoading]   = useState(false)
  const [progress,  setProgress]  = useState(0)
  const [error,     setError]     = useState('')

  const canRun = wordFile && pdfFile && selected.length > 0

  async function runValidation() {
    if (!canRun) return
    setError('')
    setLoading(true)
    setProgress(0)

    const timer = setInterval(() => {
      setProgress(p => Math.min(p + Math.random() * 12 + 3, 90))
    }, 300)

    try {
      const form = new FormData()
      form.append('word_file', wordFile)
      form.append('pdf_file',  pdfFile)
      form.append('checks',    selected.join(','))

      const API_BASE = "https://qa-tool-1oh2.onrender.com";

const { data } = await axios.post(
  `${API_BASE}/validate`,
  form,
  {
    headers: {
      'Content-Type': 'multipart/form-data'
    }
  }
);

      clearInterval(timer)
      setProgress(100)
      setTimeout(() => {
        setLoading(false)
        navigate('/report', { state: { result: data, checks: selected } })
      }, 400)
    } catch (e) {
      clearInterval(timer)
      setLoading(false)
      setError(e?.response?.data?.detail || 'Validation failed. Is the backend running?')
    }
  }

  return (
    <div className="min-h-screen bg-gray-50 py-10 px-4">
      <div className="max-w-3xl mx-auto">

        {/* Header */}
        <div className="mb-8">
          <h1 className="text-2xl font-bold text-gray-900 flex items-center gap-2">
            📋 Publishing QA Validation Tool
          </h1>
          <p className="text-gray-500 text-sm mt-1">
            Compare your Word source against the final PDF and get a structured error report — no API key required.
          </p>
        </div>

        <div className="bg-white rounded-2xl border border-gray-200 p-6 shadow-sm space-y-6">

          {/* Upload */}
          <div>
            <p className="text-xs font-semibold text-gray-500 uppercase tracking-wider mb-3">
              Upload Documents
            </p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <UploadZone
                label="Word Source Document"
                hint=".docx / .doc"
                accept=".docx,.doc"
                icon="📄"
                onFile={setWordFile}
              />
              <UploadZone
                label="Final PDF Output"
                hint=".pdf (text-extractable)"
                accept=".pdf"
                icon="📕"
                onFile={setPdfFile}
              />
            </div>
          </div>

          {/* Checks */}
          <div className="border-t border-gray-100 pt-5">
            <CheckSelector selected={selected} onChange={setSelected} />
          </div>

          {/* Error */}
          {error && (
            <div className="bg-red-50 border border-red-200 text-red-700 rounded-lg px-4 py-3 text-sm">
              ⚠️ {error}
            </div>
          )}

          {/* Progress */}
          {loading && (
            <div>
              <div className="flex justify-between text-xs text-gray-500 mb-1">
                <span>Running validation…</span>
                <span>{Math.round(progress)}%</span>
              </div>
              <div className="h-1.5 bg-gray-100 rounded-full overflow-hidden">
                <div
                  className="h-full bg-brand rounded-full transition-all duration-300"
                  style={{ width: `${progress}%` }}
                />
              </div>
            </div>
          )}

          {/* Run button */}
          <button
            onClick={runValidation}
            disabled={!canRun || loading}
            className={`w-full py-3 rounded-xl font-semibold text-white text-sm flex items-center justify-center gap-2 transition-all
              ${canRun && !loading
                ? 'bg-brand hover:bg-brand-dark active:scale-[0.99]'
                : 'bg-gray-300 cursor-not-allowed'}`}
          >
            {loading ? '⏳ Validating…' : '▶ Run QA Validation'}
          </button>

        </div>
      </div>
    </div>
  )
}
