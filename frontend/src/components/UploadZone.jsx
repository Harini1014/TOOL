import React, { useRef, useState } from 'react'

export default function UploadZone({ label, hint, accept, icon, onFile }) {
  const inputRef = useRef()
  const [filename, setFilename] = useState('')
  const [dragging, setDragging] = useState(false)

  function handleFile(file) {
    if (!file) return
    setFilename(file.name)
    onFile(file)
  }

  return (
    <div
      className={`upload-zone ${filename ? 'has-file' : ''} ${dragging ? 'active' : ''}`}
      onClick={() => inputRef.current.click()}
      onDragOver={e => { e.preventDefault(); setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); handleFile(e.dataTransfer.files[0]) }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        className="hidden"
        onChange={e => handleFile(e.target.files[0])}
      />
      <div className={`text-4xl mb-3 ${filename ? 'text-emerald-400' : 'text-slate-500'}`}>{icon}</div>
      <p className="font-semibold text-slate-200">{label}</p>
      <p className="text-xs text-slate-400 mt-1">{hint}</p>
      {filename && (
        <p className="mt-2 text-sm font-medium text-emerald-400 truncate max-w-full">✓ {filename}</p>
      )}
      {!filename && (
        <p className="mt-3 text-xs text-slate-400">Click or drag & drop</p>
      )}
    </div>
  )
}
