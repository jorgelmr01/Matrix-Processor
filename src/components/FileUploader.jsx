import React, { useRef, useState } from 'react'
import './FileUploader.css'

function FileUploader({ onFilesUploaded }) {
  const fileInputRef = useRef(null)
  const [files, setFiles] = useState([])
  const [dragging, setDragging] = useState(false)

  const handleFileChange = (e) => {
    const selectedFiles = Array.from(e.target.files)
    setFiles(prev => [...prev, ...selectedFiles])
  }

  const handleRemoveFile = (index) => {
    setFiles(prev => prev.filter((_, i) => i !== index))
  }

  const handleUpload = () => {
    if (files.length === 0) return
    onFilesUploaded(files)
  }

  const handleDragOver = (e) => {
    e.preventDefault()
    e.stopPropagation()
    setDragging(true)
  }

  const handleDragLeave = (e) => {
    e.preventDefault()
    e.stopPropagation()
    setDragging(false)
  }

  const handleDrop = (e) => {
    e.preventDefault()
    e.stopPropagation()
    setDragging(false)
    const droppedFiles = Array.from(e.dataTransfer.files).filter(file => 
      file.name.endsWith('.xlsx') || 
      file.name.endsWith('.xls') || 
      file.name.endsWith('.csv')
    )
    setFiles(prev => [...prev, ...droppedFiles])
  }

  const getFileIcon = (fileName) => {
    if (fileName.endsWith('.csv')) return '📊'
    return '📗'
  }

  const formatSize = (bytes) => {
    if (bytes < 1024) return bytes + ' B'
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB'
  }

  return (
    <div className="file-uploader">
      <h2>Upload Files</h2>
      <p className="subtitle">Add Excel or CSV files to process</p>
      
      <div 
        className={`drop-zone ${dragging ? 'dragging' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
      >
        <div className="drop-zone-content">
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="17 8 12 3 7 8"></polyline>
            <line x1="12" y1="3" x2="12" y2="15"></line>
          </svg>
          <p>Click to browse or drag files here</p>
          <p className="file-types">.xlsx, .xls, .csv</p>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept=".xlsx,.xls,.csv"
          onChange={handleFileChange}
          style={{ display: 'none' }}
        />
      </div>

      {files.length > 0 && (
        <div className="file-list">
          <h3>
            Files <span className="file-count">{files.length}</span>
          </h3>
          <ul>
            {files.map((file, index) => (
              <li key={index}>
                <span className="file-icon">{getFileIcon(file.name)}</span>
                <span className="file-name">{file.name}</span>
                <span className="file-size">{formatSize(file.size)}</span>
                <button 
                  className="remove-btn"
                  onClick={(e) => {
                    e.stopPropagation()
                    handleRemoveFile(index)
                  }}
                  title="Remove file"
                >
                  ×
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}

      <div className="button-group">
        <button 
          onClick={handleUpload} 
          className="btn btn-primary"
          disabled={files.length === 0}
        >
          Continue →
        </button>
      </div>
    </div>
  )
}

export default FileUploader
