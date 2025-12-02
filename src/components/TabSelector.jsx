import React, { useState } from 'react'
import './TabSelector.css'

function TabSelector({ fileData, onTabsSelected, onBack }) {
  const [selectedTabs, setSelectedTabs] = useState({})

  const handleTabToggle = (fileIndex, sheetName) => {
    const key = `${fileIndex}-${sheetName}`
    setSelectedTabs(prev => ({
      ...prev,
      [key]: !prev[key]
    }))
  }

  const handleContinue = () => {
    const selected = {}
    Object.keys(selectedTabs).forEach(key => {
      if (selectedTabs[key]) {
        const [fileIndex, ...rest] = key.split('-')
        const sheetName = rest.join('-')
        if (!selected[fileIndex]) {
          selected[fileIndex] = []
        }
        selected[fileIndex].push(sheetName)
      }
    })
    
    const hasSelection = Object.values(selected).some(tabs => tabs.length > 0)
    if (!hasSelection) {
      alert('Select at least one tab')
      return
    }
    
    onTabsSelected(selected)
  }

  const selectedCount = Object.values(selectedTabs).filter(Boolean).length

  return (
    <div className="tab-selector">
      <h2>Select Tabs</h2>
      <p className="subtitle">Choose which sheets to process from each file</p>

      <div className="files-container">
        {fileData.map((file, fileIndex) => (
          <div key={fileIndex} className="file-card">
            <h3 className="file-title">
              <span className="file-icon">📗</span>
              {file.fileName}
            </h3>
            <div className="tabs-list">
              {file.sheets.map((sheet, sheetIndex) => {
                const key = `${fileIndex}-${sheet.name}`
                const isSelected = selectedTabs[key] || false
                return (
                  <div
                    key={sheetIndex}
                    className={`tab-item ${isSelected ? 'selected' : ''}`}
                    onClick={() => handleTabToggle(fileIndex, sheet.name)}
                  >
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onChange={() => {}}
                      onClick={(e) => e.stopPropagation()}
                    />
                    <span className="tab-name">{sheet.name}</span>
                    <span className="tab-info">
                      {sheet.data.length}×{sheet.headers.length}
                    </span>
                  </div>
                )
              })}
            </div>
          </div>
        ))}
      </div>

      <div className="button-group">
        <button onClick={onBack} className="btn btn-secondary">
          ← Back
        </button>
        <button 
          onClick={handleContinue} 
          className="btn btn-primary"
          disabled={selectedCount === 0}
        >
          Continue ({selectedCount} selected) →
        </button>
      </div>
    </div>
  )
}

export default TabSelector
