import React, { useState } from 'react'
import './ColumnSelector.css'

function ColumnSelector({ fileData, selectedTabs, onColumnsSelected, onBack }) {
  const [selections, setSelections] = useState({})

  const handleSelection = (key, field, value) => {
    setSelections(prev => ({
      ...prev,
      [key]: {
        ...prev[key],
        [field]: value
      }
    }))
  }

  const handleContinue = () => {
    for (const fileIndex in selectedTabs) {
      for (const sheetName of selectedTabs[fileIndex]) {
        const key = `${fileIndex}-${sheetName}`
        const selection = selections[key]
        if (!selection || !selection.yAxis || !selection.xAxis) {
          alert('Select Y axis and X axis for all tabs')
          return
        }
      }
    }
    onColumnsSelected(selections)
  }

  const isComplete = () => {
    for (const fileIndex in selectedTabs) {
      for (const sheetName of selectedTabs[fileIndex]) {
        const key = `${fileIndex}-${sheetName}`
        const selection = selections[key]
        if (!selection || !selection.yAxis || !selection.xAxis) {
          return false
        }
      }
    }
    return true
  }

  return (
    <div className="column-selector">
      <h2>Select Columns</h2>
      <p className="subtitle">Define the axes for each matrix</p>

      <div className="selections-container">
        {Object.keys(selectedTabs).map(fileIndex => {
          const file = fileData[parseInt(fileIndex)]
          return selectedTabs[fileIndex].map(sheetName => {
            const sheet = file.sheets.find(s => s.name === sheetName)
            if (!sheet) return null
            
            const key = `${fileIndex}-${sheetName}`
            const selection = selections[key] || {}
            
            return (
              <div key={key} className="selection-card">
                <h3 className="selection-title">
                  <span>📊</span>
                  {file.fileName} → {sheetName}
                </h3>
                
                <div className="column-selectors">
                  <div className="selector-group">
                    <label className="required">Y Axis (Rows)</label>
                    <select
                      value={selection.yAxis || ''}
                      onChange={(e) => handleSelection(key, 'yAxis', e.target.value)}
                      className="column-select"
                    >
                      <option value="">Select column...</option>
                      {sheet.headers.map((header, idx) => (
                        <option key={idx} value={header}>{header}</option>
                      ))}
                    </select>
                  </div>

                  <div className="selector-group">
                    <label className="required">X Axis (Columns)</label>
                    <select
                      value={selection.xAxis || ''}
                      onChange={(e) => handleSelection(key, 'xAxis', e.target.value)}
                      className="column-select"
                    >
                      <option value="">Select column...</option>
                      {sheet.headers.map((header, idx) => (
                        <option key={idx} value={header}>{header}</option>
                      ))}
                    </select>
                  </div>

                  <div className="selector-group">
                    <label>Secondary X (Split by)</label>
                    <select
                      value={selection.secondaryXAxis || ''}
                      onChange={(e) => handleSelection(key, 'secondaryXAxis', e.target.value || null)}
                      className="column-select"
                    >
                      <option value="">None</option>
                      {sheet.headers.map((header, idx) => (
                        <option key={idx} value={header}>{header}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            )
          })
        })}
      </div>

      <div className="button-group">
        <button onClick={onBack} className="btn btn-secondary">
          ← Back
        </button>
        <button 
          onClick={handleContinue} 
          className="btn btn-primary"
          disabled={!isComplete()}
        >
          Continue →
        </button>
      </div>
    </div>
  )
}

export default ColumnSelector
