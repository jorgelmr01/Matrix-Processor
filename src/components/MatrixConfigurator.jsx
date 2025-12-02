import React, { useState, useEffect } from 'react'
import './MatrixConfigurator.css'

function MatrixConfigurator({ fileData, selectedTabs, columnSelections, onConfigComplete, onBack }) {
  const [configs, setConfigs] = useState([])

  useEffect(() => {
    const initialConfigs = []
    Object.keys(selectedTabs).forEach(fileIndex => {
      selectedTabs[fileIndex].forEach(sheetName => {
        const key = `${fileIndex}-${sheetName}`
        if (columnSelections[key]) {
          initialConfigs.push({
            name: `${fileData[parseInt(fileIndex)].fileName.replace(/\.[^/.]+$/, '')} - ${sheetName}`,
            sources: [{
              fileIndex: parseInt(fileIndex),
              sheetName: sheetName,
              fileName: fileData[parseInt(fileIndex)].fileName
            }],
            merge: false,
            hasSecondaryX: !!columnSelections[key].secondaryXAxis
          })
        }
      })
    })
    setConfigs(initialConfigs)
  }, [])

  const handleToggleMerge = (index) => {
    setConfigs(prev => prev.map((config, idx) => 
      idx === index ? { ...config, merge: !config.merge } : config
    ))
  }

  const handleAddMatrix = () => {
    const availableSources = []
    Object.keys(selectedTabs).forEach(fileIndex => {
      selectedTabs[fileIndex].forEach(sheetName => {
        const key = `${fileIndex}-${sheetName}`
        if (columnSelections[key]) {
          availableSources.push({
            fileIndex: parseInt(fileIndex),
            sheetName: sheetName,
            fileName: fileData[parseInt(fileIndex)].fileName
          })
        }
      })
    })

    if (availableSources.length === 0) return

    const firstSource = availableSources[0]
    const key = `${firstSource.fileIndex}-${firstSource.sheetName}`
    setConfigs(prev => [...prev, {
      name: `Matrix ${prev.length + 1}`,
      sources: [firstSource],
      merge: false,
      hasSecondaryX: !!columnSelections[key]?.secondaryXAxis
    }])
  }

  const handleRemoveMatrix = (index) => {
    setConfigs(prev => prev.filter((_, idx) => idx !== index))
  }

  const handleAddSource = (matrixIndex) => {
    const currentConfig = configs[matrixIndex]
    const availableSources = []
    
    Object.keys(selectedTabs).forEach(fileIndex => {
      selectedTabs[fileIndex].forEach(sheetName => {
        const key = `${fileIndex}-${sheetName}`
        if (columnSelections[key]) {
          const source = {
            fileIndex: parseInt(fileIndex),
            sheetName: sheetName,
            fileName: fileData[parseInt(fileIndex)].fileName
          }
          const exists = currentConfig.sources.some(s => 
            s.fileIndex === source.fileIndex && s.sheetName === source.sheetName
          )
          if (!exists) {
            availableSources.push(source)
          }
        }
      })
    })

    if (availableSources.length > 0) {
      const newSource = availableSources[0]
      const key = `${newSource.fileIndex}-${newSource.sheetName}`
      const hasSecondaryX = !!columnSelections[key]?.secondaryXAxis || currentConfig.hasSecondaryX
      
      setConfigs(prev => prev.map((config, idx) => 
        idx === matrixIndex 
          ? { ...config, sources: [...config.sources, newSource], hasSecondaryX }
          : config
      ))
    }
  }

  const handleRemoveSource = (matrixIndex, sourceIndex) => {
    setConfigs(prev => prev.map((config, idx) => 
      idx === matrixIndex 
        ? { ...config, sources: config.sources.filter((_, sIdx) => sIdx !== sourceIndex) }
        : config
    ))
  }

  const handleNameChange = (index, newName) => {
    setConfigs(prev => prev.map((config, idx) => 
      idx === index ? { ...config, name: newName } : config
    ))
  }

  const handleContinue = () => {
    if (configs.length === 0) {
      alert('Add at least one matrix')
      return
    }

    const invalid = configs.some(config => config.sources.length === 0)
    if (invalid) {
      alert('Each matrix needs at least one source')
      return
    }

    onConfigComplete(configs)
  }

  return (
    <div className="matrix-configurator">
      <h2>Configure Matrices</h2>
      <p className="subtitle">Organize and merge your data sources</p>

      <div className="matrices-list">
        {configs.map((config, index) => (
          <div key={index} className="matrix-config-card">
            <div className="matrix-header">
              <input
                type="text"
                value={config.name}
                onChange={(e) => handleNameChange(index, e.target.value)}
                className="matrix-name-input"
                placeholder="Matrix name"
              />
              <div className="matrix-actions">
                <label className="merge-toggle">
                  <input
                    type="checkbox"
                    checked={config.merge}
                    onChange={() => handleToggleMerge(index)}
                  />
                  <span>Merge sources</span>
                </label>
                {configs.length > 1 && (
                  <button
                    className="remove-matrix-btn"
                    onClick={() => handleRemoveMatrix(index)}
                  >
                    Remove
                  </button>
                )}
              </div>
            </div>

            <div className="sources-list">
              <h4>Data Sources</h4>
              {config.sources.map((source, sourceIndex) => (
                <div key={sourceIndex} className="source-item">
                  <span className="source-name">
                    {source.fileName} → {source.sheetName}
                  </span>
                  {config.sources.length > 1 && (
                    <button
                      className="remove-source-btn"
                      onClick={() => handleRemoveSource(index, sourceIndex)}
                    >
                      ×
                    </button>
                  )}
                </div>
              ))}
              <button
                className="add-source-btn"
                onClick={() => handleAddSource(index)}
              >
                + Add source
              </button>
            </div>
          </div>
        ))}

        <button onClick={handleAddMatrix} className="add-matrix-btn">
          + Add new matrix
        </button>
      </div>

      <div className="button-group">
        <button onClick={onBack} className="btn btn-secondary">
          ← Back
        </button>
        <button onClick={handleContinue} className="btn btn-primary">
          Compute Matrices →
        </button>
      </div>
    </div>
  )
}

export default MatrixConfigurator
