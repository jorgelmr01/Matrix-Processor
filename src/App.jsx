import React, { useState } from 'react'
import { processFiles, computeMatrices, exportToExcel } from './utils/fileProcessor'
import FileUploader from './components/FileUploader'
import TabSelector from './components/TabSelector'
import ColumnSelector from './components/ColumnSelector'
import MatrixConfigurator from './components/MatrixConfigurator'
import './App.css'

function App() {
  const [step, setStep] = useState(1)
  const [files, setFiles] = useState([])
  const [fileData, setFileData] = useState([])
  const [selectedTabs, setSelectedTabs] = useState({})
  const [columnSelections, setColumnSelections] = useState({})
  const [matrixConfig, setMatrixConfig] = useState([])
  const [matrices, setMatrices] = useState([])
  const [processing, setProcessing] = useState(false)

  const handleFilesUploaded = async (uploadedFiles) => {
    setFiles(uploadedFiles)
    setProcessing(true)
    try {
      const processed = await processFiles(uploadedFiles)
      setFileData(processed)
      setStep(2)
    } catch (error) {
      alert('Error: ' + error.message)
    } finally {
      setProcessing(false)
    }
  }

  const handleTabsSelected = (selections) => {
    setSelectedTabs(selections)
    setStep(3)
  }

  const handleColumnsSelected = (selections) => {
    setColumnSelections(selections)
    setStep(4)
  }

  const handleMatrixConfig = (config) => {
    setMatrixConfig(config)
    setProcessing(true)
    try {
      const computed = computeMatrices(fileData, selectedTabs, columnSelections, config)
      setMatrices(computed)
      setStep(5)
    } catch (error) {
      alert('Error: ' + error.message)
    } finally {
      setProcessing(false)
    }
  }

  const handleExport = () => {
    try {
      exportToExcel(matrices)
    } catch (error) {
      alert('Export error: ' + error.message)
    }
  }

  const resetApp = () => {
    setStep(1)
    setFiles([])
    setFileData([])
    setSelectedTabs({})
    setColumnSelections({})
    setMatrixConfig([])
    setMatrices([])
  }

  const steps = [
    { num: 1, label: 'Upload' },
    { num: 2, label: 'Tabs' },
    { num: 3, label: 'Columns' },
    { num: 4, label: 'Configure' },
    { num: 5, label: 'Export' }
  ]

  return (
    <div className="app">
      <header className="app-header">
        <h1>Matrix Processor</h1>
      </header>

      <div className="progress-bar">
        {steps.map((s) => (
          <div 
            key={s.num} 
            className={`step ${step === s.num ? 'active' : ''} ${step > s.num ? 'completed' : ''}`}
          >
            {s.num}. {s.label}
          </div>
        ))}
      </div>

      <div className="content">
        {processing && (
          <div className="loading-overlay">
            <div className="spinner"></div>
            <p>Processing...</p>
          </div>
        )}

        {step === 1 && (
          <FileUploader onFilesUploaded={handleFilesUploaded} />
        )}

        {step === 2 && (
          <TabSelector
            fileData={fileData}
            onTabsSelected={handleTabsSelected}
            onBack={() => setStep(1)}
          />
        )}

        {step === 3 && (
          <ColumnSelector
            fileData={fileData}
            selectedTabs={selectedTabs}
            onColumnsSelected={handleColumnsSelected}
            onBack={() => setStep(2)}
          />
        )}

        {step === 4 && (
          <MatrixConfigurator
            fileData={fileData}
            selectedTabs={selectedTabs}
            columnSelections={columnSelections}
            onConfigComplete={handleMatrixConfig}
            onBack={() => setStep(3)}
          />
        )}

        {step === 5 && (
          <div className="step-content">
            <div className="success-message">
              <div className="success-icon">✓</div>
              <h2>Matrices Ready</h2>
              <p className="subtitle">
                {matrices.length} matrix{matrices.length !== 1 ? 'es' : ''} computed successfully
              </p>
            </div>

            <div className="matrices-preview">
              {matrices.map((matrix, idx) => (
                <div key={idx} className="matrix-preview">
                  <h3>{matrix.name}</h3>
                  <p>{matrix.yAxis.length} rows × {matrix.xAxis.length} columns</p>
                </div>
              ))}
            </div>

            <div className="button-group">
              <button onClick={resetApp} className="btn btn-secondary">
                Start Over
              </button>
              <button onClick={handleExport} className="btn btn-success">
                Download Excel
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

export default App
