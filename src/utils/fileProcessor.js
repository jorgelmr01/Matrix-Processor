import * as XLSX from 'xlsx'

export async function processFiles(files) {
  const processedFiles = []
  
  for (const file of files) {
    const fileData = {
      fileName: file.name,
      fileType: file.name.endsWith('.csv') ? 'csv' : 'excel',
      sheets: []
    }

    try {
      const arrayBuffer = await file.arrayBuffer()
      const workbook = XLSX.read(arrayBuffer, { type: 'array' })
      
      workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName]
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
        
        // Convert to array of objects with headers
        if (jsonData.length > 0) {
          const headers = jsonData[0]
          const rows = jsonData.slice(1).map(row => {
            const obj = {}
            headers.forEach((header, idx) => {
              obj[header] = row[idx] || ''
            })
            return obj
          })
          
          fileData.sheets.push({
            name: sheetName,
            headers: headers,
            data: rows
          })
        }
      })
      
      processedFiles.push(fileData)
    } catch (error) {
      console.error(`Error processing file ${file.name}:`, error)
      throw new Error(`Failed to process ${file.name}: ${error.message}`)
    }
  }
  
  return processedFiles
}

export function computeMatrices(fileData, selectedTabs, columnSelections, matrixConfig) {
  const matrices = []
  
  matrixConfig.forEach((config) => {
    if (config.merge) {
      // Merge all sources into one matrix
      const yValues = new Set()
      const xValues = new Set()
      const secondaryXValues = new Set()
      const hasSecondaryX = config.sources.some(source => {
        const key = `${source.fileIndex}-${source.sheetName}`
        return !!columnSelections[key]?.secondaryXAxis
      })
      
      // Collect all unique values from all sources
      config.sources.forEach(source => {
        const file = fileData[source.fileIndex]
        const sheet = file.sheets.find(s => s.name === source.sheetName)
        
        if (!sheet) return
        
        const key = `${source.fileIndex}-${source.sheetName}`
        const yCol = columnSelections[key]?.yAxis
        const xCol = columnSelections[key]?.xAxis
        const secondaryXCol = columnSelections[key]?.secondaryXAxis
        
        sheet.data.forEach(row => {
          const yVal = String(row[yCol] || '').trim()
          const xVal = String(row[xCol] || '').trim()
          const secondaryXVal = secondaryXCol ? String(row[secondaryXCol] || '').trim() : null
          
          if (yVal) yValues.add(yVal)
          if (xVal) xValues.add(xVal)
          if (secondaryXVal) secondaryXValues.add(secondaryXVal)
        })
      })
      
      const sortedY = Array.from(yValues).sort()
      const sortedX = Array.from(xValues).sort()
      const sortedSecondaryX = hasSecondaryX ? Array.from(secondaryXValues).sort() : null
      
      if (sortedSecondaryX && sortedSecondaryX.length > 0) {
        // Create one matrix per secondary X value
        sortedSecondaryX.forEach(secX => {
          const matrix = {
            name: `${config.name} - ${secX}`,
            yAxis: sortedY,
            xAxis: sortedX,
            data: sortedY.map(() => sortedX.map(() => 0))
          }
          
          // Mark intersections
          config.sources.forEach(source => {
            const file = fileData[source.fileIndex]
            const sheet = file.sheets.find(s => s.name === source.sheetName)
            if (!sheet) return
            
            const key = `${source.fileIndex}-${source.sheetName}`
            const yCol = columnSelections[key]?.yAxis
            const xCol = columnSelections[key]?.xAxis
            const secondaryXCol = columnSelections[key]?.secondaryXAxis
            
            sheet.data.forEach(row => {
              const yVal = String(row[yCol] || '').trim()
              const xVal = String(row[xCol] || '').trim()
              const secondaryXVal = secondaryXCol ? String(row[secondaryXCol] || '').trim() : null
              
              if (yVal && xVal && secondaryXVal === secX) {
                const yIdx = sortedY.indexOf(yVal)
                const xIdx = sortedX.indexOf(xVal)
                if (yIdx >= 0 && xIdx >= 0) {
                  matrix.data[yIdx][xIdx] = 1
                }
              }
            })
          })
          
          matrices.push(matrix)
        })
      } else {
        // Single merged matrix
        const matrix = {
          name: config.name,
          yAxis: sortedY,
          xAxis: sortedX,
          data: sortedY.map(() => sortedX.map(() => 0))
        }
        
        // Mark intersections
        config.sources.forEach(source => {
          const file = fileData[source.fileIndex]
          const sheet = file.sheets.find(s => s.name === source.sheetName)
          if (!sheet) return
          
          const key = `${source.fileIndex}-${source.sheetName}`
          const yCol = columnSelections[key]?.yAxis
          const xCol = columnSelections[key]?.xAxis
          
          sheet.data.forEach(row => {
            const yVal = String(row[yCol] || '').trim()
            const xVal = String(row[xCol] || '').trim()
            
            if (yVal && xVal) {
              const yIdx = sortedY.indexOf(yVal)
              const xIdx = sortedX.indexOf(xVal)
              if (yIdx >= 0 && xIdx >= 0) {
                matrix.data[yIdx][xIdx] = 1
              }
            }
          })
        })
        
        matrices.push(matrix)
      }
    } else {
      // Create independent matrix for each source
      config.sources.forEach(source => {
        const file = fileData[source.fileIndex]
        const sheet = file.sheets.find(s => s.name === source.sheetName)
        if (!sheet) return
        
        const key = `${source.fileIndex}-${source.sheetName}`
        const yCol = columnSelections[key]?.yAxis
        const xCol = columnSelections[key]?.xAxis
        const secondaryXCol = columnSelections[key]?.secondaryXAxis
        
        const yValues = new Set()
        const xValues = new Set()
        const secondaryXValues = secondaryXCol ? new Set() : null
        
        sheet.data.forEach(row => {
          const yVal = String(row[yCol] || '').trim()
          const xVal = String(row[xCol] || '').trim()
          const secondaryXVal = secondaryXCol ? String(row[secondaryXCol] || '').trim() : null
          
          if (yVal) yValues.add(yVal)
          if (xVal) xValues.add(xVal)
          if (secondaryXVal && secondaryXValues) secondaryXValues.add(secondaryXVal)
        })
        
        const sortedY = Array.from(yValues).sort()
        const sortedX = Array.from(xValues).sort()
        const sortedSecondaryX = secondaryXValues ? Array.from(secondaryXValues).sort() : null
        
        if (sortedSecondaryX && sortedSecondaryX.length > 0) {
          // Create one matrix per secondary X value
          sortedSecondaryX.forEach(secX => {
            const matrix = {
              name: `${source.fileName} - ${source.sheetName} - ${secX}`,
              yAxis: sortedY,
              xAxis: sortedX,
              data: sortedY.map(() => sortedX.map(() => 0))
            }
            
            sheet.data.forEach(row => {
              const yVal = String(row[yCol] || '').trim()
              const xVal = String(row[xCol] || '').trim()
              const secondaryXVal = secondaryXCol ? String(row[secondaryXCol] || '').trim() : null
              
              if (yVal && xVal && secondaryXVal === secX) {
                const yIdx = sortedY.indexOf(yVal)
                const xIdx = sortedX.indexOf(xVal)
                if (yIdx >= 0 && xIdx >= 0) {
                  matrix.data[yIdx][xIdx] = 1
                }
              }
            })
            
            matrices.push(matrix)
          })
        } else {
          // Single matrix
          const matrix = {
            name: `${source.fileName} - ${source.sheetName}`,
            yAxis: sortedY,
            xAxis: sortedX,
            data: sortedY.map(() => sortedX.map(() => 0))
          }
          
          sheet.data.forEach(row => {
            const yVal = String(row[yCol] || '').trim()
            const xVal = String(row[xCol] || '').trim()
            
            if (yVal && xVal) {
              const yIdx = sortedY.indexOf(yVal)
              const xIdx = sortedX.indexOf(xVal)
              if (yIdx >= 0 && xIdx >= 0) {
                matrix.data[yIdx][xIdx] = 1
              }
            }
          })
          
          matrices.push(matrix)
        }
      })
    }
  })
  
  return matrices
}

export function exportToExcel(matrices) {
  const workbook = XLSX.utils.book_new()
  
  matrices.forEach(matrix => {
    // Create header row: empty cell + X axis values
    const headerRow = ['', ...matrix.xAxis]
    
    // Create data rows: Y axis value + matrix data
    const rows = [headerRow]
    matrix.yAxis.forEach((yVal, yIdx) => {
      const row = [yVal, ...matrix.data[yIdx]]
      rows.push(row)
    })
    
    // Convert to worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(rows)
    
    // Set column widths
    const maxWidth = 15
    worksheet['!cols'] = [
      { wch: maxWidth }, // Y axis column
      ...matrix.xAxis.map(() => ({ wch: maxWidth }))
    ]
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, matrix.name.substring(0, 31)) // Excel sheet name limit
  })
  
  // Generate filename with timestamp
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5)
  const filename = `matrices_${timestamp}.xlsx`
  
  // Write file
  XLSX.writeFile(workbook, filename)
}

