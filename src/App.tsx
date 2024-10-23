'use client'

import React, { useState } from 'react'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import Input from "./components/ui/input"
import './App.css';

function App() {
  const [files, setFiles] = useState<File[]>([])
  const [tables, setTables] = useState<any[]>([])
  const [fields, setFields] = useState<{ [key: string]: string[] }>({})
  const [selectedFields, setSelectedFields] = useState<{ [key: string]: string[] }>({})
  const [keyFields, setKeyFields] = useState<{ [key: string]: string }>({})
  const [mergedData, setMergedData] = useState<any[] | null>(null)
  const [sheets, setSheets] = useState<{ [key: string]: string[] }>({})
  const [selectedSheets, setSelectedSheets] = useState<{ [key: string]: string }>({})
  const [mergedPreview, setMergedPreview] = useState<any[] | null>(null)
  const [selectedFieldsOrder, setSelectedFieldsOrder] = useState<string[]>([]) // Добавляем новое состояние

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const newFiles = Array.from(event.target.files || [])
    setFiles([...files, ...newFiles])

    newFiles.forEach((file) => {
      const reader = new FileReader()
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const workbook = XLSX.read(data, { type: 'array' })
        const sheetNames = workbook.SheetNames
        
        setSheets(prevSheets => ({
          ...prevSheets,
          [file.name]: sheetNames
        }))

        if (sheetNames.length === 1) {
          handleSheetSelection(file.name, sheetNames[0])
        }
      }
      reader.readAsArrayBuffer(file)
    })
  }

  const handleSheetSelection = (fileName: string, sheetName: string) => {
    setSelectedSheets(prevSelected => ({
      ...prevSelected,
      [fileName]: sheetName
    }))

    const file = files.find(f => f.name === fileName)
    if (!file) return

    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer)
      const workbook = XLSX.read(data, { type: 'array' })
      const worksheet = workbook.Sheets[sheetName]
      
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')
      const endRow = Math.min(range.e.r, 49)
      const tempRange = { ...range, e: { ...range.e, r: endRow } }
      const partialJson = XLSX.utils.sheet_to_json(worksheet, { range: tempRange, header: 1 }) as Array<Array<any>>

      // Функция для проверки, содержит ли строка буквы
      const containsLetters = (str: string) => /[a-zA-Z]/.test(str)

      // Функция для подсчета значимых ячеек в строке
      const countSignificantCells = (row: Array<any>) => 
        row.filter(cell => 
          cell && 
          typeof cell === 'string' && 
          containsLetters(cell)
        ).length

      // Находим строку с наибольшим количеством значимых ячеек
      let headerRowIndex = 0
      let maxSignificantCells = 0

      partialJson.forEach((row, index) => {
        const significantCells = countSignificantCells(row)
        if (significantCells > maxSignificantCells) {
          maxSignificantCells = significantCells
          headerRowIndex = index
        }
      })

      // Используем найденную строку как заголовки
      const headerRow = partialJson[headerRowIndex]
      const headers: string[] = headerRow.map(cell => String(cell || '').trim())

      // Получаем все данные после заголовков
      const fullRange = { 
        ...range, 
        s: { ...range.s, r: headerRowIndex + 1 }
      }
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        range: fullRange,
        header: headers
      })

      setTables(prevTables => [...prevTables, jsonData])
      setFields(prevFields => ({
        ...prevFields,
        [fileName]: headers
      }))
      setSelectedFields(prevSelected => ({
        ...prevSelected,
        [fileName]: [],
      }))
      setKeyFields(prevKeys => ({
        ...prevKeys,
        [fileName]: '',
      }))
    }
    reader.readAsArrayBuffer(file)
  }

  const handleFieldSelection = (fileName: string, field: string) => {
    setSelectedFields((prevFields) => {
      const updatedFields = prevFields[fileName].includes(field)
        ? prevFields[fileName].filter((f) => f !== field)
        : [...prevFields[fileName], field]
      return {
        ...prevFields,
        [fileName]: updatedFields,
      }
    })
  }

  const handleKeyFieldSelection = (fileName: string, field: string) => {
    setKeyFields((prevKeys) => ({
      ...prevKeys,
      [fileName]: field,
    }))
  }

  const mergeTables = () => {
    if (tables.length < 2) {
      alert('Please upload both tables to merge.')
      return
    }

    const keyFieldSet = new Set(Object.values(keyFields))
    if (keyFieldSet.size === 0) {
      alert('Please select at least one key field for merging.')
      return
    }

    // Сохраняем порядок полей в том порядке, в котором они были выбраны
    const orderedFields = Object.entries(selectedFields).reduce((acc: string[], [_, fields]) => {
      fields.forEach(field => {
        if (!acc.includes(field)) acc.push(field)
      })
      return acc
    }, [])

    setSelectedFieldsOrder(orderedFields) // Сохраняем порядок полей

    let merged = tables[0]
    
    for (let i = 1; i < tables.length; i++) {
      const currentFile = files[i]
      const previousFile = files[i - 1]
      
      if (!currentFile || !previousFile) {
        alert('Error: Some files are missing. Please upload all required files.')
        return
      }

      const currentKeyField = keyFields[currentFile.name]
      const previousKeyField = keyFields[previousFile.name]

      if (!currentKeyField || !previousKeyField) {
        alert('Error: Key fields are not selected for all files. Please select key fields for all files.')
        return
      }

      // eslint-disable-next-line no-loop-func
      merged = merged.flatMap((row: any) => {
        const matchingRows = tables[i].filter((r: any) => r[currentKeyField] === row[previousKeyField])
        if (matchingRows.length > 0) {
          return matchingRows.map((match: any) => {
            const mergedRow: any = {}
            orderedFields.forEach(field => {
              if (match[field] !== undefined) {
                mergedRow[field] = match[field]
              } else if (row[field] !== undefined) {
                mergedRow[field] = row[field]
              }
            })
            return mergedRow
          })
        }
        return []
      })
    }

    setMergedData(merged)
    setMergedPreview(merged.slice(0, 10))
  }

  const downloadMergedFile = () => {
    if (!mergedData) return

    const worksheet = XLSX.utils.json_to_sheet(mergedData)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Merged')
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    const data = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    saveAs(data, 'merged_tables.xlsx')
  }

  return (
    <div className="App">
      <header className="App-header">
        <h1 className="text-3xl font-bold mb-6">Excel Table Merger</h1>
        <div className="file-container-wrapper">
          {[0, 1].map((index) => (
            <div key={index} className="file-container">
              <h2 className="text-xl font-semibold mb-4">File {index + 1}</h2>
              <label htmlFor={`file-input-${index}`} className="mb-2 block">Choose Excel file:</label>
              <Input 
                id={`file-input-${index}`}
                type="file" 
                accept=".xlsx,.xls" 
                onChange={handleFileUpload} 
                className="mb-4 w-full p-2 border border-gray-300 rounded"
              />
              {!files[index] && (
                <p className="text-gray-500 mb-4">No file selected</p>
              )}
              {files[index] && sheets[files[index].name] && sheets[files[index].name].length > 0 && (
                <div className="mb-4" style={{ width: '100%' }}>
                  <select
                    value={selectedSheets[files[index].name] || ''}
                    onChange={(e) => handleSheetSelection(files[index].name, e.target.value)}
                    style={{
                      width: '100%',
                      padding: '8px',
                      border: '1px solid #ccc',
                      borderRadius: '4px',
                      backgroundColor: 'white',
                      color: 'black',
                      fontSize: '14px'
                    }}
                  >
                    <option value="">Select a sheet</option>
                    {sheets[files[index].name].map((sheet) => (
                      <option key={sheet} value={sheet}>{sheet}</option>
                    ))}
                  </select>
                  {selectedSheets[files[index].name] && (
                    <p style={{ marginTop: '8px', fontSize: '14px', color: 'black' }}>
                      Selected sheet: {selectedSheets[files[index].name]}
                    </p>
                  )}
                </div>
              )}
              {files[index] && selectedSheets[files[index].name] && (
                <div className="file-content">
                  <div className="fields-column">
                    <h3 className="font-medium mb-2">Fields:</h3>
                    {fields[files[index].name]?.map((field) => (
                      <div key={field} className="field-item">
                        {field}
                      </div>
                    ))}
                  </div>
                  <div className="checkbox-column">
                    <h3 className="font-medium mb-2">Select:</h3>
                    {fields[files[index].name]?.map((field) => (
                      <div key={field} className="checkbox-container">
                        <input
                          type="checkbox"
                          id={`field-${files[index].name}-${field}`}
                          className="checkbox"
                          checked={selectedFields[files[index].name]?.includes(field)}
                          onChange={() => handleFieldSelection(files[index].name, field)}
                        />
                      </div>
                    ))}
                  </div>
                  <div className="key-column">
                    <h3 className="font-medium mb-2">Key field:</h3>
                    <select
                      value={keyFields[files[index].name] || ''}
                      onChange={(e) => handleKeyFieldSelection(files[index].name, e.target.value)}
                      style={{
                        width: '100%',
                        padding: '8px',
                        border: '1px solid #ccc',
                        borderRadius: '4px',
                        backgroundColor: 'white',
                        color: 'black',
                        fontSize: '14px'
                      }}
                    >
                      <option value="">Select a key field</option>
                      {fields[files[index].name]?.map((field) => (
                        <option key={field} value={field}>{field}</option>
                      ))}
                    </select>
                  </div>
                </div>
              )}
            </div>
          ))}
        </div>
        <div className="button-container">
          <button 
            onClick={mergeTables} 
            disabled={files.length < 2} 
            style={{
              padding: '8px 16px',
              border: '1px solid #ccc',
              borderRadius: '4px',
              backgroundColor: 'white',
              color: 'black',
              fontSize: '14px',
              cursor: 'pointer',
              marginRight: '10px'
            }}
          >
            Merge
          </button>
          <button 
            onClick={downloadMergedFile} 
            disabled={!mergedData}
            style={{
              padding: '8px 16px',
              border: '1px solid #ccc',
              borderRadius: '4px',
              backgroundColor: 'white',
              color: 'black',
              fontSize: '14px',
              cursor: 'pointer'
            }}
          >
            Download
          </button>
        </div>
      </header>
      
      {mergedPreview && mergedPreview.length > 0 && (
        <div className="merged-preview" style={{ margin: '20px 0' }}>
          <h2 className="text-xl font-semibold mb-4">Merged Data Preview</h2>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ 
              width: '100%', 
              borderCollapse: 'collapse',
              fontSize: '14px'
            }}>
              <thead>
                <tr>
                  {selectedFieldsOrder.map((field: string) => (
                    <th key={field} style={{ 
                      padding: '12px 8px',
                      borderBottom: '2px solid #ddd',
                      textAlign: 'left'
                    }}>
                      {field}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {mergedPreview.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {selectedFieldsOrder.map((field: string, cellIndex: number) => (
                      <td key={`${rowIndex}-${cellIndex}`} style={{ 
                        padding: '8px',
                        borderBottom: '1px solid #ddd'
                      }}>
                        {row[field] !== undefined ? String(row[field]) : ''}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {mergedData && mergedData.length > 10 && (
            <p style={{ marginTop: '10px', color: '#666' }}>
              Showing first 10 of {mergedData.length} rows
            </p>
          )}
        </div>
      )}
    </div>
  );
}

export default App;
