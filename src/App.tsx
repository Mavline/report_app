"use client";

import React, { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import Input from "./components/ui/input"
import './App.css';


// В начале файла добавьте определение типа
interface GroupInfo {
  level: number;
  group: number[];
  hidden: boolean;
  parent?: number;
}

// В начале файла добавим проверку окружения
const isProduction = process.env.NODE_ENV === 'production';
console.log('Environment:', process.env.NODE_ENV);

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
  const [selectedFieldsOrder, setSelectedFieldsOrder] = useState<string[]>([])
  const [isGrouped, setIsGrouped] = useState<{ [key: string]: boolean }>({})
  const [groupingStructure, setGroupingStructure] = useState<{ [key: string]: { [key: string]: GroupInfo } }>({})

  // Добавим эффект для отслеживания жизненного цикла
  useEffect(() => {
    console.log('Component lifecycle:', {
      mergedData: !!mergedData,
      selectedFieldsOrder: !!selectedFieldsOrder,
      files: files.length,
      tables: tables.length,
      isProduction
    });
  }, [mergedData, selectedFieldsOrder, files, tables]);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    console.log("File upload started");
    const newFiles = Array.from(event.target.files || [])
    console.log("New files:", newFiles.map(f => f.name));

    for (const file of newFiles) {
      console.log(`Processing file: ${file.name}`);
      const reader = new FileReader()
      reader.onload = async (e) => {
        console.log(`File ${file.name} loaded`);
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const workbook = XLSX.read(data, { type: 'array' })
        const sheetNames = workbook.SheetNames
        console.log(`Sheets in ${file.name}:`, sheetNames);
        
        setFiles(prevFiles => [...prevFiles, file]);
        setSheets(prevSheets => ({
          ...prevSheets,
          [file.name]: sheetNames
        }))

        // Убираем автоматическую обработку листа
        // if (sheetNames.length === 1) {
        //   await processSheet(file, sheetNames[0])
        // }
      }
      reader.readAsArrayBuffer(file)
    }
  }

  const processSheet = async (file: File, sheetName: string) => {
    try {
      const arrayBuffer = await file.arrayBuffer()
      console.log(`ArrayBuffer obtained for ${file.name}`);
      const zip = new JSZip()
      const zipContents = await zip.loadAsync(arrayBuffer)
      
      console.log('Files in ZIP:', Object.keys(zipContents.files));
      
      let sheetXmlPath = `xl/worksheets/sheet${sheetName}.xml`;
      if (!zipContents.files[sheetXmlPath]) {
        const sheetIndex = 1;
        sheetXmlPath = `xl/worksheets/sheet${sheetIndex}.xml`;
      }
      
      console.log(`Trying to access sheet XML at path: ${sheetXmlPath}`);
      const sheetXml = await zipContents.file(sheetXmlPath)?.async('string')

      // Обработка XML и поиск заоловка
      const workbook = XLSX.read(arrayBuffer, { type: 'array' })
      const worksheet = workbook.Sheets[sheetName]
      
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')
      const endRow = Math.min(range.e.r, 49)
      const tempRange = { ...range, e: { ...range.e, r: endRow } }
      const partialJson = XLSX.utils.sheet_to_json(worksheet, { range: tempRange, header: 1 }) as Array<Array<any>>

      // Функция для пррки, содержт и трока буквы
      const containsLetters = (str: string) => /[a-zA-Z]/.test(str)


      // Функция для подсчета значимых ячеек в строке
      const countSignificantCells = (row: Array<any>) =>
        row.filter(
          (cell) => cell && typeof cell === "string" && containsLetters(cell),
        ).length;

      // Находим строку с наибольшим количеством значимых ячеек
      let headerRowIndex = 0;
      let maxSignificantCells = 0;

      partialJson.forEach((row, index) => {
        const significantCells = countSignificantCells(row);
        if (significantCells > maxSignificantCells) {
          maxSignificantCells = significantCells;
          headerRowIndex = index;
        }
      });


      // Используем найденную строку как заголовки, но берем ТОЛЬКО названия столбцов
      const headerRow = partialJson[headerRowIndex];
      const headers: string[] = headerRow.map(cell => String(cell || '').trim());


      // Получаем все данные после заголовков
      const fullRange = {
        ...range,
        s: { ...range.s, r: headerRowIndex + 1 },
      };
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        range: fullRange,
        header: headers,
      });

      console.log('Header row index:', headerRowIndex);
      console.log('JSON Data length:', jsonData.length);
      console.log('First few rows:', jsonData.slice(0, 5));

      setTables(prevTables => {
        console.log('Setting table data:', jsonData);
        return [...prevTables, jsonData];
      });

      setFields(prevFields => ({
        ...prevFields,
        [file.name]: headers
      }))
      setSelectedFields(prevSelected => ({
        ...prevSelected,
        [file.name]: [],
      }))
      setKeyFields(prevKeys => ({
        ...prevKeys,
        [file.name]: '',
      }))

      setSelectedSheets(prevSelected => ({
        ...prevSelected,
        [file.name]: sheetName
      }))

      // Обработка XML л группировки
      if (sheetXml) {
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
        const parsedXml = parser.parse(sheetXml);
        
        if (parsedXml.worksheet && parsedXml.worksheet.sheetData && parsedXml.worksheet.sheetData.row) {
          const rows = parsedXml.worksheet.sheetData.row;
          const groupingInfo = extractGroupingInfo(rows, headerRowIndex);
          
          setGroupingStructure(prevStructure => ({
            ...prevStructure,
            [file.name]: groupingInfo
          }));

          setIsGrouped(prevGrouped => ({
            ...prevGrouped,
            [file.name]: Object.keys(groupingInfo).length > 0
          }));
        }
      }

    } catch (error) {
      console.error('Error in processSheet:', error);
    }
  }

  const extractGroupingInfo = (rows: any[], headerOffset: number) => {
    const groupingInfo: { [key: string]: GroupInfo } = {};
    
    rows.forEach((row: any) => {
      const rowIndex = parseInt(row['@_r']);
      
      // Пропускаем строки до заголовка
      if (rowIndex <= headerOffset) {
        return;
      }

      // Получаем уровень группировки из XML
      const outlineLevel = parseInt(row['@_outlineLevel'] || '0');
      
      // Схраняем информацию о группировке со скорректированны индексом
      const adjustedIndex = rowIndex - headerOffset;
      groupingInfo[adjustedIndex.toString()] = {
        level: outlineLevel,
        group: [adjustedIndex],
        hidden: row['@_hidden'] === '1'
      };
    });

    return groupingInfo;
  };

  const handleSheetSelection = (fileName: string, sheetName: string) => {
    const file = files.find(f => f.name === fileName)
    if (file) {
      processSheet(file, sheetName)
    } else {
      console.error(`File not found: ${fileName}`);
    }
  }

  const handleFieldSelection = (fileName: string, field: string) => {
    setSelectedFields((prevFields) => {
      const updatedFields = prevFields[fileName].includes(field)
        ? prevFields[fileName].filter((f) => f !== field)
        : [...prevFields[fileName], field];
      return {
        ...prevFields,
        [fileName]: updatedFields,
      };
    });
  };

  const handleKeyFieldSelection = (fileName: string, field: string) => {
    setKeyFields((prevKeys) => ({
      ...prevKeys,
      [fileName]: field,
    }));
  };

  const mergeTables = () => {
    console.log('Starting merge process...');
    
    if (tables.length < 2) {
      alert("Please upload both tables to merge.");
      return;
    }

    const keyFieldSet = new Set(Object.values(keyFields));
    if (keyFieldSet.size === 0) {
      alert("Please select at least one key field for merging.");
      return;
    }

    // Получаем информацию о группировке
    const groupedFile = files[0];
    const groupInfo = groupingStructure[groupedFile.name];
    const maxLevel = groupInfo ? Math.max(...Object.values(groupInfo).map(info => info.level)) : 0;

    // Создаем заголовки для уровней группировки
    const groupHeaders = Array.from({ length: maxLevel + 1 }, (_, i) => `Level_${i + 1}`);

    // Создаем заголовки данных, сохраняя порядок полей из исходных таблиц
    const dataHeaders: string[] = [];
    files.forEach(file => {
      const fileFields = fields[file.name].filter(field => selectedFields[file.name].includes(field));
      dataHeaders.push(...fileFields);
    });

    // Все заголовки
    const allHeaders = [...groupHeaders, 'LevelValue', ...dataHeaders];

    // Функция для создания базовой строки с учетом группировки
    const createBaseRow = (rowIndex: number) => {
      const row: Record<string, any> = {};

      if (groupInfo) {
        // Инициализируем все уровни группировки пустыи строками
        groupHeaders.forEach((header) => {
          row[header] = '';
        });

        // Получаем информацию о группировке для текущей строки
        const groupData = groupInfo[rowIndex + 2]; // +1 для учета индексации
        if (groupData) {
          const level = groupData.level;
          if (level >= 0 && level < groupHeaders.length) {
            row[groupHeaders[level]] = groupData.level + 1;
            row['LevelValue'] = groupData.level + 1; // Добавляем LevelValue
          }
        }
      }

      return row;
    };

    // Функция для создания строк с совпадениями без удаления повторов
    const createRowsWithMatches = (
      firstTableRow: Record<string, any>, 
      rowIndex: number
    ): any[] => {
      const firstTableKeyField = keyFields[files[0].name];
      const secondTable = tables[1];
      const secondKeyField = keyFields[files[1].name];
      
      const matchingRows = secondTable.filter((r: Record<string, any>) => 
        r[secondKeyField] === firstTableRow[firstTableKeyField]
      );

      // Если нет совпадений, возвращаем одну строку с данными из первой таблицы
      if (matchingRows.length === 0) {
        const baseRow = createBaseRow(rowIndex);
        // Используем fields вместо selectedFields для сохранения порядка
        fields[files[0].name].forEach(field => {
          if (selectedFields[files[0].name].includes(field)) {
            baseRow[field] = firstTableRow[field];
          }
        });
        return [baseRow];
      }

      // Создаем строки для каждого совпадения без удаления повторов
      return matchingRows.map((matchingRow: Record<string, any>, matchIndex: number) => {
        const baseRow = createBaseRow(rowIndex);

        // Добавляем данные из первой таблицы только в первую строку
        if (matchIndex === 0) {
          fields[files[0].name].forEach(field => {
            if (selectedFields[files[0].name].includes(field)) {
              baseRow[field] = firstTableRow[field];
            }
          });
        } else {
          // Оставляем поля из первой таблицы пустыми в последующих строках
          fields[files[0].name].forEach(field => {
            if (selectedFields[files[0].name].includes(field)) {
              baseRow[field] = '';
            }
          });
        }

        // Добавляем данные из второй таблицы, сохраняя исходный порядок полей
        fields[files[1].name].forEach(field => {
          if (selectedFields[files[1].name].includes(field)) {
            baseRow[field] = matchingRow[field];
          }
        });

        return baseRow;
      });
    };

    // Создаем объединенные данные
    const merged = tables[0]
      .flatMap((row: Record<string, any>, index: number) => 
        createRowsWithMatches(row, index)
      );

    // Обрабатываем LevelValue с использованием selectedFieldsOrder
    const processedData = merged.map((row: { [key: string]: any }) => {
      const entries = Object.entries(row);
      const levelValueIndex = entries.findIndex(([key]) => key === 'LevelValue');
      
      if (levelValueIndex !== -1) {
        const nextFieldEntry = entries[levelValueIndex + 1];
        if (nextFieldEntry) {
          const [, nextFieldValue] = nextFieldEntry;
          if (!nextFieldValue || nextFieldValue === '') {
            entries[levelValueIndex][1] = '';
          }
        }
      }

      return Object.fromEntries(entries);
    });

    // Обновляем состояния с обработанными данными
    setMergedData(processedData);
    setMergedPreview(processedData.slice(0, 10));
    setSelectedFieldsOrder(allHeaders);

    console.log('Final headers:', allHeaders);
    console.log('Final merged data:', processedData);
  };

  const downloadMergedFile = () => {
    if (!mergedData || mergedData.length === 0) {
      alert('No data to download. Please merge tables first.')
      return
    }

    const worksheet = XLSX.utils.json_to_sheet(mergedData)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Merged')

    // Добавляем группировку
    const groupedFile = files.find(file => isGrouped[file.name])
    if (groupedFile) {
      const groupInfo = groupingStructure[groupedFile.name] as { [key: string]: GroupInfo }
      const maxLevel = Math.max(...Object.values(groupInfo).map(info => info.level), 0)
      
      for (let i = 0; i <= maxLevel; i++) {
        worksheet['!outline'] = { ...worksheet['!outline'], [i]: 1 }
      }
    }

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
              <label htmlFor={`file-input-${index}`} className="mb-2 block">
                Choose Excel file:
              </label>
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
              {files[index] && sheets[files[index].name] && (
                <div className="mb-4" style={{ width: "100%" }}>
                  <select
                    value={selectedSheets[files[index].name] || ""}
                    onChange={(e) =>
                      handleSheetSelection(files[index].name, e.target.value)
                    }
                    style={{
                      width: "100%",
                      padding: "8px",
                      border: "1px solid #ccc",
                      borderRadius: "4px",
                      backgroundColor: "white",
                      color: "black",
                      fontSize: "14px",
                    }}
                  >
                    <option value="">Select a sheet</option>
                    {sheets[files[index].name].map((sheet, sheetIndex) => (

                      <option key={`${sheet}-${sheetIndex}`} value={sheet}>{sheet}</option>

                    ))}
                  </select>
                </div>
              )}
              
              {files[index] && selectedSheets[files[index].name] && (
                <div className="file-content">
                  <div className="fields-column">
                    <h3 className="font-medium mb-2">Fields:</h3>
                    {fields[files[index].name]?.map((field, fieldIndex) => (
                      <div key={`${field}-${fieldIndex}`} className="field-item">
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
                      value={keyFields[files[index].name] || ""}
                      onChange={(e) =>
                        handleKeyFieldSelection(
                          files[index].name,
                          e.target.value,
                        )
                      }
                      style={{
                        width: "100%",
                        padding: "8px",
                        border: "1px solid #ccc",
                        borderRadius: "4px",
                        backgroundColor: "white",
                        color: "black",
                        fontSize: "14px",
                      }}
                    >
                      <option value="">Select a key field</option>
                      {fields[files[index].name]?.map((field) => (
                        <option key={field} value={field}>
                          {field}
                        </option>
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
              padding: "8px 16px",
              border: "1px solid #ccc",
              borderRadius: "4px",
              backgroundColor: "white",
              color: "black",
              fontSize: "14px",
              cursor: "pointer",
              marginRight: "10px",
            }}
          >
            Merge
          </button>
          <button
            onClick={downloadMergedFile}
            disabled={!mergedData}
            style={{
              padding: "8px 16px",
              border: "1px solid #ccc",
              borderRadius: "4px",
              backgroundColor: "white",
              color: "black",
              fontSize: "14px",
              cursor: "pointer",
            }}
          >
            Download
          </button>
        </div>
      </header>

      {mergedPreview && mergedPreview.length > 0 && (
        <div className="merged-preview" style={{ margin: "20px 0" }}>
          <h2 className="text-xl font-semibold mb-4">Merged Data Preview</h2>
          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: "14px",
              }}
            >
              <thead>
                <tr>
                  {selectedFieldsOrder.map((field: string) => (
                    <th
                      key={field}
                      style={{
                        padding: "12px 8px",
                        borderBottom: "2px solid #ddd",
                        textAlign: "left",
                      }}
                    >
                      {field}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {mergedPreview.map((row, rowIndex) => (
                  <tr key={rowIndex}>
                    {selectedFieldsOrder.map(
                      (field: string, cellIndex: number) => (
                        <td
                          key={`${rowIndex}-${cellIndex}`}
                          style={{
                            padding: "8px",
                            borderBottom: "1px solid #ddd",
                          }}
                        >
                          {row[field] !== undefined ? String(row[field]) : ""}
                        </td>
                      ),
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {mergedData && mergedData.length > 10 && (
            <p style={{ marginTop: "10px", color: "#666" }}>
              Showing first 10 of {mergedData.length} rows
            </p>
          )}
        </div>
      )}
    </div>
  );
}

export default App;

