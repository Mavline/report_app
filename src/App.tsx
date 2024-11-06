"use client";

import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import Input from "./components/ui/input";
import './App.css';
import ExcelJS from 'exceljs';

// Define the GroupInfo type
interface GroupInfo {
  level: number;
  group: number[];
  hidden: boolean;
  parent?: number;
}

// Define the TableRow type
type TableRow = Record<string, any>;

const App: React.FC = () => {

  const [files, setFiles] = useState<File[]>([]);
  const [tables, setTables] = useState<TableRow[][]>([]);
  const [fields, setFields] = useState<{ [key: string]: string[] }>({});
  const [selectedFields, setSelectedFields] = useState<{ [key: string]: string[] }>({});
  const [keyFields, setKeyFields] = useState<{ [key: string]: string }>({});
  const [mergedData, setMergedData] = useState<TableRow[] | null>(null);
  const [sheets, setSheets] = useState<{ [key: string]: string[] }>({});
  const [selectedSheets, setSelectedSheets] = useState<{ [key: string]: string }>({});
  const [mergedPreview, setMergedPreview] = useState<TableRow[] | null>(null);
  const [selectedFieldsOrder, setSelectedFieldsOrder] = useState<string[]>([]);
  const [isGrouped, setIsGrouped] = useState<{ [key: string]: boolean }>({});
  const [groupingStructure, setGroupingStructure] = useState<{ [key: string]: { [key: string]: GroupInfo } }>({});
  const [columnToProcess, setColumnToProcess] = useState<string>('');
  const [secondColumnToProcess, setSecondColumnToProcess] = useState<string>('');

  useEffect(() => {
    // Logging component lifecycle
    console.log('Component lifecycle:', {
      mergedData: !!mergedData,
      selectedFieldsOrder: !!selectedFieldsOrder,
      files: files.length,
      tables: tables.length,
    });
  }, [mergedData, selectedFieldsOrder, files, tables]);

  useEffect(() => {
    const allSelectedFields: string[] = [];

    // Собираем все выбранные поля из обеих таблиц
    files.forEach(file => {
      const fileFields = selectedFields[file.name] || [];
      allSelectedFields.push(...fileFields);
    });

    // Обновляем selectedFieldsOrder
    setSelectedFieldsOrder(allSelectedFields);
  }, [selectedFields, files]);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    console.log("File upload started");
    const newFiles = Array.from(event.target.files || []);
    console.log("New files:", newFiles.map(f => f.name));

    for (const file of newFiles) {
      console.log(`Processing file: ${file.name}`);
      const reader = new FileReader();
      reader.onload = async (e) => {
        console.log(`File ${file.name} loaded`);
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetNames = workbook.SheetNames;
        console.log(`Sheets in ${file.name}:`, sheetNames);

        setFiles(prevFiles => [...prevFiles, file]);
        setSheets(prevSheets => ({
          ...prevSheets,
          [file.name]: sheetNames
        }));
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const processSheet = async (file: File, sheetName: string) => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      console.log(`ArrayBuffer obtained for ${file.name}`);

      const zip = new JSZip();
      const zipContents = await zip.loadAsync(arrayBuffer);

      console.log('Files in ZIP:', Object.keys(zipContents.files));

      let sheetXmlPath = `xl/worksheets/sheet${sheetName}.xml`;
      if (!zipContents.files[sheetXmlPath]) {
        const sheetIndex = 1;
        sheetXmlPath = `xl/worksheets/sheet${sheetIndex}.xml`;
      }

      console.log(`Trying to access sheet XML at path: ${sheetXmlPath}`);
      const sheetXml = await zipContents.file(sheetXmlPath)?.async('string');

      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const worksheet = workbook.Sheets[sheetName];

      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
      const endRow = Math.min(range.e.r, 49);
      const tempRange = { ...range, e: { ...range.e, r: endRow } };
      const partialJson = XLSX.utils.sheet_to_json(worksheet, { range: tempRange, header: 1 }) as any[][];

      const containsLetters = (str: string) => /[a-zA-Z]/.test(str);

      const countSignificantCells = (row: any[]) =>
        row.filter(
          (cell) => cell && typeof cell === "string" && containsLetters(cell),
        ).length;

      let headerRowIndex = 0;
      let maxSignificantCells = 0;

      partialJson.forEach((row, index) => {
        const significantCells = countSignificantCells(row);
        if (significantCells > maxSignificantCells) {
          maxSignificantCells = significantCells;
          headerRowIndex = index;
        }
      });

      const headerRow = partialJson[headerRowIndex];
      const headers: string[] = headerRow.map(cell => String(cell || '').trim());

      const fullRange = {
        ...range,
        s: { ...range.s, r: headerRowIndex + 1 },
      };
      const jsonData = XLSX.utils.sheet_to_json<TableRow>(worksheet, {
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
      }));
      setSelectedFields(prevSelected => ({
        ...prevSelected,
        [file.name]: [],
      }));
      setKeyFields(prevKeys => ({
        ...prevKeys,
        [file.name]: '',
      }));

      setSelectedSheets(prevSelected => ({
        ...prevSelected,
        [file.name]: sheetName
      }));

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
  };

  const extractGroupingInfo = (rows: any[], headerOffset: number): { [key: string]: GroupInfo } => {
    const groupingInfo: { [key: string]: GroupInfo } = {};

    rows.forEach((row: any) => {
      const rowIndex = parseInt(row['@_r']);

      if (rowIndex <= headerOffset) {
        return;
      }

      const outlineLevel = parseInt(row['@_outlineLevel'] || '0');

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
    const file = files.find(f => f.name === fileName);
    if (file) {
      processSheet(file, sheetName);
    } else {
      console.error(`File not found: ${fileName}`);
    }
  };

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

    const groupedFile = files[0];
    const groupInfo = groupingStructure[groupedFile.name];
    const maxLevel = groupInfo ? Math.max(...Object.values(groupInfo).map(info => info.level)) : 0;

    const groupHeaders = Array.from({ length: maxLevel + 1 }, (_, i) => `Level_${i + 1}`);

    const dataHeaders: string[] = [];
    files.forEach(file => {
      const fileFields = fields[file.name].filter(field => selectedFields[file.name].includes(field));
      dataHeaders.push(...fileFields);
    });

    // Add the extra column 'Note' at the end
    const allHeaders = [...groupHeaders, 'LevelValue', ...dataHeaders, 'Note'];

    const createBaseRow = (rowIndex: number): TableRow => {
      const row: TableRow = {};

      if (groupInfo) {
        groupHeaders.forEach((header) => {
          row[header] = '';
        });

        const groupData = groupInfo[(rowIndex + 2).toString()];
        if (groupData) {
          const level = groupData.level;
          if (level >= 0 && level < groupHeaders.length) {
            const levelValue = groupData.level + 1;
            row[groupHeaders[level]] = levelValue;
            // Добавляем точки перед числом, сохраняя само число
            const dots = '.'.repeat(levelValue + 1); // +1 чтобы начать с двух точек для уровня 1
            row['LevelValue'] = levelValue; // Сохраняем числовое значение
            row['LevelValue'] = `${dots}${levelValue}`; // Добавляем точки перед числом
          }
        }
      }

      return row;
    };

    const createRowsWithMatches = (
      firstTableRow: TableRow,
      rowIndex: number
    ): TableRow[] => {
      const firstTableKeyField = keyFields[files[0].name];
      const secondTable = tables[1];
      const secondKeyField = keyFields[files[1].name];

      const matchingRows = secondTable.filter((r: TableRow) =>
        r[secondKeyField] === firstTableRow[firstTableKeyField]
      );

      if (matchingRows.length === 0) {
        const baseRow = createBaseRow(rowIndex);

        // Add data from the first table
        fields[files[0].name].forEach(field => {
          if (selectedFields[files[0].name].includes(field)) {
            baseRow[field] = firstTableRow[field];
          }
        });

        // Add empty values for fields from the second table
        fields[files[1].name].forEach(field => {
          if (selectedFields[files[1].name].includes(field)) {
            baseRow[field] = '';
          }
        });

        // Add visual marker in the 'Note' column
        baseRow['Note'] = '******';

        return [baseRow];
      }

      // Remove duplicates from matchingRows, excluding Level fields
      const uniqueMatchingRowsMap = new Map<string, TableRow>();
      matchingRows.forEach((row) => {
        const fieldsToConsider = selectedFields[files[1].name].filter(field => !field.startsWith('Level'));
        const key = fieldsToConsider.map(field => row[field]).join('|');
        if (!uniqueMatchingRowsMap.has(key)) {
          uniqueMatchingRowsMap.set(key, row);
        }
      });
      const uniqueMatchingRows = Array.from(uniqueMatchingRowsMap.values());

      return uniqueMatchingRows.map((matchingRow: TableRow, matchIndex: number) => {
        const baseRow = createBaseRow(rowIndex);

        // Add data from the first table only in the first matching row
        if (matchIndex === 0) {
          fields[files[0].name].forEach(field => {
            if (selectedFields[files[0].name].includes(field)) {
              baseRow[field] = firstTableRow[field];
            }
          });
        } else {
          fields[files[0].name].forEach(field => {
            if (selectedFields[files[0].name].includes(field)) {
              baseRow[field] = '';
            }
          });
        }

        // Add data from the second table
        fields[files[1].name].forEach(field => {
          if (selectedFields[files[1].name].includes(field)) {
            baseRow[field] = matchingRow[field];
          }
        });

        // Leave 'Note' column empty for matched rows
        baseRow['Note'] = '';

        return baseRow;
      });
    };

    // Merge the tables
    const merged = tables[0]
      .flatMap((row: TableRow, index: number) =>
        createRowsWithMatches(row, index)
      );

    // Process LevelValue
    const processedData = merged.map((row: TableRow) => {
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

    // Filter out rows where all data fields are empty (excluding Level fields)
    const dataFields = allHeaders.filter(header => !header.startsWith('Level') && header !== 'Note');
    const filteredData = processedData.filter(row => {
      return dataFields.some(field => {
        const value = row[field];
        return value !== undefined && value !== null && value !== '';
      });
    });

    // Добавляем обрабоку выбранного столбца для расширения диапазонов
    if (columnToProcess || secondColumnToProcess) {
      const processedDataWithExpandedRanges = filteredData.map((row) => {
        // Обработка первой выбранной колонки
    if (columnToProcess) {
        const cellValue = row[columnToProcess];
        if (typeof cellValue === 'string' && cellValue.includes('-')) {
          row[columnToProcess] = expandRanges(cellValue);
          }
        }
        
        // Обработка второй выбранной колонки
        if (secondColumnToProcess) {
          const cellValue = row[secondColumnToProcess];
          if (typeof cellValue === 'string' && cellValue.includes('-')) {
            row[secondColumnToProcess] = expandRanges(cellValue);
          }
        }
        return row;
      });
      setMergedData(processedDataWithExpandedRanges);
      setMergedPreview(processedDataWithExpandedRanges.slice(0, 10));
    } else {
      setMergedData(filteredData);
      setMergedPreview(filteredData.slice(0, 10));
    }

    setSelectedFieldsOrder(allHeaders);
    console.log('Final headers:', allHeaders);
    console.log('Final merged data:', mergedData);
  };

  const downloadMergedFile = async () => {
    if (!mergedData || mergedData.length === 0) {
      alert('No data to download. Please merge tables first.');
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Merged');

    // Добавляем заголовки и данные
    worksheet.columns = selectedFieldsOrder.map(header => ({
      header,
      key: header
    }));
    worksheet.addRows(mergedData);

    // Применяем стили после добавления всех данных
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Пропускаем заголовки

      const rowData = mergedData[rowNumber - 2];
      const hasLevelValue = rowData && rowData['LevelValue'];

      // Если есть LevelValue, применяем стили ко всей строке
      if (hasLevelValue) {
        row.eachCell({ includeEmpty: true }, cell => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFC5' } // Бледно-желтый цвет
          };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      } else {
        // Для строк без LevelValue добавляем только вертикальные границы
        row.eachCell({ includeEmpty: true }, cell => {
          cell.border = {
            left: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    saveAs(blob, 'merged_tables.xlsx');
  };

  const expandRanges = (value: string): string => {
    const parts = value.split(',');
    const expandedParts: string[] = [];

    parts.forEach((part) => {
      part = part.trim();
      if (part.includes('-')) {
        const [start, end] = part.split('-').map((s) => s.trim());

        const startMatch = start.match(/^([A-Za-z]*)(\d+)$/);
        const endMatch = end.match(/^([A-Za-z]*)(\d+)$/);

        if (startMatch && endMatch) {
          const startPrefix = startMatch[1];
          const startNum = parseInt(startMatch[2], 10);

          const endPrefix = endMatch[1];
          const endNum = parseInt(endMatch[2], 10);

          if (startPrefix === endPrefix) {
            if (startNum <= endNum) {
              for (let i = startNum; i <= endNum; i++) {
                expandedParts.push(`${startPrefix}${i}`);
              }
            } else {
              for (let i = startNum; i >= endNum; i--) {
                expandedParts.push(`${startPrefix}${i}`);
              }
            }
          } else {
            expandedParts.push(part);
          }
        } else {
          expandedParts.push(part);
        }
      } else {
        expandedParts.push(part);
      }
    });

    return expandedParts.join(',');
  };

  // Добавим функцию для сброса состояния
  const handleReset = () => {
    // Очищаем все состояния
    setFiles([]);
    setTables([]);
    setFields({});
    setSelectedFields({});
    setKeyFields({});
    setMergedData(null);
    setSheets({});
    setSelectedSheets({});
    setMergedPreview(null);
    setSelectedFieldsOrder([]);
    setIsGrouped({});
    setGroupingStructure({});
    setColumnToProcess('');
    setSecondColumnToProcess('');
    
    // Перезагружаем страницу
    window.location.reload();
  };

  return (
    <div className="App">
      <header className="App-header">
        {/* Добавляем кнопку RESET */}
        <div className="reset-container" style={{ 
          width: '100%',
          padding: '20px',
          backgroundColor: '#015f60',
          marginBottom: '20px',
          display: 'flex',
          alignItems: 'center',
          gap: '20px'
        }}>
          <button
            onClick={handleReset}
            style={{
              padding: "12px 24px",
              backgroundColor: "#dc3545",
              color: "white",
              border: "none",
              borderRadius: "4px",
              fontSize: "16px",
              fontWeight: "bold",
              cursor: "pointer",
              boxShadow: "0 2px 4px rgba(0,0,0,0.2)",
            }}
          >
            RESET
          </button>
          <span style={{
            fontSize: "30px",
            color: "#59fafc",
            fontStyle: "Arial"
          }}>
            Start over, refresh process or clear memory
          </span>
        </div>

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
                style={{
                  backgroundColor: "#59fafc",
                  color: "black"
                }}
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
                      backgroundColor: "#59fafc",
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
                        backgroundColor: "#59fafc",
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
        <div className="controls-container">
          {/* Первый селектор столбца */}
          <div className="range-selector" style={{ marginBottom: '10px' }}>
            <select
              value={columnToProcess}
              onChange={(e) => setColumnToProcess(e.target.value)}
              style={{
                width: "100%",
                padding: "8px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                backgroundColor: "#59fafc",
                color: "black",
                fontSize: "14px",
              }}
            >
              <option value="">Select First Column to Expand Ranges</option>
              {selectedFieldsOrder.map((field) => (
                <option key={field} value={field}>
                  {field}
                </option>
              ))}
            </select>
          </div>

          {/* Второй селектор столбца */}
          <div className="range-selector" style={{ marginBottom: '20px' }}>
            <select
              value={secondColumnToProcess}
              onChange={(e) => setSecondColumnToProcess(e.target.value)}
              style={{
                width: "100%",
                padding: "8px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                backgroundColor: "#59fafc",
                color: "black",
                fontSize: "14px",
              }}
            >
              <option value="">Select Second Column to Expand Ranges</option>
              {selectedFieldsOrder.map((field) => (
                <option key={field} value={field}>
                  {field}
                </option>
              ))}
            </select>
          </div>

          {/* Кнопки управления */}
          <div className="button-container">
            <button
              onClick={mergeTables}
              disabled={files.length < 2}
              style={{
                padding: "8px 16px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                backgroundColor: "#59fafc",
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
                backgroundColor: "#59fafc",
                color: "black",
                fontSize: "14px",
                cursor: "pointer",
              }}
            >
              Download
            </button>
          </div>
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
                        backgroundColor: field === 'Note' ? '#f8d7da' : 'transparent',
                        color: field === 'Note' ? '#721c24' : 'inherit',
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
                            backgroundColor: field === 'Note' && row['Note'] ? '#f8d7da' : 'transparent',
                            color: field === 'Note' && row['Note'] ? '#721c24' : 'inherit',
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
};

export default App;
