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

// Обновляем интерфейс SelectedColumns
interface SelectedColumns {
  [key: string]: { // Добавляем индексную сигнатуру
    keyColumn: string;
    dataColumns: string[];
  };
  leftSheet: {
    keyColumn: string;
    dataColumns: string[];
  };
  rightSheet: {
    keyColumn: string;
    dataColumns: string[];
  };
}

interface TemplateColumn {
  id: string;
  title: string;
  isDateColumn?: boolean;
  isRequired?: boolean;
  isMultiple?: boolean;
}

const templateColumns: TemplateColumn[] = [
  { id: 'PO', title: 'PO' },
  { id: 'Line', title: 'Line' },
  { id: 'REF', title: 'Customer SAP#' },
  { id: 'PN', title: 'PN', isRequired: true },
  { id: 'Qty-by-date', title: 'Qty by date', isDateColumn: true, isMultiple: true },
  { id: 'Delivery-Requested', title: 'Delivery-Requested', isDateColumn: true },
  { id: 'Balance-to-Supply', title: 'Balance to Supply' }
];

// Обновляем интерфейс для маппинга полей с датами
interface DateColumnMapping {
  sourceSheet: string;
  sourceField: string;
  date: string;
}

// Обновляем интерфейс FieldMapping для поддержки множественных полей
interface FieldMapping {
  [templateField: string]: {
    sourceSheet: string;
    sourceField: string;
  } | DateColumnMapping[];
}

// Добавляем функцию форматирования дат
const formatDate = (value: any): string => {
  if (!value) return '';
  
  // Проверяем, является ли значение числом (Excel serial number)
  if (typeof value === 'number') {
    const date = new Date((value - 25569) * 86400 * 1000);
    return date.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: '2-digit'
    });
  }
  
  // Если это уже строка с датой, форматируем её
  const date = new Date(value);
  if (!isNaN(date.getTime())) {
    return date.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: '2-digit'
    });
  }
  
  return value;
};

// Добавляем интерфейс для данных листа
interface SheetData {
  name: string;
  data: TableRow[];
}

// Добавляем интерфейс для группы доставки
interface DeliveryGroup {
  pn: string;
  rows: TableRow[];
  balance: number;
  deficit: number;
}

// Функция расчета доставки с учетом новой формулы
const calculateDelivery = (rows: TableRow[]): TableRow[] => {
  // Группируем строки по PN
  const groupedByPN: { [key: string]: TableRow[] } = {};
  
  rows.forEach(row => {
    const pn = row.PN;
    if (!groupedByPN[pn]) {
      groupedByPN[pn] = [];
    }
    groupedByPN[pn].push(row);
  });

  const result: TableRow[] = [];
  
  // Обрабатываем каждую группу PN отдельно
  Object.entries(groupedByPN).forEach(([pn, groupRows]) => {
    // Сортируем строки по дате Delivery-Expected
    const sortedRows = groupRows.sort((a, b) => 
      new Date(a['Delivery-Expected']).getTime() - 
      new Date(b['Delivery-Expected']).getTime()
    );

    // Для каждой строки в группе
    sortedRows.forEach((row, index) => {
      const prevRow = index > 0 ? sortedRows[index - 1] : null;
      const nextRow = index < sortedRows.length - 1 ? sortedRows[index + 1] : null;
      
      // Приводим значения к числам
      const qtyByDates = Number(row['QTY by dates']) || 0;
      const balanceToSupply = Number(row['Balance to Supply']) || 0;

      // Рассчитываем среднее QTY by dates для текущего PN
      const avgQtyByDates = groupRows.reduce((sum, r) => sum + (Number(r['QTY by dates']) || 0), 0) / groupRows.length;
      
      // Рассчитываем сумму Balance to Supply для текущего PN
      const sumBalanceToSupply = groupRows.reduce((sum, r) => sum + (Number(r['Balance to Supply']) || 0), 0);

      // Проверяем условия согласно формуле
      if (nextRow && row.PN === nextRow.PN && row['Delivery-Expected'] === nextRow['Delivery-Expected']) {
        row.Delivery = '#';
      } else if (prevRow && row.PN === prevRow.PN && row['Delivery-Expected'] === prevRow['Delivery-Expected']) {
        row.Delivery = '#';
      } else if (avgQtyByDates > sumBalanceToSupply) {
        row.Delivery = `Missing PO ${avgQtyByDates - sumBalanceToSupply} Units`;
      } else if (groupRows.length === 1 && balanceToSupply >= qtyByDates) {
        row.Delivery = qtyByDates;
      } else if (groupRows.length === 2 && 
                groupRows.reduce((sum, r) => sum + (Number(r['QTY by dates']) || 0), 0) <= 
                groupRows.reduce((sum, r) => sum + (Number(r['Balance to Supply']) || 0), 0)) {
        row.Delivery = qtyByDates;
      } else {
        row.Delivery = '#';
      }

      result.push(row);
    });
  });

  // Сортируем результат по PN и дате
  return result.sort((a, b) => {
    const pnCompare = a.PN.localeCompare(b.PN);
    if (pnCompare !== 0) return pnCompare;
    return new Date(a['Delivery-Expected']).getTime() - 
           new Date(b['Delivery-Expected']).getTime();
  });
};

const App: React.FC = () => {

  const [file, setFile] = useState<File | null>(null);
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<{
    left: string;
    right: string;
  }>({ left: '', right: '' });
  const [sheetFields, setSheetFields] = useState<{
    [sheetName: string]: string[];
  }>({});
  const [tables, setTables] = useState<TableRow[][]>([]);
  const [fields, setFields] = useState<{ [key: string]: string[] }>({});
  const [selectedFields, setSelectedFields] = useState<{ [key: string]: string[] }>({});
  const [keyFields, setKeyFields] = useState<{ [key: string]: string }>({});
  const [mergedData, setMergedData] = useState<TableRow[] | null>(null);
  const [mergedPreview, setMergedPreview] = useState<TableRow[] | null>(null);
  const [selectedFieldsOrder, setSelectedFieldsOrder] = useState<string[]>([]);
  const [groupingStructure, setGroupingStructure] = useState<{ [key: string]: { [key: string]: GroupInfo } }>({});
  const [columnToProcess, setColumnToProcess] = useState<string>('');
  const [secondColumnToProcess, setSecondColumnToProcess] = useState<string>('');
  const [selectedColumns, setSelectedColumns] = useState<SelectedColumns>({
    leftSheet: {
      keyColumn: '',
      dataColumns: []
    },
    rightSheet: {
      keyColumn: '',
      dataColumns: []
    }
  });
  const [fieldMapping, setFieldMapping] = useState<FieldMapping>({});
  const [sheetData, setSheetData] = useState<{ [key: string]: TableRow[] }>({});

  useEffect(() => {
    // Logging component lifecycle
    console.log('Component lifecycle:', {
      mergedData: !!mergedData,
      selectedFieldsOrder: !!selectedFieldsOrder,
      files: file ? 1 : 0,
      tables: tables.length,
    });
  }, [mergedData, selectedFieldsOrder, file, tables]);

  useEffect(() => {
    const allSelectedFields: string[] = [];

    // Собираем все выбранные поля из обеих таблиц
    if (file) {
      const fileFields = selectedFields[file.name] || [];
      allSelectedFields.push(...fileFields);
    }

    // Обновляем selectedFieldsOrder
    setSelectedFieldsOrder(allSelectedFields);
  }, [selectedFields, file]);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = event.target.files?.[0];
    if (!uploadedFile) return;

    setFile(uploadedFile);
    
    // Добавляем индикатор загрузки
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        // Читаем только базовую информацию о листах
        const workbook = XLSX.read(data, { 
          type: 'array',
          bookSheets: true, // Читаем только список листов
          bookProps: false,
          cellFormula: false,
          cellHTML: false
        });
        
        setSheets(workbook.SheetNames);
        setSelectedSheets({ left: '', right: '' });
        setSheetFields({});
      } catch (error) {
        console.error('Error loading file:', error);
        alert('Error loading file. The file might be too large or corrupted.');
      }
    };

    reader.readAsArrayBuffer(uploadedFile);
  };

  const handleSheetSelection = async (side: 'left' | 'right', sheetName: string) => {
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(buffer), {
        type: 'array',
        sheets: [sheetName],
        cellFormula: false,
        cellHTML: false
      });
      
      const worksheet = workbook.Sheets[sheetName];
      const headerRowIndex = findHeaderRow(worksheet);
      const headers = filterAndFormatHeaders(worksheet, headerRowIndex);

      setSelectedSheets(prev => ({
        ...prev,
        [side]: sheetName
      }));
      
      setSheetFields(prev => ({
        ...prev,
        [sheetName]: headers
      }));

      // Добавляем принудительное обновление
      setTimeout(() => {
        window.dispatchEvent(new Event('resize'));
      }, 0);

      // Сохраняем данные листа
      const jsonData = XLSX.utils.sheet_to_json<TableRow>(worksheet, {
        range: { s: { r: headerRowIndex, c: 0 }, e: worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']).e : undefined }
      });
      setSheetData(prev => ({
        ...prev,
        [sheetName]: jsonData
      }));

    } catch (error) {
      console.error('Error processing sheet:', error);
      alert('Error processing sheet. Please try again.');
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
      let headers: string[] = headerRow.map(cell => String(cell || '').trim());

      const headerCount: { [key: string]: number } = {};
      headers = headers.map(header => {
        if (headerCount[header]) {
          headerCount[header] += 1;
          return `${header}-${headerCount[header]}`;
        } else {
          headerCount[header] = 1;
          return header;
        }
      });

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

  const handleFieldSelection = (sheetName: string, field: string) => {
    setSelectedFields(prev => {
      const currentFields = prev[sheetName] || [];
      const isFieldSelected = currentFields.includes(field);
      
      return {
        ...prev,
        [sheetName]: isFieldSelected 
          ? currentFields.filter(f => f !== field)
          : [...currentFields, field]
      };
    });
  };

  const handleKeyFieldSelection = (fileName: string, field: string) => {
    setKeyFields(prev => ({
      ...prev,
      [fileName]: field
    }));
  };

  const handleColumnToProcessChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    setColumnToProcess(e.target.value);
  };

  const handleSecondColumnToProcessChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    setSecondColumnToProcess(e.target.value);
  };

  const mergeTables = () => {
    if (!file || !selectedSheets.left || !selectedSheets.right) {
      console.error('Missing file or sheets');
      return;
    }

    try {
      const mergedRows: TableRow[] = [];
      const leftSheetData = sheetData[selectedSheets.left];
      const rightSheetData = sheetData[selectedSheets.right];

      if (!leftSheetData || !rightSheetData) return;

      // Создаем индекс для правой таблицы
      const rightSheetIndex: { [key: string]: any[] } = {};
      rightSheetData.forEach(row => {
        const pn = row['מקט'];
        if (pn) {
          if (!rightSheetIndex[pn]) {
            rightSheetIndex[pn] = [];
          }
          rightSheetIndex[pn].push(row);
        }
      });

      // Обрабатываем данные из левой таблицы
      leftSheetData.forEach(row => {
        const pn = row['ALE PN'];
        if (!pn) return;

        const rightRows = rightSheetIndex[pn] || [];
        const dateColumns = fieldMapping['Qty-by-date'] as DateColumnMapping[];

        dateColumns.forEach(dateMapping => {
          if (dateMapping.sourceSheet === selectedSheets.left) {
            const qtyField = dateMapping.sourceField.split(': ')[1];
            const normalizedField = normalizeDate(qtyField);
            const qtyValue = row[normalizedField];
            
            if (qtyValue) {
              rightRows.forEach(rightRow => {
                const balance = rightRow['יתרה לאספקה'];
                if (balance && balance !== 0) {
                  const deliveryRequested = rightRow ? 
                    new Date((rightRow['תאריך מובטח'] - 25569) * 86400 * 1000)
                      .toLocaleDateString('en-GB', {
                        day: '2-digit',
                        month: 'short',
                        year: '2-digit'
                      }) : '';

                  // Получаем значение REF из маппинга полей
                  const refMapping = fieldMapping['REF'] as { sourceSheet: string; sourceField: string };
                  const refValue = refMapping ? 
                    (refMapping.sourceSheet === selectedSheets.left ? 
                      row[refMapping.sourceField.split(': ')[1]] : 
                      rightRow[refMapping.sourceField.split(': ')[1]]) : '';

                  // Используем маппинг полей для создания новой строки
                  const newRow: TableRow = {
                    PO: rightRow ? rightRow[(fieldMapping['PO'] as { sourceField: string }).sourceField.split(': ')[1]] : '',
                    Line: rightRow ? rightRow[(fieldMapping['Line'] as { sourceField: string }).sourceField.split(': ')[1]] : '',
                    REF: refValue, // Добавляем значение REF
                    PN: pn,
                    [`Qty ${dateMapping.date}`]: qtyValue,
                    'QTY by dates': qtyValue,
                    'Delivery-Requested': deliveryRequested,
                    'Delivery-Expected': dateMapping.date,
                    'Balance to Supply': balance
                  };
                  
                  mergedRows.push(newRow);
                }
              });
            }
          }
        });
      });

      // Сортируем результат по PN и дате
      const sortedRows = mergedRows.sort((a, b) => {
        const pnCompare = a.PN.localeCompare(b.PN);
        if (pnCompare !== 0) return pnCompare;
        return new Date(a['Delivery-Expected']).getTime() - new Date(b['Delivery-Expected']).getTime();
      });

      // Группируем по PN и обрабатываем каждую группу
      const groupedByPN = sortedRows.reduce((groups: { [key: string]: TableRow[] }, row) => {
        const pn = row.PN;
        if (!groups[pn]) {
          groups[pn] = [];
        }
        groups[pn].push(row);
        return groups;
      }, {});

      // Обрабатываем каждую группу и собираем результаты
      const processedRows = Object.values(groupedByPN).flatMap(group => 
        calculateDelivery(group)
      );

      setMergedData(processedRows);
      setMergedPreview(processedRows.slice(0, 10));

    } catch (error) {
      console.error('Error during merge:', error);
      alert('Error occurred during merge. Please check console for details.');
    }
  };

  const downloadMergedFile = async () => {
    console.log('Starting download with data:', mergedData);

    if (!mergedData || mergedData.length === 0) {
      console.error('No data to download');
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Merged');

      // Обновляем список колонок, добавляя новые фиксированные колонки
      const visibleColumns = [
        { header: 'Supplier', key: 'Supplier', width: 15 },
        { header: 'Receiver', key: 'Receiver', width: 15 },
        { header: 'PO', key: 'PO', width: 15 },
        { header: 'Line', key: 'Line', width: 15 },
        { header: 'Type', key: 'Type', width: 15 },
        { header: 'Customer SAP#', key: 'REF', width: 15 },
        { header: 'PN', key: 'PN', width: 15 },
        { header: 'QTY by dates', key: 'QTY by dates', width: 15 },
        { header: 'Delivery-Requested', key: 'Delivery-Requested', width: 15 },
        { header: 'Delivery-Expected', key: 'Delivery-Expected', width: 15 },
        { header: 'Balance to Supply', key: 'Balance to Supply', width: 15 },
        { header: 'Delivery', key: 'Delivery', width: 15 }
      ];

      worksheet.columns = visibleColumns;

      // Добавляем данные, включая новые фиксированные значения
      const visibleData = mergedData.map(row => ({
        Supplier: 'A.L.Electronics',
        Receiver: 'Novocure',
        PO: row.PO,
        Line: row.Line,
        Type: 'Final assembly',
        REF: row['REF'],
        PN: row.PN,
        'QTY by dates': row['QTY by dates'],
        'Delivery-Requested': row['Delivery-Requested'],
        'Delivery-Expected': row['Delivery-Expected'],
        'Balance to Supply': row['Balance to Supply'],
        Delivery: row['Delivery']
      }));

      worksheet.addRows(visibleData);


      let previousPN = '';
    visibleData.forEach((row, index) => {
      const currentPN = row.PN;
      // Если PN изменился и это не первая строка
      if (currentPN !== previousPN && index > 0) {
        // Добавляем линию перед текущей строкой
        const rowNumber = index + 2; // +2 потому что индекс начинается с 0 и есть заголовок
        const currentRow = worksheet.getRow(rowNumber);
        // Применяем границу ко всем ячейкам в строке
        visibleColumns.forEach((_, colIndex) => {
          const cell = currentRow.getCell(colIndex + 1);
          cell.border = {
            ...cell.border,
            top: { style: 'medium' }
          };
        });
      }
      previousPN = currentPN;
    });

      // Стилизуем
      worksheet.getRow(1).font = { bold: true };
      worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFB1F0F0' }
      };

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      saveAs(blob, 'merged_tables.xlsx');
    } catch (error) {
      console.error('Error during file download:', error);
      alert('Error occurred during file download. Please check console for details.');
    }
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
    setFile(null);
    setTables([]);
    setFields({});
    setSelectedFields({});
    setKeyFields({});
    setMergedData(null);
    setSheets([]);
    setSelectedSheets({ left: '', right: '' });
    setSheetFields({});
    setMergedPreview(null);
    setSelectedFieldsOrder([]);
    setGroupingStructure({});
    setColumnToProcess('');
    setSecondColumnToProcess('');
    
    // Перезагружаем страницу
    window.location.reload();
  };

  // Обновляем JSX для выбора листа
  interface Sheet {
    name: string;
  }

  const renderSheetSelector = (index: number) => (
    <select
      value={index === 0 ? selectedSheets.left : selectedSheets.right}
      onChange={(e) =>
        handleSheetSelection(index === 0 ? 'left' : 'right', e.target.value)
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
      {sheets.map((sheetName, sheetIndex) => (
        <option key={`${sheetName}-${sheetIndex}`} value={sheetName}>
          {sheetName}
        </option>
      ))}
    </select>
  );

  const findHeaderRow = (worksheet: XLSX.WorkSheet): number => {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    let headerRowIndex = 0;
    let maxTextCells = 0;
    
    // Проверяем певые 50 строк
    const maxRow = Math.min(range.e.r, 50);
    
    for (let row = 0; row <= maxRow; row++) {
      let textCellsCount = 0;
      let hasText = false;
      
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddress];
        
        if (cell && typeof cell.v === 'string' && cell.v.trim() !== '') {
          textCellsCount++;
          if (/[a-zA-Z]/.test(cell.v)) {
            hasText = true;
          }
        }
      }
      
      if (textCellsCount > maxTextCells && hasText) {
        maxTextCells = textCellsCount;
        headerRowIndex = row;
      }
    }

    return headerRowIndex;
  };

  const AvailableFields: React.FC<{
    fields: string[];
    sheetName: string;
  }> = ({ fields, sheetName }) => {
    const handleDragStart = (field: string) => (e: React.DragEvent) => {
      e.dataTransfer.setData('text/plain', field);
      e.dataTransfer.setData('source-sheet', sheetName);
    };

    return (
      <div className="available-fields">
        <h3>{sheetName}</h3>
        {fields.map(field => (
          <div
            key={field}
            draggable
            className="field-item"
            onDragStart={handleDragStart(field)}
          >
            {field}
          </div>
        ))}
      </div>
    );
  };

  const normalizeDate = (dateStr: string): string => {
    return dateStr.replace(/^0(\d)/, '$1');  // Убираем ведущий ноль
  };

  // Обновляем обработчик drop
  const handleFieldDrop = (templateField: string, droppedField: string, sourceSheet: string) => {
    if (templateField === 'Qty-by-date') {
      const [col, value] = droppedField.split(': ');
      const normalizedValue = normalizeDate(value);
      
      setFieldMapping(prev => {
        const existing = (prev[templateField] as DateColumnMapping[]) || [];
        
        // Проверяем дубликаты
        if (existing.some(mapping => mapping.date === normalizedValue)) {
          return prev;
        }

        const newMapping = {
          sourceSheet,
          sourceField: droppedField,
          date: normalizedValue
        };

        // Сохраняем сортировку
        return {
          ...prev,
          [templateField]: [...existing, newMapping].sort((a, b) => 
            new Date(a.date).getTime() - new Date(b.date).getTime()
          )
        };
      });
    } else {
      setFieldMapping(prev => ({
        ...prev,
        [templateField]: { sourceSheet, sourceField: droppedField }
      }));
    }
  };

  // Обновляем компонент TemplateTable для корректной типизаци маппинга
  const TemplateTable: React.FC<{
    columns: TemplateColumn[];
    fieldMapping: FieldMapping;
    onDrop: (columnId: string, field: string, sourceSheet: string) => void;
  }> = ({ columns, fieldMapping, onDrop }) => {
    const [dragOverColumn, setDragOverColumn] = useState<string | null>(null);

    const handleDragOver = (columnId: string) => (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setDragOverColumn(columnId);
    };

    const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setDragOverColumn(null);
    };

    const handleDrop = (columnId: string) => (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setDragOverColumn(null);
      const field = e.dataTransfer.getData('text/plain');
      const sourceSheet = e.dataTransfer.getData('source-sheet');
      onDrop(columnId, field, sourceSheet);
    };

    return (
      <div className="template-table">
        <div className="template-header">
          {columns.map(column => (
            <div
              key={column.id}
              className={`template-column 
                ${column.isDateColumn ? 'date-column' : ''} 
                ${column.isRequired ? 'required' : ''}
                ${fieldMapping[column.id] ? 'mapped' : ''}
                ${dragOverColumn === column.id ? 'dragover' : ''}`}
              onDragOver={handleDragOver(column.id)}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop(column.id)}
            >
              <div className="column-title">{column.title}</div>
              {fieldMapping[column.id] && !Array.isArray(fieldMapping[column.id]) && (
                <div className="mapped-field">
                  {(fieldMapping[column.id] as { sourceSheet: string; sourceField: string }).sourceField}
                </div>
              )}
              {column.id === 'Qty-by-date' && Array.isArray(fieldMapping[column.id]) && (
                <div className="mapped-fields-list">
                  {(fieldMapping[column.id] as DateColumnMapping[]).map((mapping, index) => (
                    <div key={index} className="mapped-field">
                      {mapping.sourceField}
                    </div>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      </div>
    );
  };

  // Обновляем стили
  const styles = `
    body {
      background-color: #1a1a1a;
      color: #ffffff;
      margin: 0;
      padding: 0;
      overflow-y: auto;
      position: relative;
      min-height: 100vh;
    }

    .app-container {
      background-color: #2d2d2d;
      min-height: 100vh;
      padding: 20px;
      display: flex;
      flex-direction: column;
      align-items: center;
      margin-top: 100px;
    }

    .app-title {
      color: #59fafc;
      font-size: 28px;
      margin: 20px 0 40px;
      text-transform: capitalize;
      letter-spacing: 1px;
    }

    .sheets-layout {
      display: flex;
      justify-content: space-between;
      width: 100%;
      gap: 20px;
      margin: 20px 0;
    }

    .sheet-panel {
      flex: 1;
      background-color: #383838;
      padding: 15px;
      border-radius: 8px;
    }

    .sheet-selector select {
      width: 100%;
      padding: 10px;
      background-color: #2d2d2d;
      color: #ffffff;
      border: 1px solid #4a4a4a;
      border-radius: 4px;
    }

    .fields-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
      gap: 10px;
      max-height: 400px;
      overflow-y: auto;
    }

    .field-item {
      background-color: #4a4a4a;
      color: #ffffff;
      padding: 8px;
      border: 1px solid #666;
      border-radius: 4px;
      cursor: move;
    }

    .template-container {
      width: 100%;
      background-color: #383838;
      padding: 20px;
      border-radius: 8px;
      margin-top: 30px;
      overflow-x: hidden;
    }

    .template-table {
      width: 100%;
      margin-top: 20px;
    }

    .template-header {
      display: grid;
      grid-template-columns: repeat(7, 1fr);
      align-items: center;
      gap: 15px;
      padding: 10px;
      width: 100%;
    }

    .template-column {
      background-color: #2d2d2d;
      border: 2px dashed #4a4a4a;
      border-radius: 6px;
      padding: 15px;
      min-height: 100px;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      transition: all 0.3s ease;
      width: 80%;
    }

    .template-column.dragover {
      background-color: #2d4a6d;
      border-color: #59fafc;
    }

    .template-column.mapped {
      background-color: #2d4a3e;
      border-color: #4caf50;
      border-style: solid;
    }

    .template-column.required {
      border-color: #ff4444;
    }

    .column-title {
      font-weight: bold;
      color: #59fafc;
      margin-bottom: 10px;
      text-align: center;
    }

    .mapped-field {
      background-color: #383838;
      color: #ffffff;
      padding: 8px;
      border-radius: 4px;
      width: 90%;
      text-align: center;
      font-size: 0.9em;
      word-break: break-word;
    }

    .actions-container {
      display: flex;
      justify-content: center;
      gap: 20px;
      margin-top: 30px;
    }

    .merge-button, .download-button {
      padding: 12px 24px;
      border: none;
      border-radius: 6px;
      font-weight: bold;
      cursor: pointer;
      transition: all 0.3s ease;
    }

    .merge-button {
      background-color: #2196f3;
      color: white;
    }

    .download-button {
      background-color: #4caf50;
      color: white;
    }

    .merge-button:disabled, .download-button:disabled {
      background-color: #4a4a4a;
      color: #666;
      cursor: not-allowed;
    }

    .merge-button:hover:not(:disabled), .download-button:hover:not(:disabled) {
      filter: brightness(1.2);
      transform: translateY(-2px);
    }

    .sheet-panel {
      position: relative;
      z-index: 2;
    }

    .sheet-selector {
      position: relative;
      z-index: 1000;
    }

    .sheet-selector select {
      width: 100%;
      position: relative;
      z-index: 1000;
    }

    .header-container {
      display: flex;
      justify-content: space-between;
      align-items: center;
      width: 100%;
      position: relative;
      z-index: 1;
    }

    .app-title {
      margin: 0;
    }
  `;

  const getSheetHeaders = (worksheet: XLSX.WorkSheet, headerRowIndex: number): string[] => {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    const headers: { col: string; value: string }[] = [];
    
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
      const cell = worksheet[cellAddress];
      
      if (cell && cell.v) {
        let value = cell.v.toString();
        
        // Проверяем, является ли значение числом (Excel serial number)
        if (typeof cell.v === 'number' && cell.v > 1) {
          try {
            const date = new Date((cell.v - 25569) * 86400 * 1000);
            if (!isNaN(date.getTime())) {
              value = date.toLocaleDateString('en-GB', {
                day: '2-digit',
                month: 'short',
                year: '2-digit'
              });
            }
          } catch (e) {
            console.error('Error formatting date:', e);
          }
        }
        
        headers.push({
          col: XLSX.utils.encode_col(col),
          value: value
        });
      }
    }
    
    return headers.map(h => `${h.col}: ${h.value}`);
  };

  const filterAndFormatHeaders = (worksheet: XLSX.WorkSheet, headerRowIndex: number): string[] => {
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    const headers: { col: string; value: string }[] = [];
    
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
      const cell = worksheet[cellAddress];
      
      if (cell && cell.v) {
        let value = cell.v.toString();
        let shouldShow = true;
        
        // Проверяем, является ли значение датой
        if (typeof cell.v === 'number' && cell.v > 1) {
          try {
            const date = new Date((cell.v - 25569) * 86400 * 1000);
            if (!isNaN(date.getTime())) {
              value = date.toLocaleDateString('en-GB', {
                day: '2-digit',
                month: 'short',
                year: '2-digit'
              });
              
              // Скрываем даты до авгуса 2024 года
              if (date < new Date('2024-08-01')) {
                shouldShow = false;
              }
            }
          } catch (e) {
            console.error('Error formatting date:', e);
          }
        }
        
        if (shouldShow) {
          headers.push({
            col: XLSX.utils.encode_col(col),
            value: value
          });
        }
      }
    }
    
    return headers.map(h => `${h.col}: ${h.value}`);
  };

  // Добавляем useEffect для вставки стилей
  useEffect(() => {
    const styleElement = document.createElement('style');
    styleElement.textContent = styles;
    document.head.appendChild(styleElement);

    return () => {
      document.head.removeChild(styleElement);
    };
  }, []);

  const darkThemeStyles = `
    body {
      background-color: #1a1a1a;
      color: #ffffff;
    }

    .app-container {
      background-color: #2d2d2d;
    }

    .sheet-panel {
      background-color: #383838;
    }

    .sheet-selector select {
      background-color: #4a4a4a;
      color: #ffffff;
      border-color: #666;
    }

    .fields-container {
      background-color: #383838;
    }

    .field-item {
      background-color: #4a4a4a;
      color: #ffffff;
      border-color: #666;
    }

    .field-item:hover {
      background-color: #5a5a5a;
    }

    .template-container {
      background-color: #383838;
    }

    .template-column {
      background-color: #4a4a4a;
      border-color: #666;
      color: #ffffff;
    }

    .template-column.mapped {
      background-color: #2d4a3e;
      border-color: #4caf50;
    }

    .template-column.dragover {
      background-color: #2d4a6d;
      border-color: #2196f3;
    }

    .mapped-field {
      background-color: #2d4a3e;
      color: #98c99a;
    }

    .merge-button, .download-button {
      background-color: #4a4a4a;
      color: #ffffff;
      border-color: #666;
    }

    .merge-button:hover, .download-button:hover {
      background-color: #5a5a5a;
    }

    .merge-button:disabled, .download-button:disabled {
      background-color: #333;
      color: #666;
    }
  `;

  const MergedPreview: React.FC<{ data: TableRow[] }> = ({ data }) => {
    if (!data || data.length === 0) return null;

    const displayColumns = [
      'PO',
      'Line',
      'REF',
      'PN',
      'QTY by dates',
      'Delivery-Requested',
      'Delivery-Expected',
      'Balance to Supply'
    ];

    return (
      <div className="preview-container">
        <h3>Preview (first 10 rows)</h3>
        <div className="preview-table">
          <div className="preview-header">
            {displayColumns.map(col => (
              <div key={col} className="preview-cell header-cell">
                {col}
              </div>
            ))}
          </div>
          <div className="preview-body">
            {data.map((row, rowIndex) => (
              <div key={rowIndex} className="preview-row">
                {displayColumns.map(col => (
                  <div key={`${rowIndex}-${col}`} className="preview-cell">
                    {row[col]}
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  };

  const previewStyles = `
    .preview-container {
      margin-top: 20px;
      padding: 20px;
      background-color: #383838;
      border-radius: 8px;
      overflow-x: auto;
    }

    .preview-table {
      width: 100%;
      min-width: 800px;
      border-collapse: collapse;
    }

    .preview-header {
      display: grid;
      grid-template-columns: repeat(8, 1fr);
      background-color: #2d2d2d;
      padding: 10px 0;
      position: sticky;
      top: 0;
      z-index: 1;
    }

    .preview-row {
      display: grid;
      grid-template-columns: repeat(8, 1fr);
      border-bottom: 1px solid #4a4a4a;
    }

    .preview-cell {
      padding: 8px;
      text-align: left;
      min-width: 120px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .header-cell {
      font-weight: bold;
      color: #59fafc;
    }
  `;

  // Добавляем стили в useEffect
  useEffect(() => {
    const styleElement = document.createElement('style');
    styleElement.textContent = styles + previewStyles; // Добавляем стили превью
    document.head.appendChild(styleElement);

    return () => {
      document.head.removeChild(styleElement);
    };
  }, []);

  return (
    <div className="app-container">
      <div className="header-container">
        <h1 className="app-title">Manager Excel Report</h1>
        <button onClick={handleReset}>RESET</button>
      </div>

      <div className="file-input-container">
        <Input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
          className="file-input"
        />
      </div>

      {file && (
        <div className="sheets-layout">
          {/* Левая панель */}
          <div className="sheet-panel">
            <div className="sheet-selector">
              <h3>First Sheet</h3>
              <select
                value={selectedSheets.left}
                onChange={(e) => handleSheetSelection('left', e.target.value)}
              >
                <option value="">Select sheet</option>
                {sheets.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
            {selectedSheets.left && (
              <div className="fields-container">
                <div className="fields-grid">
                  {sheetFields[selectedSheets.left]?.map(field => (
                    <div
                      key={field}
                      draggable
                      className="field-item"
                      onDragStart={(e) => {
                        e.dataTransfer.setData('text/plain', field);
                        e.dataTransfer.setData('source-sheet', selectedSheets.left);
                      }}
                    >
                      {field}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Правая панель */}
          <div className="sheet-panel">
            <div className="sheet-selector">
              <h3>Second Sheet</h3>
              <select
                value={selectedSheets.right}
                onChange={(e) => handleSheetSelection('right', e.target.value)}
              >
                <option value="">Select sheet</option>
                {sheets.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>
            </div>
            {selectedSheets.right && (
              <div className="fields-container">
                <div className="fields-grid">
                  {sheetFields[selectedSheets.right]?.map(field => (
                    <div
                      key={field}
                      draggable
                      className="field-item"
                      onDragStart={(e) => {
                        e.dataTransfer.setData('text/plain', field);
                        e.dataTransfer.setData('source-sheet', selectedSheets.right);
                      }}
                    >
                      {field}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {selectedSheets.left && selectedSheets.right && (
        <div className="template-container">
          <h3>Target Table Template</h3>
          <TemplateTable 
            columns={templateColumns}
            fieldMapping={fieldMapping}
            onDrop={handleFieldDrop}
          />
          <div className="actions-container">
            <button 
              onClick={mergeTables}
              disabled={!fieldMapping['PN']}
              className="merge-button"
            >
              Merge Tables
            </button>
            <button
              onClick={downloadMergedFile}
              disabled={!mergedData || mergedData.length === 0}
              className="download-button"
            >
              Download Result
            </button>
          </div>
        </div>
      )}

      {mergedPreview && (
        <MergedPreview data={mergedPreview} />
      )}
    </div>
  );
};

export default App;
