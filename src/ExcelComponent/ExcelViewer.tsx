import React, { useState, useMemo, useCallback, useRef } from 'react';
import ReactMarkdown from 'react-markdown';
import * as XLSX from 'xlsx';
import './ExcelViewer.css';

// Minimal type definitions
interface GridSize {
  rows: number;
  cols: number;
}

interface RowData {
  [key: string]: any;
}

interface CellSelection {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

const ExcelViewer: React.FC = () => {
  const [markdownInput, setMarkdownInput] = useState<string>(`# Sample Data

| Name | Age | City | Salary |
|------|-----|------|--------|
| John Doe | 28 | New York | 50000 |
| Jane Smith | 32 | Los Angeles | 60000 |
| Bob Johnson | 25 | Chicago | 45000 |
| Alice Brown | 30 | Houston | 55000 |

This is a sample markdown table that will be converted to Excel view.`);

  const [gridSize, setGridSize] = useState<GridSize>({ rows: 20, cols: 10 });
  const [editableData, setEditableData] = useState<RowData[]>([]);
  const [selectedCells, setSelectedCells] = useState<CellSelection | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [sortConfig, setSortConfig] = useState<{column: number, direction: 'asc' | 'desc'} | null>(null);
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const [showHeaders, setShowHeaders] = useState<boolean>(true);
  const [fileName, setFileName] = useState<string>('');
  const [minRows, setMinRows] = useState<string>('20');
  const [minCols, setMinCols] = useState<string>('10');
  const gridRef = useRef<HTMLDivElement>(null);

  // Generate Excel-like column letters (A, B, C, ..., Z, AA, AB, etc.)
  const getColumnLetter = (index: number): string => {
    let result = '';
    while (index >= 0) {
      result = String.fromCharCode(65 + (index % 26)) + result;
      index = Math.floor(index / 26) - 1;
    }
    return result;
  };

  // Parse markdown tables to extract data
  const parseMarkdownTables = (markdown: string): string[][] => {
    const lines = markdown.split('\n');
    const tables: string[][][] = [];
    
    let currentTable: string[][] = [];
    let isInTable = false;
    
    for (let line of lines) {
      if (line.includes('|') && line.trim() !== '') {
        if (line.includes('---')) {
          continue;
        }
        
        const cells = line.split('|')
          .map(cell => cell.trim())
          .filter(cell => cell !== '');
        
        if (cells.length > 0) {
          currentTable.push(cells);
          isInTable = true;
        }
      } else if (isInTable && currentTable.length > 0) {
        tables.push(currentTable);
        currentTable = [];
        isInTable = false;
      }
    }
    
    if (currentTable.length > 0) {
      tables.push(currentTable);
    }
    
    return tables.length > 0 ? tables[0] : [];
  };

  // Create Excel-like grid structure
  const { columns, rows } = useMemo(() => {
    // Parse markdown first to determine required dimensions
    const table = parseMarkdownTables(markdownInput);
    
    // Get minimum values from inputs or use defaults
    const defaultMinRows = parseInt(minRows) || 20;
    const defaultMinCols = parseInt(minCols) || 10;
    
    let requiredCols = defaultMinCols;
    let requiredRows = defaultMinRows;
    
    if (table.length > 0) {
      // Calculate required columns (max columns in any row)
      const maxCols = Math.max(...table.map(row => row.length));
      requiredCols = Math.max(maxCols, defaultMinCols);
      
      // Calculate required rows (header + data rows + buffer)
      const dataRowCount = table.length; // includes header
      requiredRows = Math.max(dataRowCount + 5, defaultMinRows); // +5 for buffer
    }

    // Update column widths array to match required columns
    if (columnWidths.length !== requiredCols) {
      const newWidths = Array.from({ length: requiredCols }, (_, index) => 
        columnWidths[index] || 120 // Use existing width or default to 120
      );
      setColumnWidths(newWidths);
    }

    // Create columns with Excel-like headers (A, B, C, etc.)
    const cols: any[] = [
      {
        key: 'rowNumber',
        name: '',
        width: 50,
        frozen: true,
        resizable: false,
        sortable: false,
        renderCell: (props: any) => (
          <div className="row-number">{props.row.rowNumber}</div>
        )
      },
      ...Array.from({ length: requiredCols }, (_, index) => ({
        key: `col${index}`,
        name: getColumnLetter(index),
        width: 120,
        minWidth: 80,
        resizable: true,
        sortable: false,
        renderHeaderCell: (props: any) => (
          <div className="excel-header">{props.column.name}</div>
        )
      }))
    ];

    // Create empty grid rows
    const gridRows: RowData[] = Array.from({ length: requiredRows }, (_, rowIndex) => {
      const rowData: RowData = { 
        id: rowIndex,
        rowNumber: rowIndex + 1
      };
      // Initialize all cells as empty
      for (let colIndex = 0; colIndex < requiredCols; colIndex++) {
        rowData[`col${colIndex}`] = '';
      }
      return rowData;
    });

    // Parse markdown and populate the grid with data
    if (table.length > 0) {
      const headers = table[0];
      const dataRows = table.slice(1);
      
      // Place headers in first row if they exist
      headers.forEach((header: string, colIndex: number) => {
        if (colIndex < requiredCols && gridRows[0]) {
          gridRows[0][`col${colIndex}`] = header;
        }
      });
      
      // Place data in subsequent rows
      dataRows.forEach((dataRow: string[], rowIndex: number) => {
        const targetRowIndex = rowIndex + 1; // Start from row 2 (index 1)
        if (targetRowIndex < requiredRows) {
          dataRow.forEach((cellValue: string, colIndex: number) => {
            if (colIndex < requiredCols) {
              gridRows[targetRowIndex][`col${colIndex}`] = cellValue || '';
            }
          });
        }
      });
    }

    // Update grid size state if it has changed
    if (requiredCols !== gridSize.cols || requiredRows !== gridSize.rows) {
      // Use setTimeout to avoid state update during render
      setTimeout(() => {
        setGridSize({ rows: requiredRows, cols: requiredCols });
      }, 0);
    }

    return { columns: [], rows: gridRows };
  }, [markdownInput, gridSize, minRows, minCols, columnWidths.length]);

  // Update editable data when rows change
  React.useEffect(() => {
    setEditableData(rows);
  }, [rows]);

  // Handle cell value change
  const handleCellChange = (rowIndex: number, colKey: string, value: string) => {
    const updatedData = [...editableData];
    updatedData[rowIndex] = { ...updatedData[rowIndex], [colKey]: value };
    setEditableData(updatedData);
  };

  // Download as Excel file
  const downloadExcel = (): void => {
    if (editableData.length === 0) return;
    
    // Create headers array
    const headers = Array.from({ length: gridSize.cols }, (_, index) => getColumnLetter(index));
    
    // Convert data to worksheet format
    const wsData: any[][] = [
      headers, // Column headers
      ...editableData.map((row: RowData) => 
        headers.map((_, colIndex) => row[`col${colIndex}`] || '')
      )
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    
    // Use custom filename if provided, otherwise use default
    const finalFileName = fileName.trim() ? 
      (fileName.trim().endsWith('.xlsx') ? fileName.trim() : `${fileName.trim()}.xlsx`) : 
      'excel-data.xlsx';
    
    XLSX.writeFile(wb, finalFileName);
  };

  // Copy grid data to clipboard in Excel-compatible format
  const copyToClipboard = async (): Promise<void> => {
    if (editableData.length === 0) return;
    
    // Create headers array
    const headers = Array.from({ length: gridSize.cols }, (_, index) => getColumnLetter(index));
    
    // Convert data to tab-separated format (Excel compatible)
    const clipboardData = [
      headers.join('\t'), // Column headers
      ...editableData.map((row: RowData) => 
        headers.map((_, colIndex) => row[`col${colIndex}`] || '').join('\t')
      )
    ].join('\n');
    
    try {
      await navigator.clipboard.writeText(clipboardData);
      alert('Data copied to clipboard! You can now paste it into Excel.');
    } catch (err) {
      console.error('Failed to copy: ', err);
      // Fallback for older browsers
      fallbackCopyTextToClipboard(clipboardData);
    }
  };

  // Fallback copy method for older browsers
  const fallbackCopyTextToClipboard = (text: string): void => {
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.top = '0';
    textArea.style.left = '0';
    textArea.style.position = 'fixed';
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    
    try {
      const successful = document.execCommand('copy');
      if (successful) {
        alert('Data copied to clipboard! You can now paste it into Excel.');
      } else {
        alert('Copy failed. Please try downloading the Excel file instead.');
      }
    } catch (err) {
      console.error('Fallback: Oops, unable to copy', err);
      alert('Copy failed. Please try downloading the Excel file instead.');
    }
    
    document.body.removeChild(textArea);
  };

  // Advanced sorting function
  const sortData = (columnIndex: number) => {
    const direction = sortConfig?.column === columnIndex && sortConfig.direction === 'asc' ? 'desc' : 'asc';
    
    const sortedData = [...editableData].sort((a, b) => {
      const aVal = a[`col${columnIndex}`] || '';
      const bVal = b[`col${columnIndex}`] || '';
      
      // Try to convert to numbers for numeric sorting
      const aNum = parseFloat(aVal);
      const bNum = parseFloat(bVal);
      
      if (!isNaN(aNum) && !isNaN(bNum)) {
        return direction === 'asc' ? aNum - bNum : bNum - aNum;
      }
      
      // String sorting
      return direction === 'asc' 
        ? aVal.toString().localeCompare(bVal.toString())
        : bVal.toString().localeCompare(aVal.toString());
    });
    
    setEditableData(sortedData);
    setSortConfig({ column: columnIndex, direction });
  };

  // Cell selection handlers
  const handleCellMouseDown = (rowIndex: number, colIndex: number) => {
    setSelectedCells({
      startRow: rowIndex,
      startCol: colIndex,
      endRow: rowIndex,
      endCol: colIndex
    });
    setIsDragging(true);
  };

  const handleCellMouseEnter = (rowIndex: number, colIndex: number) => {
    if (isDragging && selectedCells) {
      setSelectedCells({
        ...selectedCells,
        endRow: rowIndex,
        endCol: colIndex
      });
    }
  };

  const handleMouseUp = () => {
    setIsDragging(false);
  };

  // Check if cell is selected
  const isCellSelected = (rowIndex: number, colIndex: number): boolean => {
    if (!selectedCells) return false;
    
    const minRow = Math.min(selectedCells.startRow, selectedCells.endRow);
    const maxRow = Math.max(selectedCells.startRow, selectedCells.endRow);
    const minCol = Math.min(selectedCells.startCol, selectedCells.endCol);
    const maxCol = Math.max(selectedCells.startCol, selectedCells.endCol);
    
    return rowIndex >= minRow && rowIndex <= maxRow && colIndex >= minCol && colIndex <= maxCol;
  };

  // Copy selected cells
  const copySelectedCells = async (): Promise<void> => {
    if (!selectedCells || editableData.length === 0) return;
    
    const minRow = Math.min(selectedCells.startRow, selectedCells.endRow);
    const maxRow = Math.max(selectedCells.startRow, selectedCells.endRow);
    const minCol = Math.min(selectedCells.startCol, selectedCells.endCol);
    const maxCol = Math.max(selectedCells.startCol, selectedCells.endCol);
    
    const selectedData: string[] = [];
    
    for (let row = minRow; row <= maxRow; row++) {
      const rowData: string[] = [];
      for (let col = minCol; col <= maxCol; col++) {
        rowData.push(editableData[row]?.[`col${col}`] || '');
      }
      selectedData.push(rowData.join('\t'));
    }
    
    const clipboardData = selectedData.join('\n');
    
    try {
      await navigator.clipboard.writeText(clipboardData);
      alert(`Selected ${maxRow - minRow + 1} rows and ${maxCol - minCol + 1} columns copied to clipboard!`);
    } catch (err) {
      console.error('Failed to copy: ', err);
      fallbackCopyTextToClipboard(clipboardData);
    }
  };

  // Column resizing
  const handleColumnResize = (colIndex: number, newWidth: number) => {
    if (colIndex >= 0 && colIndex < columnWidths.length) {
      const newWidths = [...columnWidths];
      newWidths[colIndex] = Math.max(80, newWidth);
      setColumnWidths(newWidths);
    }
  };

  // Add new row
//   const addRow = (): void => {
//     const newRowNumber = gridSize.rows + 1;
//     const newRow: RowData = { 
//       id: gridSize.rows,
//       rowNumber: newRowNumber
//     };
    
//     // Initialize all cells as empty
//     for (let colIndex = 0; colIndex < gridSize.cols; colIndex++) {
//       newRow[`col${colIndex}`] = '';
//     }
    
//     setEditableData([...editableData, newRow]);
//     setGridSize(prev => ({ ...prev, rows: prev.rows + 1 }));
//   };

  // Add new column
//   const addColumn = (): void => {
//     const newColIndex = Math.max(gridSize.cols, columns.length - 1); // -1 for row number column
//     const newColKey = `col${newColIndex}`;
    
//     // Update existing rows with new column
//     const updatedRows: RowData[] = editableData.map((row: RowData) => ({
//       ...row,
//       [newColKey]: ''
//     }));
    
//     setEditableData(updatedRows);
//     setGridSize(prev => ({ ...prev, cols: prev.cols + 1 }));
//   };

  // Clear all data
//   const clearGrid = (): void => {
//     const clearedRows: RowData[] = editableData.map((row: RowData) => {
//       const newRow: RowData = { 
//         id: row.id,
//         rowNumber: row.rowNumber
//       };
//       // Clear all data columns
//       for (let colIndex = 0; colIndex < gridSize.cols; colIndex++) {
//         newRow[`col${colIndex}`] = '';
//       }
//       return newRow;
//     });
    
//     setEditableData(clearedRows);
//   };

  const handleMarkdownChange = (e: React.ChangeEvent<HTMLTextAreaElement>): void => {
    setMarkdownInput(e.target.value);
  };

  return (
    <div className="excel-viewer">
      <div className="input-section">
        <h2>Markdown Input</h2>
        <textarea
          value={markdownInput}
          onChange={handleMarkdownChange}
          placeholder="Enter your markdown with tables here..."
          className="markdown-input"
        />
        
        <div className="grid-size-inputs">
          <div className="size-input-container">
            <label className="size-label">
              Min Rows:
              <input
                type="number"
                value={minRows}
                onChange={(e) => setMinRows(e.target.value)}
                placeholder="20"
                min="1"
                max="1000"
                className="size-input"
              />
            </label>
          </div>
          <div className="size-input-container">
            <label className="size-label">
              Min Columns:
              <input
                type="number"
                value={minCols}
                onChange={(e) => setMinCols(e.target.value)}
                placeholder="10"
                min="1"
                max="100"
                className="size-input"
              />
            </label>
          </div>
        </div>
      </div>

      <div className="excel-section">
        <div className="excel-header">
          <h2>Advanced Excel View</h2>
          <div className="excel-controls">
            <div className="filename-container">
              <label className="filename-label">
                File Name:
                <input
                  type="text"
                  value={fileName}
                  onChange={(e) => setFileName(e.target.value)}
                  placeholder="excel-data (optional)"
                  className="filename-input"
                />
              </label>
            </div>
            <div className="checkbox-container">
              <label className="checkbox-label">
                <input
                  type="checkbox"
                  checked={showHeaders}
                  onChange={(e) => setShowHeaders(e.target.checked)}
                  className="header-checkbox"
                />
                Show Column & Row Headers
              </label>
            </div>
            <button onClick={copySelectedCells} className="btn btn-info" disabled={!selectedCells}>
              Copy Selected
            </button>
            <button onClick={copyToClipboard} className="btn btn-info">
              Copy All
            </button>
            <button onClick={downloadExcel} className="btn btn-success">
              Download Excel
            </button>
          </div>
        </div>
        
        <div className="grid-info">
          <span>Grid Size: {gridSize.rows} rows × {gridSize.cols} columns</span>
          <span className="min-size-info">
            | Min Size: {parseInt(minRows) || 20} rows × {parseInt(minCols) || 10} columns
          </span>
          {selectedCells && (
            <span className="selection-info">
              | Selected: {Math.abs(selectedCells.endRow - selectedCells.startRow) + 1} rows × {Math.abs(selectedCells.endCol - selectedCells.startCol) + 1} columns
            </span>
          )}
        </div>

        <div 
          className="advanced-excel-grid"
          ref={gridRef}
          onMouseUp={handleMouseUp}
          onMouseLeave={handleMouseUp}
        >
          {/* Column Headers - Conditionally Rendered */}
          {showHeaders && (
            <div className="excel-row header-row">
              <div className="excel-cell row-header"></div>
              {Array.from({ length: gridSize.cols }, (_, colIndex) => (
                <div 
                  key={colIndex} 
                  className="excel-cell column-header sortable"
                  style={{ width: columnWidths[colIndex] }}
                  onClick={() => sortData(colIndex)}
                >
                  <span>{getColumnLetter(colIndex)}</span>
                  {sortConfig?.column === colIndex && (
                    <span className="sort-indicator">
                      {sortConfig.direction === 'asc' ? ' ↑' : ' ↓'}
                    </span>
                  )}
                  <div 
                    className="column-resizer"
                    onMouseDown={(e) => {
                      e.preventDefault();
                      const startX = e.clientX;
                      const startWidth = columnWidths[colIndex];
                      
                      const handleMouseMove = (e: MouseEvent) => {
                        const newWidth = startWidth + (e.clientX - startX);
                        handleColumnResize(colIndex, newWidth);
                      };
                      
                      const handleMouseUp = () => {
                        document.removeEventListener('mousemove', handleMouseMove);
                        document.removeEventListener('mouseup', handleMouseUp);
                      };
                      
                      document.addEventListener('mousemove', handleMouseMove);
                      document.addEventListener('mouseup', handleMouseUp);
                    }}
                  />
                </div>
              ))}
            </div>
          )}

          {/* Data Rows */}
          {editableData.map((row, rowIndex) => (
            <div key={rowIndex} className="excel-row">
              {/* Row Number - Conditionally Rendered */}
              {showHeaders && (
                <div className="excel-cell row-header">{rowIndex + 1}</div>
              )}
              
              {/* Data Cells */}
              {Array.from({ length: gridSize.cols }, (_, colIndex) => (
                <div 
                  key={colIndex} 
                  className={`excel-cell data-cell advanced-cell ${
                    isCellSelected(rowIndex, colIndex) ? 'selected' : ''
                  } ${!showHeaders ? 'no-headers' : ''}`}
                  style={{ 
                    width: columnWidths[colIndex] || 120,
                    minWidth: '80px' // Ensure minimum width for grid lines
                  }}
                  onMouseDown={() => handleCellMouseDown(rowIndex, colIndex)}
                  onMouseEnter={() => handleCellMouseEnter(rowIndex, colIndex)}
                >
                  <span className="cell-content">
                    {row[`col${colIndex}`] || ''}
                  </span>
                </div>
              ))}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default ExcelViewer;
