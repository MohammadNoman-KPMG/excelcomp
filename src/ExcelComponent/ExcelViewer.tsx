import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react';
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

// Interface for the single function parameters
interface ExcelGeneratorOptions {
  markdownData: string; // mandatory
  minRows?: number; // optional, default 20
  minCols?: number; // optional, default 10
  fileName?: string; // optional, default 'excel-data'
  showHeaders?: boolean; // optional, default true
}

// Props interface for external data
interface ExcelViewerProps {
  externalMarkdown?: string; // Optional external markdown data
  externalMinRows?: number; // Optional external min rows
  externalMinCols?: number; // Optional external min cols
  externalFileName?: string; // Optional external filename
  externalShowHeaders?: boolean; // Optional external show headers
  readOnly?: boolean; // Optional read-only mode
}

const ExcelViewer: React.FC<ExcelViewerProps> = ({
  externalMarkdown,
  externalMinRows,
  externalMinCols,
  externalFileName,
  externalShowHeaders,
  readOnly = false
}) => {
  // Use external data if provided, otherwise use internal state
  const [markdownInput, setMarkdownInput] = useState<string>(
    externalMarkdown || ``
  );

  const [gridSize, setGridSize] = useState<GridSize>({ rows: 20, cols: 10 });
  const [editableData, setEditableData] = useState<RowData[]>([]);
  const [selectedCells, setSelectedCells] = useState<CellSelection | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [sortConfig, setSortConfig] = useState<{column: number, direction: 'asc' | 'desc'} | null>(null);
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const [showHeaders, setShowHeaders] = useState<boolean>(externalShowHeaders ?? true);
  const [fileName, setFileName] = useState<string>(externalFileName || '');
  const [minRows, setMinRows] = useState<string>((externalMinRows || 20).toString());
  const [minCols, setMinCols] = useState<string>((externalMinCols || 10).toString());
  const gridRef = useRef<HTMLDivElement>(null);

  // Update internal state when external props change
  useEffect(() => {
    if (externalMarkdown !== undefined) {
      setMarkdownInput(externalMarkdown);
    }
  }, [externalMarkdown]);

  useEffect(() => {
    if (externalMinRows !== undefined) {
      setMinRows(externalMinRows.toString());
    }
  }, [externalMinRows]);

  useEffect(() => {
    if (externalMinCols !== undefined) {
      setMinCols(externalMinCols.toString());
    }
  }, [externalMinCols]);

  useEffect(() => {
    if (externalFileName !== undefined) {
      setFileName(externalFileName);
    }
  }, [externalFileName]);

  useEffect(() => {
    if (externalShowHeaders !== undefined) {
      setShowHeaders(externalShowHeaders);
    }
  }, [externalShowHeaders]);

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
    } catch (err) {
      console.error('Failed to copy: ', err);
    }
  };

  
// Enhanced sorting function that preserves headers
const sortDataArray = (data: RowData[], columnIndex: number, direction: 'asc' | 'desc'): RowData[] => {
  // Function to detect if first row is likely a header
  const isFirstRowHeader = (data: RowData[]): boolean => {
    if (data.length === 0) return false;
    
    const firstRow = data[0];
    const secondRow = data[1];
    
    
    // Check if first row values are different from numeric patterns in second row
    for (let col = 0; col < gridSize.cols; col++) {
      const firstVal = firstRow[`col${col}`] || '';
      const secondVal = secondRow[`col${col}`] || '';
      
      // If second row has numbers but first row doesn't, likely header
      if (secondVal && !isNaN(parseFloat(secondVal.toString().replace(/[,$%]/g, ''))) &&
          firstVal && isNaN(parseFloat(firstVal.toString().replace(/[,$%]/g, '')))) {
        return true;
      }
      
      // If first row has typical header words
      if (typeof firstVal === 'string' && 
          /^(name|title|id|date|amount|price|quantity|total|status|type|category)$/i.test(firstVal.trim())) {
        return true;
      }
    }
    
    return false;
  };
  if (data.length === 0) return data;
  
  const hasHeader = isFirstRowHeader(data);
  const headerRow = hasHeader ? data[0] : null;
  const dataRows = hasHeader ? data.slice(1) : data;
  
  const sortedDataRows = dataRows.sort((a, b) => {
    const aVal = a[`col${columnIndex}`] || '';
    const bVal = b[`col${columnIndex}`] || '';
    
    // Handle empty values - put them at the end
    if (!aVal && !bVal) return 0;
    if (!aVal) return 1;
    if (!bVal) return -1;
    
    // Try date parsing first
    const aDate = new Date(aVal);
    const bDate = new Date(bVal);
    const aIsValidDate = !isNaN(aDate.getTime()) && aVal.toString().match(/\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/);
    const bIsValidDate = !isNaN(bDate.getTime()) && bVal.toString().match(/\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/);
    
    if (aIsValidDate && bIsValidDate) {
      return direction === 'asc' ? aDate.getTime() - bDate.getTime() : bDate.getTime() - aDate.getTime();
    }
    
    // Try numeric parsing
    const aNum = parseFloat(aVal.toString().replace(/[,$%]/g, ''));
    const bNum = parseFloat(bVal.toString().replace(/[,$%]/g, ''));
    
    if (!isNaN(aNum) && !isNaN(bNum)) {
      return direction === 'asc' ? aNum - bNum : bNum - aNum;
    }
    
    // String sorting with locale support
    const aStr = aVal.toString().toLowerCase();
    const bStr = bVal.toString().toLowerCase();
    
    return direction === 'asc' 
      ? aStr.localeCompare(bStr, undefined, { numeric: true, sensitivity: 'base' })
      : bStr.localeCompare(aStr, undefined, { numeric: true, sensitivity: 'base' });
  });
  
  // Return header + sorted data, or just sorted data
  return headerRow ? [headerRow, ...sortedDataRows] : sortedDataRows;
};

// Improved sorting function
const sortData = (columnIndex: number) => {
  const direction = sortConfig?.column === columnIndex && sortConfig.direction === 'asc' ? 'desc' : 'asc';
  
  const sortedData = sortDataArray(editableData, columnIndex, direction);
  
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
    } catch (err) {
      console.error('Failed to copy: ', err);
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

  // Clipboard copy on Ctrl+C
  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
        // Only copy if focus is not in textarea (markdown input)
        const active = document.activeElement;
        if (active && (active.tagName === 'TEXTAREA' || (active as HTMLElement).isContentEditable)) return;

        e.preventDefault();
        if (selectedCells) {
          copySelectedCells();
        } else {
          copyToClipboard();
        }
      }
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [selectedCells, editableData, gridSize]);

  // Function to use the main function with current form values
    const handleGenerateExcel = useCallback(() => {
      downloadExcel();
    }, [editableData, gridSize, fileName]);


  return (
    <div className="excel-viewer">
      <div className="excel-section">
        <div className="excel-controls">
          {/* Only show controls if not in read-only mode */}
            <div className="checkbox-container">
              <label className="checkbox-label">
                <input
                  type="checkbox"
                  checked={showHeaders}
                  onChange={(e) => setShowHeaders(e.target.checked)}
                  className="header-checkbox"
                />
                Show Header
              </label>
            </div>
          <button onClick={handleGenerateExcel} className="btn btn-success">
            Download Excel
          </button>
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
export type { ExcelGeneratorOptions, ExcelViewerProps };
