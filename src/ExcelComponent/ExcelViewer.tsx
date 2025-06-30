import React, { useState, useMemo, useCallback } from 'react';
import ReactMarkdown from 'react-markdown';
import DataGrid from 'react-data-grid';
import 'react-data-grid/lib/styles.css';
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
    let requiredCols = gridSize.cols;
    let requiredRows = gridSize.rows;
    
    if (table.length > 0) {
      // Calculate required columns (max columns in any row)
      const maxCols = Math.max(...table.map(row => row.length));
      requiredCols = Math.max(maxCols, gridSize.cols);
      
      // Calculate required rows (header + data rows + buffer)
      const dataRowCount = table.length; // includes header
      requiredRows = Math.max(dataRowCount + 5, gridSize.rows); // +5 for buffer
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

    return { columns: cols, rows: gridRows };
  }, [markdownInput, gridSize]);

  // Update editable data when rows change
  React.useEffect(() => {
    setEditableData(rows);
  }, [rows]);

  // Handle cell editing
  const handleRowsChange = useCallback((newRows: RowData[]) => {
    setEditableData(newRows);
  }, []);

  // Download as Excel file
  const downloadExcel = (): void => {
    if (editableData.length === 0) return;
    
    // Convert data to worksheet format (exclude row number column)
    const wsData: any[][] = editableData.map((row: RowData) => 
      columns.slice(1).map((col: any) => row[col.key] || '')
    );
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    
    XLSX.writeFile(wb, "excel-data.xlsx");
  };

  // Add new row
  const addRow = (): void => {
    const newRowNumber = gridSize.rows + 1;
    const newRow: RowData = { 
      id: gridSize.rows,
      rowNumber: newRowNumber
    };
    
    // Initialize all cells as empty
    for (let colIndex = 0; colIndex < gridSize.cols; colIndex++) {
      newRow[`col${colIndex}`] = '';
    }
    
    setEditableData([...editableData, newRow]);
    setGridSize(prev => ({ ...prev, rows: prev.rows + 1 }));
  };

  // Add new column
  const addColumn = (): void => {
    const newColIndex = Math.max(gridSize.cols, columns.length - 1); // -1 for row number column
    const newColKey = `col${newColIndex}`;
    
    // Update existing rows with new column
    const updatedRows: RowData[] = editableData.map((row: RowData) => ({
      ...row,
      [newColKey]: ''
    }));
    
    setEditableData(updatedRows);
    setGridSize(prev => ({ ...prev, cols: prev.cols + 1 }));
  };

  // Clear all data
  const clearGrid = (): void => {
    const clearedRows: RowData[] = editableData.map((row: RowData) => {
      const newRow: RowData = { 
        id: row.id,
        rowNumber: row.rowNumber
      };
      // Clear all data columns
      for (let colIndex = 0; colIndex < gridSize.cols; colIndex++) {
        newRow[`col${colIndex}`] = '';
      }
      return newRow;
    });
    
    setEditableData(clearedRows);
  };

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
        
        <div className="preview-section">
          <h3>Markdown Preview</h3>
          <div className="markdown-preview">
            <ReactMarkdown>{markdownInput}</ReactMarkdown>
          </div>
        </div>
      </div>

      <div className="excel-section">
        <div className="excel-header">
          <h2>Excel View</h2>
          <div className="excel-controls">
            <button onClick={addRow} className="btn btn-primary">
              Add Row
            </button>
            <button onClick={addColumn} className="btn btn-secondary">
              Add Column
            </button>
            <button onClick={clearGrid} className="btn btn-warning">
              Clear Grid
            </button>
            <button 
              onClick={downloadExcel} 
              className="btn btn-success"
            >
              Download Excel
            </button>
          </div>
        </div>
        
        <div className="grid-info">
          <span>Grid Size: {gridSize.rows} rows × {gridSize.cols} columns</span>
        </div>

        <div className="data-grid-container">
          {/* Use spread operator with type assertion to bypass type checking */}
          <DataGrid
            {...{
              columns,
              rows: editableData,
              onRowsChange: handleRowsChange,
              style: { height: '600px' },
              defaultColumnOptions: { resizable: true },
              className: "excel-grid"
            } as any}
          />
        </div>
      </div>
    </div>
  );
};


export default ExcelViewer;
