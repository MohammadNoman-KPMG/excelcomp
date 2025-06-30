import React, { useState, useMemo, useCallback } from 'react';
import ReactMarkdown from 'react-markdown';
import { DataGrid } from 'react-data-grid';
import 'react-data-grid/lib/styles.css';
import * as XLSX from 'xlsx';
import './ExcelViewer.css';

const ExcelViewer = () => {
  const [markdownInput, setMarkdownInput] = useState(`# Sample Data

| Name | Age | City | Salary |
|------|-----|------|--------|
| John Doe | 28 | New York | 50000 |
| Jane Smith | 32 | Los Angeles | 60000 |
| Bob Johnson | 25 | Chicago | 45000 |
| Alice Brown | 30 | Houston | 55000 |

This is a sample markdown table that will be converted to Excel view.`);

  const [gridSize, setGridSize] = useState({ rows: 20, cols: 10 });
  const [editableData, setEditableData] = useState([]);

  // Generate Excel-like column letters (A, B, C, ..., Z, AA, AB, etc.)
  const getColumnLetter = (index) => {
    let result = '';
    while (index >= 0) {
      result = String.fromCharCode(65 + (index % 26)) + result;
      index = Math.floor(index / 26) - 1;
    }
    return result;
  };

  // Parse markdown tables to extract data
  const parseMarkdownTables = (markdown) => {
    const tableRegex = /\|(.+)\|/g;
    const lines = markdown.split('\n');
    const tables = [];
    
    let currentTable = [];
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
    
    return tables;
  };

  // Create Excel-like grid structure
  const { columns, rows } = useMemo(() => {
    // Create columns with Excel-like headers (A, B, C, etc.)
    const cols = [
      {
        key: 'rowNumber',
        name: '',
        width: 50,
        frozen: true,
        resizable: false,
        sortable: false,
        renderCell: ({ row }) => (
          <div className="row-number">{row.rowNumber}</div>
        )
      },
      ...Array.from({ length: gridSize.cols }, (_, index) => ({
        key: `col${index}`,
        name: getColumnLetter(index),
        width: 120,
        minWidth: 80,
        resizable: true,
        sortable: false,
        renderHeaderCell: ({ column }) => (
          <div className="excel-header">{column.name}</div>
        )
      }))
    ];

    // Create empty grid rows
    const gridRows = Array.from({ length: gridSize.rows }, (_, rowIndex) => {
      const rowData = { 
        id: rowIndex,
        rowNumber: rowIndex + 1
      };
      // Initialize all cells as empty
      for (let colIndex = 0; colIndex < gridSize.cols; colIndex++) {
        rowData[`col${colIndex}`] = '';
      }
      return rowData;
    });

    // Parse markdown and populate the grid with data
    const tables = parseMarkdownTables(markdownInput);
    if (tables.length > 0) {
      const table = tables[0];
      if (table.length > 0) {
        const headers = table[0];
        const dataRows = table.slice(1);
        
        // Place headers in first row if they exist
        headers.forEach((header, colIndex) => {
          if (colIndex < gridSize.cols && gridRows[0]) {
            gridRows[0][`col${colIndex}`] = header;
          }
        });
        
        // Place data in subsequent rows
        dataRows.forEach((dataRow, rowIndex) => {
          const targetRowIndex = rowIndex + 1; // Start from row 2 (index 1)
          if (targetRowIndex < gridSize.rows) {
            dataRow.forEach((cellValue, colIndex) => {
              if (colIndex < gridSize.cols) {
                gridRows[targetRowIndex][`col${colIndex}`] = cellValue || '';
              }
            });
          }
        });
      }
    }

    return { columns: cols, rows: gridRows };
  }, [markdownInput, gridSize]);

  // Update editable data when rows change
  React.useEffect(() => {
    setEditableData(rows);
  }, [rows]);

  // Handle cell editing
  const handleRowsChange = useCallback((newRows) => {
    setEditableData(newRows);
  }, []);

  // Download as Excel file
  const downloadExcel = () => {
    if (editableData.length === 0) return;
    
    // Convert data to worksheet format (exclude row number column)
    const wsData = editableData.map(row => 
      columns.slice(1).map(col => row[col.key] || '')
    );
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    
    XLSX.writeFile(wb, "excel-data.xlsx");
  };

  // Add new row
  const addRow = () => {
    const newRowNumber = gridSize.rows + 1;
    const newRow = { 
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
  const addColumn = () => {
    const newColIndex = gridSize.cols;
    const newColKey = `col${newColIndex}`;
    
    // Update existing rows with new column
    const updatedRows = editableData.map(row => ({
      ...row,
      [newColKey]: ''
    }));
    
    setEditableData(updatedRows);
    setGridSize(prev => ({ ...prev, cols: prev.cols + 1 }));
  };

  // Clear all data
  const clearGrid = () => {
    const clearedRows = editableData.map(row => {
      const newRow = { 
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

  return (
    <div className="excel-viewer">
      <div className="input-section">
        <h2>Markdown Input</h2>
        <textarea
          value={markdownInput}
          onChange={(e) => setMarkdownInput(e.target.value)}
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
          <DataGrid
            columns={columns}
            rows={editableData}
            onRowsChange={handleRowsChange}
            style={{ height: '600px' }}
            defaultColumnOptions={{
              resizable: true
            }}
            className="excel-grid"
          />
        </div>
      </div>
    </div>
  );
};

export default ExcelViewer;
