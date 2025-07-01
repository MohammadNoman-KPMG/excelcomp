import React, { useState } from 'react';
import './App.css';
import { generateExcelFromMarkdown } from './ExcelComponent/ExcelViewer.tsx';
import ExcelViewer from './ExcelComponent/ExcelViewer.tsx';

function App() {
  const [markdownData, setMarkdownData] = useState(`# Sample Data

| Name | Age | City | Salary |
|------|-----|------|--------|
| John Doe | 28 | New York | 50000 |
| Jane Smith | 32 | Los Angeles | 60000 |
| Bob Johnson | 25 | Chicago | 45000 |
| Alice Brown | 30 | Houston | 55000 |`);

  const [minRows, setMinRows] = useState('20');
  const [minCols, setMinCols] = useState('10');
  const [fileName, setFileName] = useState('');
  const [showHeaders, setShowHeaders] = useState(true);

  // Handle quick generation with minimal parameters
  const handleQuickGenerate = () => {
    generateExcelFromMarkdown({
      markdownData: markdownData
    });
  };

  // Handle generation with all custom parameters
  const handleCustomGenerate = () => {
    generateExcelFromMarkdown({
      markdownData: markdownData,
      minRows: parseInt(minRows) || 20,
      minCols: parseInt(minCols) || 10,
      fileName: fileName,
      showHeaders: showHeaders
    });
  };

  return (
    <div className="App">
      <header className="App-header-custom">
        <h1>Markdown to Excel Converter</h1>
        <p>Convert markdown tables to interactive Excel-like view</p>
      </header>
      <main>
        <div className="app-container">
          <div className="input-section">
            <h2>Markdown Input</h2>
            <textarea
              value={markdownData}
              onChange={(e) => setMarkdownData(e.target.value)}
              placeholder="Enter your markdown with tables here..."
              className="markdown-input"
              rows="15"
            />
          </div>

          {/* <div className="controls-section">
            <h3>Export Settings</h3>
            <div className="controls-grid">
              <div className="control-group">
                <label>
                  Min Rows:
                  <input
                    type="number"
                    value={minRows}
                    onChange={(e) => setMinRows(e.target.value)}
                    placeholder="20"
                    min="1"
                    max="1000"
                    className="number-input"
                  />
                </label>
              </div>

              <div className="control-group">
                <label>
                  Min Columns:
                  <input
                    type="number"
                    value={minCols}
                    onChange={(e) => setMinCols(e.target.value)}
                    placeholder="10"
                    min="1"
                    max="100"
                    className="number-input"
                  />
                </label>
              </div>

              <div className="control-group">
                <label>
                  File Name:
                  <input
                    type="text"
                    value={fileName}
                    onChange={(e) => setFileName(e.target.value)}
                    placeholder="excel-data (optional)"
                    className="text-input"
                  />
                </label>
              </div>

              <div className="control-group checkbox-group">
                <label>
                  <input
                    type="checkbox"
                    checked={showHeaders}
                    onChange={(e) => setShowHeaders(e.target.checked)}
                  />
                  Show Column Headers
                </label>
              </div>
            </div>

            <div className="buttons-section">
              <button 
                onClick={handleQuickGenerate} 
                className="btn btn-primary"
              >
                Quick Export (Default Settings)
              </button>
              <button 
                onClick={handleCustomGenerate} 
                className="btn btn-success"
              >
                Export with Custom Settings
              </button>
            </div>
          </div> */}

          <div className="preview-section">
            <h3>Preview & Download</h3>
            <div className="excel-preview">
              <ExcelViewer
                externalMarkdown={markdownData}
                externalMinRows={parseInt(minRows) || 20}
                externalMinCols={parseInt(minCols) || 10}
                externalFileName={fileName}
                externalShowHeaders={showHeaders}
                readOnly={true}
              />
            </div>
          </div>

          {/* How to use the function  */}
          {/* <div className="info-section">
            <h3>How to Use</h3>
            <ul>
              <li>Enter markdown text with tables in the textarea above</li>
              <li>Adjust export settings if needed</li>
              <li>Preview the table in the section below</li>
              <li>Click "Quick Export" for default settings or "Custom Export" for your settings</li>
              <li>You can also download directly from the preview using the "Download Excel" button</li>
            </ul>
          </div> */}
        </div>
      </main>
    </div>
  );
}

export default App;
