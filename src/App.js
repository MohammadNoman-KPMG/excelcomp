import React, { useState } from 'react';
import './App.css';
// import { generateExcelFromMarkdown } from './ExcelComponent/ExcelViewer.tsx';
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
