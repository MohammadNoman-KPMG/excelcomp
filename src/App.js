import React from 'react';
import './App.css';
// Import ExcelViewer component directly from the tsx file
import ExcelViewer from './ExcelComponent/ExcelViewer.tsx';

function App() {
  return (
    <div className="App">
      <header className="App-header-custom">
        <h1>Markdown to Excel Converter</h1>
        <p>Convert markdown tables to interactive Excel-like view</p>
      </header>
      <main>
        <ExcelViewer />
      </main>
    </div>
  );
}

export default App;
