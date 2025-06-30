import React from 'react';
import './App.css';
import ExcelViewer from './ExcelComponent/index.ts';

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
