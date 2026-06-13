import React, { useState } from 'react';
import FileUploader from './components/FileUploader';
import SheetSelector from './components/SheetSelector';
import ColumnSelector from './components/ColumnSelector';
import ColumnComparer from './components/ColumnComparer';
import axios from 'axios';

function App() {
  const [uploadId, setUploadId] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [columns, setColumns] = useState([]);
  const [selectedColumn, setSelectedColumn] = useState('');
  const [mode, setMode] = useState('abbrev');
  const [error, setError] = useState('');

  const handleUploadSuccess = (data) => {
    setUploadId(data.uploadId);
    setSheets(data.sheets);
    setError('');
  };

  const handleSheetSelect = async (sheetName) => {
    setSelectedSheet(sheetName);
    try {
      const res = await axios.get(`https://cbre-advanced-fuzzy-lookup-uyqn.onrender.com/api/columns/${uploadId}/${sheetName}`);
      setColumns(Array.isArray(res.data.columns) ? res.data.columns : []);
    } catch (err) {
      setError('Failed to load columns');
    }
  };

  return (
    <div style={{ padding: 20, fontFamily: 'Arial' }}>
      <h1>Excel Fuzzy Lookup</h1>
      {!uploadId ? (
        <FileUploader onUploadSuccess={handleUploadSuccess} onError={setError} />
      ) : (
        <>
          <div style={{ marginBottom: 20 }}>
            <button onClick={() => setMode('abbrev')} style={{ marginRight: 10, padding: '5px 10px', background: mode === 'abbrev' ? '#007bff' : '#ccc', color: '#fff', border: 'none' }}>Abbreviation Mode</button>
            <button onClick={() => setMode('compare')} style={{ padding: '5px 10px', background: mode === 'compare' ? '#007bff' : '#ccc', color: '#fff', border: 'none' }}>Cross‑Sheet Match</button>
          </div>
          {error && <div style={{ color: 'red' }}>{error}</div>}
          {mode === 'abbrev' && (
            <>
              {!selectedSheet && <SheetSelector sheets={sheets} onSelect={handleSheetSelect} />}
              {selectedSheet && <ColumnSelector columns={columns} selectedColumn={selectedColumn} onSelect={setSelectedColumn} />}
            </>
          )}
          {mode === 'compare' && <ColumnComparer uploadId={uploadId} sheets={sheets} />}
        </>
      )}
    </div>
  );
}

export default App;