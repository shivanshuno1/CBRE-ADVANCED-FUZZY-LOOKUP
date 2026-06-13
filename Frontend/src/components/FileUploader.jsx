import React, { useState } from 'react';
import axios from 'axios';

const FileUploader = ({ onUploadSuccess, onError }) => {
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const handleUpload = async () => {
    if (!file) return onError('Select a file');
    const fd = new FormData();
    fd.append('file', file);
    setUploading(true);
    try {
      const res = await axios.post('https://cbre-advanced-fuzzy-lookup-uyqn.onrender.com/api/upload', fd);
      onUploadSuccess(res.data);
    } catch (err) {
      onError(err.message);
    } finally {
      setUploading(false);
    }
  };
  return (
    <div style={{ border: '1px solid #ccc', padding: 20 }}>
      <h2>Upload Excel</h2>
      <input type="file" accept=".xlsx,.xls" onChange={e => setFile(e.target.files[0])} />
      <button onClick={handleUpload} disabled={uploading} style={{ marginLeft: 10 }}>Upload</button>
    </div>
  );
};
export default FileUploader;