import React, { useState, useEffect } from 'react';
import axios from 'axios';

const ColumnComparer = ({ uploadId, sheets }) => {
  const [leftSheet, setLeftSheet] = useState('');
  const [rightSheet, setRightSheet] = useState('');
  const [leftCols, setLeftCols] = useState([]);
  const [rightCols, setRightCols] = useState([]);
  const [matchLeft, setMatchLeft] = useState('');
  const [matchRight, setMatchRight] = useState('');
  const [threshold, setThreshold] = useState(60);
  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState(null);
  const [error, setError] = useState('');
  const [abortController, setAbortController] = useState(null);

  const API_BASE = 'https://cbre-advanced-fuzzy-lookup-uyqn.onrender.com';

  // Fetch left columns
  useEffect(() => {
    if (!leftSheet || !uploadId) {
      setLeftCols([]);
      return;
    }
    axios.get(`${API_BASE}/api/columns/${uploadId}/${encodeURIComponent(leftSheet)}`)
      .then(res => {
        const columns = res.data && Array.isArray(res.data.columns) ? res.data.columns : [];
        setLeftCols(columns);
      })
      .catch(err => {
        console.error('Error fetching left columns:', err);
        setLeftCols([]);
        setError(`Failed to load columns from left sheet: ${err.message}`);
      });
  }, [leftSheet, uploadId]);

  // Fetch right columns
  useEffect(() => {
    if (!rightSheet || !uploadId) {
      setRightCols([]);
      return;
    }
    axios.get(`${API_BASE}/api/columns/${uploadId}/${encodeURIComponent(rightSheet)}`)
      .then(res => {
        const columns = res.data && Array.isArray(res.data.columns) ? res.data.columns : [];
        setRightCols(columns);
      })
      .catch(err => {
        console.error('Error fetching right columns:', err);
        setRightCols([]);
        setError(`Failed to load columns from right sheet: ${err.message}`);
      });
  }, [rightSheet, uploadId]);

  // Run fuzzy match with abort support
  const runMatch = async () => {
    if (!leftSheet || !rightSheet || !matchLeft || !matchRight) {
      setError('Please select both sheets and match columns');
      return;
    }

    // Cancel any previous ongoing request
    if (abortController) {
      abortController.abort();
    }

    const controller = new AbortController();
    setAbortController(controller);
    setLoading(true);
    setError('');
    setResults(null);

    try {
      const res = await axios.post(
        `${API_BASE}/api/fuzzy-match-preview`,
        {
          sheetLeft: leftSheet,
          sheetRight: rightSheet,
          columnLeft: matchLeft,
          columnRight: matchRight,
          threshold
        },
        { signal: controller.signal }
      );
      
      // Debug logging
      console.log('=== API Response ===');
      console.log('Full response:', res.data);
      console.log('Matched count:', res.data.matchedCount);
      console.log('Matched array length:', res.data.matched?.length);
      console.log('All matches:', res.data.matched);
      
      // Set results directly from response
      setResults({
        matched: res.data.matched || [],
        matchedCount: res.data.matchedCount || 0,
        totalLeft: res.data.totalLeft || 0,
        totalRight: res.data.totalRight || 0
      });
      
    } catch (err) {
      if (err.name === 'AbortError') {
        setError('Matching was cancelled.');
      } else {
        console.error('Match error:', err);
        setError(err.response?.data?.error || err.message);
      }
    } finally {
      setLoading(false);
      setAbortController(null);
    }
  };

  const cancelMatch = () => {
    if (abortController) {
      abortController.abort();
      setLoading(false);
    }
  };

  const downloadExcel = async () => {
    try {
      const res = await axios.post(
        `${API_BASE}/api/fuzzy-compare-cross-sheet`,
        {
          sheetLeft: leftSheet,
          sheetRight: rightSheet,
          columnLeft: matchLeft,
          columnRight: matchRight,
          threshold
        },
        { responseType: 'blob' }
      );
      const url = URL.createObjectURL(res.data);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'fuzzy_result.xlsx';
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Download error:', err);
      setError('Download failed: ' + err.message);
    }
  };

  // Safe render for column options
  const renderColumnOptions = (cols, type) => {
    if (!cols || !Array.isArray(cols) || cols.length === 0) {
      return <option value="" disabled>No columns available</option>;
    }
    return cols.map((col) => {
      const colName = col && col.name ? col.name : col;
      if (!colName) return null;
      return (
        <option key={`${type}-${colName}`} value={colName}>
          {colName}
        </option>
      );
    }).filter(Boolean);
  };

  const safeSheets = Array.isArray(sheets) ? sheets : [];

  return (
    <div style={{ border: '1px solid #ccc', padding: 20, marginTop: 20 }}>
      <h2>Cross-Sheet Fuzzy Match</h2>

      {error && (
        <div style={{ color: 'red', marginBottom: 15, padding: 10, backgroundColor: '#ffeeee', borderRadius: 4 }}>
          {error}
        </div>
      )}

      <div style={{ display: 'flex', gap: 20, flexWrap: 'wrap' }}>
        {/* Left Sheet Section */}
        <div style={{ flex: 1, minWidth: 250 }}>
          <div><strong>Left Sheet</strong></div>
          <select
            value={leftSheet}
            onChange={e => setLeftSheet(e.target.value)}
            style={{ width: '100%', padding: 8, marginTop: 5 }}
          >
            <option value="">Select a sheet</option>
            {safeSheets.map(s => <option key={s.name} value={s.name}>{s.name}</option>)}
          </select>

          {leftSheet && (
            <div style={{ marginTop: 15 }}>
              <div><strong>Match Column (Left)</strong></div>
              <select
                value={matchLeft}
                onChange={e => setMatchLeft(e.target.value)}
                style={{ width: '100%', padding: 8, marginTop: 5 }}
                disabled={!leftCols || leftCols.length === 0}
              >
                <option value="">Select column</option>
                {renderColumnOptions(leftCols, 'left')}
              </select>
              {leftCols && leftCols.length === 0 && leftSheet && (
                <div style={{ fontSize: 12, color: '#666', marginTop: 5 }}>
                  Loading columns or no columns found...
                </div>
              )}
            </div>
          )}
        </div>

        {/* Right Sheet Section */}
        <div style={{ flex: 1, minWidth: 250 }}>
          <div><strong>Right Sheet</strong></div>
          <select
            value={rightSheet}
            onChange={e => setRightSheet(e.target.value)}
            style={{ width: '100%', padding: 8, marginTop: 5 }}
          >
            <option value="">Select a sheet</option>
            {safeSheets.map(s => <option key={s.name} value={s.name}>{s.name}</option>)}
          </select>

          {rightSheet && (
            <div style={{ marginTop: 15 }}>
              <div><strong>Match Column (Right)</strong></div>
              <select
                value={matchRight}
                onChange={e => setMatchRight(e.target.value)}
                style={{ width: '100%', padding: 8, marginTop: 5 }}
                disabled={!rightCols || rightCols.length === 0}
              >
                <option value="">Select column</option>
                {renderColumnOptions(rightCols, 'right')}
              </select>
              {rightCols && rightCols.length === 0 && rightSheet && (
                <div style={{ fontSize: 12, color: '#666', marginTop: 5 }}>
                  Loading columns or no columns found...
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Threshold Slider */}
      <div style={{ marginTop: 20 }}>
        <div><strong>Similarity Threshold: {threshold}%</strong></div>
        <input
          type="range"
          min="0"
          max="100"
          value={threshold}
          onChange={e => setThreshold(parseInt(e.target.value))}
          style={{ width: '100%', marginTop: 5 }}
        />
        <div style={{ fontSize: 12, color: '#666', marginTop: 5 }}>
          Higher = stricter matching, Lower = more matches
        </div>
      </div>

      {/* Action Buttons */}
      <div style={{ marginTop: 20, display: 'flex', gap: 10 }}>
        {!loading ? (
          <button
            onClick={runMatch}
            disabled={!leftSheet || !rightSheet || !matchLeft || !matchRight}
            style={{
              padding: '8px 16px',
              backgroundColor: '#007bff',
              color: 'white',
              border: 'none',
              borderRadius: 4,
              cursor: (!leftSheet || !rightSheet || !matchLeft || !matchRight) ? 'not-allowed' : 'pointer',
              opacity: (!leftSheet || !rightSheet || !matchLeft || !matchRight) ? 0.6 : 1
            }}
          >
            Run Fuzzy Match
          </button>
        ) : (
          <button
            onClick={cancelMatch}
            style={{
              padding: '8px 16px',
              backgroundColor: '#dc3545',
              color: 'white',
              border: 'none',
              borderRadius: 4,
              cursor: 'pointer'
            }}
          >
            Cancel Matching
          </button>
        )}

        {results && results.matched && results.matched.length > 0 && (
          <button
            onClick={downloadExcel}
            style={{
              padding: '8px 16px',
              backgroundColor: '#28a745',
              color: 'white',
              border: 'none',
              borderRadius: 4,
              cursor: 'pointer'
            }}
          >
            Download Excel
          </button>
        )}
      </div>

      {/* Results Display */}
      {results && results.matched && (
        <div style={{ marginTop: 20 }}>
          <h3>
            Results: {results.matchedCount || results.matched.length} matches out of {results.totalLeft || 0} left rows
            {results.matched.length === 0 && " — Try lowering the similarity threshold"}
          </h3>

          {results.matched.length > 0 && (
            <div style={{ overflowX: 'auto' }}>
              <table border="1" cellPadding="8" style={{ borderCollapse: 'collapse', width: '100%', marginTop: 10 }}>
                <thead style={{ backgroundColor: '#f0f0f0' }}>
                  <tr>
                    <th>Left {matchLeft || 'Column'}</th>
                    <th>Right {matchRight || 'Column'}</th>
                    <th>Similarity</th>
                  </tr>
                </thead>
                <tbody>
                  {/* Display ALL matches - removed .slice(0,20) */}
                  {results.matched.map((m, idx) => (
                    <tr key={idx}>
                      <td style={{ border: '1px solid #ddd', padding: 8 }}>
                        {m.left && m.left[matchLeft] ? String(m.left[matchLeft]) : '—'}
                      </td>
                      <td style={{ border: '1px solid #ddd', padding: 8 }}>
                        {m.right && m.right[matchRight] ? String(m.right[matchRight]) : '—'}
                      </td>
                      <td style={{ border: '1px solid #ddd', padding: 8, textAlign: 'center' }}>
                        {m.similarity}%
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default ColumnComparer;