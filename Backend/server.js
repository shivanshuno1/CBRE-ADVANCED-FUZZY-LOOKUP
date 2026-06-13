import express from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import * as fuzzball from 'fuzzball';
import cors from 'cors';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);




const app = express();
app.use(cors());
app.use(express.json());

// Ensure uploads directory exists
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const upload = multer({ dest: uploadDir });

let workbook = null;
let currentUploadId = null;

// ============ ADD THIS HELPER FUNCTION ============
function escapeRegex(string) {
  if (!string) return '';
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ============ UNIVERSAL TEXT NORMALIZATION ============

// Generic abbreviation expansion (can be extended by user)
const universalAbbrMap = {
  // Legal
  'pvt': 'private', 'ltd': 'limited', 'corp': 'corporation', 'inc': 'incorporated',
  'co': 'company', 'llc': 'limited liability company', 'llp': 'limited liability partnership',
  'plc': 'public limited company', 'pc': 'professional corporation',
  
  // Business
  'ent': 'enterprise', 'grp': 'group', 'hldgs': 'holdings', 'intl': 'international',
  'tech': 'technologies', 'sols': 'solutions', 'svcs': 'services', 'mgmt': 'management',
  'assoc': 'associates', 'inst': 'institute', 'fdn': 'foundation', 'inds': 'industries',
  'mfg': 'manufacturing', 'dist': 'distribution', 'whs': 'wholesale', 'ret': 'retail',
  
  // Financial
  'cap': 'capital', 'fin': 'financial', 'inv': 'investment', 'adv': 'advisors',
  'part': 'partners', 'cons': 'consulting',
  
  // Geographic
  'natl': 'national', 'nat': 'national', 'amer': 'american', 'glob': 'global',
  'east': 'eastern', 'west': 'western', 'north': 'northern', 'south': 'southern',
  'central': 'central', 'metro': 'metropolitan',
  
  // Common misspellings
  'logistix': 'logistics', 'logistik': 'logistics',
  'soloutions': 'solutions', 'solutons': 'solutions',
  'incorperated': 'incorporated', 'corperation': 'corporation',
  'finacial': 'financial', 'teh': 'the'
};

// Clean text for comparison
function cleanText(text) {
  if (!text || typeof text !== 'string') return '';
  
  // Convert to string and lowercase
  let cleaned = String(text).toLowerCase();
  
  // Remove special characters but keep letters, numbers, and spaces
  cleaned = cleaned.replace(/[^a-z0-9\s]/g, ' ');
  
  // Remove extra spaces
  cleaned = cleaned.replace(/\s+/g, ' ').trim();
  
  return cleaned;
}

// Expand abbreviations in text
function expandAbbreviations(text, customAbbrMap = {}) {
  if (!text) return '';
  
  const combinedMap = { ...universalAbbrMap, ...customAbbrMap };
  let result = ' ' + cleanText(text) + ' ';
  
  // Sort by length (longest first) to avoid partial matches
  const sortedAbbrs = Object.keys(combinedMap).sort((a, b) => b.length - a.length);
  
  for (const abbr of sortedAbbrs) {
    const regex = new RegExp(`\\b${abbr.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'gi');
    result = result.replace(regex, ` ${combinedMap[abbr]} `);
  }
  
  return result.trim();
}

// Calculate similarity score (0-100)
function calculateSimilarity(str1, str2, customAbbrMap = {}) {
  if (!str1 || !str2) return 0;
  
  const s1_raw = String(str1);
  const s2_raw = String(str2);
  
  // Quick exact match after basic cleaning
  if (cleanText(s1_raw) === cleanText(s2_raw)) {
    return 100;
  }
  
  // Expand abbreviations
  const s1 = expandAbbreviations(s1_raw, customAbbrMap);
  const s2 = expandAbbreviations(s2_raw, customAbbrMap);
  
  // Use multiple fuzzy matching algorithms
  const tokenSetRatio = fuzzball.token_set_ratio(s1, s2);
  const tokenSortRatio = fuzzball.token_sort_ratio(s1, s2);
  const partialRatio = fuzzball.partial_ratio(s1, s2);
  const wRatio = fuzzball.WRatio(s1, s2);
  
  // Weighted combination - token_set_ratio is best for word order independence
  let score = Math.max(tokenSetRatio, wRatio, partialRatio, tokenSortRatio);
  
  // Boost score for high partial matches (important for addresses)
  if (partialRatio > 85 && score < partialRatio) {
    score = partialRatio;
  }
  
  return Math.min(100, Math.max(0, Math.round(score)));
}

// Build prefix index for performance
function buildIndex(data, column) {
  const index = new Map();
  for (let i = 0; i < data.length; i++) {
    const val = data[i][column] ? String(data[i][column]).toLowerCase() : '';
    const cleaned = cleanText(val);
    if (cleaned.length >= 2) {
      const prefix = cleaned.substring(0, 2);
      if (!index.has(prefix)) index.set(prefix, []);
      index.get(prefix).push(i);
    }
  }
  return index;
}

// ============ API ENDPOINTS ============

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    message: 'Universal Fuzzy Lookup API',
    version: '1.0.0',
    status: 'running',
    endpoints: {
      upload: 'POST /api/upload',
      sheets: 'GET /api/sheets/:uploadId',
      columns: 'GET /api/columns/:uploadId/:sheetName',
      match: 'POST /api/fuzzy-match',
      download: 'POST /api/download-results'
    }
  });
});

// Health check
app.get('/health', (req, res) => {
  res.json({ status: 'healthy', timestamp: new Date().toISOString() });
});

// Upload Excel file
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    const wb = XLSX.readFile(req.file.path);
    workbook = wb;
    currentUploadId = req.file.filename;
    
    // Get sheet info with row counts
    const sheets = wb.SheetNames.map(name => {
      const sheet = wb.Sheets[name];
      const data = XLSX.utils.sheet_to_json(sheet);
      return { name, rowCount: data.length };
    });
    
    res.json({ 
      success: true, 
      uploadId: req.file.filename, 
      sheets,
      totalSheets: sheets.length
    });
    
    // Clean up file
    fs.unlinkSync(req.file.path);
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Get all sheets
app.get('/api/sheets/:uploadId', (req, res) => {
  if (!workbook) {
    return res.status(404).json({ error: 'No workbook uploaded' });
  }
  
  const sheets = workbook.SheetNames.map(name => ({ name }));
  res.json({ sheets });
});

// Get columns from a sheet
app.get('/api/columns/:uploadId/:sheetName', (req, res) => {
  if (!workbook) {
    return res.status(404).json({ error: 'No workbook uploaded' });
  }
  
  try {
    const sheetName = req.params.sheetName;
    const sheet = workbook.Sheets[sheetName];
    
    if (!sheet) {
      return res.status(404).json({ error: `Sheet '${sheetName}' not found` });
    }
    
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (!data || !data.length) {
      return res.json({ columns: [] });
    }
    
    const headers = data[0].map(h => String(h));
    const columns = headers.map((name, idx) => ({ 
      name, 
      index: idx,
      sampleData: data.slice(1, 4).map(row => row[idx]).filter(v => v)
    }));
    
    res.json({ columns });
  } catch (err) {
    console.error('Columns error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Perform fuzzy matching
app.post('/api/fuzzy-match', async (req, res) => {
  const { 
    sheetLeft, 
    sheetRight, 
    columnLeft, 
    columnRight, 
    threshold = 60,
    maxResultsPerLeft = 1,  // Max matches per left row
    caseSensitive = false,
    customAbbreviations = {}
  } = req.body;
  
  if (!workbook) {
    return res.status(404).json({ error: 'No workbook uploaded' });
  }
  
  const startTime = Date.now();
  
  try {
    const leftSheet = workbook.Sheets[sheetLeft];
    const rightSheet = workbook.Sheets[sheetRight];
    
    if (!leftSheet || !rightSheet) {
      return res.status(404).json({ error: 'One or both sheets not found' });
    }
    
    let leftData = XLSX.utils.sheet_to_json(leftSheet);
    let rightData = XLSX.utils.sheet_to_json(rightSheet);
    
    console.log(`Matching ${leftData.length} left rows against ${rightData.length} right rows...`);
    console.log(`Left column: ${columnLeft}, Right column: ${columnRight}`);
    console.log(`Threshold: ${threshold}, Max results per left: ${maxResultsPerLeft}`);
    
    // Build index for performance
    const rightIndex = buildIndex(rightData, columnRight);
    
    const results = [];
    let totalComparisons = 0;
    let matchedCount = 0;
    
    for (let li = 0; li < leftData.length; li++) {
      const leftRow = leftData[li];
      const leftVal = leftRow[columnLeft] ? String(leftRow[columnLeft]) : '';
      const leftClean = cleanText(leftVal);
      const leftPrefix = leftClean.length >= 2 ? leftClean.substring(0, 2) : '';
      
      // Get candidates using prefix index
      let candidates = rightIndex.get(leftPrefix) || [];
      
      // Fallback: if no candidates, sample random rows for large datasets
      if (candidates.length === 0 && rightData.length > 0) {
        const sampleSize = Math.min(500, rightData.length);
        candidates = Array.from({ length: sampleSize }, () => Math.floor(Math.random() * rightData.length));
      }
      
      // Calculate scores for all candidates
      const scores = [];
      for (const ri of candidates) {
        const rightVal = rightData[ri][columnRight] ? String(rightData[ri][columnRight]) : '';
        const score = calculateSimilarity(leftVal, rightVal, customAbbreviations);
        totalComparisons++;
        
        if (score >= threshold) {
          scores.push({ index: ri, score, rightRow: rightData[ri] });
        }
      }
      
      // Sort by score (highest first) and take top N
      scores.sort((a, b) => b.score - a.score);
      const topMatches = scores.slice(0, maxResultsPerLeft);
      
      if (topMatches.length > 0) {
        matchedCount++;
        results.push({
          leftRowIndex: li,
          leftRow,
          matches: topMatches.map(m => ({
            rightRowIndex: m.index,
            rightRow: m.rightRow,
            similarityScore: m.score
          }))
        });
      }
      
      // Progress logging
      if ((li + 1) % 1000 === 0) {
        console.log(`Processed ${li + 1}/${leftData.length} rows (${totalComparisons} comparisons, ${Date.now() - startTime}ms)`);
      }
    }
    
    const elapsedTime = Date.now() - startTime;
    console.log(`Match complete: ${matchedCount}/${leftData.length} left rows matched`);
    console.log(`Total comparisons: ${totalComparisons}, time: ${elapsedTime}ms`);
    
    res.json({
      success: true,
      matchedCount,
      totalLeftRows: leftData.length,
      totalRightRows: rightData.length,
      threshold,
      maxResultsPerLeft,
      comparisons: totalComparisons,
      elapsedTimeMs: elapsedTime,
      results
    });
    
  } catch (err) {
    console.error('Fuzzy match error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Download results as Excel
app.post('/api/download-results', (req, res) => {
  const { results, leftData, rightData, columnLeft, columnRight } = req.body;
  
  try {
    const wb = XLSX.utils.book_new();
    
    // Sheet 1: Matched results
    const matchedRows = [];
    for (const result of results) {
      for (const match of result.matches) {
        matchedRows.push({
          ['Left_' + columnLeft]: result.leftRow[columnLeft],
          ...result.leftRow,
          ['Right_' + columnRight]: match.rightRow[columnRight],
          ...match.rightRow,
          'Similarity_Score': match.similarityScore
        });
      }
    }
    
    if (matchedRows.length) {
      const matchedSheet = XLSX.utils.json_to_sheet(matchedRows);
      XLSX.utils.book_append_sheet(wb, matchedSheet, 'Matched_Results');
    }
    
    // Sheet 2: Summary
    const summary = [
      { Metric: 'Total Left Rows', Value: results.reduce((sum, r) => sum + 1, 0) },
      { Metric: 'Total Matches Found', Value: matchedRows.length },
      { Metric: 'Average Similarity Score', Value: matchedRows.length ? 
        (matchedRows.reduce((sum, r) => sum + r.Similarity_Score, 0) / matchedRows.length).toFixed(2) : 0 }
    ];
    
    const summarySheet = XLSX.utils.json_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');
    
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    
    res.setHeader('Content-Disposition', 'attachment; filename=fuzzy_match_results.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
    
  } catch (err) {
    console.error('Download error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Start server
const PORT = process.env.PORT || 5000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`========================================`);
  console.log(`Universal Fuzzy Lookup API`);
  console.log(`Running on port ${PORT}`);
  console.log(`Health check: http://localhost:${PORT}/health`);
  console.log(`========================================`);
});