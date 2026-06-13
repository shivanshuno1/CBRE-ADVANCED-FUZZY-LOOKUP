import express from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import fs from 'fs';
import * as fuzzball from 'fuzzball';
import cors from 'cors';

const app = express();
app.use(cors());
app.use(express.json());
const upload = multer({ dest: 'uploads/' });

let workbook = null;

// ABBREVIATION MAP - Only true abbreviations, NOT semantic synonyms
const abbrMap = {
  // Legal entity types (true abbreviations)
  pvt: 'private',
  ltd: 'limited',
  corp: 'corporation',
  inc: 'incorporated',
  co: 'company',
  llc: 'limited liability company',
  llp: 'limited liability partnership',
  plc: 'public limited company',
  lllp: 'limited liability limited partnership',
  
  // Business suffixes (true abbreviations)
  ent: 'enterprise',
  grp: 'group',
  hldgs: 'holdings',
  hldg: 'holding',
  intl: 'international',
  tech: 'technologies',
  sols: 'solutions',
  svcs: 'services',
  mgmt: 'management',
  assoc: 'associates',
  assn: 'association',
  inst: 'institute',
  fdn: 'foundation',
  inds: 'industries',
  mfg: 'manufacturing',
  dist: 'distribution',
  
  // Business terms (true abbreviations)
  cap: 'capital',
  fin: 'financial',        // 'Fin' is abbreviation for Financial
  inv: 'investment',
  advisors: 'advisors',
  partners: 'partners',
  consulting: 'consulting',
  digital: 'digital',
  creative: 'creative',
  media: 'media',
  comms: 'communications',
  telecom: 'telecommunications',
  software: 'software',
  systems: 'systems',
  networks: 'networks',
  logistics: 'logistics',
  
  // Common prefixes
  natl: 'national',
  nat: 'national',
  amer: 'american',
  glob: 'global',
  eu: 'european',
  asia: 'asian',
  
  // Government/Public sector
  gov: 'government',
  auth: 'authority',
  dept: 'department',
  
  // Industry specific
  air: 'airlines',
  airways: 'airways',
  rail: 'railway',
  shipping: 'shipping',
  maritime: 'maritime',
  energy: 'energy',
  power: 'power',
  utilities: 'utilities',
  health: 'healthcare',
  pharma: 'pharmaceutical',
  biotech: 'biotechnology',
  auto: 'automotive',
  aero: 'aerospace',
  defense: 'defense',
  
  // Regional
  midwest: 'midwest',
  northeast: 'northeast',
  southeast: 'southeast',
  southwest: 'southwest',
  northwest: 'northwest',
  
  // Misspellings / common typos (fix these, but don't merge different words)
  logistix: 'logistics',     // Fix misspelling
  logistik: 'logistics',     // Fix misspelling
  sol: 'solutions',          // Common abbreviation
  soln: 'solutions',         // Common abbreviation
};

// Helper: Check if a word is a true abbreviation (shortened form)
function isTrueAbbreviation(word) {
  const abbreviations = Object.keys(abbrMap);
  return abbreviations.includes(word);
}

// Expand abbreviations ONLY (preserve semantic differences like Finance vs Financial)
function expandAbbr(s) {
  if (!s) return '';
  
  let result = String(s).toLowerCase();
  
  // Split into words and expand only true abbreviations
  const words = result.split(/\s+/);
  const expandedWords = words.map(word => {
    // Remove trailing punctuation
    const cleanWord = word.replace(/[.,!?;:]$/, '');
    
    // Only expand if it's in the abbreviation map
    if (abbrMap[cleanWord]) {
      return abbrMap[cleanWord];
    }
    return cleanWord;
  });
  
  result = expandedWords.join(' ');
  
  // Handle punctuation patterns (but preserve word distinctions)
  result = result.replace(/corp\.?/g, 'corporation');
  result = result.replace(/inc\.?/g, 'incorporated');
  result = result.replace(/ltd\.?/g, 'limited');
  result = result.replace(/co\.?/g, 'company');
  
  // IMPORTANT: Do NOT convert 'finance' to 'financial'
  // These remain as different words for similarity scoring
  
  return result;
}

// Enhanced similarity with multiple strategies
function calculateSimilarity(a, b) {
  if (!a || !b) return 0;
  
  const s1_raw = String(a).toLowerCase();
  const s2_raw = String(b).toLowerCase();
  
  // Quick exact match after basic cleaning
  if (s1_raw.replace(/[^a-z]/g, '') === s2_raw.replace(/[^a-z]/g, '')) {
    return 100;
  }
  
  // Expand abbreviations
  const s1 = expandAbbr(s1_raw);
  const s2 = expandAbbr(s2_raw);
  
  // Use multiple fuzzy matching algorithms
  const tokenSetRatio = fuzzball.token_set_ratio(s1, s2);
  const tokenSortRatio = fuzzball.token_sort_ratio(s1, s2);
  const partialRatio = fuzzball.partial_ratio(s1, s2);
  const wRatio = fuzzball.WRatio(s1, s2);
  
  // Weighted combination - token_set_ratio is best for word order independence
  let score = Math.max(tokenSetRatio, wRatio, partialRatio);
  
  // Adjust for specific edge cases
  // Finance vs Financial should NOT be 100%
  if ((s1.includes('finance') && s2.includes('financial')) ||
      (s1.includes('financial') && s2.includes('finance'))) {
    // Cap the score at 85% for related but different words
    score = Math.min(score, 85);
  }
  
  // Logistics vs Logistix (misspelling) can go higher
  if ((s1.includes('logistics') && s2.includes('logistix')) ||
      (s1.includes('logistix') && s2.includes('logistics'))) {
    // Allow up to 95% for misspellings
    score = Math.min(score, 95);
  }
  
  return Math.round(score);
}

// Build prefix index for performance
function buildIndex(data, column) {
  const index = new Map();
  for (let i = 0; i < data.length; i++) {
    const val = data[i][column] ? String(data[i][column]).toLowerCase() : '';
    if (val.length >= 3) {
      const prefix = val.substring(0, 3);
      if (!index.has(prefix)) index.set(prefix, []);
      index.get(prefix).push(i);
    }
  }
  return index;
}

// API Endpoints

app.post('/api/upload', upload.single('file'), (req, res) => {
  const wb = XLSX.readFile(req.file.path);
  workbook = wb;
  const sheets = wb.SheetNames.map(name => ({ name }));
  res.json({ success: true, uploadId: req.file.filename, sheets });
  fs.unlinkSync(req.file.path);
});

app.get('/api/columns/:uploadId/:sheetName', (req, res) => {
  if (!workbook) return res.status(404).json({ error: 'No workbook' });
  const sheet = workbook.Sheets[req.params.sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  if (!data.length) return res.json({ columns: [] });
  const headers = data[0].map(h => String(h));
  const columns = headers.map((name, idx) => ({ name, index: idx }));
  res.json({ columns });
});

app.post('/api/fuzzy-match-preview', async (req, res) => {
  const { sheetLeft, sheetRight, columnLeft, columnRight, threshold } = req.body;
  const startTime = Date.now();
  
  try {
    const leftSheet = workbook.Sheets[sheetLeft];
    const rightSheet = workbook.Sheets[sheetRight];
    let leftData = XLSX.utils.sheet_to_json(leftSheet);
    let rightData = XLSX.utils.sheet_to_json(rightSheet);
    
    console.log(`Matching ${leftData.length} left rows against ${rightData.length} right rows...`);
    
    // Cache expanded values for performance
    const rightCache = new Map();
    for (let i = 0; i < rightData.length; i++) {
      const val = rightData[i][columnRight] ? String(rightData[i][columnRight]) : '';
      rightCache.set(i, {
        raw: val,
        expanded: expandAbbr(val)
      });
    }
    
    const rightIndex = buildIndex(rightData, columnRight);
    const matched = [];
    const usedRight = new Set();
    let comparisons = 0;
    
    for (let li = 0; li < leftData.length; li++) {
      const leftRow = leftData[li];
      const leftVal = leftRow[columnLeft] ? String(leftRow[columnLeft]) : '';
      const leftExpanded = expandAbbr(leftVal);
      const leftPrefix = leftVal.length >= 3 ? leftVal.substring(0, 3).toLowerCase() : '';
      
      let candidates = rightIndex.get(leftPrefix) || [];
      if (candidates.length === 0) {
        candidates = Array.from({ length: Math.min(500, rightData.length) }, (_, i) => i);
      }
      
      let bestScore = 0;
      let bestIdx = -1;
      
      for (const ri of candidates) {
        if (usedRight.has(ri)) continue;
        
        const rightExpanded = rightCache.get(ri).expanded;
        const score = calculateSimilarity(leftExpanded, rightExpanded);
        comparisons++;
        
        if (score > bestScore && score >= threshold) {
          bestScore = score;
          bestIdx = ri;
        }
      }
      
      if (bestIdx !== -1) {
        matched.push({
          left: leftRow,
          right: rightData[bestIdx],
          similarity: bestScore
        });
        usedRight.add(bestIdx);
      }
      
      if ((li + 1) % 1000 === 0) {
        console.log(`Processed ${li + 1}/${leftData.length} rows (${comparisons} comparisons, ${Date.now() - startTime}ms)`);
      }
    }
    
    console.log(`Total comparisons: ${comparisons}, time: ${Date.now() - startTime}ms`);
    res.json({ 
      matched, 
      matchedCount: matched.length, 
      totalLeft: leftData.length, 
      totalRight: rightData.length 
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/fuzzy-compare-cross-sheet', (req, res) => {
  const { sheetLeft, sheetRight, columnLeft, columnRight, threshold } = req.body;
  
  const leftSheet = workbook.Sheets[sheetLeft];
  const rightSheet = workbook.Sheets[sheetRight];
  let leftData = XLSX.utils.sheet_to_json(leftSheet);
  let rightData = XLSX.utils.sheet_to_json(rightSheet);
  
  // Cache expanded values
  const rightCache = new Map();
  for (let i = 0; i < rightData.length; i++) {
    const val = rightData[i][columnRight] ? String(rightData[i][columnRight]) : '';
    rightCache.set(i, expandAbbr(val));
  }
  
  const matched = [];
  const unmatchedLeft = [];
  const unmatchedRight = [];
  const usedRight = new Set();
  const rightIndex = buildIndex(rightData, columnRight);
  
  // Match left to right
  for (const leftRow of leftData) {
    const leftVal = leftRow[columnLeft] ? String(leftRow[columnLeft]) : '';
    const leftExpanded = expandAbbr(leftVal);
    const leftPrefix = leftVal.length >= 3 ? leftVal.substring(0, 3).toLowerCase() : '';
    
    let candidates = rightIndex.get(leftPrefix) || [];
    if (candidates.length === 0) {
      candidates = Array.from({ length: Math.min(500, rightData.length) }, (_, i) => i);
    }
    
    let bestScore = 0;
    let bestIdx = -1;
    
    for (const ri of candidates) {
      if (usedRight.has(ri)) continue;
      const rightExpanded = rightCache.get(ri);
      const score = calculateSimilarity(leftExpanded, rightExpanded);
      
      if (score > bestScore && score >= threshold) {
        bestScore = score;
        bestIdx = ri;
      }
    }
    
    if (bestIdx !== -1) {
      matched.push({
        ...leftRow,
        ...rightData[bestIdx],
        _similarityScore: bestScore
      });
      usedRight.add(bestIdx);
    } else {
      unmatchedLeft.push({ ...leftRow, _matchStatus: 'No match found' });
    }
  }
  
  // Collect unmatched right rows
  for (let i = 0; i < rightData.length; i++) {
    if (!usedRight.has(i)) {
      unmatchedRight.push({ ...rightData[i], _matchStatus: 'No match found' });
    }
  }
  
  // Create Excel output
  const wb = XLSX.utils.book_new();
  
  if (matched.length) {
    const matchedSheet = XLSX.utils.json_to_sheet(matched);
    XLSX.utils.book_append_sheet(wb, matchedSheet, 'Matched');
  }
  
  if (unmatchedLeft.length) {
    const unmatchedLeftSheet = XLSX.utils.json_to_sheet(unmatchedLeft);
    XLSX.utils.book_append_sheet(wb, unmatchedLeftSheet, 'Unmatched_Left');
  }
  
  if (unmatchedRight.length) {
    const unmatchedRightSheet = XLSX.utils.json_to_sheet(unmatchedRight);
    XLSX.utils.book_append_sheet(wb, unmatchedRightSheet, 'Unmatched_Right');
  }
  
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  
  res.setHeader('Content-Disposition', 'attachment; filename=fuzzy_result.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));