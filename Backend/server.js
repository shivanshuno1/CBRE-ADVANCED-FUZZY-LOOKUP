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

function expandAbbr(s) {
  const map = {
    // Legal entity types (most common)
    pvt: 'private',
    ltd: 'limited',
    corp: 'corporation',
    inc: 'incorporated',
    co: 'company',
    llc: 'limited liability company',
    llp: 'limited liability partnership',
    plc: 'public limited company',
    lllp: 'limited liability limited partnership',
    
    // Business suffixes
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
    
    // Common prefixes
    natl: 'national',
    nat: 'national',
    amer: 'american',
    intl: 'international',
    glob: 'global',
    eu: 'european',
    asia: 'asian',
    
    // Business terms
    cap: 'capital',
    fin: 'financial',
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
    supply: 'supply chain',
    
    // Government/Public sector
    gov: 'government',
    corp: 'corporation',
    auth: 'authority',
    dept: 'department',
    
    // Brand/Industry specific
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
    northwest: 'northwest'
  };
  
  return s.toLowerCase().replace(/\b(\w+)\b/g, w => map[w] || w);
}

function similarity(a, b) {
  if (!a || !b) return 0;
  const s1 = expandAbbr(String(a));
  const s2 = expandAbbr(String(b));
  if (s1 === s2) return 100;
  return fuzzball.token_set_ratio(s1, s2);
}

// Helper: build prefix index (first 3 letters)
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
    
    // Build prefix index for right data
    const rightIndex = buildIndex(rightData, columnRight);
    
    const matched = [];
    const usedRight = new Set();
    let comparisons = 0;
    
    for (let li = 0; li < leftData.length; li++) {
      const leftRow = leftData[li];
      const leftVal = leftRow[columnLeft] ? String(leftRow[columnLeft]) : '';
      const leftPrefix = leftVal.length >= 3 ? leftVal.substring(0, 3).toLowerCase() : '';
      
      let candidates = [];
      if (rightIndex.has(leftPrefix)) {
        candidates = rightIndex.get(leftPrefix);
      } else {
        // fallback: if no prefix match, try all (but limit to first 1000 for speed)
        candidates = Array.from({ length: Math.min(1000, rightData.length) }, (_, i) => i);
      }
      
      let bestScore = 0;
      let bestIdx = -1;
      for (const ri of candidates) {
        if (usedRight.has(ri)) continue;
        const rightVal = rightData[ri][columnRight] ? String(rightData[ri][columnRight]) : '';
        const score = similarity(leftVal, rightVal);
        comparisons++;
        if (score > bestScore && score >= threshold) {
          bestScore = score;
          bestIdx = ri;
        }
      }
      if (bestIdx !== -1) {
        matched.push({ left: leftRow, right: rightData[bestIdx], similarity: bestScore });
        usedRight.add(bestIdx);
      }
      
      // Log every 1000 rows
      if ((li + 1) % 1000 === 0) {
        console.log(`Processed ${li + 1}/${leftData.length} rows (${comparisons} comparisons, ${Date.now() - startTime}ms)`);
      }
    }
    
    console.log(`Total comparisons: ${comparisons}, time: ${Date.now() - startTime}ms`);
    res.json({ matched, matchedCount: matched.length, totalLeft: leftData.length, totalRight: rightData.length });
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
  const matched = [], unmatchedLeft = [], unmatchedRight = [];
  const usedRight = new Set();
  const rightIndex = buildIndex(rightData, columnRight);
  
  for (const leftRow of leftData) {
    const leftVal = leftRow[columnLeft] ? String(leftRow[columnLeft]) : '';
    const leftPrefix = leftVal.length >= 3 ? leftVal.substring(0, 3).toLowerCase() : '';
    let candidates = rightIndex.get(leftPrefix) || [];
    if (candidates.length === 0) candidates = Array.from({ length: Math.min(1000, rightData.length) }, (_, i) => i);
    
    let bestScore = 0, bestIdx = -1;
    for (const ri of candidates) {
      if (usedRight.has(ri)) continue;
      const rightVal = rightData[ri][columnRight] ? String(rightData[ri][columnRight]) : '';
      const score = similarity(leftVal, rightVal);
      if (score > bestScore && score >= threshold) {
        bestScore = score;
        bestIdx = ri;
      }
    }
    if (bestIdx !== -1) {
      matched.push({ ...leftRow, ...rightData[bestIdx], _similarityScore: bestScore });
      usedRight.add(bestIdx);
    } else {
      unmatchedLeft.push({ ...leftRow, _matchStatus: 'No match' });
    }
  }
  for (let i = 0; i < rightData.length; i++) {
    if (!usedRight.has(i)) unmatchedRight.push({ ...rightData[i], _matchStatus: 'No match' });
  }
  const wb = XLSX.utils.book_new();
  if (matched.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(matched), 'Matched');
  if (unmatchedLeft.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(unmatchedLeft), 'Unmatched_Left');
  if (unmatchedRight.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(unmatchedRight), 'Unmatched_Right');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename=fuzzy_result.xlsx');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));