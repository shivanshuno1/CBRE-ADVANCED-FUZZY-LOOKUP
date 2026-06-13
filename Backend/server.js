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

// ============ GENERIC ABBREVIATION DICTIONARY ============
const abbreviationMap = {
  // Legal
  'pvt': 'private', 'ltd': 'limited', 'corp': 'corporation', 'inc': 'incorporated',
  'co': 'company', 'llc': 'limited liability company', 'llp': 'limited liability partnership',
  'plc': 'public limited company', 'pc': 'professional corporation', 'pa': 'professional association',
  
  // Business suffixes
  'ent': 'enterprise', 'grp': 'group', 'hldgs': 'holdings', 'hldg': 'holding',
  'intl': 'international', 'tech': 'technology', 'sols': 'solutions', 'svcs': 'services',
  'mgmt': 'management', 'assoc': 'associates', 'assn': 'association', 'inst': 'institute',
  'fdn': 'foundation', 'inds': 'industries', 'mfg': 'manufacturing', 'dist': 'distribution',
  'whs': 'wholesale', 'ret': 'retail', 'div': 'division', 'sub': 'subsidiary',
  
  // Common business terms
  'cap': 'capital', 'fin': 'financial', 'inv': 'investment', 'adv': 'advisors',
  'part': 'partners', 'cons': 'consulting', 'dig': 'digital', 'creat': 'creative',
  'media': 'media', 'comms': 'communications', 'telecom': 'telecommunications',
  'soft': 'software', 'sys': 'systems', 'net': 'networks', 'log': 'logistics',
  'supply': 'supply chain', 'prod': 'products', 'serv': 'services', 'soln': 'solutions',
  
  // Geographic
  'natl': 'national', 'nat': 'national', 'amer': 'american', 'glob': 'global',
  'eu': 'european', 'asia': 'asian', 'pac': 'pacific', 'atl': 'atlantic',
  'midwest': 'midwest', 'northeast': 'northeast', 'southeast': 'southeast',
  'southwest': 'southwest', 'northwest': 'northwest',
  
  // Government
  'gov': 'government', 'govt': 'government', 'auth': 'authority', 'dept': 'department',
  'agcy': 'agency', 'bur': 'bureau', 'comm': 'commission',
  
  // Industry specific
  'air': 'airlines', 'airways': 'airways', 'rail': 'railway', 'shipping': 'shipping',
  'maritime': 'maritime', 'energy': 'energy', 'power': 'power', 'util': 'utilities',
  'health': 'healthcare', 'pharma': 'pharmaceutical', 'biotech': 'biotechnology',
  'auto': 'automotive', 'aero': 'aerospace', 'def': 'defense', 'sec': 'security',
  'it': 'information technology', 'ai': 'artificial intelligence', 'ml': 'machine learning',
  
  // Common prefixes/suffixes
  'east': 'eastern', 'west': 'western', 'north': 'northern', 'south': 'southern',
  'central': 'central', 'metro': 'metropolitan', 'city': 'city', 'town': 'town',
  'village': 'village', 'heights': 'heights', 'park': 'park', 'plaza': 'plaza',
  'tower': 'tower', 'center': 'center', 'centre': 'centre', 'square': 'square',
};

// ============ GENERIC MISSPELLING CORRECTIONS ============
const misspellingMap = {
  // Logistics variations
  'logistix': 'logistics', 'logistik': 'logistics', 'logistc': 'logistics',
  'lodgistics': 'logistics', 'logisitcs': 'logistics',
  
  // Solutions variations
  'soloutions': 'solutions', 'solutons': 'solutions', 'solns': 'solutions',
  
  // Incorporated variations
  'incorperated': 'incorporated', 'incorp': 'incorporated', 'incorprated': 'incorporated',
  
  // Corporation variations
  'corperation': 'corporation', 'corparation': 'corporation', 'corpration': 'corporation',
  
  // Financial variations
  'finacial': 'financial', 'finanical': 'financial', 'fincial': 'financial',
  'finence': 'finance', 'finical': 'financial',
  
  // Common typos
  'teh': 'the', 'withe': 'with',
};

// ============ HELPER FUNCTIONS ============

// Escape regex special characters (FIXED - standalone function)
function escapeRegex(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Generic abbreviation expansion
function expandAbbreviations(text) {
  if (!text || typeof text !== 'string') return '';
  
  let result = ' ' + text.toLowerCase().replace(/[.,!?;:]$/g, '') + ' ';
  
  // Sort by length (longest first) to avoid partial replacements
  const sortedAbbrs = Object.keys(abbreviationMap).sort((a, b) => b.length - a.length);
  
  for (const abbr of sortedAbbrs) {
    const regex = new RegExp(`\\b${escapeRegex(abbr)}\\b`, 'gi');
    if (regex.test(result)) {
      result = result.replace(regex, ` ${abbreviationMap[abbr]} `);
    }
  }
  
  return result.trim();
}

// Generic misspelling correction
function correctMisspellings(text) {
  if (!text || typeof text !== 'string') return text;
  
  let result = ' ' + text.toLowerCase() + ' ';
  
  const sortedMisspellings = Object.keys(misspellingMap).sort((a, b) => b.length - a.length);
  
  for (const wrong of sortedMisspellings) {
    const regex = new RegExp(`\\b${escapeRegex(wrong)}\\b`, 'gi');
    if (regex.test(result)) {
      result = result.replace(regex, ` ${misspellingMap[wrong]} `);
    }
  }
  
  return result.trim();
}

// Calculate Levenshtein distance for character-level similarity
function levenshteinDistance(a, b) {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  
  const matrix = [];
  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }
  
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      const cost = a[j - 1] === b[i - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }
  
  return matrix[b.length][a.length];
}

// Character-level similarity (0-100)
function charSimilarity(a, b) {
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 100;
  const distance = levenshteinDistance(a, b);
  return Math.round((1 - distance / maxLen) * 100);
}

// GENERIC SIMILARITY CALCULATION
function calculateSimilarity(str1, str2) {
  if (!str1 || !str2) return 0;
  
  const s1 = String(str1);
  const s2 = String(str2);
  
  // Quick exact match after cleaning
  const clean1 = s1.toLowerCase().replace(/[^a-z0-9]/g, '');
  const clean2 = s2.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (clean1 === clean2) return 100;
  
  // Step 1: Expand abbreviations
  const expanded1 = expandAbbreviations(s1);
  const expanded2 = expandAbbreviations(s2);
  
  // Step 2: Correct misspellings
  const corrected1 = correctMisspellings(expanded1);
  const corrected2 = correctMisspellings(expanded2);
  
  // Step 3: Calculate multiple similarity metrics
  const tokenSetRatio = fuzzball.token_set_ratio(corrected1, corrected2);
  const tokenSortRatio = fuzzball.token_sort_ratio(corrected1, corrected2);
  const partialRatio = fuzzball.partial_ratio(corrected1, corrected2);
  const wRatio = fuzzball.WRatio(corrected1, corrected2);
  
  // Step 4: Character-level similarity (catches typos)
  const charSim = charSimilarity(corrected1, corrected2);
  
  // Step 5: Word overlap ratio
  const words1 = new Set(corrected1.split(/\s+/));
  const words2 = new Set(corrected2.split(/\s+/));
  const intersection = new Set([...words1].filter(x => words2.has(x)));
  const union = new Set([...words1, ...words2]);
  const wordOverlap = union.size === 0 ? 0 : (intersection.size / union.size) * 100;
  
  // Step 6: Combine scores with weights
  let finalScore = Math.max(
    tokenSetRatio * 0.4,
    wRatio * 0.3,
    charSim * 0.2,
    wordOverlap * 0.1
  );
  
  // Step 7: Apply dynamic penalties
  
  // Penalty for different string lengths
  const lengthRatio = Math.min(corrected1.length, corrected2.length) / 
                      Math.max(corrected1.length, corrected2.length);
  if (lengthRatio < 0.5) {
    finalScore *= 0.8; // 20% penalty for very different lengths
  } else if (lengthRatio < 0.7) {
    finalScore *= 0.9; // 10% penalty for moderately different lengths
  }
  
  // Penalty for missing key business words
  const keyWords = ['company', 'corporation', 'limited', 'incorporated', 'llc', 'inc', 'corp', 'ltd'];
  let missingKeyWords = 0;
  for (const word of keyWords) {
    const hasIn1 = corrected1.includes(word);
    const hasIn2 = corrected2.includes(word);
    if (hasIn1 !== hasIn2) missingKeyWords++;
  }
  if (missingKeyWords > 0) {
    finalScore *= (1 - Math.min(0.25, missingKeyWords * 0.05));
  }
  
  // Cap at 100 and round
  finalScore = Math.min(100, Math.max(0, Math.round(finalScore)));
  
  return finalScore;
}

// ============ PERFORMANCE OPTIMIZATIONS ============

// Build prefix index for fast candidate lookup
function buildIndex(data, column) {
  const index = new Map();
  for (let i = 0; i < data.length; i++) {
    const val = data[i][column] ? String(data[i][column]).toLowerCase() : '';
    if (val.length >= 2) {
      const prefix = val.substring(0, 2);
      if (!index.has(prefix)) index.set(prefix, []);
      index.get(prefix).push(i);
    }
  }
  return index;
}

// Cache expanded values for performance
function buildExpansionCache(data, column) {
  const cache = new Map();
  for (let i = 0; i < data.length; i++) {
    const val = data[i][column] ? String(data[i][column]) : '';
    cache.set(i, {
      raw: val,
      expanded: expandAbbreviations(val),
      corrected: correctMisspellings(expandAbbreviations(val))
    });
  }
  return cache;
}

// ============ API ENDPOINTS ============

app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    const wb = XLSX.readFile(req.file.path);
    workbook = wb;
    const sheets = wb.SheetNames.map(name => ({ name }));
    res.json({ success: true, uploadId: req.file.filename, sheets });
    fs.unlinkSync(req.file.path);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/columns/:uploadId/:sheetName', (req, res) => {
  if (!workbook) return res.status(404).json({ error: 'No workbook' });
  try {
    const sheet = workbook.Sheets[req.params.sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (!data.length) return res.json({ columns: [] });
    const headers = data[0].map(h => String(h));
    const columns = headers.map((name, idx) => ({ name, index: idx }));
    res.json({ columns });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
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
    
    const rightCache = buildExpansionCache(rightData, columnRight);
    const rightIndex = buildIndex(rightData, columnRight);
    
    const matched = [];
    const usedRight = new Set();
    let comparisons = 0;
    
    for (let li = 0; li < leftData.length; li++) {
      const leftRow = leftData[li];
      const leftVal = leftRow[columnLeft] ? String(leftRow[columnLeft]) : '';
      const leftPrefix = leftVal.length >= 2 ? leftVal.substring(0, 2).toLowerCase() : '';
      
      let candidates = rightIndex.get(leftPrefix) || [];
      
      if (candidates.length === 0 && rightData.length > 0) {
        const sampleSize = Math.min(200, rightData.length);
        candidates = Array.from({ length: sampleSize }, () => Math.floor(Math.random() * rightData.length));
      }
      
      let bestScore = 0;
      let bestIdx = -1;
      
      for (const ri of candidates) {
        if (usedRight.has(ri)) continue;
        
        const rightCached = rightCache.get(ri);
        const score = calculateSimilarity(leftVal, rightCached.raw);
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
  
  try {
    const leftSheet = workbook.Sheets[sheetLeft];
    const rightSheet = workbook.Sheets[sheetRight];
    let leftData = XLSX.utils.sheet_to_json(leftSheet);
    let rightData = XLSX.utils.sheet_to_json(rightSheet);
    
    const rightCache = buildExpansionCache(rightData, columnRight);
    const rightIndex = buildIndex(rightData, columnRight);
    
    const matched = [];
    const unmatchedLeft = [];
    const unmatchedRight = [];
    const usedRight = new Set();
    
    for (const leftRow of leftData) {
      const leftVal = leftRow[columnLeft] ? String(leftRow[columnLeft]) : '';
      const leftPrefix = leftVal.length >= 2 ? leftVal.substring(0, 2).toLowerCase() : '';
      
      let candidates = rightIndex.get(leftPrefix) || [];
      if (candidates.length === 0 && rightData.length > 0) {
        const sampleSize = Math.min(200, rightData.length);
        candidates = Array.from({ length: sampleSize }, () => Math.floor(Math.random() * rightData.length));
      }
      
      let bestScore = 0;
      let bestIdx = -1;
      
      for (const ri of candidates) {
        if (usedRight.has(ri)) continue;
        const rightCached = rightCache.get(ri);
        const score = calculateSimilarity(leftVal, rightCached.raw);
        
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
    
    for (let i = 0; i < rightData.length; i++) {
      if (!usedRight.has(i)) {
        unmatchedRight.push({ ...rightData[i], _matchStatus: 'No match found' });
      }
    }
    
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
    
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));