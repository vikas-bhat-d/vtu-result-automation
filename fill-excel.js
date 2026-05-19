'use strict';

/**
 * fill-excel.js
 * Reads all output/<usn>_*.json files and fills CIE + SEE marks
 * into the "VII Semester Marks" sheet of the template Excel.
 * 
 * Only the two raw input columns per subject are written —
 * all totals, grade points, grades, SGPA, %, result are
 * already formula-driven in the template and will auto-recalculate.
 */

const ExcelJS = require('exceljs');
const fs      = require('fs');
const path    = require('path');

const TEMPLATE = path.join(__dirname, 'Copy of 4SU22-Sem marks.xlsx');
const OUT_DIR  = path.join(__dirname, 'output');
const SHEET    = 'VII Semester Marks';
const DATA_START_ROW = 7;   // first student row

// subject code → { cie column letter, see column letter }
const SUBJECT_COLS = {
  'BIS701' : { cie: 'D',  see: 'E'  },
  'BCS702' : { cie: 'I',  see: 'J'  },
  'BIS703' : { cie: 'N',  see: 'O'  },
  'BCS714A': { cie: 'S',  see: 'T'  },
  'BTE755C': { cie: 'X',  see: 'Y'  },
  'BIS786' : { cie: 'AC', see: 'AD' },
};

async function fillExcel(resultsMap) {
  // resultsMap: { '4SU22IS062': { usn, name, semester, subjects: [...] }, ... }

  const wb    = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE);

  const sheet = wb.getWorksheet(SHEET);
  if (!sheet) throw new Error(`Sheet "${SHEET}" not found in workbook.`);

  // Build a USN → row number index from the sheet
  const usnRowIndex = {};
  sheet.eachRow({ includeEmpty: false }, (row, rNum) => {
    if (rNum < DATA_START_ROW) return;
    const cellVal = row.getCell('B').value;
    if (cellVal) usnRowIndex[String(cellVal).toUpperCase().trim()] = rNum;
  });

  let filled = 0, skipped = 0;

  for (const [usnKey, result] of Object.entries(resultsMap)) {
    const usn   = usnKey.toUpperCase();
    const rowNum = usnRowIndex[usn];

    if (!rowNum) {
      console.log(`  [SKIP] ${usn} — not found in Excel sheet (check USN or sheet).`);
      skipped++;
      continue;
    }

    const row = sheet.getRow(rowNum);
    let subjectsFilled = 0;

    for (const subj of result.subjects) {
      const cols = SUBJECT_COLS[subj.code.trim().toUpperCase()];
      if (!cols) {
        console.log(`  [WARN] ${usn} — unknown subject code "${subj.code}", skipping column mapping.`);
        continue;
      }

      row.getCell(cols.cie).value = subj.internal;
      row.getCell(cols.see).value = subj.external;
      subjectsFilled++;
    }

    row.commit();
    console.log(`  ${usn.padEnd(15)} row ${rowNum} — filled ${subjectsFilled} subject(s)`);
    filled++;
  }

  // Save alongside template with a timestamp suffix
  const ts      = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const outPath = path.join(__dirname, `output/SEM7_Marks_Filled_${ts}.xlsx`);
  await wb.xlsx.writeFile(outPath);

  console.log(`\nExcel saved → ${outPath}`);
  console.log(`Filled: ${filled}  |  Skipped: ${skipped}`);
  return outPath;
}

// ── Standalone mode: load all JSON files from output/ and fill ────────────────
async function main() {
  const jsonFiles = fs.readdirSync(OUT_DIR)
    .filter(f => f.endsWith('.json') && !f.includes('raw'));

  if (!jsonFiles.length) {
    console.error('No JSON result files found in output/. Run index.js first.');
    process.exit(1);
  }

  const resultsMap = {};
  for (const file of jsonFiles) {
    const data = JSON.parse(fs.readFileSync(path.join(OUT_DIR, file), 'utf8'));
    if (data.usn) resultsMap[data.usn.toUpperCase()] = data;
  }

  console.log(`Loaded ${Object.keys(resultsMap).length} result(s): ${Object.keys(resultsMap).join(', ')}`);
  await fillExcel(resultsMap);
}

module.exports = { fillExcel };

if (require.main === module) main().catch(err => {
  console.error('[ERROR]', err.message);
  process.exit(1);
});
