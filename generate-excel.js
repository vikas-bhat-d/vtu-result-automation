/**
 * VTU Result Automation — generate-excel.js
 * --------------------------------------------------
 * Reads a CSV of USNs, fetches each student's result
 * (with per-USN CAPTCHA), and generates a fully-formatted
 * Excel report matching the original college template.
 *
 * Usage:  node generate-excel.js [path-to-usn-list.csv]
 */

'use strict';

const axios    = require('axios');
const cheerio  = require('cheerio');
const https    = require('https');
const ExcelJS  = require('exceljs');
const fs       = require('fs');
const path     = require('path');
const readline = require('readline');
const { exec } = require('child_process');

// ── Config ─────────────────────────────────────────────────────────────────────
const EXAM_PATH = 'D25J26Ecbcs';
const BASE_HOST = 'results.vtu.ac.in';
const BASE_URL  = `https://${BASE_HOST}`;
const OUT_DIR   = path.join(__dirname, 'output');
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

// ── Axios ──────────────────────────────────────────────────────────────────────
const cookieJar = {};
const client = axios.create({
  baseURL    : BASE_URL,
  httpsAgent : new https.Agent({ rejectUnauthorized: false }),
  headers    : {
    'User-Agent'        : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept-Language'   : 'en-US,en;q=0.9',
    'Referer'           : `${BASE_URL}/${EXAM_PATH}/index.php`,
  },
  withCredentials: true,
  maxRedirects   : 5
});
client.interceptors.response.use(res => {
  const raw = res.headers['set-cookie'];
  if (raw) {
    (Array.isArray(raw) ? raw : [raw]).forEach(entry => {
      const [kv] = entry.split(';');
      const idx  = kv.indexOf('=');
      if (idx > 0) cookieJar[kv.slice(0, idx).trim()] = kv.slice(idx + 1).trim();
    });
  }
  return res;
});
client.interceptors.request.use(cfg => {
  const cookie = Object.entries(cookieJar).map(([k, v]) => `${k}=${v}`).join('; ');
  if (cookie) cfg.headers['Cookie'] = cookie;
  return cfg;
});

// ── CLI helpers ────────────────────────────────────────────────────────────────
const rl  = readline.createInterface({ input: process.stdin, output: process.stdout });
const ask = q => new Promise(r => rl.question(q, r));

// ── VTU helpers ────────────────────────────────────────────────────────────────
async function getToken() {
  process.stdout.write('Fetching Token … ');
  const { data: html } = await client.get(`/${EXAM_PATH}/index.php`);
  const patterns = [
    /name=["']Token["'][^>]*value=["']([a-f0-9]{20,})["']/i,
    /value=["']([a-f0-9]{20,})["'][^>]*name=["']Token["']/i,
    /["']Token["']\s*[=:]\s*["']([a-f0-9]{20,})["']/i,
    /var\s+[Tt]oken\s*=\s*["']([a-f0-9]{20,})["']/,
    /[Tt]oken[=\s:]+([a-f0-9]{20,})/
  ];
  for (const p of patterns) {
    const m = html.match(p);
    if (m) { console.log(`found → ${m[1]}`); return m[1]; }
  }
  console.log('not found automatically.');
  return (await ask('  Paste Token value here: ')).trim();
}

function buildT() {
  return `${Math.random().toFixed(8)} ${Math.floor(Date.now() / 1000)}`;
}

async function getCaptcha() {
  const t = buildT();
  process.stdout.write('Fetching CAPTCHA … ');
  const { data: buf } = await client.get(
    `/captcha/vtu_captcha.php?_CAPTCHA&t=${encodeURIComponent(t)}`,
    { responseType: 'arraybuffer', headers: { 'Accept': 'image/*' } }
  );
  console.log('done.');
  const imgPath = path.join(OUT_DIR, `captcha_${Date.now()}.png`);
  fs.writeFileSync(imgPath, buf);
  return imgPath;
}

function openImage(imgPath) {
  return new Promise(r => exec(`start "" "${imgPath}"`, () => r()));
}

async function fetchResult(usn, token, code) {
  process.stdout.write(`  Submitting for ${usn} … `);
  const { data: html, status } = await client.post(
    `/${EXAM_PATH}/resultpage.php`,
    new URLSearchParams({ Token: token, lns: usn, captchacode: code }).toString(),
    { headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Origin': BASE_URL } }
  );
  console.log(`HTTP ${status}`);
  return html;
}

function parseResult(html) {
  const $ = cheerio.load(html);
  let usn = '', name = '';
  $('table tr').each((_, tr) => {
    const tds = $(tr).find('td');
    if (tds.length >= 2) {
      const label = $(tds[0]).text().trim();
      const value = $(tds[1]).text().replace(/^[\s:]+/, '').trim();
      if (/university\s*seat\s*number/i.test(label)) usn  = value;
      if (/student\s*name/i.test(label))              name = value;
    }
  });
  const subjects = [];
  $('.divTableBody .divTableRow').each((_, row) => {
    const cells = $(row).find('.divTableCell').map((_, td) => $(td).text().trim()).get();
    if (!cells[0] || /subject\s*code/i.test(cells[0]) || /->/.test(cells[0])) return;
    subjects.push({
      code    : cells[0],
      name    : cells[1],
      internal: isNaN(+cells[2]) ? cells[2] : +cells[2],
      external: isNaN(+cells[3]) ? cells[3] : +cells[3],
    });
  });
  return { usn, name, subjects };
}

// ── CSV reader ─────────────────────────────────────────────────────────────────
function readUSNsFromCSV(csvPath) {
  const raw   = fs.readFileSync(csvPath, 'utf8').replace(/^\uFEFF/, ''); // strip BOM
  const lines = raw.split(/\r?\n/).filter(l => l.trim());
  const usns  = [];

  for (const line of lines) {
    // Support both "one per row" and "all on one row, comma-separated"
    const cells = line.split(',').map(c => c.trim().replace(/^["']|["']$/g, ''));
    for (const cell of cells) {
      if (!cell) continue;
      if (/^(usn|sl\.?\s*no|number|name)/i.test(cell)) continue; // skip header words
      usns.push(cell.toUpperCase());
    }
  }
  return usns;
}

// ── Load results from existing JSON files ─────────────────────────────────────
function loadResultsFromJSON(usns) {
  const resultsMap  = {};
  const subjectMeta = [];

  for (const usn of usns) {
    // Find latest JSON for this USN in output dir
    const files = fs.readdirSync(OUT_DIR)
      .filter(f => f.startsWith(usn + '_') && f.endsWith('.json'))
      .sort();
    if (files.length === 0) {
      console.log(`  [!] No JSON found for ${usn}, will be left blank in Excel.`);
      resultsMap[usn] = { usn, name: '', subjects: [] };
      continue;
    }
    const latest = files[files.length - 1];
    const data   = JSON.parse(fs.readFileSync(path.join(OUT_DIR, latest), 'utf8'));
    resultsMap[usn] = data;
    console.log(`  Loaded ${latest}`);

    if (subjectMeta.length === 0 && data.subjects && data.subjects.length > 0) {
      data.subjects.forEach(s => subjectMeta.push({ code: s.code, name: s.name }));
    }
  }
  return { resultsMap, subjectMeta };
}

// ── Excel column letter helper ─────────────────────────────────────────────────
function colL(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ── Excel generation ──────────────────────────────────────────────────────────
async function generateExcel(resultsMap, subjects, credits, collegeName, batchName, examName, semLabel) {
  const N = subjects.length;
  // Col layout: 1=Sl, 2=USN, 3=Name, then 5 cols per subject, then SGPA(%,Result), AK,AL,AM
  // Subject i (0-based): CIE=4+i*5, SEE=5+i*5, Total=6+i*5, GrPt=7+i*5, Grade=8+i*5
  const CIE_COL  = i => 4 + i * 5;
  const SEE_COL  = i => 5 + i * 5;
  const TOT_COL  = i => 6 + i * 5;
  const GRP_COL  = i => 7 + i * 5;
  const GRD_COL  = i => 8 + i * 5;
  const SGPA_COL = 4 + N * 5;
  const PCT_COL  = SGPA_COL + 1;
  const RES_COL  = SGPA_COL + 2;
  const PEND_COL = SGPA_COL + 3;
  const PENDL_COL= SGPA_COL + 4;
  const PASA_COL = SGPA_COL + 5;
  const TOTAL_COLS = PASA_COL;

  const totalCredits = credits.reduce((a, c) => a + c, 0);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'VTU Result Automation';
  const ws = wb.addWorksheet(semLabel + ' Marks', { pageSetup: { fitToPage: true } });

  // Freeze columns A, B, C so Sl.No / USN / Name stay visible when scrolling right
  ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 0, topLeftCell: 'D1' }];

  // ── Column widths ────────────────────────────────────────────────────────────
  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 14.63;
  ws.getColumn(3).width = 35.63;
  for (let i = 0; i < N; i++) {
    for (let c = 0; c < 5; c++) ws.getColumn(CIE_COL(i) + c).width = 5.75;
  }
  ws.getColumn(SGPA_COL).width = 9.38;
  ws.getColumn(PCT_COL).width  = 6.38;
  ws.getColumn(RES_COL).width  = 7.38;
  ws.getColumn(PEND_COL).width = 9.75;
  ws.getColumn(PENDL_COL).width= 37.63;
  ws.getColumn(PASA_COL).width = 37.63;

  // ── Style helpers ────────────────────────────────────────────────────────────
  const fillPurple = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF351C75' } };
  const fillPink   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6B8AF' } };
  const fillYellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  const fillOrange = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE5CD' } };
  const fillGreen  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAD3' } };
  const fillPink2  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAD1DC' } };

  const thin  = { style: 'thin' };
  const border = { top: thin, left: thin, bottom: thin, right: thin };

  function styleCell(cell, opts = {}) {
    const { fill, bold, size = 12, color, halign = 'center', valign = 'middle', wrap = false } = opts;
    if (fill) cell.fill = fill;
    cell.font = { bold: !!bold, size, color: color ? { argb: color } : undefined };
    cell.alignment = { horizontal: halign, vertical: valign, wrapText: wrap };
    cell.border = border;
  }

  // ── Row 1 — College name merged A1:C1 only; D1 onwards left empty ─────────
  ws.getRow(1).height = 30;
  ws.mergeCells(1, 1, 1, 3);
  const r1 = ws.getCell(1, 1);
  r1.value = collegeName.toUpperCase();
  styleCell(r1, { fill: fillPurple, bold: true, size: 17, color: 'FFFFFFFF' });

  // ── Row 2 — Batch + Exam title ────────────────────────────────────────────
  ws.getRow(2).height = 30;
  ws.mergeCells(2, 1, 2, 3);
  const r2a = ws.getCell(2, 1);
  r2a.value = batchName.toUpperCase();
  styleCell(r2a, { fill: fillPurple, bold: true, size: 14, color: 'FFEDFA4A' });

  ws.mergeCells(2, 4, 2, TOTAL_COLS);
  const r2b = ws.getCell(2, 4);
  r2b.value = examName.toUpperCase();
  styleCell(r2b, { fill: fillPurple, bold: true, size: 14, color: 'FFFFFFFF', halign: 'left', valign: 'middle' });

  // ── Rows 3-6: A,B,C merged vertically ────────────────────────────────────
  ws.getRow(3).height = 21;
  ws.mergeCells(3, 1, 6, 1);
  ws.mergeCells(3, 2, 6, 2);
  ws.mergeCells(3, 3, 6, 3);
  const labels3 = ['Sl. No.', 'USN', 'Name'];
  [1, 2, 3].forEach((c, idx) => {
    const cell = ws.getCell(3, c);
    cell.value = labels3[idx];
    styleCell(cell, { fill: fillPink, bold: true });
  });

  // ── Row 3 — Subject codes ─────────────────────────────────────────────────
  for (let i = 0; i < N; i++) {
    ws.mergeCells(3, CIE_COL(i), 3, GRD_COL(i));
    const cell = ws.getCell(3, CIE_COL(i));
    cell.value = subjects[i].code;
    styleCell(cell, { fill: fillPink, bold: true });
  }
  // Results header
  ws.mergeCells(3, SGPA_COL, 3, RES_COL);
  const resHdr = ws.getCell(3, SGPA_COL);
  resHdr.value = 'Results';
  styleCell(resHdr, { fill: fillPink, bold: true });
  // AK,AL,AM merged rows 3-6
  ws.mergeCells(3, PEND_COL, 6, PEND_COL);
  ws.mergeCells(3, PENDL_COL, 6, PENDL_COL);
  ws.mergeCells(3, PASA_COL, 6, PASA_COL);
  ['Number of Subjects Pending', 'List of Pending Subjects', 'PASSED ARREAR Subjects'].forEach((v, idx) => {
    const cell = ws.getCell(3, PEND_COL + idx);
    cell.value = v;
    styleCell(cell, { fill: fillGreen, bold: true, size: 11, wrap: true });
  });

  // ── Row 4 — Subject names ─────────────────────────────────────────────────
  ws.getRow(4).height = 15.75;
  for (let i = 0; i < N; i++) {
    ws.mergeCells(4, CIE_COL(i), 4, GRD_COL(i));
    const cell = ws.getCell(4, CIE_COL(i));
    cell.value = subjects[i].name;
    styleCell(cell, { fill: fillPink, bold: true });
  }
  // fill merged results area rows 4-5-6 (just border/fill)
  for (let c = SGPA_COL; c <= RES_COL; c++) {
    for (let r = 4; r <= 6; r++) {
      ws.getCell(r, c).border = border;
    }
  }

  // ── Row 5 — Credits ───────────────────────────────────────────────────────
  ws.getRow(5).height = 15.75;
  for (let i = 0; i < N; i++) {
    // "Credits:" label at Total col, value at GrPt col
    const lblCell = ws.getCell(5, TOT_COL(i));
    lblCell.value = 'Credits:';
    styleCell(lblCell, { fill: fillPink, bold: true });

    const valCell = ws.getCell(5, GRP_COL(i));
    valCell.value = credits[i];
    styleCell(valCell, { fill: fillPink, bold: true });

    // Fill CIE, SEE, Grade cols in row 5 with pink and border
    [CIE_COL(i), SEE_COL(i), GRD_COL(i)].forEach(c => {
      styleCell(ws.getCell(5, c), { fill: fillPink, bold: false });
    });
  }

  // ── Row 6 — Sub-headers CIE/SEE/Total/Gr.Pt./Grade ───────────────────────
  ws.getRow(6).height = 15.75;
  const subHdrs = ['CIE', 'SEE', 'Total', 'Gr. Pt.', 'Grade'];
  for (let i = 0; i < N; i++) {
    subHdrs.forEach((h, off) => {
      const cell = ws.getCell(6, CIE_COL(i) + off);
      cell.value = h;
      styleCell(cell, { fill: fillPink, bold: true });
    });
  }
  ['SGPA', '%', 'Result'].forEach((h, off) => {
    const cell = ws.getCell(6, SGPA_COL + off);
    cell.value = h;
    styleCell(cell, { fill: fillPink, bold: true });
  });

  // ── Data rows ─────────────────────────────────────────────────────────────
  const usnList = Object.keys(resultsMap);

  // Determine which subjects are "project" (marks out of 200)
  const isProject = subjects.map(s =>
    /project|dissertation/i.test(s.name) || /project|dissertation/i.test(s.code)
  );

  for (let idx = 0; idx < usnList.length; idx++) {
    const usn = usnList[idx];
    const r   = 7 + idx;
    ws.getRow(r).height = 26.25;
    const student = resultsMap[usn];

    // A — Sl. No.
    const aCell = ws.getCell(r, 1);
    aCell.value = idx + 1;
    styleCell(aCell, { fill: fillPink });

    // B — USN
    const bCell = ws.getCell(r, 2);
    bCell.value = student ? student.usn || usn : usn;
    styleCell(bCell, { fill: fillPink });

    // C — Name
    const cCell = ws.getCell(r, 3);
    cCell.value = student ? student.name : '';
    styleCell(cCell, { fill: fillPink, halign: 'left' });

    for (let i = 0; i < N; i++) {
      const cieLetter = colL(CIE_COL(i));
      const seeLetter = colL(SEE_COL(i));
      const totLetter = colL(TOT_COL(i));
      const grpLetter = colL(GRP_COL(i));
      const grdLetter = colL(GRD_COL(i));

      // Write CIE + SEE from result
      const subjData = student && student.subjects ? student.subjects.find(s => s.code === subjects[i].code) : null;
      const cieCell = ws.getCell(r, CIE_COL(i));
      const seeCell = ws.getCell(r, SEE_COL(i));
      cieCell.value = subjData && !isNaN(subjData.internal) ? subjData.internal : null;
      seeCell.value = subjData && !isNaN(subjData.external) ? subjData.external : null;
      cieCell.border = border;
      seeCell.border = border;
      cieCell.font = { size: 11, color: { argb: 'FF000000' } };
      seeCell.font = { size: 11, color: { argb: 'FF000000' } };
      cieCell.alignment = { horizontal: 'center', vertical: 'middle' };
      seeCell.alignment = { horizontal: 'center', vertical: 'middle' };

      // Total formula
      const totCell = ws.getCell(r, TOT_COL(i));
      const fillUsed = isProject[i] ? fillOrange : fillYellow;
      totCell.value = { formula: `IF(ISBLANK(${cieLetter}${r}),"",${cieLetter}${r}+${seeLetter}${r})` };
      styleCell(totCell, { fill: fillUsed });

      // Grade Points formula
      const grpCell = ws.getCell(r, GRP_COL(i));
      let grpFmla;
      if (isProject[i]) {
        grpFmla = `IF(ISBLANK(${cieLetter}${r}),"",IF(OR(${cieLetter}${r}<40,${seeLetter}${r}<35,${totLetter}${r}<80),0,IF(${totLetter}${r}>=180,10,IF(${totLetter}${r}>=160,9,IF(${totLetter}${r}>=140,8,IF(${totLetter}${r}>=120,7,IF(${totLetter}${r}>=110,6,IF(${totLetter}${r}>=100,5,IF(${totLetter}${r}>=80,4,0)))))))))`;
      } else {
        grpFmla = `IF(ISBLANK(${cieLetter}${r}),"",IF(OR(${cieLetter}${r}<20,${seeLetter}${r}<18,${totLetter}${r}<40),0,IF(${totLetter}${r}>=90,10,IF(${totLetter}${r}>=80,9,IF(${totLetter}${r}>=70,8,IF(${totLetter}${r}>=60,7,IF(${totLetter}${r}>=55,6,IF(${totLetter}${r}>=50,5,IF(${totLetter}${r}>=40,4,0)))))))))`;
      }
      grpCell.value = { formula: grpFmla };
      styleCell(grpCell, { fill: fillUsed, valign: 'bottom' });

      // Grade formula
      const grdCell = ws.getCell(r, GRD_COL(i));
      let grdFmla;
      if (isProject[i]) {
        grdFmla = `IF(ISBLANK(${cieLetter}${r}),"",IF(OR(${cieLetter}${r}<40,${seeLetter}${r}<35,${totLetter}${r}<80),"F",IF(${totLetter}${r}>=180,"O",IF(${totLetter}${r}>=160,"A+",IF(${totLetter}${r}>=140,"A",IF(${totLetter}${r}>=120,"B+",IF(${totLetter}${r}>=110,"B",IF(${totLetter}${r}>=100,"C","P"))))))))`;
      } else {
        grdFmla = `IF(ISBLANK(${cieLetter}${r}),"",IF(${cieLetter}${r}<20,"NE",IF(OR(${seeLetter}${r}<18,${totLetter}${r}<40),"F",IF(${totLetter}${r}>=90,"O",IF(${totLetter}${r}>=80,"A+",IF(${totLetter}${r}>=70,"A",IF(${totLetter}${r}>=60,"B+",IF(${totLetter}${r}>=55,"B",IF(${totLetter}${r}>=50,"C",IF(${totLetter}${r}>=40,"P","F"))))))))))`;
      }
      grdCell.value = { formula: grdFmla };
      styleCell(grdCell, { fill: fillUsed, valign: 'bottom' });
    }

    // SGPA: SUM(c0*GP0, c1*GP1, ...) / totalCredits
    const sgpaLetter = colL(SGPA_COL);
    const pctLetter  = colL(PCT_COL);
    const sgpaParts  = subjects.map((_, i) => `${credits[i]}*${colL(GRP_COL(i))}${r}`).join(',');
    const sgpaCell   = ws.getCell(r, SGPA_COL);
    sgpaCell.value = { formula: `(SUM(${sgpaParts})/${totalCredits})` };
    styleCell(sgpaCell, { fill: fillYellow, valign: 'bottom' });

    // %
    const pctCell = ws.getCell(r, PCT_COL);
    pctCell.value = { formula: `${sgpaLetter}${r}*10` };
    styleCell(pctCell, { fill: fillYellow });

    // Result
    const gradeCols = subjects.map((_, i) => `${colL(GRD_COL(i))}${r}`);
    const orParts   = gradeCols.flatMap(g => [`${g}="F"`, `${g}="NE"`]).join(',');
    const resFmla   = `IF(IF(OR(${orParts}),"Fail","Pass")="Pass",IF(${pctLetter}${r}>=70,"FCD",IF(${pctLetter}${r}>=60,"FC",IF(${pctLetter}${r}>=40,"SC"))),"Fail")`;
    const resCell   = ws.getCell(r, RES_COL);
    resCell.value = { formula: resFmla };
    styleCell(resCell, { fill: fillYellow });

    // Pending count
    const gradeSet = `{${gradeCols.join(',')}}`;
    const passList = ['O','A+','A','B+','B','C','P'].map(g => `COUNTIF(${gradeSet},"${g}")`).join('+');
    const pendCell = ws.getCell(r, PEND_COL);
    pendCell.value = { formula: `${N}-(${passList})` };
    styleCell(pendCell, { fill: fillPink2 });

    // List of Pending (blank, user fills)
    const pendlCell = ws.getCell(r, PENDL_COL);
    pendlCell.value = '               ';
    styleCell(pendlCell, { fill: fillPink2, halign: 'left', wrap: true });

    // PASSED ARREAR (blank)
    ws.getCell(r, PASA_COL).border = border;
  }

  // ── Summary rows ─────────────────────────────────────────────────────────
  const fillBlue = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC9DAF8' } };
  const D_START  = 7;
  const D_END    = 6 + usnList.length;
  const SR       = D_END + 1; // first summary row

  const ROW_AVG  = SR,     ROW_NSTU = SR+1,  ROW_NE   = SR+2,  ROW_ABS  = SR+3,
        ROW_APP  = SR+4,  ROW_PASS = SR+5,  ROW_FAIL = SR+6,
        ROW_FCD  = SR+7,  ROW_FC   = SR+8,  ROW_SC   = SR+9,
        ROW_PG   = SR+10, ROW_PCT  = SR+11;

  const sumLabels = [
    [ROW_AVG,  'Average Marks (%)'], [ROW_NSTU, 'No. of Students'],
    [ROW_NE,   'Not Eligible'],      [ROW_ABS,  'Absent'],
    [ROW_APP,  'Students Appeared'], [ROW_PASS, 'Passed'],
    [ROW_FAIL, 'Fail'],              [ROW_FCD,  'FCD'],
    [ROW_FC,   'FC'],                [ROW_SC,   'SC'],
    [ROW_PG,   'Pass'],              [ROW_PCT,  '% Pass Result'],
  ];

  function sCell(r, c, val) {
    const cell = ws.getCell(r, c);
    cell.value = (typeof val === 'string' && val.startsWith('='))
      ? { formula: val.slice(1) } : val;
    cell.fill      = fillBlue;
    cell.font      = { size: 12 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border    = border;
  }

  // Labels in col C
  sumLabels.forEach(([r, label]) => {
    ws.getRow(r).height = 26.25;
    const lc = ws.getCell(r, 3);
    lc.value = label; lc.fill = fillBlue;
    lc.font  = { size: 12 };
    lc.alignment = { horizontal: 'left', vertical: 'middle' };
    lc.border = border;
  });

  const sgpaL = colL(SGPA_COL);
  const resL  = colL(RES_COL);
  const dRange = (col, s, e) => `${col}${s}:${col}${e}`;

  for (let i = 0; i < N; i++) {
    const cL  = colL(CIE_COL(i)), sL  = colL(SEE_COL(i));
    const tL  = colL(TOT_COL(i)), gpL = colL(GRP_COL(i)), grL = colL(GRD_COL(i));
    const dr  = (col) => dRange(col, D_START, D_END);

    // Average Marks (%): CIE and SEE as % of 50, Total as raw avg
    sCell(ROW_AVG, CIE_COL(i), `=TEXT(IFERROR(AVERAGE(${dr(cL)})*100/50,""),"0.0")`);
    sCell(ROW_AVG, SEE_COL(i), `=TEXT(IFERROR(AVERAGE(${dr(sL)})*100/50,""),"0.0")`);
    ws.mergeCells(ROW_AVG, TOT_COL(i), ROW_AVG, GRD_COL(i));
    sCell(ROW_AVG, TOT_COL(i), `=TEXT(IFERROR(AVERAGE(${dr(tL)}),""),"0.0")`);

    // No. of Students
    sCell(ROW_NSTU, GRP_COL(i), `=COUNTA($B$${D_START}:$B$${D_END})`);

    // Not Eligible
    sCell(ROW_NE,   GRP_COL(i), `=COUNTIF(${dr(grL)},"NE")`);

    // Absent (SEE = 0)
    sCell(ROW_ABS,  GRP_COL(i), `=COUNTIF(${dr(sL)},"0")`);

    // Students Appeared
    sCell(ROW_APP,  GRP_COL(i), `=COUNTA(${dr(grL)})-${gpL}${ROW_NE}-${gpL}${ROW_ABS}`);

    // Passed
    const passF = ['O','A+','A','B+','B','C','P'].map(g => `COUNTIFS(${dr(grL)},"${g}")`).join('+');
    sCell(ROW_PASS, GRP_COL(i), `=${passF}`);

    // Fail
    sCell(ROW_FAIL, GRP_COL(i), `=COUNTIF(${dr(grL)},"F")`);

    // FCD / FC / SC / Pass-grade
    sCell(ROW_FCD, GRP_COL(i), `=COUNTIF(${dr(grL)},"=O")+COUNTIF(${dr(grL)},"=A+")+COUNTIF(${dr(grL)},"=A")`);
    sCell(ROW_FC,  GRP_COL(i), `=COUNTIF(${dr(grL)},"=B+")`);
    sCell(ROW_SC,  GRP_COL(i), `=COUNTIF(${dr(grL)},"=B")+COUNTIF(${dr(grL)},"=C")`);
    sCell(ROW_PG,  GRP_COL(i), `=COUNTIF(${dr(grL)},"=P")`);

    // % Pass Result: merged GrPt:Grade
    ws.mergeCells(ROW_PCT, GRP_COL(i), ROW_PCT, GRD_COL(i));
    sCell(ROW_PCT, GRP_COL(i), `=TEXT(IFERROR(${gpL}${ROW_PASS}*100/${gpL}${ROW_APP},""),"0.00")`);
  }

  // Overall stats at SGPA_COL
  const dro = (col) => dRange(col, D_START, D_END);
  sCell(ROW_NSTU, SGPA_COL, `=COUNTA($B$${D_START}:$B$${D_END})`);
  sCell(ROW_NE,   SGPA_COL, `=COUNTIF(${dro(resL)},"NE")`);
  sCell(ROW_APP,  SGPA_COL, `=SUM(${sgpaL}${ROW_PASS},${sgpaL}${ROW_FAIL})`);
  sCell(ROW_PASS, SGPA_COL, `=SUM(${sgpaL}${ROW_FCD}:${sgpaL}${ROW_SC})`);
  sCell(ROW_FAIL, SGPA_COL, `=COUNTIFS(${dro(resL)},"Fail")`);
  sCell(ROW_FCD,  SGPA_COL, `=COUNTIFS(${dro(resL)},"FCD")`);
  sCell(ROW_FC,   SGPA_COL, `=COUNTIFS(${dro(resL)},"FC")`);
  sCell(ROW_SC,   SGPA_COL, `=COUNTIFS(${dro(resL)},"SC")`);
  // % Pass overall: merged SGPA:Result
  ws.mergeCells(ROW_PCT, SGPA_COL, ROW_PCT, RES_COL);
  sCell(ROW_PCT, SGPA_COL, `=TEXT(IFERROR(${sgpaL}${ROW_PASS}*100/${sgpaL}${ROW_APP},""),"0.00")`);

  // ── Analytics block polish: fill + structural borders ────────────────────
  // Border weight helpers
  const mediumSide = { style: 'medium' };
  const thickSide  = { style: 'thick' };

  // Pass 1: paint EVERY cell in the analytics block blue with a thin grid border
  // (sCell already styled data cells; this covers the empty/gap cells too)
  for (let r = SR; r <= ROW_PCT; r++) {
    for (let c = 1; c <= TOTAL_COLS; c++) {
      const cell = ws.getCell(r, c);
      cell.fill   = fillBlue;
      // Keep any existing border (data cells already have thin all-around),
      // or apply the default thin border to empty cells
      if (!cell.border) cell.border = { top: thin, left: thin, bottom: thin, right: thin };
    }
  }

  // Helper: override specific border sides on a cell without losing the others
  function setBorderSide(r, c, overrides) {
    const cell = ws.getCell(r, c);
    const b    = cell.border || { top: thin, left: thin, bottom: thin, right: thin };
    cell.border = Object.assign({}, b, overrides);
  }

  // Pass 2: structural dividers ──────────────────────────────────────────────

  // (a) Top edge of entire analytics block — medium top on first summary row
  for (let c = 1; c <= TOTAL_COLS; c++) setBorderSide(SR, c, { top: mediumSide });

  // (b) Bottom edge of entire analytics block — thick bottom on last row
  for (let c = 1; c <= TOTAL_COLS; c++) setBorderSide(ROW_PCT, c, { bottom: thickSide });

  // (c) Average Marks row — medium bottom (separates avg from count block)
  for (let c = 1; c <= TOTAL_COLS; c++) setBorderSide(ROW_AVG, c, { bottom: mediumSide });

  // (d) Frozen-pane divider — medium right on col 3 (Name) for all summary rows
  for (let r = SR; r <= ROW_PCT; r++) setBorderSide(r, 3, { right: mediumSide });

  // (e) Subject-group dividers — medium right on each subject's last col (Grade)
  for (let i = 0; i < N; i++) {
    for (let r = SR; r <= ROW_PCT; r++) setBorderSide(r, GRD_COL(i), { right: mediumSide });
  }

  // (f) Student-summary / Pending divider — medium right on Result col
  for (let r = SR; r <= ROW_PCT; r++) setBorderSide(r, RES_COL, { right: mediumSide });

  // ── Save ──────────────────────────────────────────────────────────────────
  const ts       = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const outPath  = path.join(OUT_DIR, `SEM_Marks_Generated_${ts}.xlsx`);
  await wb.xlsx.writeFile(outPath);
  console.log(`\nExcel saved → ${outPath}`);
  return outPath;
}

// ── Main ──────────────────────────────────────────────────────────────────────
// Usage:
//   node generate-excel.js usns.csv                         ← live fetch, all prompts
//   node generate-excel.js usns.csv --from-json             ← load output/*.json, all prompts
//   node generate-excel.js usns.csv config.json             ← live fetch, no prompts
//   node generate-excel.js usns.csv config.json --from-json ← fully automated, zero prompts
async function main() {
  const args = process.argv.slice(2);

  // ── Parse positional args (non-flag) ────────────────────────────────────────
  const positional = args.filter(a => !a.startsWith('--'));
  const fromJson   = args.includes('--from-json');

  // First positional = CSV, second positional = config JSON
  let csvPath    = positional[0];
  let configPath = positional[1];

  if (!csvPath) {
    csvPath = (await ask('Enter path to USN CSV file: ')).trim().replace(/^["']|["']$/g, '');
  }
  if (!fs.existsSync(csvPath)) { console.error('File not found:', csvPath); process.exit(1); }

  const usns = readUSNsFromCSV(csvPath);
  console.log(`Read ${usns.length} USN(s) from CSV.`);

  // ── Load config JSON if provided ─────────────────────────────────────────────
  let cfg = {};
  if (configPath) {
    if (!fs.existsSync(configPath)) { console.error('Config file not found:', configPath); process.exit(1); }
    cfg = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    console.log('Config loaded from', configPath);
  }

  // ── Metadata — use config values, fall back to prompt ────────────────────────
  const collegeName = cfg.collegeName
    || (await ask('College name [SDM Institute of Technology, Ujire]: ')).trim()
    || 'SDM Institute of Technology, Ujire';

  const batchName = cfg.batchName
    || (await ask('Batch/class label [FOURTH YEAR B.E. IS (2024-25)]: ')).trim()
    || 'FOURTH YEAR B.E. IS (2024-25)';

  const examName = cfg.examName
    || (await ask('Exam title [Dec/Jan-2026 Examination Results]: ')).trim()
    || 'Dec/Jan-2026 Examination Results';

  const semLabel = cfg.semLabel
    || (await ask('Sheet label [VII Semester]: ')).trim()
    || 'VII Semester';

  // ── Results ────────────────────────────────────────────────────────────────
  let resultsMap  = {};
  let subjectMeta = [];

  if (fromJson) {
    console.log('\nLoading results from existing JSON files in output/ …');
    ({ resultsMap, subjectMeta } = loadResultsFromJSON(usns));
  } else {
    for (let idx = 0; idx < usns.length; idx++) {
      const usn = usns[idx];
      console.log(`\n[${idx + 1}/${usns.length}] USN: ${usn}`);
      let parsed = null;

      while (!parsed) {
        // Clear cookies and re-fetch index.php for every attempt.
        // This gives the server a clean session, a fresh token, and ensures
        // the CAPTCHA endpoint returns a valid image (not a corrupted one from
        // a consumed/stale PHP session).
        Object.keys(cookieJar).forEach(k => delete cookieJar[k]);
        const token = await getToken();
        const captchaPath = await getCaptcha();
        await openImage(captchaPath);
        const code = (await ask('Enter CAPTCHA code (or "skip" to skip this USN): ')).trim();
        try { fs.unlinkSync(captchaPath); } catch (_) {}

        if (code.toLowerCase() === 'skip') { console.log(`  Skipping ${usn}`); break; }

        const html = await fetchResult(usn, token, code);

        if (/wrong captcha|invalid captcha|captcha error/i.test(html)) {
          console.log('  Wrong CAPTCHA, retrying…'); continue;
        }
        if (/no result|result not found/i.test(html) || !html.includes('divTableRow')) {
          console.log(`  No result found for ${usn} (saved raw HTML).`);
          fs.writeFileSync(path.join(OUT_DIR, `${usn}_raw.html`), html);
          parsed = { usn, name: '', subjects: [] };
          break;
        }

        parsed = parseResult(html);
        fs.writeFileSync(path.join(OUT_DIR, `${usn}_raw.html`), html);
      }

      if (parsed) {
        resultsMap[usn] = parsed;
        if (subjectMeta.length === 0 && parsed.subjects.length > 0) {
          subjectMeta = parsed.subjects.map(s => ({ code: s.code, name: s.name }));
        }
      }
    }
  }

  if (Object.keys(resultsMap).length === 0) {
    console.log('No results fetched. Exiting.'); rl.close(); return;
  }

  // ── Subject config — use cfg.subjects if provided, else prompt ───────────────
  let credits = [];

  if (cfg.subjects && cfg.subjects.length > 0) {
    // Config fully defines subject order, names, and credits — no prompts needed
    subjectMeta = cfg.subjects.map(s => ({ code: s.code, name: s.name }));
    credits     = cfg.subjects.map(s => s.credit);
    console.log('\nSubjects loaded from config:');
    subjectMeta.forEach((s, i) => console.log(`  ${i+1}. ${s.code} — ${s.name}  (${credits[i]} credits)`));
  } else {
    // Interactive subject config
    console.log('\n── Subject Configuration ──────────────────────────────');
    if (subjectMeta.length === 0) {
      console.log('Could not auto-detect subjects. Enter manually:');
      const n = parseInt(await ask('Number of subjects: '));
      for (let i = 0; i < n; i++) {
        const code = (await ask(`  Subject ${i+1} code: `)).trim();
        const name = (await ask(`  Subject ${i+1} name: `)).trim();
        subjectMeta.push({ code, name });
      }
    } else {
      console.log('Detected subjects:');
      subjectMeta.forEach((s, i) => console.log(`  ${i+1}. ${s.code} — ${s.name}`));
      const reorder = (await ask('Change subject order/count? [y/N]: ')).trim().toLowerCase();
      if (reorder === 'y') {
        const indices = (await ask('Enter 1-based indices in order (e.g. 1,2,3,4,5,6): '))
          .split(',').map(s => parseInt(s.trim()) - 1);
        subjectMeta = indices.map(i => subjectMeta[i]);
      }
    }
    for (const s of subjectMeta) {
      const c = parseInt(await ask(`Credits for ${s.code} (${s.name}): `));
      credits.push(isNaN(c) ? 4 : c);
    }
  }

  // ── Generate Excel ────────────────────────────────────────────────────────
  await generateExcel(resultsMap, subjectMeta, credits, collegeName, batchName, examName, semLabel);

  rl.close();
}

main().catch(err => { console.error(err); rl.close(); process.exit(1); });
