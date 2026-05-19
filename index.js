/**
 * VTU Result Automation Script
 * -------------------------------------------------
 * 1. Asks for USN
 * 2. Fetches Token from index.php (auto-parsed or manual)
 * 3. Generates t= param  →  Math.random().toFixed(8) + " " + unixSeconds
 * 4. Downloads CAPTCHA PNG, saves it to ./output/, opens it
 * 5. Accepts CAPTCHA code from user
 * 6. POSTs to resultpage.php
 * 7. Parses HTML → structured JSON (subjects only)
 * 8. Saves result as .json + raw .html in ./output/
 */

'use strict';

const axios    = require('axios');
const cheerio  = require('cheerio');
const https    = require('https');
const fs       = require('fs');
const path     = require('path');
const readline = require('readline');
const { exec } = require('child_process');
const { fillExcel } = require('./fill-excel');

// ── Configuration ──────────────────────────────────────────────────────────────
const EXAM_PATH = 'D25J26Ecbcs';          // ← change per exam (visible in URL)
const BASE_HOST = 'results.vtu.ac.in';
const BASE_URL  = `https://${BASE_HOST}`;
const OUT_DIR   = path.join(__dirname, 'output');

if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

// ── Axios instance ─────────────────────────────────────────────────────────────
// VTU server has a broken cert chain; rejectUnauthorized: false is intentional —
// we are targeting a known government result portal only.
const client = axios.create({
  baseURL    : BASE_URL,
  httpsAgent : new https.Agent({ rejectUnauthorized: false }),
  headers    : {
    'User-Agent'        : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36',
    'Accept-Language'   : 'en-US,en;q=0.9',
    'Referer'           : `${BASE_URL}/${EXAM_PATH}/index.php`,
    'sec-ch-ua'         : '"Chromium";v="148", "Brave";v="148", "Not/A)Brand";v="99"',
    'sec-ch-ua-mobile'  : '?0',
    'sec-ch-ua-platform': '"Windows"'
  },
  withCredentials: true,
  maxRedirects   : 5
});

// persist cookies across requests
const cookieJar = {};

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
const ask = (prompt) => new Promise(resolve => rl.question(prompt, resolve));

// ── Step 1 – Token ─────────────────────────────────────────────────────────────
async function getToken() {
  process.stdout.write('Fetching index page for Token … ');

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
  console.log('  Tip: DevTools → Network → resultpage.php → Payload → Token field.');
  return (await ask('  Paste Token value here: ')).trim();
}

// ── Step 2 – CAPTCHA ───────────────────────────────────────────────────────────
// t = "<Math.random().toFixed(8)> <unix-seconds>"
// URL-encoded space becomes + → t=0.75061200+1779202124
function buildT() {
  return `${Math.random().toFixed(8)} ${Math.floor(Date.now() / 1000)}`;
}

async function getCaptcha() {
  const t = buildT();
  process.stdout.write('Fetching CAPTCHA image … ');

  const { data: imgBuffer } = await client.get(
    `/captcha/vtu_captcha.php?_CAPTCHA&t=${encodeURIComponent(t)}`,
    {
      responseType: 'arraybuffer',
      headers     : {
        'Accept'        : 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
        'Sec-Fetch-Dest': 'image',
        'Sec-Fetch-Mode': 'no-cors',
        'Sec-Fetch-Site': 'same-origin'
      }
    }
  );

  console.log('done.');
  const imgPath = path.join(OUT_DIR, `captcha_${Date.now()}.png`);
  fs.writeFileSync(imgPath, imgBuffer);

  const base64 = Buffer.from(imgBuffer).toString('base64');
  console.log(`Saved  : ${imgPath}`);
  console.log(`Base64 : data:image/png;base64,${base64.slice(0, 60)}…`);

  return { imgPath, base64 };
}

// ── Step 3 – Open image ────────────────────────────────────────────────────────
function openImage(imgPath) {
  return new Promise(resolve => {
    exec(`start "" "${imgPath}"`, err => {
      if (err) console.log(`  [!] Could not auto-open. Please open: ${imgPath}`);
      resolve();
    });
  });
}

// ── Step 4 – Fetch result HTML ─────────────────────────────────────────────────
async function fetchResult(usn, token, captchaCode) {
  process.stdout.write('\nSubmitting to resultpage.php … ');

  const { data: html, status } = await client.post(
    `/${EXAM_PATH}/resultpage.php`,
    new URLSearchParams({ Token: token, lns: usn, captchacode: captchaCode }).toString(),
    {
      headers: {
        'Content-Type'  : 'application/x-www-form-urlencoded',
        'Origin'        : BASE_URL,
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin'
      }
    }
  );

  console.log(`HTTP ${status}`);
  return html;
}

// ── Step 5 – Parse HTML → structured JSON ─────────────────────────────────────
function parseResult(html) {
  const $ = cheerio.load(html);

  // USN and Name from the info table at the top
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

  // Semester from the centered bold label
  const semText  = $('div b').filter((_, el) => /Semester\s*:/i.test($(el).text())).first().text();
  const semester = parseInt((semText.match(/\d+/) || [])[0]) || null;

  // Subject rows — VTU uses div-based table (.divTableRow / .divTableCell)
  const subjects = [];
  $('.divTableBody .divTableRow').each((_, row) => {
    const cells = $(row).find('.divTableCell').map((_, td) => $(td).text().trim()).get();

    // Skip the header row ("Subject Code") and the abbreviation legend row
    if (!cells[0] || /subject\s*code/i.test(cells[0]) || /->/.test(cells[0])) return;

    subjects.push({
      code    : cells[0],
      name    : cells[1],
      internal: isNaN(+cells[2]) ? cells[2] : +cells[2],
      external: isNaN(+cells[3]) ? cells[3] : +cells[3],
      total   : isNaN(+cells[4]) ? cells[4] : +cells[4],
      result  : cells[5],
      date    : cells[6]
    });
  });

  return { usn, name, semester, subjects };
}

// ── Step 6 – Save ─────────────────────────────────────────────────────────────
function saveResult(html, parsed, usn) {
  const ts       = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const jsonPath = path.join(OUT_DIR, `${usn}_${ts}.json`);

  fs.writeFileSync(jsonPath, JSON.stringify(parsed, null, 2), 'utf8');

  return { jsonPath };
}

// ── Main ───────────────────────────────────────────────────────────────────────
async function main() {
  console.log('╔══════════════════════════════════════════╗');
  console.log('║     VTU Result Automation Script         ║');
  console.log(`║     Exam: ${EXAM_PATH.padEnd(30)}║`);
  console.log('╚══════════════════════════════════════════╝\n');

  try {
    // Accept one or many USNs (comma-separated)
    const usnInput = (await ask('Enter USN(s) — comma-separated for multiple (e.g. 4su22is062,4su22is063): ')).trim();
    if (!usnInput) throw new Error('At least one USN is required.');

    const usnList = usnInput.split(',').map(u => u.trim().toLowerCase()).filter(Boolean);
    console.log(`\n${usnList.length} USN(s) queued: ${usnList.join(', ')}`);

    let token = await getToken();
    if (!token) throw new Error('Token is required.');

    // ── Loop through all USNs — fresh captcha per USN, retry on wrong code ──
    const summary    = [];
    const resultsMap = {};   // collected for Excel fill at the end

    for (const usn of usnList) {
      console.log(`\n${'─'.repeat(44)}`);
      console.log(`USN: ${usn.toUpperCase()}  (${usnList.indexOf(usn) + 1}/${usnList.length})`);
      console.log('─'.repeat(44));

      let html = null;

      // Inner loop: retry with fresh captcha until code is accepted
      while (true) {
        // Clear cookies and re-fetch index.php so every attempt gets a clean
        // session, a fresh token, and a valid CAPTCHA image.
        Object.keys(cookieJar).forEach(k => delete cookieJar[k]);
        token = await getToken();
        const { imgPath } = await getCaptcha();
        console.log('\nOpening CAPTCHA image …');
        await openImage(imgPath);

        const captchaCode = (await ask('Enter CAPTCHA code: ')).trim();

        // Delete the captcha image right after the user enters the code
        try { fs.unlinkSync(imgPath); } catch {}

        if (!captchaCode) { console.log('  Empty input — retrying with new CAPTCHA.'); continue; }

        const response = await fetchResult(usn, token, captchaCode);

        // Save raw HTML always
        const ts      = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        const rawPath = path.join(OUT_DIR, `${usn}_${ts}_raw.html`);
        fs.writeFileSync(rawPath, response, 'utf8');
        console.log(`Raw HTML   : ${rawPath}`);

        if (/invalid\s*captcha/i.test(response) || /wrong\s*captcha/i.test(response)) {
          console.log('  Wrong CAPTCHA — fetching a new one…');
          continue;  // retry with fresh captcha
        }

        html = response;
        break;
      }

      try {
        if (/invalid\s*usn/i.test(html) || /no\s*result\s*found/i.test(html)) {
          console.log(`  [SKIP] No result found for ${usn.toUpperCase()}`);
          summary.push({ usn: usn.toUpperCase(), status: 'no result' });
          continue;
        }

        const parsed       = parseResult(html);
        const { jsonPath } = saveResult(html, parsed, usn);

        console.log(`JSON saved : ${jsonPath}`);
        console.log('\n── Subjects ────────────────────────────────');
        parsed.subjects.forEach(s => {
          console.log(`  ${s.code.padEnd(10)} ${s.name.padEnd(35)} ${String(s.total).padStart(3)}  ${s.result}`);
        });

        resultsMap[parsed.usn.toUpperCase()] = parsed;
        summary.push({ usn: usn.toUpperCase(), status: 'ok', subjects: parsed.subjects.length, jsonPath });

      } catch (err) {
        console.error(`  [ERROR] ${err.message}`);
        summary.push({ usn: usn.toUpperCase(), status: 'error', reason: err.message });
      }
    }

    // ── Final summary ────────────────────────────────────────────────────────
    console.log(`\n${'═'.repeat(44)}`);
    console.log('SUMMARY');
    console.log('═'.repeat(44));
    summary.forEach(s => {
      const tag = s.status === 'ok' ? `✓ ${s.subjects} subjects` : `✗ ${s.status}`;
      console.log(`  ${s.usn.padEnd(15)} ${tag}`);
    });

    // ── Fill Excel if we have any successful results ───────────────────────
    if (Object.keys(resultsMap).length > 0) {
      console.log(`\nFilling ${Object.keys(resultsMap).length} result(s) into Excel…`);
      await fillExcel(resultsMap);
    }

  } catch (err) {
    console.error(`\n[ERROR] ${err.message}`);
    process.exit(1);
  } finally {
    rl.close();
  }
}

main();
