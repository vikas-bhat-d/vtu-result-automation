'use strict';

/**
 * VTU Result Automation — Web Server
 * Serves the UI and provides a REST API for:
 *  - Config management
 *  - Session/captcha flow
 *  - Result caching (configurable TTL)
 *  - Excel generation + download
 *  - Cache management
 *  - Log access
 */

const express   = require('express');
const multer    = require('multer');
const axios     = require('axios');
const cheerio   = require('cheerio');
const https     = require('https');
const ExcelJS   = require('exceljs');
const fs        = require('fs');
const path      = require('path');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });

// ── Directories ────────────────────────────────────────────────────────────────
const CONFIG_PATH        = path.join(__dirname, 'config.json');
const SYSTEM_CONFIG_PATH = path.join(__dirname, 'system.config.json');

function readSystemConfig() {
  try { return JSON.parse(fs.readFileSync(SYSTEM_CONFIG_PATH, 'utf8')); } catch { return {}; }
}

// ── Directories (from system config, with defaults) ────────────────────────
const _sysDirs   = (readSystemConfig().dirs || {});
const CACHE_DIR  = path.resolve(__dirname, _sysDirs.cache  || 'cache');
const LOG_DIR    = path.resolve(__dirname, _sysDirs.logs   || 'logs');
const OUT_DIR    = path.resolve(__dirname, _sysDirs.output || 'output');
const PUBLIC_DIR = path.join(__dirname, 'public');

[CACHE_DIR, LOG_DIR, OUT_DIR, PUBLIC_DIR].forEach(d => {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

// ── Middleware ─────────────────────────────────────────────────────────────────
app.use(express.json({ limit: '2mb' }));
app.use(express.static(PUBLIC_DIR));

// ── Logger ─────────────────────────────────────────────────────────────────────
function log(level, msg, data) {
  const now     = new Date();
  const dateStr = now.toISOString().slice(0, 10);
  const entry   = { t: now.toISOString(), level, msg, ...(data != null ? { data } : {}) };
  try {
    fs.appendFileSync(path.join(LOG_DIR, `${dateStr}.log`), JSON.stringify(entry) + '\n');
  } catch { /* non-fatal */ }
  const colors = { info: '\x1b[36m', warn: '\x1b[33m', error: '\x1b[31m', debug: '\x1b[90m' };
  const c = colors[level] || '';
  console.log(`${c}[${level.toUpperCase().padEnd(5)}]\x1b[0m ${msg}`, data != null ? data : '');
}

// ── Config ─────────────────────────────────────────────────────────────────────
function readConfig() {
  try { return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')); } catch { return {}; }
}
function writeConfig(cfg) {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2), 'utf8');
}

// ── Cache ──────────────────────────────────────────────────────────────────────
function getCacheTtlMs() {
  return (readSystemConfig().cache?.ttlHours || 24) * 60 * 60 * 1000;
}

// Cache filename: USN_EXAMCODE.json so results from different exams don't collide
function cacheKey(usn, examPath) {
  const safeExam = (examPath || 'default').replace(/[^a-zA-Z0-9]/g, '');
  return `${usn.toUpperCase()}_${safeExam}`;
}

function cacheGet(usn, examPath) {
  const fp = path.join(CACHE_DIR, `${cacheKey(usn, examPath)}.json`);
  if (!fs.existsSync(fp)) return null;
  try {
    const cached = JSON.parse(fs.readFileSync(fp, 'utf8'));
    if (Date.now() - new Date(cached.fetchedAt).getTime() < getCacheTtlMs()) {
      return cached.data;
    }
    fs.unlinkSync(fp); // expired
  } catch { /* corrupt */ }
  return null;
}

function cacheSet(usn, examPath, data) {
  const fp = path.join(CACHE_DIR, `${cacheKey(usn, examPath)}.json`);
  fs.writeFileSync(fp, JSON.stringify({ fetchedAt: new Date().toISOString(), data }, null, 2), 'utf8');
  log('debug', `Cached: ${usn} [${examPath}]`);
}

function cacheClearOne(usn) {
  // clear any exam variant for this USN
  const prefix = `${usn.toUpperCase()}_`;
  fs.readdirSync(CACHE_DIR)
    .filter(f => f.endsWith('.json') && f.toUpperCase().startsWith(prefix))
    .forEach(f => fs.unlinkSync(path.join(CACHE_DIR, f)));
}

function cacheClearAll() {
  fs.readdirSync(CACHE_DIR)
    .filter(f => f.endsWith('.json'))
    .forEach(f => fs.unlinkSync(path.join(CACHE_DIR, f)));
}

function cacheStatus() {
  const ttl = getCacheTtlMs();
  return fs.readdirSync(CACHE_DIR)
    .filter(f => f.endsWith('.json'))
    .map(f => {
      const basename = f.replace('.json', '');
      const idx      = basename.indexOf('_');
      const usn      = idx > 0 ? basename.slice(0, idx) : basename;
      const examCode = idx > 0 ? basename.slice(idx + 1) : '';
      const fp  = path.join(CACHE_DIR, f);
      try {
        const c       = JSON.parse(fs.readFileSync(fp, 'utf8'));
        const age     = Date.now() - new Date(c.fetchedAt).getTime();
        const expired = age >= ttl;
        const expiresInMin = expired ? 0 : Math.round((ttl - age) / 60000);
        return { usn, examCode, fetchedAt: c.fetchedAt, expired, expiresInMin };
      } catch {
        return { usn, examCode, corrupt: true };
      }
    })
    .sort((a, b) => a.usn.localeCompare(b.usn));
}

// ── VTU Scraper ────────────────────────────────────────────────────────────────
const BASE_URL = 'https://results.vtu.ac.in';

function createVtuClient(examPath) {
  const cookieJar = {};

  const client = axios.create({
    baseURL   : BASE_URL,
    httpsAgent: new https.Agent({ rejectUnauthorized: false }),
    headers   : {
      'User-Agent'        : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36',
      'Accept-Language'   : 'en-US,en;q=0.9',
      'Referer'           : `${BASE_URL}/${examPath}/index.php`,
      'sec-ch-ua'         : '"Chromium";v="148", "Brave";v="148", "Not/A)Brand";v="99"',
      'sec-ch-ua-mobile'  : '?0',
      'sec-ch-ua-platform': '"Windows"'
    },
    withCredentials: true,
    maxRedirects   : 5,
    timeout        : 15000
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

  return client;
}

async function fetchToken(client, examPath) {
  const { data: html } = await client.get(`/${examPath}/index.php`);
  const patterns = [
    /name=["']Token["'][^>]*value=["']([a-f0-9]{20,})["']/i,
    /value=["']([a-f0-9]{20,})["'][^>]*name=["']Token["']/i,
    /["']Token["']\s*[=:]\s*["']([a-f0-9]{20,})["']/i,
    /var\s+[Tt]oken\s*=\s*["']([a-f0-9]{20,})["']/,
    /[Tt]oken[=\s:]+([a-f0-9]{20,})/
  ];
  for (const p of patterns) {
    const m = html.match(p);
    if (m) return m[1];
  }
  throw new Error('Token not found. VTU portal may be down or the exam path is incorrect.');
}

async function fetchCaptchaBase64(client) {
  const t = `${Math.random().toFixed(8)} ${Math.floor(Date.now() / 1000)}`;
  const { data: imgBuf } = await client.get(
    `/captcha/vtu_captcha.php?_CAPTCHA&t=${encodeURIComponent(t)}`,
    {
      responseType: 'arraybuffer',
      headers: {
        'Accept'        : 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
        'Sec-Fetch-Dest': 'image',
        'Sec-Fetch-Mode': 'no-cors',
        'Sec-Fetch-Site': 'same-origin'
      }
    }
  );
  return `data:image/png;base64,${Buffer.from(imgBuf).toString('base64')}`;
}

async function submitCaptcha(client, examPath, usn, token, captchaCode) {
  const { data: html } = await client.post(
    `/${examPath}/resultpage.php`,
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

  const semText  = $('div b').filter((_, el) => /Semester\s*:/i.test($(el).text())).first().text();
  const semester = parseInt((semText.match(/\d+/) || [])[0]) || null;

  const subjects = [];
  $('.divTableBody .divTableRow').each((_, row) => {
    const cells = $(row).find('.divTableCell').map((_, td) => $(td).text().trim()).get();
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

// ── In-memory session ──────────────────────────────────────────────────────────
let session = {
  status    : 'idle',     // idle | running | captcha-wait | done | error
  usns      : [],
  examPath  : '',
  currentIdx: 0,
  results   : {},         // USN → parsed result
  skipped   : [],         // USNs with no result
  errors    : [],         // { usn, reason }
  pending   : null        // { usn, token, captchaImg, client } — awaiting captcha input
};

function sessionProgress() {
  const total   = session.usns.length;
  const done    = Object.keys(session.results).length + session.skipped.length + session.errors.length;
  return {
    total,
    done,
    remaining : total - done,
    currentIdx: session.currentIdx,
    usns      : session.usns,
    results   : Object.fromEntries(
      Object.entries(session.results).map(([u, r]) => [u, { name: r.name, subjects: r.subjects.length }])
    ),
    skipped   : session.skipped,
    errors    : session.errors
  };
}

// Retry wrapper — retries on network errors and HTTP 5xx, not on logical errors
async function withRetry(fn, maxAttempts = 3, baseDelay = 1500) {
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      const status = err.response?.status;
      const retryable = !status || status >= 500;
      if (!retryable || attempt === maxAttempts) throw err;
      const delay = baseDelay * attempt;
      log('warn', `VTU request failed (attempt ${attempt}/${maxAttempts}), retrying in ${delay}ms — ${err.message}`);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
  throw lastErr;
}

async function prepareCaptcha(usn) {
  log('info', `Preparing captcha for ${usn}`);
  const cfg      = readConfig();
  const examPath = session.examPath || cfg.examPath;
  if (!examPath) throw new Error('examPath not set. Configure it in Setup.');

  const client     = createVtuClient(examPath);
  const token      = await withRetry(() => fetchToken(client, examPath));
  const captchaImg = await withRetry(() => fetchCaptchaBase64(client));

  session.pending = { usn, token, captchaImg, client };
  session.status  = 'captcha-wait';
  log('debug', `Captcha ready for ${usn}`);
}

async function advanceSession() {
  while (session.currentIdx < session.usns.length) {
    const usn    = session.usns[session.currentIdx];
    const cached = cacheGet(usn, session.examPath);
    if (cached) {
      log('info', `Cache hit: ${usn} — ${cached.name}`);
      session.results[usn.toUpperCase()] = cached;
      session.currentIdx++;
      continue;
    }
    // Needs captcha
    await prepareCaptcha(usn);
    return;
  }
  // All done
  session.status  = 'done';
  session.pending = null;
  log('info', 'Session complete', {
    results: Object.keys(session.results).length,
    skipped: session.skipped.length,
    errors : session.errors.length
  });
}

// ── Startup Cleanup ──────────────────────────────────────────────────────────
function runCleanup() {
  const sys     = readSystemConfig();
  const cleanup = sys.cleanup || {};
  const now     = Date.now();

  // Delete output files older than outputMaxAgeDays (default 2)
  const outputMaxAgeMs = (cleanup.outputMaxAgeDays ?? 2) * 24 * 60 * 60 * 1000;
  try {
    fs.readdirSync(OUT_DIR).forEach(f => {
      const fp = path.join(OUT_DIR, f);
      try {
        if (fs.statSync(fp).isFile() && (now - fs.statSync(fp).mtimeMs) > outputMaxAgeMs) {
          fs.unlinkSync(fp);
          log('info', `Cleanup: removed old output file: ${f}`);
        }
      } catch { /* skip */ }
    });
  } catch { /* dir may not exist yet */ }

  // Delete log files older than logsMaxAgeDays (default 3)
  const logsMaxAgeMs = (cleanup.logsMaxAgeDays ?? 3) * 24 * 60 * 60 * 1000;
  try {
    fs.readdirSync(LOG_DIR).forEach(f => {
      const fp = path.join(LOG_DIR, f);
      try {
        if (fs.statSync(fp).isFile() && (now - fs.statSync(fp).mtimeMs) > logsMaxAgeMs) {
          fs.unlinkSync(fp);
          log('info', `Cleanup: removed old log file: ${f}`);
        }
      } catch { /* skip */ }
    });
  } catch { /* dir may not exist yet */ }

  // Delete expired cache entries
  const ttlMs = getCacheTtlMs();
  try {
    fs.readdirSync(CACHE_DIR).filter(f => f.endsWith('.json')).forEach(f => {
      const fp = path.join(CACHE_DIR, f);
      try {
        const cached = JSON.parse(fs.readFileSync(fp, 'utf8'));
        if (Date.now() - new Date(cached.fetchedAt).getTime() >= ttlMs) {
          fs.unlinkSync(fp);
          log('info', `Cleanup: removed expired cache: ${f}`);
        }
      } catch {
        fs.unlinkSync(fp);
        log('warn', `Cleanup: removed corrupt cache: ${f}`);
      }
    });
  } catch { /* dir may not exist yet */ }
}

// ── API Routes ─────────────────────────────────────────────────────────────────

// Config
app.get('/api/config', (_req, res) => res.json(readConfig()));

app.post('/api/config', (req, res) => {
  try {
    const current = readConfig();
    const updated = { ...current, ...req.body };
    writeConfig(updated);
    log('info', 'Config saved');
    res.json({ ok: true, config: updated });
  } catch (err) {
    log('error', 'Config save failed', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Upload CSV → USN list
app.post('/api/upload/csv', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const text = req.file.buffer.toString('utf8');
  // Accept comma, newline, semicolon delimiters; each token must look like a USN
  const usns = text
    .split(/[\r\n,;]+/)
    .map(s => s.trim().toUpperCase())
    .filter(s => /^[A-Z0-9]{7,15}$/.test(s));
  if (!usns.length) return res.status(400).json({ error: 'No valid USNs found in CSV' });
  log('info', `CSV parsed: ${usns.length} USNs`);
  res.json({ usns });
});

// Session start
app.post('/api/session/start', async (req, res) => {
  const { usns, examPath } = req.body;
  if (!Array.isArray(usns) || !usns.length) return res.status(400).json({ error: 'No USNs provided' });
  if (!examPath) return res.status(400).json({ error: 'examPath is required' });

  const deduped = [...new Set(usns.map(u => u.trim().toUpperCase()).filter(Boolean))];

  session = {
    status    : 'running',
    usns      : deduped,
    examPath  : examPath.trim(),
    currentIdx: 0,
    results   : {},
    skipped   : [],
    errors    : [],
    pending   : null
  };

  log('info', `Session started: ${deduped.length} USNs`, { examPath });

  try {
    await advanceSession();
    res.json({
      status    : session.status,
      progress  : sessionProgress(),
      captchaImg: session.pending?.captchaImg || null,
      currentUSN: session.pending?.usn || null
    });
  } catch (err) {
    session.status = 'error';
    log('error', 'Session start failed', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Session status
app.get('/api/session/status', (_req, res) => {
  res.json({
    status    : session.status,
    progress  : sessionProgress(),
    captchaImg: session.pending?.captchaImg || null,
    currentUSN: session.pending?.usn || null
  });
});

// Get new captcha image for current USN (retry)
app.post('/api/captcha/retry', async (req, res) => {
  if (session.status !== 'captcha-wait' || !session.pending) {
    return res.status(400).json({ error: 'No captcha pending' });
  }
  const usn = session.pending.usn;
  try {
    log('info', `Captcha retry: ${usn}`);
    await prepareCaptcha(usn);
    res.json({ captchaImg: session.pending.captchaImg, currentUSN: usn });
  } catch (err) {
    log('error', `Captcha retry failed for ${usn}`, err.message);
    res.status(500).json({ error: err.message });
  }
});

// Skip current USN
app.post('/api/captcha/skip', async (req, res) => {
  if (session.status !== 'captcha-wait' || !session.pending) {
    return res.status(400).json({ error: 'No captcha pending' });
  }
  const usn = session.pending.usn;
  log('warn', `Skipped: ${usn}`);
  session.skipped.push(usn);
  session.currentIdx++;
  session.pending = null;
  session.status  = 'running';

  try {
    await advanceSession();
    res.json({
      status    : session.status,
      progress  : sessionProgress(),
      captchaImg: session.pending?.captchaImg || null,
      currentUSN: session.pending?.usn || null,
      skipped   : usn
    });
  } catch (err) {
    session.status = 'error';
    log('error', 'Advance failed after skip', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Submit captcha code
app.post('/api/captcha/submit', async (req, res) => {
  const { captchaCode } = req.body;
  if (!captchaCode || !captchaCode.trim()) {
    return res.status(400).json({ error: 'captchaCode is required' });
  }
  if (session.status !== 'captcha-wait' || !session.pending) {
    return res.status(400).json({ error: 'No captcha pending' });
  }

  const { usn, token, client } = session.pending;
  const examPath = session.examPath;

  log('info', `Submitting captcha for ${usn}`);

  try {
    const html = await withRetry(() => submitCaptcha(client, examPath, usn, token, captchaCode.trim()));

    // Always save raw HTML for debugging
    const ts      = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const rawPath = path.join(OUT_DIR, `${usn}_${ts}_raw.html`);
    fs.writeFileSync(rawPath, html, 'utf8');
    // Save raw JSON after successful parse (written later, path captured now)
    const rawJsonPath = path.join(OUT_DIR, `${usn}_${ts}_raw.json`);

    // Wrong captcha
    if (/invalid\s*captcha|wrong\s*captcha/i.test(html)) {
      log('warn', `Wrong captcha for ${usn} — retrying`);
      await prepareCaptcha(usn); // fetches fresh token + captcha
      return res.json({
        status    : 'wrong-captcha',
        captchaImg: session.pending.captchaImg,
        currentUSN: usn
      });
    }

    // USN not found
    if (/invalid\s*usn|no\s*result\s*found/i.test(html)) {
      log('warn', `No result for ${usn}`);
      session.skipped.push(usn);
      session.currentIdx++;
      session.pending = null;
      session.status  = 'running';
      await advanceSession();
      return res.json({
        status    : session.status,
        progress  : sessionProgress(),
        captchaImg: session.pending?.captchaImg || null,
        currentUSN: session.pending?.usn || null,
        event     : { type: 'no-result', usn }
      });
    }

    // Parse and cache
    const parsed = parseResult(html);
    if (!parsed.usn && !parsed.name) {
      // Unexpected response — treat as error but keep going
      log('warn', `Unexpected response for ${usn} — could not parse result`);
      session.errors.push({ usn, reason: 'Could not parse result page' });
      session.currentIdx++;
      session.pending = null;
      session.status  = 'running';
      await advanceSession();
      return res.json({
        status    : session.status,
        progress  : sessionProgress(),
        captchaImg: session.pending?.captchaImg || null,
        currentUSN: session.pending?.usn || null,
        event     : { type: 'parse-error', usn }
      });
    }

    cacheSet(usn, session.examPath, parsed);
    session.results[usn.toUpperCase()] = parsed;
    fs.writeFileSync(rawJsonPath, JSON.stringify(parsed, null, 2), 'utf8');

    log('info', `Result: ${usn} → ${parsed.name} (${parsed.subjects.length} subjects)`);

    session.currentIdx++;
    session.pending = null;
    session.status  = 'running';
    await advanceSession();

    res.json({
      status    : session.status,
      progress  : sessionProgress(),
      captchaImg: session.pending?.captchaImg || null,
      currentUSN: session.pending?.usn || null,
      event     : { type: 'result', usn, name: parsed.name, subjects: parsed.subjects.length }
    });

  } catch (err) {
    log('error', `Submit error for ${usn}`, err.message);
    // Mark as error, try to move on
    session.errors.push({ usn, reason: err.message });
    session.currentIdx++;
    session.pending = null;
    session.status  = 'running';
    try { await advanceSession(); } catch { session.status = 'error'; }
    res.status(500).json({ error: err.message });
  }
});

// ── Excel column letter helper (1-based) ──────────────────────────────────────
function colL(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ── Grade helpers (used when VTU result overrides formula) ───────────────────
function gradePointFromTotal(total, isProject) {
  if (typeof total !== 'number' || isNaN(total)) return 0;
  if (isProject) {
    if (total >= 180) return 10;
    if (total >= 160) return 9;
    if (total >= 140) return 8;
    if (total >= 120) return 7;
    if (total >= 110) return 6;
    if (total >= 100) return 5;
    if (total >= 80)  return 4;
    return 0;
  }
  if (total >= 90) return 10;
  if (total >= 80) return 9;
  if (total >= 70) return 8;
  if (total >= 60) return 7;
  if (total >= 55) return 6;
  if (total >= 50) return 5;
  if (total >= 40) return 4;
  return 0;
}
function gradeFromPoint(gp) {
  return ({10:'O',9:'A+',8:'A',7:'B+',6:'B',5:'C',4:'P',0:'F'})[gp] ?? 'F';
}

// ── Generate Excel from scratch ────────────────────────────────────────────────
async function generateExcel(resultsMap, subjects, credits, collegeName, batchName, examName, semLabel, examPath) {
  const N = subjects.length;
  const CIE_COL  = i => 4 + i * 5;
  const SEE_COL  = i => 5 + i * 5;
  const TOT_COL  = i => 6 + i * 5;
  const GRP_COL  = i => 7 + i * 5;
  const GRD_COL  = i => 8 + i * 5;
  const SGPA_COL  = 4 + N * 5;
  const PCT_COL   = SGPA_COL + 1;
  const RES_COL   = SGPA_COL + 2;
  const PEND_COL  = SGPA_COL + 3;
  const PENDL_COL = SGPA_COL + 4;
  const PASA_COL  = SGPA_COL + 5;
  const TOTAL_COLS = PASA_COL;
  const totalCredits = credits.reduce((a, c) => a + c, 0);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'VTU Result Automation';
  const ws = wb.addWorksheet(semLabel + ' Marks', { pageSetup: { fitToPage: true } });
  ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 0, topLeftCell: 'D1' }];

  ws.getColumn(1).width = 10;
  ws.getColumn(2).width = 14.63;
  ws.getColumn(3).width = 35.63;
  for (let i = 0; i < N; i++) {
    for (let c = 0; c < 5; c++) ws.getColumn(CIE_COL(i) + c).width = 5.75;
  }
  ws.getColumn(SGPA_COL).width  = 9.38;
  ws.getColumn(PCT_COL).width   = 6.38;
  ws.getColumn(RES_COL).width   = 7.38;
  ws.getColumn(PEND_COL).width  = 9.75;
  ws.getColumn(PENDL_COL).width = 37.63;
  ws.getColumn(PASA_COL).width  = 37.63;

  const fillPurple = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF351C75' } };
  const fillPink   = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFE6B8AF' } };
  const fillYellow = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF2CC' } };
  const fillOrange = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFCE5CD' } };
  const fillGreen  = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFD9EAD3' } };
  const fillPink2  = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFEAD1DC' } };
  const fillBlue   = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFC9DAF8' } };
  const thin   = { style:'thin' };
  const border = { top:thin, left:thin, bottom:thin, right:thin };

  function styleCell(cell, opts = {}) {
    const { fill, bold, size = 12, color, halign = 'center', valign = 'middle', wrap = false } = opts;
    if (fill) cell.fill = fill;
    cell.font = { bold: !!bold, size, color: color ? { argb: color } : undefined };
    cell.alignment = { horizontal: halign, vertical: valign, wrapText: wrap };
    cell.border = border;
  }

  // Row 1
  ws.getRow(1).height = 30;
  ws.mergeCells(1, 1, 1, 3);
  const r1 = ws.getCell(1, 1);
  r1.value = collegeName.toUpperCase();
  styleCell(r1, { fill: fillPurple, bold: true, size: 17, color: 'FFFFFFFF' });

  // Row 2
  ws.getRow(2).height = 30;
  ws.mergeCells(2, 1, 2, 3);
  const r2a = ws.getCell(2, 1);
  r2a.value = batchName.toUpperCase();
  styleCell(r2a, { fill: fillPurple, bold: true, size: 14, color: 'FFEDFA4A' });
  ws.mergeCells(2, 4, 2, TOTAL_COLS);
  const r2b = ws.getCell(2, 4);
  r2b.value = examName.toUpperCase();
  styleCell(r2b, { fill: fillPurple, bold: true, size: 14, color: 'FFFFFFFF', halign: 'left', valign: 'middle' });

  // Rows 3-6 fixed cols
  ws.getRow(3).height = 21;
  ws.mergeCells(3, 1, 6, 1);
  ws.mergeCells(3, 2, 6, 2);
  ws.mergeCells(3, 3, 6, 3);
  ['Sl. No.', 'USN', 'Name'].forEach((v, idx) => {
    const cell = ws.getCell(3, idx + 1);
    cell.value = v;
    styleCell(cell, { fill: fillPink, bold: true });
  });

  for (let i = 0; i < N; i++) {
    ws.mergeCells(3, CIE_COL(i), 3, GRD_COL(i));
    const cell = ws.getCell(3, CIE_COL(i));
    cell.value = Array.isArray(subjects[i].code) ? subjects[i].code.join('/') : subjects[i].code;
    styleCell(cell, { fill: fillPink, bold: true });
  }
  ws.mergeCells(3, SGPA_COL, 3, RES_COL);
  styleCell(ws.getCell(3, SGPA_COL), { fill: fillPink, bold: true ,valign:'middle'});
  ws.getCell(3, SGPA_COL).value = 'Results';
  ws.mergeCells(3, PEND_COL, 6, PEND_COL);
  ws.mergeCells(3, PENDL_COL, 6, PENDL_COL);
  ws.mergeCells(3, PASA_COL, 6, PASA_COL);
  ['Number of Subjects Pending', 'List of Pending Subjects', 'PASSED ARREAR Subjects'].forEach((v, idx) => {
    const cell = ws.getCell(3, PEND_COL + idx);
    cell.value = v;
    styleCell(cell, { fill: fillGreen, bold: true, size: 11, wrap: true });
  });

  // Row 4
  ws.getRow(4).height = 15.75;
  for (let i = 0; i < N; i++) {
    ws.mergeCells(4, CIE_COL(i), 4, GRD_COL(i));
    const cell = ws.getCell(4, CIE_COL(i));
    cell.value = subjects[i].name;
    styleCell(cell, { fill: fillPink, bold: true });
  }
  for (let c = SGPA_COL; c <= RES_COL; c++) {
    for (let r = 4; r <= 6; r++) ws.getCell(r, c).border = border;
  }

  // Row 5
  ws.getRow(5).height = 15.75;
  for (let i = 0; i < N; i++) {
    const lblCell = ws.getCell(5, TOT_COL(i));
    lblCell.value = 'Credits:';
    styleCell(lblCell, { fill: fillPink, bold: true });
    const valCell = ws.getCell(5, GRP_COL(i));
    valCell.value = credits[i];
    styleCell(valCell, { fill: fillPink, bold: true });
    [CIE_COL(i), SEE_COL(i), GRD_COL(i)].forEach(c => styleCell(ws.getCell(5, c), { fill: fillPink }));
  }

  // Row 6
  ws.getRow(6).height = 15.75;
  ['CIE', 'SEE', 'Total', 'Gr. Pt.', 'Grade'].forEach((h, off) => {
    for (let i = 0; i < N; i++) {
      const cell = ws.getCell(6, CIE_COL(i) + off);
      cell.value = h;
      styleCell(cell, { fill: fillPink, bold: true });
    }
  });
  ['SGPA', '%', 'Result'].forEach((h, off) => {
    const cell = ws.getCell(6, SGPA_COL + off);
    cell.value = h;
    styleCell(cell, { fill: fillPink, bold: true });
  });

  // Data rows
  const usnList = Object.keys(resultsMap);
  const isProject = subjects.map(s => {
    const codes = Array.isArray(s.code) ? s.code : String(s.code).split(',').map(c => c.trim());
    return /project|dissertation/i.test(s.name) || codes.some(c => /project|dissertation/i.test(c));
  });

  for (let idx = 0; idx < usnList.length; idx++) {
    const usn     = usnList[idx];
    const r       = 7 + idx;
    const student = resultsMap[usn];
    ws.getRow(r).height = 26.25;

    const aCell = ws.getCell(r, 1);
    aCell.value = idx + 1;
    styleCell(aCell, { fill: fillPink });
    const bCell = ws.getCell(r, 2);
    bCell.value = student ? (student.usn || usn) : usn;
    styleCell(bCell, { fill: fillPink });
    const cCell = ws.getCell(r, 3);
    cCell.value = student ? student.name : '';
    styleCell(cCell, { fill: fillPink, halign: 'left' });

    for (let i = 0; i < N; i++) {
      const cieLetter = colL(CIE_COL(i));
      const seeLetter = colL(SEE_COL(i));
      const totLetter = colL(TOT_COL(i));
      const grpLetter = colL(GRP_COL(i));
      const grdLetter = colL(GRD_COL(i));
      const cfgCodes  = Array.isArray(subjects[i].code)
        ? subjects[i].code
        : String(subjects[i].code).split(',').map(c => c.trim()).filter(Boolean);
      const subjData  = student && student.subjects ? student.subjects.find(s => cfgCodes.includes(s.code)) : null;
      const fillUsed  = isProject[i] ? fillOrange : fillYellow;

      const cieCell = ws.getCell(r, CIE_COL(i));
      const seeCell = ws.getCell(r, SEE_COL(i));
      cieCell.value = subjData && !isNaN(subjData.internal) ? subjData.internal : null;
      seeCell.value = subjData && !isNaN(subjData.external) ? subjData.external : null;
      cieCell.border = border; cieCell.font = { size: 11 }; cieCell.alignment = { horizontal: 'center', vertical: 'middle' };
      seeCell.border = border; seeCell.font = { size: 11 }; seeCell.alignment = { horizontal: 'center', vertical: 'middle' };

      const totCell = ws.getCell(r, TOT_COL(i));
      totCell.value = { formula: `IF(ISBLANK(${cieLetter}${r}),"",${cieLetter}${r}+${seeLetter}${r})` };
      styleCell(totCell, { fill: fillUsed, size: 11 });

      const grpCell = ws.getCell(r, GRP_COL(i));
      const grpFmla = isProject[i]
        ? `IF(ISBLANK(${cieLetter}${r}),"",IF(OR(${cieLetter}${r}<40,${seeLetter}${r}<35,${totLetter}${r}<80),0,IF(${totLetter}${r}>=180,10,IF(${totLetter}${r}>=160,9,IF(${totLetter}${r}>=140,8,IF(${totLetter}${r}>=120,7,IF(${totLetter}${r}>=110,6,IF(${totLetter}${r}>=100,5,IF(${totLetter}${r}>=80,4,0)))))))))` 
        : `IF(ISBLANK(${cieLetter}${r}),"",IF(OR(${cieLetter}${r}<20,${seeLetter}${r}<18,${totLetter}${r}<40),0,IF(${totLetter}${r}>=90,10,IF(${totLetter}${r}>=80,9,IF(${totLetter}${r}>=70,8,IF(${totLetter}${r}>=60,7,IF(${totLetter}${r}>=55,6,IF(${totLetter}${r}>=50,5,IF(${totLetter}${r}>=40,4,0)))))))))`;

      const grdCell = ws.getCell(r, GRD_COL(i));
      const grdFmla = isProject[i]
        ? `IF(ISBLANK(${cieLetter}${r}),"",IF(OR(${cieLetter}${r}<40,${seeLetter}${r}<35,${totLetter}${r}<80),"F",IF(${totLetter}${r}>=180,"O",IF(${totLetter}${r}>=160,"A+",IF(${totLetter}${r}>=140,"A",IF(${totLetter}${r}>=120,"B+",IF(${totLetter}${r}>=110,"B",IF(${totLetter}${r}>=100,"C","P"))))))))))` 
        : `IF(ISBLANK(${cieLetter}${r}),"",IF(${cieLetter}${r}<20,"NE",IF(OR(${seeLetter}${r}<18,${totLetter}${r}<40),"F",IF(${totLetter}${r}>=90,"O",IF(${totLetter}${r}>=80,"A+",IF(${totLetter}${r}>=70,"A",IF(${totLetter}${r}>=60,"B+",IF(${totLetter}${r}>=55,"B",IF(${totLetter}${r}>=50,"C",IF(${totLetter}${r}>=40,"P","F"))))))))))`;
      // If VTU explicitly reports this subject as passed, use static computed values
      // (ignores SEE minimum — handles subjects with no external exam, e.g. NSS, Project)
      const vtuPass = subjData && /^P$/i.test(String(subjData.result ?? '').trim());
      if (vtuPass) {
        const tot = typeof subjData.total === 'number' && !isNaN(subjData.total)
          ? subjData.total
          : (typeof subjData.internal === 'number' ? subjData.internal : 0);
        const gp = gradePointFromTotal(tot, isProject[i]);
        grpCell.value = gp;
        grdCell.value = gradeFromPoint(gp);
      } else {
        grpCell.value = { formula: grpFmla };
        grdCell.value = { formula: grdFmla };
      }
      styleCell(grpCell, { fill: fillUsed, halign: 'center',valign:"middle", size: 11 });
      styleCell(grdCell, { fill: fillUsed, halign: 'center',valign:"middle", size: 11 });
    }

    const sgpaLetter = colL(SGPA_COL);
    const pctLetter  = colL(PCT_COL);
    const sgpaParts  = subjects.map((_, i) => `${credits[i]}*${colL(GRP_COL(i))}${r}`).join(',');
    const sgpaCell   = ws.getCell(r, SGPA_COL);
    sgpaCell.value = { formula: `(SUM(${sgpaParts})/${totalCredits})` };
    styleCell(sgpaCell, { fill: fillYellow, valign: 'middle', size: 11 });

    const pctCell = ws.getCell(r, PCT_COL);
    pctCell.value = { formula: `${sgpaLetter}${r}*10` };
    styleCell(pctCell, { fill: fillYellow, size: 11 });

    const gradeCols = subjects.map((_, i) => `${colL(GRD_COL(i))}${r}`);
    // Universal: OR across all grade cells — works in Excel 2007+
    const orParts = gradeCols.flatMap(g => [`${g}="F"`, `${g}="NE"`]).join(',');
    const resFmla = `IF(OR(${orParts}),"Fail",IF(${pctLetter}${r}>=70,"FCD",IF(${pctLetter}${r}>=60,"FC",IF(${pctLetter}${r}>=40,"SC","Fail"))))`;
    const resCell = ws.getCell(r, RES_COL);
    resCell.value = { formula: resFmla };
    styleCell(resCell, { fill: fillYellow, size: 11 });

    // Pending count: boolean arithmetic, no COUNTIF array constant needed
    const pendFmla = subjects.map((_, i) => {
      const g = colL(GRD_COL(i));
      return `(${g}${r}="F")+(${g}${r}="NE")`;
    }).join('+');
    const pendCell = ws.getCell(r, PEND_COL);
    pendCell.value = { formula: pendFmla };
    styleCell(pendCell, { fill: fillPink2, size: 11 });

    // List of pending: MID+& concatenation — no TEXTJOIN, works in Excel 2007+
    const pendlParts = subjects.map((_, i) => {
      const grL     = colL(GRD_COL(i));
      const codeRef = `${colL(CIE_COL(i))}3`;   // subject code is in row 3
      return `IF(OR(${grL}${r}="F",${grL}${r}="NE"),", "&${codeRef},"")`;
    });
    // Prefix each failing subject with ", " then MID(result,3,1000) strips the leading ", "
    const pendlCell = ws.getCell(r, PENDL_COL);
    pendlCell.value = { formula: `MID(${pendlParts.join('&')},3,1000)` };
    styleCell(pendlCell, { fill: fillPink2, halign: 'left', wrap: true, size: 11 });

    ws.getCell(r, PASA_COL).border = border;
  }

  // Summary rows
  const D_START = 7;
  const D_END   = 6 + usnList.length;
  const SR      = D_END + 1;
  const ROW_AVG=SR, ROW_NSTU=SR+1, ROW_NE=SR+2, ROW_ABS=SR+3,
        ROW_APP=SR+4, ROW_PASS=SR+5, ROW_FAIL=SR+6,
        ROW_FCD=SR+7, ROW_FC=SR+8, ROW_SC=SR+9,
        ROW_PG=SR+10, ROW_PCT=SR+11;

  const sumLabels = [
    [ROW_AVG,'Average Marks (%)'], [ROW_NSTU,'No. of Students'],
    [ROW_NE,'Not Eligible'], [ROW_ABS,'Absent'],
    [ROW_APP,'Students Appeared'], [ROW_PASS,'Passed'],
    [ROW_FAIL,'Fail'], [ROW_FCD,'FCD'],
    [ROW_FC,'FC'], [ROW_SC,'SC'],
    [ROW_PG,'Pass'], [ROW_PCT,'% Pass Result']
  ];

  function sCell(r, c, val) {
    const cell = ws.getCell(r, c);
    cell.value = (typeof val === 'string' && val.startsWith('=')) ? { formula: val.slice(1) } : val;
    cell.fill = fillBlue; cell.font = { size: 12 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = border;
  }

  sumLabels.forEach(([r, label]) => {
    ws.getRow(r).height = 26.25;
    const lc = ws.getCell(r, 3);
    lc.value = label; lc.fill = fillBlue; lc.font = { size: 12 };
    lc.alignment = { horizontal: 'left', vertical: 'middle' };
    lc.border = border;
  });

  const sgpaL = colL(SGPA_COL);
  const resL  = colL(RES_COL);

  for (let i = 0; i < N; i++) {
    const cL = colL(CIE_COL(i)), sL = colL(SEE_COL(i));
    const tL = colL(TOT_COL(i)), gpL = colL(GRP_COL(i)), grL = colL(GRD_COL(i));
    const dr = col => `${col}${D_START}:${col}${D_END}`;

    sCell(ROW_AVG, CIE_COL(i), `=TEXT(IFERROR(AVERAGE(${dr(cL)})*100/50,""),"0.0")`);
    sCell(ROW_AVG, SEE_COL(i), `=TEXT(IFERROR(AVERAGE(${dr(sL)})*100/50,""),"0.0")`);
    ws.mergeCells(ROW_AVG, TOT_COL(i), ROW_AVG, GRD_COL(i));
    sCell(ROW_AVG, TOT_COL(i), `=TEXT(IFERROR(AVERAGE(${dr(tL)}),""),"0.0")`);

    sCell(ROW_NSTU, GRP_COL(i), `=COUNTA($B$${D_START}:$B$${D_END})`);
    sCell(ROW_NE,   GRP_COL(i), `=COUNTIF(${dr(grL)},"NE")`);
    sCell(ROW_ABS,  GRP_COL(i), `=COUNTIF(${dr(sL)},"0")`);
    sCell(ROW_APP,  GRP_COL(i), `=COUNTA(${dr(grL)})-${gpL}${ROW_NE}-${gpL}${ROW_ABS}`);

    const passF = ['O','A+','A','B+','B','C','P'].map(g => `COUNTIFS(${dr(grL)},"${g}")`).join('+');
    sCell(ROW_PASS, GRP_COL(i), `=${passF}`);
    sCell(ROW_FAIL, GRP_COL(i), `=COUNTIF(${dr(grL)},"F")`);
    sCell(ROW_FCD,  GRP_COL(i), `=COUNTIF(${dr(grL)},"=O")+COUNTIF(${dr(grL)},"=A+")+COUNTIF(${dr(grL)},"=A")`);
    sCell(ROW_FC,   GRP_COL(i), `=COUNTIF(${dr(grL)},"=B+")`);
    sCell(ROW_SC,   GRP_COL(i), `=COUNTIF(${dr(grL)},"=B")+COUNTIF(${dr(grL)},"=C")`);
    sCell(ROW_PG,   GRP_COL(i), `=COUNTIF(${dr(grL)},"=P")`);
    ws.mergeCells(ROW_PCT, GRP_COL(i), ROW_PCT, GRD_COL(i));
    sCell(ROW_PCT, GRP_COL(i), `=TEXT(IFERROR(${gpL}${ROW_PASS}*100/${gpL}${ROW_APP},""),"0.00")`);
  }

  const dro = col => `${col}${D_START}:${col}${D_END}`;
  sCell(ROW_NSTU, SGPA_COL, `=COUNTA($B$${D_START}:$B$${D_END})`);
  sCell(ROW_NE,   SGPA_COL, `=COUNTIF(${dro(resL)},"NE")`);
  sCell(ROW_APP,  SGPA_COL, `=SUM(${sgpaL}${ROW_PASS},${sgpaL}${ROW_FAIL})`);
  sCell(ROW_PASS, SGPA_COL, `=SUM(${sgpaL}${ROW_FCD}:${sgpaL}${ROW_SC})`);
  sCell(ROW_FAIL, SGPA_COL, `=COUNTIFS(${dro(resL)},"Fail")`);
  sCell(ROW_FCD,  SGPA_COL, `=COUNTIFS(${dro(resL)},"FCD")`);
  sCell(ROW_FC,   SGPA_COL, `=COUNTIFS(${dro(resL)},"FC")`);
  sCell(ROW_SC,   SGPA_COL, `=COUNTIFS(${dro(resL)},"SC")`);
  ws.mergeCells(ROW_PCT, SGPA_COL, ROW_PCT, RES_COL);
  sCell(ROW_PCT, SGPA_COL, `=TEXT(IFERROR(${sgpaL}${ROW_PASS}*100/${sgpaL}${ROW_APP},""),"0.00")`);

  // Analytics block: paint all cells, then add structural dividers
  for (let r = SR; r <= ROW_PCT; r++) {
    for (let c = 1; c <= TOTAL_COLS; c++) {
      const cell = ws.getCell(r, c);
      cell.fill = fillBlue;
      if (!cell.border) cell.border = border;
    }
  }
  const medSide  = { style: 'medium' };
  const thickSide = { style: 'thick' };
  function setBorderSide(r, c, overrides) {
    const cell = ws.getCell(r, c);
    cell.border = Object.assign({}, cell.border || border, overrides);
  }
  for (let c = 1; c <= TOTAL_COLS; c++) setBorderSide(SR,      c, { top:    medSide  });
  for (let c = 1; c <= TOTAL_COLS; c++) setBorderSide(ROW_PCT, c, { bottom: thickSide });
  for (let c = 1; c <= TOTAL_COLS; c++) setBorderSide(ROW_AVG, c, { bottom: medSide  });
  for (let r = SR; r <= ROW_PCT; r++)   setBorderSide(r, 3,        { right:  medSide  });
  for (let i = 0; i < N; i++) {
    for (let r = SR; r <= ROW_PCT; r++) setBorderSide(r, GRD_COL(i), { right: medSide });
  }
  for (let r = SR; r <= ROW_PCT; r++) setBorderSide(r, RES_COL, { right: medSide });

  const safeLabel = (semLabel  || 'SEM' ).replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '');
  const safeExam  = (examPath  || 'exam').replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '');
  const ts        = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const outPath   = path.join(OUT_DIR, `${safeLabel}_${safeExam}_result_${ts}.xlsx`);
  await wb.xlsx.writeFile(outPath);
  log('info', `Excel saved: ${outPath}`);
  return outPath;
}

// ── Generate Excel
app.post('/api/excel/generate', async (_req, res) => {
  const count = Object.keys(session.results).length;
  if (!count) return res.status(400).json({ error: 'No results to export. Process some USNs first.' });
  const cfg      = readConfig();
  const subjects = (cfg.subjects || []);
  const credits  = subjects.map(s => s.credit || 4);
  if (!subjects.length) return res.status(400).json({ error: 'No subjects configured. Add subjects in Setup first.' });
  try {
    log('info', `Generating Excel for ${count} students`);
    const outPath = await generateExcel(
      session.results,
      subjects,
      credits,
      cfg.collegeName  || '',
      cfg.batchName    || '',
      cfg.examName     || '',
      cfg.semLabel     || 'Semester',
      session.examPath || cfg.examPath || ''
    );
    res.json({ ok: true, filename: path.basename(outPath) });
  } catch (err) {
    log('error', 'Excel generation failed', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Download latest Excel
app.get('/api/excel/download', (_req, res) => {
  const files = fs.readdirSync(OUT_DIR)
    .filter(f => f.endsWith('.xlsx') && !f.startsWith('~$'))
    .map(f => ({ name: f, mtime: fs.statSync(path.join(OUT_DIR, f)).mtime }))
    .sort((a, b) => b.mtime - a.mtime);

  if (!files.length) return res.status(404).json({ error: 'No Excel file found. Generate one first.' });

  const latest = path.join(OUT_DIR, files[0].name);
  res.download(latest, files[0].name);
});

// Cache
app.get('/api/cache', (_req, res) => {
  res.json({ ttlHours: readSystemConfig().cache?.ttlHours || 24, entries: cacheStatus() });
});

app.delete('/api/cache/all', (_req, res) => {
  cacheClearAll();
  log('info', 'Cache cleared (all)');
  res.json({ ok: true });
});

app.delete('/api/cache/:usn', (req, res) => {
  const usn = req.params.usn.toUpperCase();
  cacheClearOne(usn);
  log('info', `Cache cleared: ${usn}`);
  res.json({ ok: true });
});

// Logs
app.get('/api/logs', (req, res) => {
  const limit  = Math.min(parseInt(req.query.limit) || 150, 500);
  const date   = req.query.date || new Date().toISOString().slice(0, 10);
  const fp     = path.join(LOG_DIR, `${date}.log`);
  if (!fs.existsSync(fp)) return res.json({ entries: [], date });

  const lines   = fs.readFileSync(fp, 'utf8').trim().split('\n').filter(Boolean);
  const entries = lines.slice(-limit).map(l => {
    try { return JSON.parse(l); } catch { return { t: null, level: 'raw', msg: l }; }
  });
  res.json({ entries, date });
});

// Log file listing
app.get('/api/logs/dates', (_req, res) => {
  const dates = fs.readdirSync(LOG_DIR)
    .filter(f => f.endsWith('.log'))
    .map(f => f.replace('.log', ''))
    .sort()
    .reverse();
  res.json({ dates });
});

// ── Start ──────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || readSystemConfig().port || 4000;
app.listen(PORT, () => {
  log('info', `Server running → http://localhost:${PORT}`);
  console.log(`\n  Open  http://localhost:${PORT}  in your browser\n`);
  runCleanup();
});
