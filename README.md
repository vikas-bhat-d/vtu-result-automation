# VTU Result Automation

Automated VTU exam result fetcher with intelligent CAPTCHA solving and Excel report generation. Fetches results from [results.vtu.ac.in](https://results.vtu.ac.in), parses exam data, and generates formatted Excel workbooks with grade calculations.

**Key Features:**
- ✅ Bulk fetch results from CSV (USN list)
- ✅ Automatic CAPTCHA solving via 2captcha/anticaptcha
- ✅ Intelligent HTML parsing (handles layout changes)
- ✅ Dynamic grade/grade-point formulas with per-subject normalization
- ✅ Excel workbooks with styling, SGPA, percentage, results classification
- ✅ Web UI for configuration (no CLI args needed)
- ✅ Result caching (configurable TTL)
- ✅ Works **on Windows** (packaged as `.exe`)

---

## Installation & Usage

Choose one of the three methods below:

### Option 1: Run Directly with Node.js (Development)

**Requirements:** Node.js 18+ and npm

**Steps:**

```bash
# Clone or download the project
git clone https://github.com/your-org/vtu-result-automation.git
cd vtu-result-automation

# Install dependencies
npm install

# Start the server
npm run server
```

Open your browser to **http://localhost:4000** (or the configured port in `system.config.json`).

**Advantages:**
- Fastest iteration during development
- Easy to inspect logs and debug
- No build step needed

---

### Option 2: Build and Run the Standalone Executable

**Requirements:** Node.js 18+ and npm (only for building; NOT needed to run the `.exe`)

**Steps:**

```bash
# Clone or download the project
git clone https://github.com/your-org/vtu-result-automation.git
cd vtu-result-automation

# Install dependencies
npm install

# Build the Windows executable
npm run build
```

The executable will be generated at: **`dist/vtu-result.exe`**

**Run the executable:**

```bash
# Double-click dist/vtu-result.exe
# OR from command line:
.\dist\vtu-result.exe
```

Then open your browser to **http://localhost:4000**.

**Advantages:**
- Single `.exe` file — no Node.js installation needed on client machines
- Portable — copy and run anywhere on Windows
- Production-ready deployment

---

### Option 3: Download from GitHub Release

**Requirements:** None (Windows machine only)

**Steps:**

1. Go to **[GitHub Releases](https://github.com/your-org/vtu-result-automation/releases)**
2. Download the latest `vtu-result.exe` from the release assets
3. Save it anywhere (e.g., `C:\vtu-tools\`)
4. Double-click to run (or run from command line: `vtu-result.exe`)

Open your browser to **http://localhost:4000**.

**Advantages:**
- Quickest setup — no building or cloning required
- No dependencies to install
- Just download and run

---

## Configuration

### `config.json` – Exam & Subject Settings

Located in the **project root** (or in `dist/` if running the `.exe`).

**Structure:**

```json
{
  "examPath": "MJ26cbcs",
  "collegeName": "SDM Institute of Technology, Ujire",
  "batchName": "FOURTH YEAR B.E. IS (2025-2026)",
  "examName": "MAY/JUNE-2026 Examination Results",
  "semLabel": "VIII Semester",
  "subjects": [
    {
      "code": "BIS801",
      "name": "PROFESSIONAL ELECTIVE COURSE",
      "credit": 3
    },
    {
      "code": "BINT803B",
      "name": "INDUSTRY INTERNSHIP",
      "credit": 10,
      "cie": 100,
      "see": 100
    }
  ]
}
```

**Fields:**
- `examPath` – VTU exam identifier (e.g., `MJ26cbcs`, `JJEcbcs25`)
  - Find this in the URL: `results.vtu.ac.in/**examPath**/index.php`
- `collegeName`, `batchName`, `examName`, `semLabel` – Metadata for the Excel report
- `subjects[].code`, `.name`, `.credit` – Course details
- `subjects[].cie`, `.see` *(optional)* – Max CIE/SEE marks (default: 50/50 each)

**Demo config:** See `config.demo.json` for reference.

### `system.config.json` – System & Server Settings

Located next to the `.exe` or in the project root.

**Structure:**

```json
{
  "port": 4000,
  "dirs": {
    "output": "output",
    "logs": "logs",
    "cache": "cache"
  },
  "cache": {
    "ttlHours": 24
  },
  "cleanup": {
    "outputMaxAgeDays": 2,
    "logsMaxAgeDays": 3
  }
}
```

**Fields:**
- `port` – HTTP server port (default: 4000)
- `dirs` – Output directories for Excel files, logs, and cached results
- `cache.ttlHours` – How long to keep cached result JSON (default: 24 hours)
- `cleanup.*` – Auto-delete old files (set to 0 to disable)

---

## Web UI Workflow

1. **Step 1 – Configure Exam**
   - Upload or manually enter exam metadata (college, batch, exam, semester)
   - Add subjects with code, name, credits, and optional CIE/SEE max marks
   - Click **Save & Continue**

2. **Step 2 – Upload CAPTCHA Credentials**
   - Choose a CAPTCHA solver: `2captcha` or `anticaptcha`
   - Enter your API key
   - Upload the CSV with USNs (or drag-and-drop)
   - CSV format: one USN per line (e.g., `4SU22IS044`)

3. **Step 3 – Fetch & Generate**
   - Click **Fetch Results** to start
   - Monitor progress in real-time
   - Generated Excel files appear in `output/`

4. **Download Excel**
   - Each USN gets one Excel workbook: `<usn>_<timestamp>.xlsx`
   - Contains parsed results, grades, SGPA, and overall classification

---

## Grade Calculation

Grades are computed dynamically in Excel using formulas. Each subject normalizes total marks to a 100-point scale based on its configured CIE/SEE max marks.

**Grade Point Scale (normalized):**
- 90–100: **10** (O)
- 80–89: **9** (A+)
- 70–79: **8** (A)
- 60–69: **7** (B+)
- 55–59: **6** (B)
- 50–54: **5** (C)
- 40–49: **4** (P)
- <40: **0** (F)

**Pass Conditions (per subject):**
- CIE ≥ 40% of max CIE
- SEE ≥ 35% of max SEE  
- Total ≥ 40 (normalized)

**Overall Result:**
- **FCD** (First Class Distinction): ≥70% SGPA
- **FC** (First Class): 60–69%
- **SC** (Second Class): 40–59%
- **Fail**: <40% or any subject failed

---

## Output Files

### Excel Workbooks (`output/`)

Generated Excel files are named: `<USN>_<timestamp>.xlsx`

**Contents:**
- Student header (USN, name, batch)
- Subject-wise CIE/SEE/Total/Grade columns
- Grade points and grades (formulas)
- SGPA calculation (weighted by credits)
- Percentage (SGPA × 10)
- Overall result (FCD/FC/SC/Fail)

### Logs (`logs/`)

Server activity logged to `logs/app-<timestamp>.log`

### Cache (`cache/`)

Result JSON cached for `ttlHours` to speed up repeated requests.

---

## Troubleshooting

### Port Already in Use

Change `port` in `system.config.json`:

```json
{ "port": 4001 }
```

Then restart.

### CAPTCHA Solver Errors

- Ensure API key is correct
- Check internet connectivity
- Verify CAPTCHA service has balance/credits

### Excel File Opens with Recovery Warning

Indicates a formula syntax error. Check server logs for details. This should not happen with current version.

### Missing Output Directory

Directories are auto-created on first run. If not:

```bash
# Create manually
mkdir output logs cache
```

### Results Not Fetching

1. Verify `examPath` in `config.json` is correct (check VTU website URL)
2. Check network connectivity
3. Inspect browser console (F12) and server logs for error messages
4. Try a different CAPTCHA solver

---

## Development

### Project Structure

```
vtu-result-automation/
├── server.js               # Express app, scraper, Excel generator
├── index.js                # Legacy CLI (deprecated)
├── generate-excel.js       # Excel workbook generation
├── config.json             # Exam config (edit via UI)
├── config.demo.json        # Example config
├── system.config.json      # System settings
├── package.json
├── public/
│   └── index.html          # Web UI (single-page app)
├── cache/                  # Cached results
├── output/                 # Generated Excel files
├── logs/                   # Server logs
└── dist/
    └── vtu-result.exe      # Packaged executable
```

### Build & Package

```bash
# Build the Windows executable (requires pkg and Node 20)
npm run build

# Output: dist/vtu-result.exe
```

The executable bundles Node.js, public assets, and all dependencies.

---

## License

[Your License Here]

---

## Support

For issues, feature requests, or questions:
- **GitHub Issues:** [your-repo/issues](https://github.com/your-org/vtu-result-automation/issues)
- **Documentation:** See this README and `config.demo.json` for examples
