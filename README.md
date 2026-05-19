# VTU Result Automation

Fetches VTU exam results automatically with CAPTCHA support. **No external dependencies** — pure Node.js built-ins only.

## Usage

```bash
node index.js
```

The script will:
1. Ask for your **USN**
2. Try to auto-extract the **Token** from `index.php`; if it can't, ask you to paste it
3. Download the **CAPTCHA** image → save it to `./output/captcha_<timestamp>.png` → open it automatically
4. Accept the **CAPTCHA code** you type in
5. POST to `resultpage.php` and save the result as `.txt` and `.html` in `./output/`

## Changing the exam

Open `index.js` and update `EXAM_PATH` at the top:

```js
const EXAM_PATH = 'D25J26Ecbcs';   // ← change this for a different exam
```

The exam path is the segment between `results.vtu.ac.in/` and `/index.php` in your browser URL.

## How the `t=` parameter works

```
t = <Math.random().toFixed(8)> <Unix timestamp in seconds>
// e.g.  0.75061200 1779202124
// URL:  t=0.75061200+1779202124   (space encoded as + in query strings)
```

## Output files (`./output/`)

| File | Contents |
|------|----------|
| `captcha_<ts>.png` | CAPTCHA image |
| `<usn>_<ts>.txt`   | Plain-text result (stripped HTML) |
| `<usn>_<ts>.html`  | Raw HTML response |
