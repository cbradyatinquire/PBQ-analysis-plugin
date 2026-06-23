# PBQ Analysis — Google Sheets Add-on

A Google Sheets sidebar add-on for visualizing student response patterns from **Pattern-Based Question (PBQ)** checkbox questions in Google Forms.

After students submit responses, the sidebar reads the linked form, lets you mark the correct answer(s), and generates a column chart showing all answer-combination frequencies sorted by Hamming distance from the correct answer — most-wrong combinations on the left.

---

## What it does

- **Discovers answer options** directly from the linked Google Form (exact text, correct order, quiz-mode auto-detection) with a fallback to inferring options from response text
- **Encodes responses** as bitmasks (A=bit 0, B=bit 1, C=bit 2, D=bit 3) and counts all 15 non-empty combinations
- **Sorts by Hamming distance** from the correct answer so the chart immediately highlights the most common misconceptions
- **Persists per-tab form links** — in a workbook with one form per class session, each tab remembers its own form URL permanently

---

## Setup (for a new spreadsheet)

1. Open your Google Sheet containing Form response data
2. Go to **Extensions → Apps Script**
3. Copy in `PBQCode.gs`, `Sidebar.html`, and `appsscript.json` from this repo
4. In the Apps Script editor, enable **Google Sheets API** under Services (Advanced Services)
5. Save and reload the spreadsheet — a **PBQ Analysis** menu will appear

---

## Usage

1. **PBQ Analysis → Open PBQ Panel** to open the sidebar
2. For each response tab, click **Link form** and paste the Google Form URL — this is stored permanently per tab
3. Select the PBQ question column from the dropdown
4. Verify the answer options (auto-filled from the form)
5. Select the correct answer(s) (A/B/C/D tiles)
6. Click **Generate Chart**

When switching tabs, click **↺ Refresh for this tab** to reload the sidebar for the new tab's form and columns.

---

## Development workflow (clasp)

This project uses [clasp](https://github.com/google/clasp) to sync between local files and the Apps Script project.

```bash
npm install -g @google/clasp
clasp login
```

Create `.clasp.json` in the project root (not committed — see `.gitignore`):
```json
{
  "scriptId": "<your-script-id>",
  "rootDir": "."
}
```

Your script ID is in the Apps Script editor under **Project Settings → Script ID**.

```bash
clasp push          # push local changes to Apps Script
clasp push --watch  # auto-push on every file save
clasp pull          # pull changes made in the Apps Script editor
```

---

## Files

| File | Description |
|------|-------------|
| `PBQCode.gs` | All server-side logic: column detection, FormApp option discovery, bitmask encoding, chart generation, PropertiesService tracking |
| `Sidebar.html` | Full sidebar UI — HTML, CSS, and client-side JS |
| `appsscript.json` | Apps Script manifest: OAuth scopes, Google Sheets Advanced API |

---

## Requirements

- Google Sheets with one or more Google Forms routing responses to tabs
- The account running the add-on must own (or have editor access to) the linked Google Forms

---

## Legal

[Privacy Policy](PRIVACY.md) | [Terms of Service](TERMS.md)
