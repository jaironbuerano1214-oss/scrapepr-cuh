# Alumni Employment Tracer Automation

Reads a list of alumni names from `list of names.xlsx`, cross-references them against an existing Google Form responses file, and writes a fully formatted `alumni_tracer_output.xlsx` with all tracer questions answered (or marked **"No data found"** if the person has not submitted the form yet).

---

## Files

| File | Purpose |
|------|---------|
| `list of names.xlsx` | Input — alumni names + batch year |
| `College of Computing Studies - Alumni Employment Tracer Form (Responses) (2).xlsx` | Input — existing Google Form responses |
| `alumni_tracer_output.xlsx` | **Output** — generated automatically |
| `alumni_tracer.py` | Main automation script |
| `requirements.txt` | Python dependencies |

---

## Quick Start

### 1. Install dependencies
```
pip install -r requirements.txt
```

### 2. Run the script
```
python alumni_tracer.py
```

The output file `alumni_tracer_output.xlsx` will be created/updated in the same folder.

---

## Output Excel Structure

The output has **two sheets**:

### Sheet 1: `Alumni Tracer Results`
One row per alumni from the names list. Color-coded:

| Row Color | Meaning |
|-----------|---------|
| 🟩 Light green | Data found in existing form responses |
| 🟨 Light yellow | Data found via web search *(if enabled)* |
| 🟥 Light red | **No data found** — person has not submitted the form |

Columns match all tracer form questions:
- Last Name, Given Name, Middle Initial, Year Batch
- Present Address, Contact Number, Program Taken, Year of Graduation
- Employment Status, Self-Employed Business Details
- Pursuing Further Studies, Further Studies Details
- Sector, Employer, Work Location, Employer Classification
- Challenges in Job Search, Skills/Competencies Lacking
- Consent, Data Source

### Sheet 2: `Summary`
Quick count of found vs. not-found alumni.

---

## Enabling Web Search (LinkedIn / Google)

> ⚠️ LinkedIn actively blocks automated scrapers. This feature works best for public Google search snippets.

1. Install the extra packages:
   ```
   pip install selenium webdriver-manager
   ```

2. Open `alumni_tracer.py` and change line 20:
   ```python
   ENABLE_WEB_SCRAPING = True
   ```

3. Run the script again. It will search Google for each "not found" alumni and attempt to extract job/employer info from snippets.

---

## How Name Matching Works

The script normalizes names (removes accents, lowercases, trims whitespace) and requires:
1. **Exact last name match**
2. **First given-name token match** (e.g. "John" from "John Vincent")

This prevents false positives between alumni who share only a last name (e.g., siblings with "Cruz" or "Navarro").

---

## Updating / Re-running

- Add more responses to the responses `.xlsx` file and re-run — the script will pick them up automatically.
- The output file is **overwritten** each run, so any manual edits will be lost.
