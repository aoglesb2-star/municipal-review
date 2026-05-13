# Municipal Code Review Tool

Compares a city's municipal code against the MTAS model and a 142-item standard checklist, then writes a completed Excel review file.

## Setup (one time)

### 1. Install dependencies
```
pip3 install anthropic openpyxl
```

### 2. Set your Anthropic API key
```
export ANTHROPIC_API_KEY=sk-ant-...
```
Add that line to `~/.zshrc` to make it permanent.

### 3. Add your MTAS model files
Copy or paste each title into `mtas/`:
```
mtas/title_1.txt
mtas/title_2.txt
...
mtas/title_20.txt
```
These are reused for every city review — set up once.

---

## Per-review workflow

### 1. Create a folder for the city
```
mkdir -p cities/memphis
```

### 2. Paste each title's code into a text file
```
cities/memphis/title_1.txt
cities/memphis/title_3.txt
...
```
You don't need all 20 — the tool skips titles with no file.
Copy from Word/PDF and paste into a plain .txt file. Formatting doesn't matter.

### 3. Run the review
```
# All titles found in the city folder:
python3 review.py --city Memphis

# Specific titles only:
python3 review.py --city Memphis --titles 1,3,6,11

# Custom output path:
python3 review.py --city Memphis --output ~/Desktop/Memphis_review.xlsx
```

### 4. Open the output
Results are written to `output/Memphis_review_YYYYMMDD.xlsx` with:
- **Review sheet** — all findings with Title, Chapter, Section, Section Title, Section Text, Comment, Source
- **Summary sheet** — finding counts by title

Yellow rows = new findings not in the standard checklist.

---

## Notes

- The standard checklist (142 issues across all 20 titles) is baked into the script — no spreadsheet needed at run time.
- The tool processes one title at a time, so token usage stays low.
- MTAS model files are optional per title; the review still runs without them using only the checklist.
- Typical cost: ~$0.05–0.15 per full 20-title review depending on code length.
