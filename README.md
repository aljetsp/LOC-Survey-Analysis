# LOC Survey Analysis Tool

Turns a raw Qualtrics survey export into a finished analysis workbook: cleaned and
coded data, crosstabs by population quintile and region, and a validation sheet —
all with live Excel formulas.

## Install

Python 3.10+ and one dependency:

```bash
pip install openpyxl
```

Put the `survey_tool/` folder and `city_reference.csv` in the same directory.

## Use

```bash
python -m survey_tool.cli build MySurvey_2027.xlsx --reference city_reference.csv
```

It reads the export, writes a codebook, then asks about the three things that
change every year. Afterwards:

```bash
python -m survey_tool.cli recalc output/MySurvey_2027.analysis.xlsx
```

That step populates the cached values via LibreOffice and reports any formula
errors. Opening and re-saving in Excel does the same thing — Excel recalculates
on open. **Until you do one or the other, the formula cells will look empty in
previewers and in pandas**, because a freshly written file has formulas but no
stored results.

To see how columns would be treated without building anything:

```bash
python -m survey_tool.cli profile MySurvey_2027.xlsx
```

To rebuild without prompts, once the reference is already correct:

```bash
python -m survey_tool.cli build MySurvey_2027.xlsx --reference city_reference.csv --no-prompt
```

## What it asks you

**City names.** Reports how many matched exactly, asks you to confirm each
approximate match, and asks where to place names it cannot resolve. Approximate
matches are never applied silently: a wrong match quietly reassigns a city's
quintile and region, and nothing downstream would look wrong.

**Populations.** Point it at a CSV with `city,population` columns — the PSU
Population Research Center certified estimates work directly — or edit individual
cities. Anything you change is written back to the reference file.

**Quintile breakpoints.** Quintiles are derived from population, never stored
separately, so the two cannot drift apart. It shows the resulting distribution and
lists every city sitting within 3% of a boundary, since those are the ones whose
quintile flips year to year.

**Regions.** Asks only for responding cities that have no region on file.

## Output

| Sheet | Contents |
|---|---|
| **Cleaned** | One row per respondent: city, population, quintile, region, then coded answers. Qualtrics metadata and respondent contact details are dropped. |
| **Analysis** | Per question: counts and percentages by quintile and by region. Numeric questions also get n, mean, median, min, max, and group means. |
| **Validation** | Reconciliation checks, each reading OK or FAIL. |
| **Cleaning Log** | Every row dropped, answer blanked, unit stripped, and approximate city match. |

Each range covers the respondents plus 40 blank rows, so next year's data can be
pasted in without editing formulas.

## The codebook

`build` writes `output/<name>.codebook.json` before doing anything else. It maps
each answer label to the integer code used in the crosstabs, and those codes drive
the column order of every table built from that question.

Columns whose note begins with `REVIEW` had codes assigned alphabetically because
the answer scale was not recognised. **Review those before publishing.** Yes/No,
Yes/No/Unsure, and Monthly/Bi-Monthly/Quarterly are recognised and coded to match
the existing workbooks; an "Other (Please Specify)" option is always coded last.

Edit the JSON and re-run `build` to apply it. Setting a column's `kind` to `skip`
drops it entirely. The codebook carries forward, so the review is once per survey,
not once per year.

## Design decisions worth knowing

**Every range derives from two constants.** `DATA_START` and `data_end` in
`build.py` generate every formula range in the output. The criteria-range
misalignment found in the hand-built 2025 workbooks — where the quintile range
started one row below the answer range, so each city was matched against the next
city's answer — cannot occur here.

**Numbers are counted before they are parsed.** An answer with one number is
recoverable (`39 miles` → 39). Two or more is ambiguous and gets left blank and
logged. This matters: stripping punctuation from Portland's answer of
`$39 (20% of SFR customers pay $23/mnth, 60% pay $39/mnth and 20% pay $56/mnth)`
and keeping the digits yields 39,202,360,392,056, which recalculates without
error and puts a $39 trillion mean in a published table. Magnitude words are
applied too, since `13 million` silently read as 13 is a millionfold error that
looks entirely at home in a column of gallon figures.

**Multi-selects become indicator columns.** Each check-all-that-apply option gets
its own 0/1 column, matched on exact tokens. A wildcard `COUNTIFS` would count
"Late Fee" inside "Late Fee and Interest".

**Only discontinuous outliers are flagged** — a value 25× the next largest. A
plain spread rule fires on Portland and Salem in every population column, and a
log that cries wolf catches nothing.

## Verification

Verified against all six 2025/2026 surveys: **89,335 formulas, zero errors**. Every
one of the 26,656 generated `COUNTIFS` cells was recomputed independently in Python
and matched exactly; all `AVERAGEIF` cells matched to within 5×10⁻¹⁵, which is
floating-point display precision.

A clean recalculation only proves formulas *evaluate*. The independent
recomputation is what shows they are *right*, and it is worth re-running after any
change to `build.py`:

```bash
python verify_all.py
```

## Known limits

- The reference holds **200 of Oregon's 241 cities**, drawn from your own
  workbooks. The rest get added through the prompts as they respond.
- The **Q4/Q5 breakpoint moved between the 2025 and 2026 surveys** (Cottage Grove:
  10,879 → Q5 in 2025, 10,909 → Q4 in 2026), so no fixed set of breakpoints
  reproduces every historical file. The default is 500 / 1,400 / 3,300 / 10,700;
  confirm it each year.
- Analysis blocks are **stacked vertically**, one per question, rather than
  arranged horizontally as in the hand-built files. Easier to audit, but it will
  not look like what members are used to.
- **Free-text answers are carried but not tabulated.**
- Duplicate submissions from one city are **retained and flagged**, not merged —
  which one is authoritative is a judgement call.

## Layout

```
survey_tool/
  reference.py   city table, quintile bands, city-name matching
  profile.py     column classification, number parsing, codebook
  clean.py       cleaning: city resolution, coding, quality logging
  build.py       workbook and formula generation
  validate.py    validation sheet and cleaning log
  pipeline.py    profile -> clean -> build -> validate
  cli.py         command line and prompts
city_reference.csv
verify_all.py    independent recomputation of every generated formula
```
