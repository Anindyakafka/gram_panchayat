# Gram Manchitra Gram Panchayat Extraction

This repository contains a reproducible pipeline for extracting Gram Panchayat
(GP) profile and ward-election information from the Government of India's
[Gram Manchitra](https://grammanchitra.gov.in/gm4MVC) website and writing it to
an analysis-ready Excel workbook.

The current extraction covers 19 surveyed GPs:

| State | District | Block | Gram Panchayat | Survey GP ID |
|---|---|---|---|---:|
| Madhya Pradesh | Barwani | Barwani | Bhandarda | 132840 |
| Madhya Pradesh | Barwani | Pati | Budi | 132970 |
| Madhya Pradesh | Barwani | Rajpur | Danod | 133021 |
| Madhya Pradesh | Barwani | Barwani | Dhamnai | 132850 |
| Madhya Pradesh | Barwani | Barwani | Kalyanpura | 132856 |
| Madhya Pradesh | Barwani | Pati | Osada | 132988 |
| Madhya Pradesh | Barwani | Pati | Pakhalya | 132989 |
| Madhya Pradesh | Barwani | Pati | Semlet (F) | 132999 |
| Madhya Pradesh | Barwani | Pati | Semli | 133000 |
| Madhya Pradesh | Barwani | Barwani | Silawad | 132874 |
| Madhya Pradesh | Barwani | Pati | Valan | 133006 |
| Bihar | Nawada | Mescaur | Ankri Pandeybigha | 98342 |
| Bihar | Nawada | Mescaur | Badosar | 98344 |
| Bihar | Nawada | Mescaur | Barat | 98343 |
| Bihar | Nawada | Narhat | Jamuara | 98364 |
| Bihar | Nawada | Narhat | Konibar | 98366 |
| Bihar | Nawada | Mescaur | Meskaur | 98347 |
| Bihar | Nawada | Narhat | Pali Khurd | 98368 |
| Bihar | Nawada | Narhat | Punaul | 98369 |

The generated workbook is `output/gram_manchitra_selected_gps.xlsx`. It
contains a GP-level information sheet and a long-format ward-election sheet.
Limited raw JSON snapshots are saved so every workbook value can be audited
against the source responses.

## Contents

- [Project structure](#project-structure)
- [Data sources](#data-sources)
- [Output workbook](#output-workbook)
- [Data-handling rules](#data-handling-rules)
- [Validation safeguards](#validation-safeguards)
- [Installation](#installation)
- [Running the extraction](#running-the-extraction)
- [Adding more Gram Panchayats](#adding-more-gram-panchayats)
- [Raw audit files](#raw-audit-files)
- [Known limitations](#known-limitations)
- [Troubleshooting](#troubleshooting)

## Project Structure

```text
gram_panchayat/
|-- README.md
|-- PROJECT_STRUCTURE.md
|-- requirements.txt
|-- gram_panchayat_reservation.xlsx
|-- config/
|   `-- targets.csv
|-- scripts/
|   `-- scrape_gram_manchitra.py
|-- data/
|   `-- raw/
|       |-- <gp_code>_<gp_name>_basic.json
|       |-- <gp_code>_<gp_name>_glance.json
|       `-- <gp_code>_<gp_name>_villages.json
`-- output/
    `-- gram_manchitra_selected_gps.xlsx
```

### Important Files

| Path | Purpose |
|---|---|
| `gram_panchayat_reservation.xlsx` | Supplied survey workbook. Its `gp_code` sheet maps survey GP names to GP IDs. |
| `config/targets.csv` | Administrative hierarchy and GP names to extract. |
| `scripts/scrape_gram_manchitra.py` | Downloads, validates, reshapes, and writes the data. |
| `data/raw/` | Generated audit snapshots from Gram Manchitra's public responses. |
| `output/` | Generated analysis-ready Excel workbooks. |
| `requirements.txt` | Pinned Python dependency used to read and write Excel files. |

## Data Sources

The script uses three inputs.

### 1. Survey GP-Code Workbook

The `gp_code` sheet in `gram_panchayat_reservation.xlsx` contains:

| Column | Meaning |
|---|---|
| `code` | Survey GP identifier / Gram Panchayat code |
| `gp_name` | GP name used to match a requested target |

The script trims surrounding whitespace and performs case-insensitive matching
only for validation and lookup. It does not use this normalization to alter text
copied into the output workbook.

### 2. Gram Manchitra Profile Endpoint

The public `Home/GetPanchAreaDetails` endpoint supplies:

- the GP identity and administrative hierarchy;
- Sarpanch and Secretary names;
- coded person attributes;
- the `GP at a glance` Election Details records.

The extraction uses the same response consumed by the website's Election
Details table.

### 3. Gram Manchitra Administrative GIS Layer

The public Census Villages layer supplies:

- village names belonging to the selected GP;
- village codes;
- GP code and GP name;
- Block, District, and State names.

This layer also verifies that the selected GP belongs to the requested State,
District, and Block before any output is written.

## Output Workbook

The workbook contains two sheets with one header row, no merged cells, and one
variable per column.

### GP Information

This sheet contains one row per Gram Panchayat.

| Column | Description |
|---|---|
| `State` | Requested and verified State name |
| `District` | Requested and verified District name |
| `Block` | Requested and verified Block name |
| `GP name` | GP name from the target list |
| `Our GP ID` | Code matched from the survey workbook's `gp_code` sheet |
| `Village 1`, `Village 2`, ... | Village names returned by Gram Manchitra; columns expand to fit the GP with the most villages |
| `Sarpanch name` | Name returned by the GP profile response |
| `Sarpanch gender` | Blank unless a human-readable gender label is available |
| `Secretary name` | Name returned by the GP profile response |
| `Secretary gender` | Blank unless a human-readable gender label is available |
| `Notes` | Explanation for requested fields that could not be populated |

Current village coverage:

| District | GPs | Maximum villages in one GP |
|---|---:|---:|
| Barwani | 11 | 8 |
| Nawada | 8 | 12 |

The workbook dynamically expands through `Village 12` for the current target
set. Exact village lists are stored in the GP Information sheet and the
corresponding `_villages.json` audit files.

### Ward Election Details

This sheet is in long format: each API Election Details record occupies one row.
State, District, Block, GP name, and GP ID are repeated on every row.

The current website exposes these election fields:

| Output column | Gram Manchitra field | Type / handling |
|---|---|---|
| `Ward name` | `wardName` | Text copied as returned |
| `Member/elected representative name` | `name` | Text copied as returned |
| `Ward type` | `wardType` | Text; blank when unavailable |
| `Gender` | `gender` | Human-readable website category |
| `Caste/category` | `caste` | Website category |
| `Educational qualification` | `qualification` | Website category; blank when unavailable |
| `Age` | `age` | Numeric when displayed; API value `0` is blank because the live website renders it as `N/A` |
| `Election term start date` | `electionTermStart` | Excel date formatted as `YYYY-MM-DD` |
| `Election term end date` | `electionTermEnddate` | Excel date formatted as `YYYY-MM-DD` |
| `Election type` | `electionType` | Text; blank when unavailable |
| `Notes` | Generated | Lists fields displayed as unavailable for that record |

The current workbook contains 336 Election Details records:

| GP | Records |
|---|---:|
| Bhandarda | 21 |
| Budi | 21 |
| Danod | 21 |
| Dhamnai | 15 |
| Kalyanpura | 19 |
| Osada | 21 |
| Pakhalya | 19 |
| Semlet (F) | 21 |
| Semli | 21 |
| Silawad | 21 |
| Valan | 21 |
| Ankri Pandeybigha | 17 |
| Badosar | 13 |
| Barat | 15 |
| Jamuara | 14 |
| Konibar | 15 |
| Meskaur | 13 |
| Pali Khurd | 14 |
| Punaul | 14 |

The script retains final records where Gram Manchitra returns a null
representative name, `wardName = N/A`, and `wardType = S`. These are source
records and are not silently discarded.

If Gram Manchitra adds keys to an Election Details response, the script
automatically appends them as columns using the original API field names. This
prevents newly available fields from being silently dropped.

## Data-Handling Rules

1. Names and categories retain their original spelling, capitalization, and
   spacing from the relevant Gram Manchitra response.
2. Village, Sarpanch, Secretary, and representative names are not standardized.
3. Missing fields remain blank. The script does not infer missing values.
4. Gender is never inferred from a person's name.
5. The profile currently supplies numeric Sarpanch and Secretary gender codes
   without human-readable labels. These gender cells remain blank, and the
   codes are mentioned in the GP-level `Notes` column.
6. Election `age = 0` remains blank because the live table displays it as
   `N/A`.
7. Election dates are stored as Excel date values, not combined text fields.
8. API nulls become blank Excel cells with a short explanation in `Notes`.
9. No cells are merged. Each sheet has exactly one header row.
10. GP identifiers are repeated on every ward-election row.

## Validation Safeguards

The script stops instead of producing a potentially misidentified workbook when:

- the source workbook lacks a `gp_code` sheet;
- the expected `code` or `gp_name` column is missing;
- a target GP has no survey-workbook match;
- a target GP name maps to more than one distinct code;
- the GIS layer returns no village for the GP code;
- village records disagree about State, District, Block, or GP;
- the GIS hierarchy does not match `config/targets.csv`;
- the profile response's GP code or GP name does not match the target;
- a required public endpoint repeatedly fails.

Text comparisons for these checks are case-insensitive and whitespace tolerant.
Output values are not normalized by that process.

Network requests use a 90-second timeout and up to four attempts with bounded
exponential backoff because the government GIS service can be intermittent.

## Installation

### Requirements

- Python 3.10 or newer
- Internet access to the Gram Manchitra application and GIS service
- The supplied `.xlsx` workbook

### Optional Virtual Environment

PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### Install the Dependency

```powershell
python -m pip install -r requirements.txt
```

The only external Python package is `openpyxl`, pinned in
`requirements.txt` for reproducible Excel handling. HTTP requests use Python's
standard library.

## Running the Extraction

From the repository root:

```powershell
python scripts/scrape_gram_manchitra.py
```

With the default configuration, this command:

1. reads GP codes from `gram_panchayat_reservation.xlsx`;
2. reads requested GPs from `config/targets.csv`;
3. downloads basic profile, GP-at-a-glance, and village responses;
4. validates each administrative hierarchy and GP identity;
5. writes audit JSON files under `data/raw/`;
6. writes the two-sheet workbook under `output/`.

A successful run prints:

```text
Wrote ...\output\gram_manchitra_selected_gps.xlsx with 19 GPs and 336 election rows
```

### Command-Line Options

All default paths can be overridden:

```powershell
python scripts/scrape_gram_manchitra.py `
  --source path\to\survey_workbook.xlsx `
  --targets path\to\targets.csv `
  --output path\to\result.xlsx `
  --raw-dir path\to\raw_json
```

| Option | Default | Purpose |
|---|---|---|
| `--source` | `gram_panchayat_reservation.xlsx` | Workbook containing `gp_code` |
| `--targets` | `config/targets.csv` | Administrative hierarchies and GP names |
| `--output` | `output/gram_manchitra_selected_gps.xlsx` | Generated workbook |
| `--raw-dir` | `data/raw` | Generated audit JSON directory |

The program replaces files with the same generated output names. It does not
modify the supplied survey workbook.

## Adding More Gram Panchayats

Add one line per GP to `config/targets.csv`:

```csv
state,district,block,gp_name,site_gp_name
Madhya Pradesh,Barwani,Barwani,Bhandarda,
Madhya Pradesh,Barwani,Pati,Semlet,Semlet (F)
```

Keep the five column names unchanged. `gp_name` is the name used to match the
survey workbook. `site_gp_name` is optional and should normally be blank. Use it
only for a verified Gram Manchitra label difference, such as survey name
`Semlet` versus site name `Semlet (F)`. Case differences are tolerated during
validation, but correct State, District, Block, and GP selections are required.

Confirm that each target appears in the source workbook's `gp_code` sheet. The
current sheet contains only GP name and code, not District or Block. If the same
GP name maps to multiple codes, the script deliberately stops because it cannot
safely choose a code from name alone. Resolve that case by making the input
mapping unambiguous or extending it with administrative identifiers.

For a larger batch, use a general output filename:

```powershell
python scripts/scrape_gram_manchitra.py `
  --output output\gram_manchitra_all_surveyed_gps.xlsx
```

## Raw Audit Files

Each GP produces three JSON files.

| Filename suffix | Contents |
|---|---|
| `_basic.json` | Task-relevant GP identity, office-holder names, and unlabeled person codes |
| `_glance.json` | Full GP-at-a-glance response, including Election Details |
| `_villages.json` | Census Villages response used for names and hierarchy verification |

Example:

```text
132840_bhandarda_basic.json
132840_bhandarda_glance.json
132840_bhandarda_villages.json
```

The basic snapshot is restricted to task-relevant fields. Large profile
photographs and unrelated contact details are not retained. Raw files are UTF-8
JSON and preserve source values for auditing.

## Known Limitations

### No Separate Ward Number

For these GPs, Gram Manchitra exposes `wardName` but no separate ward-number
field. The administrative GIS service also has no public ward layer for these
records. A ward number is not inferred from row order or house ranges.

### No Additional Election Variables

The live Election Details table and API expose the same ten fields listed above.
They do not provide reservation status, political party, vote count, winning
margin, contact information, or a separate member ID for these two GPs.

### Unlabeled Person Gender Codes

The basic profile response returns numeric values for Sarpanch and Secretary
gender, but the public GP view provides no verified mapping from those codes to
labels. The output leaves both fields blank rather than inferring a category
from a name or undocumented code.

### Website Availability and Change

Gram Manchitra is an external government system. Endpoints, response fields,
availability, or categories may change. Retry logic handles temporary failures,
and unknown election fields are retained automatically, but structural changes
may require a script update.

### Source Data Are Not Cleaned

The output intentionally preserves source spelling and capitalization. Similar
names may appear in different cases or with different spellings. Cleaning and
harmonization should be performed in a separate analysis step so the extracted
source record remains reproducible.

## Troubleshooting

### `ModuleNotFoundError: No module named 'openpyxl'`

```powershell
python -m pip install -r requirements.txt
```

### `No gp_code match`

Check that the GP exists in the `gp_code` sheet and that `gp_name` is
populated. Matching ignores case and repeated whitespace but does not use fuzzy
matching or spelling correction.

### `Ambiguous gp_code matches`

The same GP name maps to multiple distinct codes in the source sheet. The
current mapping lacks District and Block columns, so the script cannot safely
disambiguate the records. Correct or enrich the mapping before rerunning.

### `Hierarchy mismatch`

The State, District, Block, or GP returned by the village layer does not match
the target row. Review `config/targets.csv` and the GP-code mapping. Do not
bypass this check without manually establishing the correct GP identity.

### Request Timeout or Endpoint Failure

The script retries each request four times. If all attempts fail:

1. confirm that Gram Manchitra opens in a browser;
2. wait and rerun because the GIS service can be intermittent;
3. confirm that a firewall is not blocking the two hosts;
4. inspect the final exception to identify the failed endpoint.

### Workbook Is Open in Excel

Close the generated workbook before rerunning. Excel may lock the file and
prevent the script from replacing it.

## Reproducibility Notes

- Keep the original survey workbook unchanged.
- Commit changes to `config/targets.csv` and the extraction script together.
- Retain raw JSON files associated with any workbook used for analysis.
- Record the extraction date in the surrounding research workflow or data log.
- Treat cleaning, standardization, fuzzy matching, and derived variables as a
  separate downstream stage.

This separation preserves a direct, auditable path from Gram Manchitra's source
responses to the Excel values used by the research team.
