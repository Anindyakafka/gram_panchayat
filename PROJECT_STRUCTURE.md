# Gram Manchitra extraction

This folder keeps the supplied survey workbook at the project root and adds a
small reproducible extraction pipeline:

- `gram_panchayat_reservation.xlsx`: supplied workbook; `gp_code` is the ID source.
- `config/targets.csv`: requested State/District/Block/GP selections.
- `scripts/scrape_gram_manchitra.py`: validated extraction and workbook builder.
- `data/raw/`: audit JSON returned by the public endpoints (generated).
- `output/`: final analysis-ready workbook (generated).

Install and run from PowerShell:

```powershell
python -m pip install -r requirements.txt
python scripts/scrape_gram_manchitra.py
```

The script stops on a GP-code or administrative-hierarchy mismatch. It copies
person and village names without case or spelling changes. Null website fields
remain blank. Person gender stays blank when Gram Manchitra supplies only an
unlabeled numeric code, and the reason is recorded in `Notes`. The election
extract includes every field returned by the endpoint; any fields added by the
site beyond the known table columns are appended automatically using their raw
field names. Values rendered as `N/A` by the live table, including an age of
zero, remain blank and are documented in `Notes`.
