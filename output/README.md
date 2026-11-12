# eSAP: metadata-driven SAP for PIONEER 1

This folder turns a Word SAP template and a metadata workbook into a draft
Statistical Analysis Plan. The template is fixed. The workbook holds the study
facts. The script copies the facts into the template and checks the result.

## How to run it

Double-click **`run_esap.bat`**.

- **First run.** It creates a private Python environment in `.venv` and installs two
  packages (`openpyxl` reads Excel, `lxml` edits Word). This needs Python 3.11 or
  newer and an internet connection, once.
- **Every run.** It checks the metadata. If anything is wrong, it **stops and lists the
  problems in plain English**, and no SAP is written. If all is well, it writes the SAP
  and a run log to `sap\`.

From a command prompt you can also run:

    run_esap.bat                          build PIONEER 1 (uses pioneer1.toml)
    run_esap.bat pioneer1.toml --check-only   check only, build nothing
    run_esap.bat otherstudy.toml          build another study

When you open the SAP, Word asks whether to update fields. Say **Yes**, and the
table of contents fills in.

## What is here

| File | What it is |
|---|---|
| `esap.py` | The script. |
| `pioneer1.toml` | The study settings: which template and workbook to use, the output name, and how the sheets link. |
| `run_esap.bat` | Sets up Python the first time, then runs the script. |
| `requirements.txt` | The two Python packages the script needs. |
| `input\` | Copies of the template and the metadata workbook. The script never changes them. |
| `sap\PIONEER1_SAP_v0.1.docx` | The draft SAP. |
| `sap\PIONEER1_SAP_v0.1_run_log.txt` | What was checked, what was built, and fingerprints (SHA-256) of every input and the output. |

## How the template is filled (the rubber stamp)

Each content control in the template is named after where its text comes from.

- **`Study.Protocol title`**: one field from the Study sheet.
- **`Objectives.Objective (Primary)`**: a field. The script fills it from the current
  row and renames it **`Objectives.Objective [OBJ1]`**.
- **`Objective (Primary)`** (no dot): a **group**, which works like a stamp. The script
  prints one copy per matching row, for example per Primary objective, and renames
  each copy with its row ID: **`Objective [OBJ1]`**. Estimands are printed inside their
  objective, and intercurrent events inside their estimand. `(each)` means every row.
- **`Objective number`**: worked out by the script. Objectives are numbered 1, 2, 3 …
  within each level, in sheet order.
- A level with no rows (here, Exploratory) shows **"Not applicable"**.
- Estimands with Role *Supplementary* go under the "Supplementary Estimand(s) for …"
  subheading. The others go under "Estimand(s) for …".
- Bookmarks are kept exactly as in the template.

To find where a value came from, search the SAP's content controls for the name in
the workbook's green *SAP fields* column, e.g. `Estimands.Population [E1]`.

## What is checked

**Before building** (any failure stops the run):

1. Every sheet and column the template needs is in the workbook, with no duplicate headings.
2. No data cell is a formula Excel has not yet calculated.
3. No hidden control characters (these often come from copy-and-paste).
4. Each Study field is listed once.
5. Every row has an ID, and IDs are unique on each sheet.
6. Every Objective ID and Estimand ID link points to a real row.
7. Every template control name is understood.
8. No blank cells in columns the SAP needs. Blank Study fields are allowed only if listed
   as optional in the config.
9. Level, Role and Strategy hold only the values in the workbook's drop-down lists.
10. The template has a stamp for every level used.
11. Each estimand's Role matches its objective's Level. Supplementary estimands can sit
    under any level.
12. The SAP version on the Study sheet is in the Version history, with the same date.
13. The *SAP fields* column is filled in on every row and refers to its own row.
14. The output file name can be made.

**Notes** (reported, but they do not stop the run): empty levels, objectives without
estimands, estimands without intercurrent events, an estimand endpoint that differs
from its objective's endpoint, and metadata that this template does not use.

**After building, before saving** (any failure means the SAP is not saved):

1. The file is a complete, readable Word package.
2. Every metadata value appears in the SAP word for word, under its `Sheet.Column [ID]` name.
3. Objective numbers run 1, 2, 3 … within each level.
4. No unfilled stamps or placeholders remain, except optional blanks.
5. All template bookmarks are kept, once each, and correctly paired.
6. Every content control has a unique id.
7. Every name in the *SAP fields* column can be found in the SAP.

The same inputs always give a byte-identical SAP. The run log records the SHA-256
fingerprints, so you can show that a SAP came from a given template and workbook.

## Using it for another study

1. Put the new study's workbook (same layout) and template in `input\`.
2. Copy `pioneer1.toml` to e.g. `mystudy.toml`. Change `[study]` and `[files]`. If needed,
   also change `optional_fields`.
3. Run `run_esap.bat mystudy.toml`.

The sheet links, the numbering and the Supplementary split are also set in the TOML
file, so a template with other groups can be handled by editing the config, not the
script.
