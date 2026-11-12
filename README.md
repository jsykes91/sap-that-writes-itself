# The SAP That Writes Itself

**Bring your estimands: building an eSAP.** The companion code to the STAT-BLOG post
[The SAP That Writes Itself](https://stat-blog.com/posts/the-sap-that-writes-itself/).

An **eSAP** (electronic, or metadata-driven, statistical analysis plan) keeps the
structured parts of a SAP (objectives, endpoints, estimands) as data in one place, and a
script writes the SAP text from that data. This repository builds the title page and
section 1.1 (Objectives, Endpoints and Estimands) of a draft SAP for **PIONEER 1** (oral
semaglutide vs placebo, NCT02906930), with each estimand laid out as an ICH M11 table of
estimand characteristics.

## What's here

```
input/                          what I brought
├── TEMPLATE_SAP_001.docx       the SAP template, every field a named Word content control
└── pioneer1-metadata.xlsx      the PIONEER 1 metadata (read the "Read me" sheet first)
output/                         what Claude built from one prompt, untouched
├── esap.py                     the Python script
├── run_esap.bat                double-click to run (Windows)
├── pioneer1.toml               the settings for this study
├── requirements.txt
├── README.md                   how it works, and every check it makes
├── input/                      copies of the two input files
└── sap/
    ├── PIONEER1_SAP_v0.1.docx  the draft SAP
    └── PIONEER1_SAP_v0.1_run_log.txt
```

## Try it

1. Install [Python](https://www.python.org/downloads/) 3.11 or newer.
2. Download this repository (green **Code** button > **Download ZIP**) and unzip it.
3. Double-click `output/run_esap.bat`. The first run sets up Python's add-ons (needs the
   internet once); then it checks the metadata and builds `output/sap/PIONEER1_SAP_v0.1.docx`.
4. Open the SAP and click **Yes** when Word asks to update fields.

On Mac or Linux: `pip install -r output/requirements.txt`, then
`python output/esap.py --config output/pioneer1.toml`.

## The prompt

Everything in `output/` was built by Claude Code from this one prompt (about 15 minutes,
34 turns), with `input/` as the starting point:

> You're an expert statistician who loves to program with Python. Build a metadata-driven
> SAP (eSAP) for PIONEER 1 in one go, without asking me questions; tell me at the end about
> any choices you made.
>
> In C:\eSAP\input there's our SAP template (TEMPLATE_SAP_001.docx) and its metadata
> (pioneer1-metadata.xlsx). Read the spreadsheet's "Read me" sheet and the template's
> content control titles first. Don't change anything in input.
>
> In C:\eSAP\output, create:
> - esap.py, which turns the template into draft version 0.1 of the PIONEER 1 SAP
> - run_esap.bat, which sets up Python the first time and then runs the script
> - a TOML config file, so the same script works for other studies
> - copies of the inputs, a short README, and sap\PIONEER1_SAP_v0.1.docx with a run log
>
> Each control is named after where its data comes from: Sheet.Column (Level). Controls
> without a dot are groups: copy them once per matching row, number the objectives within
> each level, and rename each copy with its row ID, e.g. [OBJ1]. Show "Not applicable" for
> empty levels, and keep the bookmarks.
>
> Check the metadata before building, and stop with a plain-English list of any problems.
> Then run the script and tell me what you checked.

## Worth knowing

- **The output is exactly what the one-shot produced.** I haven't edited it. For example,
  the sponsor's legal address is blank in the metadata, and Claude chose to leave the
  template's grey "Enter sponsor legal address" prompt in the SAP so it's easy to spot
  (see `optional_fields` in `pioneer1.toml`). I'd have chosen differently. It runs, but
  "it runs" isn't the same as "it's right": review everything.
- **"Byte-identical" means on the same computer.** Rebuilt on another operating system,
  the SAP's contents are identical but the file's compression can differ, so its SHA-256
  fingerprint changes.
- **The template is our own**, made for this demo.
- **The metadata comes from public sources**: ClinicalTrials.gov (NCT02906930) and Aroda
  et al., *Diabetes Obes Metab* 2019;21:2203-2210. Objective and strategy wording is
  paraphrased. It is a demonstration, not the sponsor's SAP.
- **A demo, not a validated tool.** The output is a draft for a statistician to review.

## Licence

MIT. See [LICENSE](LICENSE).
