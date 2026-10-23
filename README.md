# tsv-to-sysmlv2: deterministic SysML 1.7 to SysML v2 converter

[![DOI](https://zenodo.org/badge/1430612443.svg)](https://doi.org/10.5281/zenodo.23587841)

Converts a SysML 1.7 model exported from CATIA Magic / MSOSA 2026x into
SysML v2 textual notation, and writes a migration report listing every
element, property, and diagram that could not be migrated automatically.
Same input in, same output out, every run. No language model is involved,
so results are reproducible and auditable.

Supporting software for the doctoral praxis "Metamodel Scope Using Model-Based Decision on Intended-Use & Requirements Engineering Categorization Tool (MB-DIRECT)"
(Evan Matthews, The George Washington University, School of Engineering and Applied Science, 2026).

## What is here

| Path | Contents |
|---|---|
| `src/tsv_to_sysmlv2.py` | The converter. Python 3.8+, needs openpyxl for the report. |
| `templates/SysML17_Migration_Export_w_Diagrams.vm` | Report Wizard template (v6.1) that exports the 33-column TSV the converter reads. |
| `examples/Homegrown_Export.tsv` | A small exported model for a smoke test. |
| `examples/expected_Homegrown_pilot.sysml` | The converter's output for that TSV, pilot dialect. |
| `examples/expected_Homegrown_MigrationReport.xlsx` | The companion migration report for that run. |
| `docs/data_dictionary.md` | Column-by-column description of the TSV input and the report workbook. |

## Install

```
python -m venv .venv
.venv\Scripts\activate        (Windows)    or    source .venv/bin/activate   (macOS, Linux)
pip install -r requirements.txt
```

## Step 1: export the model from MSOSA

1. In CATIA Magic / MSOSA 2026x open the SysML 1.7 project.
2. Tools > Report Wizard > New template, and point it at
   `templates/SysML17_Migration_Export_w_Diagrams.vm`. Output format is plain
   text; give the output file a `.tsv` extension.
3. Select as scope exactly the packages you want converted. Do not include
   profiles or libraries.
4. Generate. The result is a tab-separated file with 33 columns and one row
   per element or diagram.

## Step 2: convert

```
python src/tsv_to_sysmlv2.py path/to/Export.tsv
```

This writes `Export_pilot.sysml` and `Export_pilot_MigrationReport.xlsx`
next to the TSV. Options:

| Option | Meaning |
|---|---|
| `-o OUT`, `--out OUT` | Output `.sysml` path (default: `<tsv name>_<dialect>.sysml` next to the TSV). |
| `--dialect pilot` | Default. Output loads cleanly in the SysML v2 Pilot Implementation. |
| `--dialect msosa` | CATIA Magic / MSOSA 2026x textual conventions (colon-typed state actions). |
| `--report R.xlsx` | Migration report path (default: `<out>_MigrationReport.xlsx`). |
| `--no-report` | Skip the report. |
| `--validate` | Run the output through the SysML v2 Pilot kernel and print ERROR and WARNING lines. Needs Jupyter and the `sysml` kernel installed. |
| `--jupyter PATH` | Jupyter executable for `--validate` (default: the first `jupyter` on the PATH). |

## Smoke test

```
python src/tsv_to_sysmlv2.py examples/Homegrown_Export.tsv -o examples/Homegrown_pilot.sysml
```

The new `examples/Homegrown_pilot.sysml` must be identical to
`examples/expected_Homegrown_pilot.sysml`, and the console summary must read:

```
elements: 30 migrated, 6 TODO-stubbed (in report), 2 diagrams (in report), 1 consumed by parent constructs
```

## How the output is organized

Every SysML 1.7 element with a clean SysML v2 equivalent is emitted as that
construct (for example Block to `part def`, Requirement to `requirement`
usage, StateMachine to `state def`). Anything without a clean v2 home
becomes a `// TODO-MIGRATE` comment at the right place in the text and a
row in the report workbook, so nothing is dropped silently. The report uses
the same sheet layout as the CATIA Magic migration plugin's own
`MigrationReportOf<project>.xlsx`, so the two can be compared directly.
Column definitions are in `docs/data_dictionary.md`.

The companion package `compare-migration` audits the result of a migration
(this converter's or the CATIA plugin's) against the original model.

## AI assistance

Portions of `src/tsv_to_sysmlv2.py` were developed with assistance from Anthropic's
Claude (Claude Code). All AI-assisted code was specified, reviewed, tested,
and validated by the author. This statement is provided in accordance with
the AI Policy for George Washington University Online Engineering Programs.

## License

Code: MIT (see `LICENSE`). Templates and example data: CC BY 4.0 (see
`DATA_LICENSE.md`).

## How to cite

See `CITATION.cff`, or use the "Cite this repository" button on GitHub.
The archived release is at https://doi.org/10.5281/zenodo.XXXXXXX.

## Author

Evan Matthews, ORCID https://orcid.org/0009-0005-2265-0754
