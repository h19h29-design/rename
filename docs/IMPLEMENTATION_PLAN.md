# RE:NAME — Windows document pseudonymization

Goal: A local Windows desktop utility for reviewing and replacing names, resident-registration numbers, birthdays, and optional amounts / other dates in Excel and Hancom documents.

Architecture: Tkinter desktop UI → reviewable, in-memory detection session → format-preserving ZIP/XML adapters. No network client, telemetry, database, or original-value report. XLS and HWP use separately installed desktop Office applications only for conversion; final formats are XLSX and HWPX.

## Decisions
- Primary formats: XLSX / HWPX. Legacy XLS / HWP require Windows + installed Microsoft Excel / Hancom Hangul respectively.
- Names use explicit labels, table-column hints, and user-supplied exact names. Unlabelled arbitrary names are not guaranteed to be detected.
- Replacement modes: unique placeholders, full masking, deletion; every replacement remains editable before saving.
- Default categories: name, resident-registration number, birthday. Amount / general date require opt-in.
- Originals never overwritten. Neutral output filenames. Session mappings live in memory only.
- XLSX output is a sharing copy: formulas become cached values, comments/properties/link targets are removed, and shared strings are materialized to prevent unused originals surviving.
- HWPX output regenerates text preview and blanks image preview, and strips document metadata. Images require an explicit manual-review acknowledgement. Embedded objects, active content, revision histories, and unsupported data caches are blocked rather than falsely declared safe.
- Native Windows + Hancom integration needs a real-device acceptance test; synthetic XML tests cannot establish native fidelity.

## Tasks
- [x] Add detection tests; implement category and replacement rules, date validation, overlap precedence and repeat consistency (`rename_app/rules.py`).
- [x] Add ZIP/XML and document tests; implement guarded input, XLSX / HWPX adapters, metadata cleanup and output verification (`rename_app/documents.py`).
- [x] Add session tests; implement preview, atomic no-overwrite export and counts-only reporting (`rename_app/session.py`).
- [x] Add Windows converters with disabled Excel macros and no global Hancom security bypass (`rename_app/windows.py`).
- [x] Add desktop interface, preview table, batch files, manual names and safe acknowledgement (`rename_app/gui.py`).
- [ ] Add Windows source launcher / portable build and CI workflow; run core + offscreen UI tests, inspect screenshot, commit/push and verify remote SHA.

## Acceptance
`python -m pytest -q` passes; synthetic sensitive text is absent from unpacked outputs; original hashes are unchanged; deselected options remain unchanged; duplicate outputs cannot overwrite; a GUI session can load, inspect, edit and export without network access. Windows EXE / installed Hancom claims must be based on actual Windows evidence, not inferred from Linux tests.
