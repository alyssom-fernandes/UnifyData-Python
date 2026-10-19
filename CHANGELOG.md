# Changelog

## Unreleased

### Closing
- The app now reads the *Total geral* line of each Linx report and
  compares it with the sum of the people it read. The tape gets a
  *Total geral no relatório* line, and a warning appears (also in the
  printed report) when they differ. A report without that line is not
  checked, and the app says so.

### Code and tests
- Removed the unused `escolhas` parameter from `consolidar` and
  `divergencia_repetido`; the namesake test now uses `fixar_linha`, the
  path the app really takes.
- `parse_linx_files` also returns the *Total geral* of each report.
- New test that runs the screen in sample mode (Streamlit's `AppTest`)
  and checks the final total.
- `requirements.txt` pins the tested versions; CI installs from it.

## 1.1.0 (2026-10-13)

### Closing
- The result opens with a reconciliation tape, like an adding-machine
  roll: each store's total in Linx (kept before any card is ignored),
  their sum, what was ignored and what goes to payroll.
- Each step's number is part of its title.
- On wide screens (1200 px or more) the column is wider and the company
  table sits beside the tape.

### Demo
- The sample-mode banner sits at the top of the page, above the sidebar
  and Streamlit's bar, in the brand color.

### Outputs
- Files are named after the month of the due dates
  (`Fechamento_UnifyData_2026-09.xlsx`).
- In the sample mode, the Excel and the report say the data is made up,
  and the file names (Excel, CSV) end in `_DEMO`. The CSV itself carries
  no notice.
- The report and the Excel header use the brand teal instead of burgundy,
  and the report summary shows the same reconciliation as the tape.

### Privacy and security
- The fonts are served by the app itself (`static/fonts`): nothing is
  fetched from Google Fonts.
- A file that cannot be read gets a plain notice; the technical detail
  goes to the console.

### Project
- CI runs the tests and compiles the app on Python 3.10, 3.12 and 3.14.
  Tests: 13 → 17 (unreadable file, file names, the demo Excel and the
  reconciliation).
- The README no longer states what could not be checked.

## 1.0.0 (2026-10-02)

- New interface: the four steps with a sidebar, light and dark themes, a
  phone layout and a sample mode.
- Name matching with review cards; Excel, CSV and a printable report; the
  rules in `nucleo.py`, tested with `unittest`.
