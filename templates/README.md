# TMC Report Templates

`four_leg_tmc_report_template.xlsx` is the source layout template and should be the primary template candidate for now.

Do not use `four_leg_tmc_report_template_cleaned.xlsx` as the default template unless it has been opened in Microsoft Excel and confirmed to save/open without a repair prompt. Complex Excel templates with charts, drawings, and shapes should be cleaned manually in Excel. `openpyxl` should not be used to clean and resave shape-heavy templates because unsupported drawings or chart parts may be dropped or corrupted.

Run the read-only audit before using a template:

```powershell
python scripts/audit_template.py
```

The audit reports worksheet formulas containing `#REF!`, external workbook references such as `[file.xlsx]`, missing sheet references, and detectable native chart sources containing `#REF!`. Unsafe worksheet formulas or missing sheets should block template-driven export and fall back to the generated report workbook. Broken native chart sources can be ignored by template export code as long as export inserts generated PNG charts at mapped anchors instead of editing existing chart objects.

Default template-driven export rules:

- Write only to cells defined in `four_leg_tmc_report_template_map.json`.
- Insert generated PNG charts from the app at mapped chart anchors.
- Do not edit, repair, or depend on existing native chart objects.
- Review mapped cell addresses after any layout edit, especially merged cells, movement diagram cells, table anchors, or chart placement.

This openpyxl-based export with PNG charts is the safe default because it opens reliably without Excel repair warnings.

Native template charts are supported only by the optional Microsoft Excel COM export on Windows. COM export requires Microsoft Excel and pywin32. It opens a copy of `four_leg_tmc_report_template.xlsx`, writes values/formulas into mapped cells and native chart source ranges, lets Excel recalculate, and saves a new workbook. The source template must not be overwritten.

COM export must not create, delete, repair, patch, or directly edit chart objects. The native charts should update only through their existing linked ranges:

- Hourly PCU: `Summary!V10:V21` and `Summary!AM10:AM21`
- Vehicle composition: `Summary!AP39:BA39` and `Summary!AP40:BA40`

If an exported workbook opens with an Excel repair warning, do not use that output. Report the issue and fall back to the default openpyxl export with PNG charts.

The safest template design is mostly cell-based with minimal shapes and native chart objects.
