# TMC Processor — Windows user guide

## Setup and launch

1. Use Windows with Python **3.10 or newer** installed. When installing Python,
   select **Add python.exe to PATH**.
2. Download `TMC-Processor-vX.Y.Z-Windows.zip` for the stable version from the project's
   [GitHub Releases](https://github.com/bokoboss/tmc-processor/releases/latest).
3. Extract the entire ZIP into a folder where you can create files. Open the
   extracted `TMC-Processor-vX.Y.Z-Windows` folder (X.Y.Z is the version number).
4. Double-click **start_tmc_processor.bat**. Keep its window open while using
   the app. Your browser opens the local Streamlit application.

This ZIP contains the application source and resources. **It requires Python;
it is not a standalone executable.** You do not need Git or a repository clone.

The first launch creates `.venv` in the extracted folder and downloads Python
packages, so **internet access is required for first setup**. Setup can take
several minutes. Subsequent launches reuse `.venv`. To stop the app, press
**Ctrl+C** in the launcher window.

Microsoft Excel is **optional** for legacy Excel integration. The launcher
attempts to install its integration package and checks for Excel; a warning
from that optional check does not prevent supported Standard or Safe PNG exports.

## Process a survey

Follow **Data → Mapping → Analyze → Review → Export**. Upload a TMC workbook,
check the movement mapping, analyze it, review QC and explicitly confirm AM/PM
Peak periods before exporting. Batch mode processes multiple days of the same
survey point with a shared mapping.

**Standard** preserves the supplied Excel report template when it supports the
confirmed Peak. **Safe PNG** puts chart images into a generated workbook and
is also the fallback when Standard cannot represent the chosen Peak. The exact
confirmed Peak is preserved. Standard does not normally require Excel installed.

Use the app's **download buttons** to save reports, charts, project sessions or
result ZIPs. Files go to your browser's download location (usually Downloads)
or the location you choose. The app does not automatically save them in an
`outputs` folder beside the program.

Optional synthetic examples are in the separate
`TMC-Processor-vX.Y.Z-Demo-Files.zip`. They are not required to run the app.
Extract that ZIP and select the demo files when uploading in the app.

## Privacy and troubleshooting

Survey workbooks, project sessions and generated results may contain private
project information. Keep them in a suitable local folder and inspect result
packages before sharing. Default result ZIPs exclude raw input workbooks;
other files can still reveal project details.

If startup stops, read the message in the launcher window. Check your Python
version, internet connection and write permission to the extracted folder.
Keep `app.py`, `pyproject.toml`, `src`, `templates` and `.streamlit` together;
moving individual files can break imports or template lookup.
