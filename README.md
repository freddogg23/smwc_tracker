# SMW ROM Hack Tracker — Community Edition

This repository contains the approved Community Edition workbook template, the GitHub catalog updater, and the VBA builder used to create the public `.xlsm` release.

## Included features

- Incremental **Refresh Hacks** updates: downloaded workbooks normally fetch only `version.json` and missing delta CSVs.
- Full-catalog recovery command if a delta is missing.
- Personal Tracker entries and custom hacks are preserved.
- Hack Finder filters for:
  - Difficulty
  - Type
  - SMWCentral Rating
- Random Hack button that selects a matching hack.
- SMWCentral Rating column support on Hack Database, Tracker, and Hack Finder.
- Only Dashboard, Tracker, Hack Database, and Hack Finder are visible in the built `.xlsm`; support sheets are Very Hidden.

## Repository structure

```text
.github/workflows/update-smwcentral.yml
scripts/Fetch_SMWCentral_All_Hacks.py
data/SMWCentral_All_Moderated_Hacks.csv
data/SMWCentral_All_Moderated_Hacks.json
data/version.json
data/deltas/.gitkeep
community/SMW_ROM_Hack_Tracker_Community.xlsx
community/SMWCommunity.bas
community/Build_Community_Tracker.ps1
community/BUILD_COMMUNITY_TRACKER.bat
```

## Public download

Users should download `SMW_ROM_Hack_Tracker_Community.xlsm` from GitHub Releases, not the source `.xlsx` in the repository.

## Rating note

The updater writes SMWCentral Rating values when the SMWCentral catalog response exposes a rating field. Ratings remain blank when the source response does not provide one.
