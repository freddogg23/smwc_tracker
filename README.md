# SMW ROM Hack Tracker — Community Edition

This repository contains the public data feed and Excel workbook source for the SMW ROM Hack Tracker Community Edition.

## What users download

Users should download the macro-enabled workbook from **GitHub Releases**:

`SMW_ROM_Hack_Tracker_Community.xlsm`

The workbook lets users:
- track their own SMW ROM hack progress locally;
- click **Refresh Hacks** to pull the newest SMWCentral catalog;
- add custom hacks that are never removed by refreshes;
- use **Open Page** and **Download** links without seeing raw URLs;
- keep personal dates, ratings, playtime, and notes private in their own copy.

## Repository structure

```text
.github/workflows/update-smwcentral.yml
scripts/Fetch_SMWCentral_All_Hacks.py
data/SMWCentral_All_Moderated_Hacks.csv
data/SMWCentral_All_Moderated_Hacks.json
data/version.json
community/SMW_ROM_Hack_Tracker_Community.xlsx
community/SMWCommunity.bas
community/Build_Community_Tracker.ps1
community/BUILD_COMMUNITY_TRACKER.bat
```

## Daily catalog updates

GitHub Actions updates the public SMWCentral catalog daily. It can also be run manually from the Actions tab.
