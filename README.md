# SMW ROM Hack Tracker — Community Edition

A blank, community-ready Excel tracker backed by a daily-updated SMWCentral catalog.

## What users get

- A clean **Dashboard** and **Tracker**.
- A searchable dropdown containing the moderated SMWCentral hack catalog.
- Automatic title, creator, exits, difficulty, type, page link, and patch link.
- Personal dates, status, ratings, playtime, and notes.
- A **Refresh Hacks** button that updates only the shared official catalog.
- An **Add Custom Hack** button for hacks outside the official catalog.
- Custom hacks and personal progress remain local to each user's copy.

Only the `Dashboard` and `Tracker` sheets are visible. Support sheets are Very Hidden and the workbook structure is protected to prevent accidental damage.

## Repository layout

```text
.github/workflows/update-smwcentral.yml   Daily cloud updater
data/                                     Public CSV, JSON, and version files
scripts/update_catalog.py                 SMWCentral catalog updater
community/SMW_ROM_Hack_Tracker_Community.xlsx
community/SMWCommunity.bas
community/Build_Community_Tracker.ps1
community/BUILD_COMMUNITY_TRACKER.bat
```

## Maintainer setup

See [SETUP.md](SETUP.md). The one-click Windows builder creates the final macro-enabled file:

```text
SMW_ROM_Hack_Tracker_Community.xlsm
```

That `.xlsm` file is the Community Edition release users should download.

## Privacy and ownership

Every downloaded workbook is independent. One user cannot change another user's Tracker. The GitHub repository supplies only the shared hack catalog; personal progress and custom hacks stay inside the user's local workbook.

## Catalog source

The catalog updater reads SMWCentral's currently moderated Super Mario World hacks section. The workbook links to SMWCentral patch archives and does not contain copyrighted ROM files.
