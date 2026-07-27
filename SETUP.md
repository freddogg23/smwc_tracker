# Setup Guide

## 1. Upload this project to GitHub

Upload the **contents** of this folder to the root of your repository.

Do not upload the ZIP file itself. GitHub needs to see:

```text
.github/
community/
data/
scripts/
README.md
SETUP.md
```

## 2. Run the updater once

In GitHub:

1. Click **Actions**.
2. Click **Update SMWCentral Catalog**.
3. Click **Run workflow**.
4. Wait for the green check mark.

## 3. Build the workbook

On your PC:

1. Download/clone the repository.
2. Open the `community` folder.
3. Close every Excel window.
4. Double-click `BUILD_COMMUNITY_TRACKER.bat`.
5. When prompted, enter your repo as `owner/repository`.

Example:

```text
FredDOGG23/smwc_tracker
```

The builder creates:

```text
community/SMW_ROM_Hack_Tracker_Community.xlsm
```

## 4. Test the workbook

Open the generated `.xlsm` and test:

- Only Dashboard, Tracker, and Hack Database are visible.
- Refresh Hacks works.
- Selecting a hack fills the Tracker row.
- Rating dropdown contains 1–10.
- Hours dropdown reaches 1000 Hours.
- Minutes and Seconds reach 59.
- Hack Database links display as Open Page and Download.
- Custom hacks survive a refresh.

## 5. Publish a Release

Use GitHub Releases to share the generated `.xlsm`.

Suggested release note:

```text
SMW ROM Hack Tracker — Community Edition

Requirements:
- Microsoft Excel desktop for Windows
- Internet connection when clicking Refresh Hacks

After downloading:
1. Right-click the XLSM file.
2. Click Properties.
3. Check Unblock.
4. Open the workbook in Excel.
```
