# One-Time Maintainer Setup

These steps use the existing GitHub repository:

```text
https://github.com/freddogg23/smwc_tracker
```

The workbook is already configured to refresh from that repository.

## 1. Upload the project to GitHub

1. Extract `SMW_Community_Tracker_Project.zip`.
2. Open the extracted `SMW_Community_Tracker_Project` folder.
3. In the `smwc_tracker` repository, choose **Add file → Upload files**.
4. Upload the **contents** of the project folder, not the outer folder and not only the ZIP.
5. Allow GitHub to replace the existing workflow, script, data, README, and setup files.
6. Commit the changes.

The repository root should contain:

```text
.github/
community/
data/
scripts/
README.md
SETUP.md
CHANGELOG.md
COMMUNITY_USER_GUIDE.md
```

## 2. Test the daily GitHub updater

1. Open the repository's **Actions** tab.
2. Select **Update SMWCentral Catalog**.
3. Choose **Run workflow → Run workflow**.
4. Wait for the green check mark.

The workflow runs daily at 09:15 UTC and can also be run manually.

## 3. Build the actual Community Edition workbook

This is a one-time maintainer build step. Community users do not run the builder.

Requirements:

- Windows
- Microsoft Excel desktop
- All Excel windows closed

Steps:

1. Open the local `community` folder.
2. Double-click `BUILD_COMMUNITY_TRACKER.bat`.
3. Wait for the builder to create and open:

```text
SMW_ROM_Hack_Tracker_Community.xlsm
```

The builder automatically:

- Converts the clean `.xlsx` template to `.xlsm`.
- Imports the Refresh and Custom Hack VBA code.
- Adds the two Dashboard buttons.
- Locks formula cells while leaving user-entry cells editable.
- Makes all support sheets Very Hidden.
- Protects workbook structure.

### If the builder reports a VBA project access error

1. Open Excel.
2. Go to **File → Options → Trust Center → Trust Center Settings**.
3. Open **Macro Settings**.
4. Enable **Trust access to the VBA project object model**.
5. Close every Excel window.
6. Run `BUILD_COMMUNITY_TRACKER.bat` again.

The builder normally enables this setting temporarily and restores its previous value, but some managed Office installations require the manual setting.

## 4. Test the finished XLSM

1. Open `SMW_ROM_Hack_Tracker_Community.xlsm`.
2. Enable macros if Excel asks.
3. Confirm that only **Dashboard** and **Tracker** are visible.
4. On Dashboard, click **Refresh Hacks**.
5. Add a test custom hack and confirm it appears in the Tracker dropdown.
6. Click **Refresh Hacks** again and confirm the custom hack remains.
7. Delete the test custom hack by rebuilding the XLSM from the clean template before publishing, or remove it through the hidden Manual Database while maintaining the file.

## 5. Publish the Community Edition

Recommended method:

1. In GitHub, open **Releases**.
2. Choose **Draft a new release**.
3. Tag it `v1.0.0`.
4. Title it `SMW ROM Hack Tracker — Community Edition v1.0.0`.
5. Attach `SMW_ROM_Hack_Tracker_Community.xlsm`.
6. Publish the release.
7. Share the release link in Discord.

Users download the workbook once. Later catalog updates arrive through the **Refresh Hacks** button, so their personal Tracker does not need to be replaced.
