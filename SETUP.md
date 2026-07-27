# GitHub Setup

## 1. Upload this package

Extract the ZIP, then upload the **contents** to the root of your existing GitHub repository. The repository should show:

```text
.github/
community/
data/
scripts/
README.md
SETUP.md
CHANGELOG.md
```

Do not upload the ZIP itself as the repository content.

## 2. Commit the files

Use a commit message such as:

```text
Add incremental refresh and Random Hack Finder
```

## 3. Run GitHub Actions once

Open:

```text
Actions → Update SMWCentral Catalog → Run workflow
```

Wait for a green check mark. The workflow maintains the full CSV/JSON and creates numbered delta files under `data/deltas/` whenever the catalog changes.

## 4. Build the final macro-enabled workbook

On your Windows PC:

1. Download or clone the repository.
2. Open the `community` folder.
3. Close every Excel window.
4. Double-click `BUILD_COMMUNITY_TRACKER.bat`.
5. Enter your repository as `owner/repository`, for example:

```text
freddogg23/smwc_tracker
```

The builder creates:

```text
SMW_ROM_Hack_Tracker_Community.xlsm
```

If VBA import is blocked, enable this Excel option temporarily:

```text
File → Options → Trust Center → Trust Center Settings → Macro Settings
→ Trust access to the VBA project object model
```

## 5. Test before publishing

Verify:

- Dashboard, Tracker, Hack Database, and Hack Finder are visible.
- Support sheets are not visible in Excel's normal Unhide dialog.
- Refresh Hacks reports the catalog is current after a fresh build.
- Difficulty, Type, and Rating filters appear on Hack Finder.
- Random Hack selects a matching hack.
- Tracker and Hack Finder pull the SMWCentral Rating value when available.
- Custom hacks remain after Refresh Hacks.

## 6. Publish a GitHub Release

Create a release such as `v1.2.0`, attach the generated `.xlsm`, and publish it. Share the release page in Discord.

Users may need to right-click the downloaded `.xlsm`, choose **Properties**, and select **Unblock** before Excel allows its macros to run.
