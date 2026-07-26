# One-time setup

## 1. Upload this project to GitHub

1. Create a **public** GitHub repository.
2. Upload the contents of this folder to the repository root.
3. Open **Actions** and enable the `Update SMWCentral Catalog` workflow.
4. Run it manually once using **Run workflow**.
5. Open `data/SMWCentral_All_Moderated_Hacks.csv`, click **Raw**, and copy the raw URL.

The workflow also runs daily at 09:15 UTC.

## 2. Configure the workbook

1. Open `excel/Hacklist_Cleaned_and_Styled(1).xlsx`.
2. Paste the raw CSV URL into `Dashboard!J10`.

The workbook already contains:

- `Online Database`: replaceable official catalog.
- `Manual Database`: local custom hacks; never cleared by refresh.
- `Hack Database`: combined lookup data used by the existing tracker formulas.
- Existing `Tracker`, `Dashboard`, `Lists`, `Hack Finder`, and `Read Me` names.

## 3. Optional in-sheet buttons

Excel `.xlsx` cannot store VBA code. To use the actual **Refresh Hacks** and **Add Custom Hack** buttons:

1. Open the prepared `.xlsx` file.
2. Use **File > Save As** and select **Excel Macro-Enabled Workbook (`.xlsm`)**.
3. Press `Alt+F11`.
4. Select **File > Import File** and import `RefreshHacks.bas`.
5. Press `Alt+F8`, select `InstallButtons`, and click **Run**.
6. Save the workbook.

The base filename and every sheet/column name remain unchanged; only the extension changes from `.xlsx` to `.xlsm` because Microsoft does not allow VBA inside `.xlsx` files.

## What the refresh does

- Downloads the latest official CSV.
- Replaces only `Online Database`.
- Preserves `Manual Database`.
- Rebuilds `Hack Database` from both sources.
- Refreshes the dropdown list.
- Does not touch Tracker dates, ratings, notes, playtime, or completion data.

## Add a custom hack

Complete the form in Dashboard cells `J13:J20`, then click **Add Custom Hack**. The macro assigns a `MANUAL-...` ID and stores the entry in `Manual Database`.
