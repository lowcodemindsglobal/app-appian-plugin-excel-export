# Smart Service Icon Rendering & Multi-Sheet Parameter Notes

Date: 2026-07-22

Notes from investigating a broken palette/canvas icon after Appian Cloud
deployment, and clarifying how the `SQLSheetDataList` and `editableColumns`
parameters behave for multi-sheet exports.

## 1. Smart service icon breaks after deployment

**Icon resolution mechanism:** `AdvanceExcelExport.java` has no `@PaletteInfo`
annotation. Appian resolves the palette/canvas icons purely by file-path
convention under:

```
src/main/resources/com/lowcodeminds/plugins/advance-excel-export/advanceExcelExport/images/
  palette-icon.svg   (27x19 px)
  canvas-icon.svg    (60x40 px)
```

The path segments map to `<plugin-key>/<smart-service-key>/images/...`, per
the keys declared in `src/main/resources/appian-plugin.xml`.

**Root cause found:** both SVG files open with a leading `<!-- ... -->`
comment block *before* the `<svg>` root element:

```xml
<!--
  Placeholder palette icon - replace with real branded artwork before release.
  Required size per Appian docs: 27x19 px. The leading <?xml ...?> declaration
  must NOT be present, or the icon fails to render in the Process Modeler palette.
-->
<svg xmlns="http://www.w3.org/2000/svg" width="27" height="19" viewBox="0 0 27 19">
```

The file's own comment warns that nothing may precede `<svg>` (referring to
an XML declaration), but a leading comment is itself content preceding
`<svg>` — the same category of problem. Appian's icon loader appears not to
tolerate this, even though most SVG viewers/browsers skip leading comments
silently, which is why the defect can go unnoticed until the icon is viewed
inside actual Appian UI (palette / canvas), e.g. after a fresh deployment.

**Fix:** remove the leading comment block from both `palette-icon.svg` and
`canvas-icon.svg` so the file's first bytes are `<svg ...>`. Move any
documentation about sizing requirements into the README or `pom.xml` instead
of inside the SVG. Rebuild and redeploy the plugin, then verify the icon
renders in both the Process Modeler palette and the canvas.

Also confirmed (not the cause, but ruled out during investigation):
- No Maven resource filtering is enabled on `src/main/resources` in
  `pom.xml`, so binary/text corruption of the SVGs during the build is not
  the issue.
- `maven-shade-plugin`'s excludes only target `META-INF/*.SF|.DSA|.RSA`
  signature files, not `images/*.svg`.
- The SVGs are explicitly flagged in `README.md` as placeholder artwork,
  never confirmed production-ready — worth replacing with real branded
  icons regardless of this fix.

**Status:** fix identified, not yet applied as of this writing — pending
confirmation to edit the two SVG files.

## 2. Multiple sheets via `SQLSheetDataList`

`SQLSheetDataList` is declared as `SQLSheetData[]` on the smart service
(`AdvanceExcelExport.java`), i.e. an array, not a single object — capped at
20 entries (`MAX_SHEETS = 20`). Each element is a `{sheetName, sqlQuery}`
pair (`SQLSheetData.java`).

This means passing an array of dictionaries is fully supported and is the
intended way to export multiple sheets in one call:

```
{
  {
    sheetName: "Testing1",
    sqlQuery: rule!HHS_dynamicSqlForExcelExport()
  },
  {
    sheetName: "Testing2",
    sqlQuery: rule!HHS_dynamicSqlForExcelExport()
  },
  {
    sheetName: "Testing3",
    sqlQuery: rule!HHS_dynamicSqlForExcelExport()
  }
}
```

## 3. `editableColumns` is global, not per-sheet

`editableColumns` is a single flat `String[]` parameter on the smart
service — it is **not** nested inside `SQLSheetData`, and it is **not** a
parallel list indexed against the sheets.

Matching behavior: `ExcelExportConfig.isEditable(columnName)` matches by
column name, case-insensitively, across **every** sheet. So for a 3-sheet
export, supply one combined list of every column name that should be
editable anywhere, e.g. `{"Amount", "Comments", "Status"}` — any sheet
containing a column with a matching name gets it unlocked.

**Caveat:** because matching is name-only with no sheet scoping, if two
sheets share a column name (e.g. both have "Amount") but only one sheet
should have it editable, this cannot currently be expressed — the column
will be unlocked in both sheets wherever the name matches.

**Making it sheet-specific (not yet implemented):** would require:
1. Adding an `editableColumns` field to the `SQLSheetData` CDT itself, so
   each sheet dictionary carries its own list, e.g.
   `{sheetName: "Testing1", sqlQuery: ..., editableColumns: {"Amount"}}`.
2. Updating `ExcelExportConfig`/`ExcelGenerator` so `isEditable(...)` checks
   the current sheet's own list instead of one global list.
3. Deciding whether to keep the existing global `editableColumns` parameter
   as a fallback for sheets that don't specify their own (backward
   compatibility with existing process models), or remove it.

**Status:** design discussed, not yet implemented — open decision on
whether to keep the global parameter as a fallback.
