# app-appian-plugin-excel-export

Appian Smart Service plug-in that runs one or more SQL queries against a
JNDI-configured data source and exports the combined results to a
multi-sheet `.xlsx` file with column-level cell protection: designated
columns are editable, all others (including headers) are locked, and
row/column insertion and deletion are blocked.

This project follows the structure and conventions described in Appian's
official plug-in documentation and best-practices guide, cross-checked
against `lowcodemindsglobal/appian-header-match-plugin` (a real,
successfully deployed sibling plug-in from this org) wherever the two
disagreed:
- https://docs.appian.com/suite/help/26.6/Custom_Smart_Service_Plug-ins.html
- https://community.appian.com/success/w/article/3273/plugin-development-best-practices
- https://github.com/lowcodemindsglobal/appian-header-match-plugin

## Build

```
mvn clean package
```

Produces the deployable shaded JAR at
`target/advance-excel-export-1.0.0-SNAPSHOT.jar` (POI is relocated to
`com.lowcodeminds.appian.plugins.excel.shaded.poi` to avoid classpath
conflicts with whatever version Appian's own server bundles internally).

## Appian SDK dependency

`com.appiancorp:appian-plugin-sdk:25.0.0-local` is a local-only Maven
coordinate installed from `lib/appian-plug-in-sdk.jar` (not a published Appian
release version). Re-install it into a fresh local `.m2` with:

```
mvn install:install-file -Dfile=lib/appian-plug-in-sdk.jar \
  -DgroupId=com.appiancorp -DartifactId=appian-plugin-sdk \
  -Dversion=25.0.0-local -Dpackaging=jar
```

## Compatibility

Requires **Appian 24.4 or later** (declared as
`<application-version min="24.4"/>` in `appian-plugin.xml`). This is a
confirmed floor, not a guess: an earlier `min="25.1"` failed to deploy
against a real Appian instance running 24.4.460.0 with "requires a minimum
version of 25.1.0. The installed version of Appian is 24.4.460.0" - the
plugin loader rejects the version gate before loading a single class, so
setting this too high hard-blocks deployment outright. The
`appian-plugin-sdk:25.0.0-local` coordinate this project compiles against
(see below) is just a placeholder version string with no real bearing on the
minimum supported release - don't infer a floor from it. Documenting this is
also required by Appian's
[AppMarket Submission Policies for Plug-Ins](https://docs.appian.com/suite/help/26.6/Shared_Components.html):
plug-ins without clearly documented version compatibility are removed from
the AppMarket, and the policy applies even for private Appian Cloud
deployment, not just public AppMarket listing.

## Logging

The plug-in logs via SLF4J (`org.slf4j.Logger`), `provided` scope - Appian's
application server supplies the actual logging backend. `log4j` is
explicitly banned by the `maven-enforcer-plugin` rule in `pom.xml`, matching
`appian-header-match-plugin`'s own enforcer configuration.

## Test

```
mvn test
```

27 JUnit 5 tests across three classes: `AdvanceExcelExportTest` (smart-service
`validate()` rules - sheet count/name/SQL checks), `ExcelExportConfigTest`
(config building and editable-column matching), and `ExcelGeneratorTest`
(protection logic and encryption, including a 50,000-row generation test).

## Dependency vulnerability scan

```
mvn dependency-check:check
```

Not run automatically as part of `package` (it needs network access to
refresh the CVE database, which would otherwise slow down or break offline
builds) - run it manually before releases, per Appian's plugin best practices
around third-party libraries like Apache POI.

`maven-enforcer-plugin` (bound to the build, so it always runs) additionally
bans known-risky dependencies outright (log4j, netty, jetty, junit 4.x) if
they show up anywhere in the dependency tree, transitively or not.

## Project layout (mirrors Appian's documented plug-in structure)

```
src/main/resources/
  appian-plugin.xml                                        # plug-in descriptor, JAR root
  com/lowcodeminds/plugins/advance-excel-export/
    advanceExcelExport_en_US.properties                     # i18n: labels, tooltips - locale suffix required, see file header comment
    advanceExcelExport/images/
      palette-icon.svg                                      # 27x19, placeholder artwork
      canvas-icon.svg                                        # 60x40, placeholder artwork
```

These resource paths are derived from the plug-in's `key`
(`com.lowcodeminds.plugins.advance-excel-export`), not its display `name`.
Appian resolves the i18n bundle the same way
`java.util.ResourceBundle.getBundle()` always has - by joining
`<plugin-key>.<smart-service-key>` into a single basename and treating
**every dot as a directory separator**, so the path must be fully nested
(`com/lowcodeminds/plugins/advance-excel-export/...`), not a single folder
literally named with dots. See the header comment in
`advanceExcelExport_en_US.properties` for the real deployment failure
(`APNX-1-4200-000`) that a flat, literal-dot folder produced - do not revert
to that layout.

## Deploy

1. Upload `target/advance-excel-export-1.0.0-SNAPSHOT.jar` via the Appian
   Admin Console (Plug-ins section).
2. Confirm the "Excel Export" node appears in the Process Model Designer's
   palette, under Automation Smart Services > Document Generation.
3. Replace the placeholder `palette-icon.svg` / `canvas-icon.svg` with real
   branded artwork before a production release (optional - the reference
   sibling plug-in shipped successfully with no icons at all).
4. The i18n bundle's resource-path convention (see the header comment in
   `advanceExcelExport_en_US.properties`) is confirmed by a real deployment
   failure, not a guess: an earlier version placed the file directly under a
   single folder literally named `com.lowcodeminds.plugins.advance-excel-export`
   (dots kept as literal characters, modeled on an unverified assumption about
   a sibling plug-in's layout) and deployment failed with `APNX-1-4200-000:
   missing internationalization bundle(s) for Locale en_US`. The bundle must
   instead live at the fully nested path
   `com/lowcodeminds/plugins/advance-excel-export/advanceExcelExport_en_US.properties`,
   matching how `java.util.ResourceBundle.getBundle()` always treats dots in a
   basename - every dot is a directory separator, none are literal. If labels
   ever regress to auto-generated fallbacks again, check this file's path first.

## Input contract

`AdvanceExcelExport` takes one query per output sheet, run directly against a
JNDI-configured data source - it does not take pre-fetched rows:

| Input | Required | Description |
|---|---|---|
| `SQLSheetDataList` (`SQLSheetData[]`) | Always | One entry per output sheet: a `sqlQuery` (`SELECT` only, no semicolons) and its `sheetName`. Max 20 entries, unique sheet names ≤31 chars, no illegal Excel characters. |
| `jndiName` | Always | JNDI name of the pre-configured data source each query runs against. |
| `editableColumns` (`String[]`) | Optional | Column names (matched case-insensitively across all sheets) left unlocked; everything else, including headers, is locked. |
| `targetFolder` | Always | Appian folder to save the generated document in. |
| `NewDocumentName` | Always | Document name (no extension). |
| `DateFormat` / `DateTimeFormat` | Optional | Excel number formats for DATE/TIMESTAMP columns (default `dd-mm-yyyy` / `dd-mm-yyyy hh:mm:ss`). |
| `NoneEditableHeaderColor` | Optional | Hex fill color for locked-column headers. |
| `ExcelPassword` | Optional | If set, encrypts the whole file for opening, in addition to the always-on cell/structural protection. |

`SQLSheetData` is a plug-in-defined CDT (`<datatype key="SQLSheetData">` in
`appian-plugin.xml`, backed by the JAXB-annotated
`com.lowcodeminds.appian.plugins.excel.SQLSheetData`), not a JSON string -
Appian's plugin loader supports custom CDT arrays as an `@Input` type
directly, no serialization workaround needed.

This wasn't the first design tried. An early version accepted pre-fetched
rows as a raw `Map<String, Object>[]` (Dictionary array), which passed local
unit tests and `mvn package` cleanly but **failed to deploy** against a real
Appian server with:

```
com.appiancorp.suiteapi.type.exceptions.InvalidTypeException: ...
Invalid Type: Unsupported type [Ljava.util.Map; (APNX-1-4047-000)
```

Appian's plugin loader scans every `@Input`/`@Output` method against a fixed
table of Java-to-Appian type mappings, and a raw `Map[]` isn't in it - one bad
input type disables the *entire* node at load time, regardless of the other
inputs being fine. That design was replaced first with a JSON-encoded
`dataRowsJson` string (parsed with Jackson - the same `String`-input
workaround `appian-header-match-plugin`'s `existingMappingsJson` uses), and
then, once the current SQL/JNDI, multi-sheet requirement was settled, with
the `SQLSheetData[]` CDT design described above.

## Known caveat

`AdvanceExcelExport` uses `ContentService.uploadDocument()` /
`ContentUploadOutputStream` to save the generated file. An earlier version used
the now-deprecated `ContentService.upload()` / `ContentOutputStream` pair,
based on the mistaken assumption that `Document.write(InputStream)` /
`getOutputStream()` was the only non-deprecated alternative (that path
delegates to an internal `documentHelper` field the Appian engine only
populates on framework-returned `Document` instances, so it throws
`NullPointerException` on a plugin-constructed `Document`). An Appian AppMarket
reviewer pointed out `uploadDocument()` as the correct, non-deprecated
replacement, and the code was switched over accordingly.
