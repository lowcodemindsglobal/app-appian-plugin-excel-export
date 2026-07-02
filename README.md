# app-appian-plugin-excel-export

Appian Smart Service plug-in that exports data rows to a `.xlsx` file with
column-level cell protection: designated columns are editable, all others
(including headers) are locked, and row/column insertion and deletion are
blocked.

## Build

```
mvn clean package
```

Produces the deployable shaded JAR at
`target/excel-export-plugin-1.0.0-SNAPSHOT.jar` (POI is relocated to
`com.lowcodeminds.appian.plugins.excel.shaded.*` to avoid classpath conflicts
with Appian's bundled POI version).

## Appian SDK dependency

`com.appiancorp:appian-plugin-sdk:25.0.0-local` is a local-only Maven
coordinate installed from `lib/appian-plug-in-sdk.jar` (not a published Appian
release version). Re-install it into a fresh local `.m2` with:

```
mvn install:install-file -Dfile=lib/appian-plug-in-sdk.jar \
  -DgroupId=com.appiancorp -DartifactId=appian-plugin-sdk \
  -Dversion=25.0.0-local -Dpackaging=jar
```

## Test

```
mvn test
```

13 JUnit 5 tests cover `ColumnProtectionConfig` validation and `ExcelGenerator`
protection logic, including a 50,000-row generation test.

## Deploy

1. Upload `target/excel-export-plugin-1.0.0-SNAPSHOT.jar` via the Appian
   Admin Console (Plug-ins section).
2. Confirm the "Excel Export with Protection" node appears in the Process
   Model Designer's Smart Service palette.
3. **Before first deploy**, verify `src/main/resources/META-INF/appian-plugin.xml`
   against your Appian version's plugin documentation - its schema was not
   independently verified against the SDK jar (see comments in the file).

## Known caveat

`ExcelExportSmartService` uses `ContentService.upload()` /
`ContentOutputStream` to save the generated file, even though the SDK marks
both `@Deprecated`. The non-deprecated alternative
(`Document.write(InputStream)` / `getOutputStream()`) delegates to an
internal `documentHelper` field that the Appian engine only populates on
framework-returned `Document` instances - calling it on a plugin-constructed
`Document` throws `NullPointerException`. `upload()` remains the only working
path for plugins to push byte content in this SDK version.
