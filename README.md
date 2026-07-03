# app-appian-plugin-excel-export

Appian Smart Service plug-in that exports data rows to a `.xlsx` file with
column-level cell protection: designated columns are editable, all others
(including headers) are locked, and row/column insertion and deletion are
blocked.

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
`target/excel-export-plugin-1.0.1-SNAPSHOT.jar` (POI and Jackson are both
relocated to `com.lowcodeminds.appian.plugins.excel.shaded.*` to avoid
classpath conflicts with whatever versions of those libraries Appian's own
server bundles internally).

## Appian SDK dependency

`com.appiancorp:appian-plugin-sdk:25.0.0-local` is a local-only Maven
coordinate installed from `lib/appian-plug-in-sdk.jar` (not a published Appian
release version). Re-install it into a fresh local `.m2` with:

```
mvn install:install-file -Dfile=lib/appian-plug-in-sdk.jar \
  -DgroupId=com.appiancorp -DartifactId=appian-plugin-sdk \
  -Dversion=25.0.0-local -Dpackaging=jar
```

## Logging

The plug-in logs via SLF4J (`org.slf4j.Logger`), `provided` scope - Appian's
application server supplies the actual logging backend. `log4j` is
explicitly banned by the `maven-enforcer-plugin` rule in `pom.xml`, matching
`appian-header-match-plugin`'s own enforcer configuration.

## Test

```
mvn test
```

13 JUnit 5 tests cover `ColumnProtectionConfig` validation and `ExcelGenerator`
protection logic, including a 50,000-row generation test.

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
  com/lowcodeminds/plugins/excel-export-plugin-v2/
    excelExportSmartService.properties                      # i18n: labels, tooltips (no locale suffix - see file header comment)
  com.lowcodeminds.plugins.excel-export-plugin-v2/
    excelExportSmartService/images/
      palette-icon.svg                                      # 27x19, placeholder artwork
      canvas-icon.svg                                        # 60x40, placeholder artwork
```

## Deploy

1. Upload `target/excel-export-plugin-1.0.1-SNAPSHOT.jar` via the Appian
   Admin Console (Plug-ins section).
2. Confirm the "Excel Export with Protection" node appears in the Process
   Model Designer's palette, under Automation Smart Services > Document
   Generation.
3. Replace the placeholder `palette-icon.svg` / `canvas-icon.svg` with real
   branded artwork before a production release (optional - the reference
   sibling plug-in shipped successfully with no icons at all).
4. The i18n bundle's resource-path convention (see comments in
   `excelExportSmartService.properties`) was cross-checked against
   `appian-header-match-plugin`'s real, working bundle, not independently
   verified against the SDK jar's bytecode the way the Java classes were. If
   input/output labels don't render in the Process Modeler, check this file
   first.

## Input contract: `dataRowsJson` is a JSON string, not a Dictionary array

The calling process model must pass the export rows as a JSON-encoded string
(e.g. `a!toJson(rows)`), not as a native Appian "List of Dictionary" value.

This was discovered the hard way: an earlier version declared this input as
`Map<String, Object>[]`, which passed local unit tests and `mvn package`
cleanly but **failed to deploy** against a real Appian server with:

```
com.appiancorp.suiteapi.type.exceptions.InvalidTypeException: ...
Invalid Type: Unsupported type [Ljava.util.Map; (APNX-1-4047-000)
```

Appian's plugin loader scans every `@Input`/`@Output` method against a fixed
table of Java-to-Appian type mappings, and a raw `Map[]` isn't in it - nor
does the SDK expose an annotation plugin authors can use to declare one
(only pre-built convenience annotations like `@DocumentDataType` exist; none
covers Dictionary). One bad input type disables the *entire* node at load
time, regardless of the other inputs being fine.

`ExcelExportSmartService` parses `dataRowsJson` with Jackson
(`ObjectMapper.readValue`) inside `run()` instead - `String` is unambiguously
supported, sidestepping the problem entirely. Same pattern as
`appian-header-match-plugin`'s `existingMappingsJson` input.

## Known caveat

`ExcelExportSmartService` uses `ContentService.upload()` /
`ContentOutputStream` to save the generated file, even though the SDK marks
both `@Deprecated`. The non-deprecated alternative
(`Document.write(InputStream)` / `getOutputStream()`) delegates to an
internal `documentHelper` field that the Appian engine only populates on
framework-returned `Document` instances - calling it on a plugin-constructed
`Document` throws `NullPointerException`. `upload()` remains the only working
path for plugins to push byte content in this SDK version.
