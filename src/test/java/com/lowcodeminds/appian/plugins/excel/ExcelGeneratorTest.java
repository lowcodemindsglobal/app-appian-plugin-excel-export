package com.lowcodeminds.appian.plugins.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelGeneratorTest {

  private static final List<String> ALL_COLUMNS =
      Arrays.asList("EmployeeID", "Name", "Department", "Status", "Comments");
  private static final List<String> EDITABLE_COLUMNS = Arrays.asList("Status", "Comments");

  // XSSFWorkbook parses the entire OOXML tree into memory upfront, so the returned
  // sheet's XMLBeans objects remain valid after the OPCPackage backing the workbook
  // is closed here.
  private static XSSFSheet readBack(byte[] excelBytes) throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excelBytes))) {
      return workbook.getSheetAt(0);
    }
  }

  private static List<Map<String, Object>> sampleRows(int count) {
    List<Map<String, Object>> rows = new ArrayList<>();
    for (int i = 0; i < count; i++) {
      Map<String, Object> row = new LinkedHashMap<>();
      row.put("EmployeeID", "EMP" + i);
      row.put("Name", "Employee " + i);
      row.put("Department", "Engineering");
      row.put("Status", "Active");
      row.put("Comments", "");
      rows.add(row);
    }
    return rows;
  }

  private static ColumnProtectionConfig.Builder baseConfig() {
    return ColumnProtectionConfig.builder()
        .allColumns(ALL_COLUMNS)
        .fileName("export.xlsx")
        .targetFolderId(1L);
  }

  @Test
  void testBasicGeneration() throws IOException {
    ColumnProtectionConfig config = baseConfig().editableColumns(EDITABLE_COLUMNS).build();
    byte[] excelBytes = new ExcelGenerator(config).generate(sampleRows(10));

    assertTrue(excelBytes.length > 0);

    XSSFSheet sheet = readBack(excelBytes);
    assertEquals(10, sheet.getLastRowNum());
    assertEquals(ALL_COLUMNS.size(), sheet.getRow(0).getLastCellNum());
  }

  @Test
  void testEditableColumnsAreUnlocked() throws IOException {
    ColumnProtectionConfig config = baseConfig().editableColumns(EDITABLE_COLUMNS).build();
    byte[] excelBytes = new ExcelGenerator(config).generate(sampleRows(10));
    XSSFSheet sheet = readBack(excelBytes);

    for (int r = 1; r <= sheet.getLastRowNum(); r++) {
      Row row = sheet.getRow(r);
      for (String editableColumn : EDITABLE_COLUMNS) {
        int colIndex = ALL_COLUMNS.indexOf(editableColumn);
        Cell cell = row.getCell(colIndex);
        assertFalse(cell.getCellStyle().getLocked(),
            "Expected " + editableColumn + " to be unlocked at row " + r);
      }
    }
  }

  @Test
  void testNonEditableColumnsAreLocked() throws IOException {
    ColumnProtectionConfig config = baseConfig().editableColumns(EDITABLE_COLUMNS).build();
    byte[] excelBytes = new ExcelGenerator(config).generate(sampleRows(10));
    XSSFSheet sheet = readBack(excelBytes);

    List<String> nonEditable = Arrays.asList("EmployeeID", "Name", "Department");
    for (int r = 1; r <= sheet.getLastRowNum(); r++) {
      Row row = sheet.getRow(r);
      for (String column : nonEditable) {
        int colIndex = ALL_COLUMNS.indexOf(column);
        Cell cell = row.getCell(colIndex);
        assertTrue(cell.getCellStyle().getLocked(),
            "Expected " + column + " to be locked at row " + r);
      }
    }
  }

  @Test
  void testHeadersAlwaysLocked() throws IOException {
    ColumnProtectionConfig config = baseConfig().editableColumns(EDITABLE_COLUMNS).build();
    byte[] excelBytes = new ExcelGenerator(config).generate(sampleRows(5));
    XSSFSheet sheet = readBack(excelBytes);

    Row headerRow = sheet.getRow(0);
    for (int c = 0; c < ALL_COLUMNS.size(); c++) {
      assertTrue(headerRow.getCell(c).getCellStyle().getLocked(),
          "Expected header cell " + c + " to be locked, including editable-column headers");
    }
  }

  @Test
  void testSheetIsProtected() throws IOException {
    ColumnProtectionConfig config = baseConfig().editableColumns(EDITABLE_COLUMNS).build();
    byte[] excelBytes = new ExcelGenerator(config).generate(sampleRows(5));
    XSSFSheet sheet = readBack(excelBytes);

    assertTrue(sheet.getProtect());
  }

  @Test
  void testRowInsertionBlocked() throws IOException {
    ColumnProtectionConfig config = baseConfig().editableColumns(EDITABLE_COLUMNS).build();
    byte[] excelBytes = new ExcelGenerator(config).generate(sampleRows(5));
    XSSFSheet sheet = readBack(excelBytes);

    CTSheetProtection protection = sheet.getCTWorksheet().getSheetProtection();
    assertTrue(protection.getInsertRows());
    assertTrue(protection.getDeleteRows());
    assertTrue(protection.getInsertColumns());
    assertTrue(protection.getDeleteColumns());
  }

  @Test
  void testLargeDataset() throws IOException {
    int rowCount = 50_000;
    ColumnProtectionConfig config = baseConfig().editableColumns(EDITABLE_COLUMNS).build();
    List<Map<String, Object>> rows = sampleRows(rowCount);

    long start = System.nanoTime();
    byte[] excelBytes = new ExcelGenerator(config).generate(rows);
    long elapsedMillis = (System.nanoTime() - start) / 1_000_000;

    assertTrue(elapsedMillis < 60_000, "Generation took too long: " + elapsedMillis + "ms");
    assertTrue(excelBytes.length < 20L * 1024 * 1024, "File size exceeded 20MB: " + excelBytes.length);

    XSSFSheet sheet = readBack(excelBytes);
    assertEquals(rowCount, sheet.getLastRowNum());
    assertEquals("EMP0", sheet.getRow(1).getCell(0).getStringCellValue());
    assertEquals("EMP" + (rowCount - 1), sheet.getRow(rowCount).getCell(0).getStringCellValue());
  }

  @Test
  void testEmptyEditableColumns() throws IOException {
    ColumnProtectionConfig config = baseConfig().editableColumns(List.of()).build();
    byte[] excelBytes = new ExcelGenerator(config).generate(sampleRows(5));
    XSSFSheet sheet = readBack(excelBytes);

    for (int r = 1; r <= sheet.getLastRowNum(); r++) {
      Row row = sheet.getRow(r);
      for (int c = 0; c < ALL_COLUMNS.size(); c++) {
        assertTrue(row.getCell(c).getCellStyle().getLocked(),
            "Expected all cells locked when editableColumns is empty");
      }
    }
  }

  @Test
  void testNullPassword() throws IOException {
    ColumnProtectionConfig config = baseConfig()
        .editableColumns(EDITABLE_COLUMNS)
        .protectionPassword(null)
        .build();
    byte[] excelBytes = new ExcelGenerator(config).generate(sampleRows(5));
    XSSFSheet sheet = readBack(excelBytes);

    assertTrue(sheet.getProtect());
  }
}
