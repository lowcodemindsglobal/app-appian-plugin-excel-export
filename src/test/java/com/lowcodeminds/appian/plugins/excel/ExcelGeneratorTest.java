package com.lowcodeminds.appian.plugins.excel;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelGeneratorTest {

  private Connection connection;

  @BeforeEach
  void setUp() throws SQLException {
    connection = DriverManager.getConnection("jdbc:h2:mem:excelgen_" + UUID.randomUUID() + ";DB_CLOSE_DELAY=-1");
    try (Statement statement = connection.createStatement()) {
      statement.execute("CREATE TABLE employees (" +
          "employee_id VARCHAR(20), name VARCHAR(100), department VARCHAR(50), " +
          "status VARCHAR(20), hire_date DATE, last_login TIMESTAMP)");
      for (int i = 0; i < 10; i++) {
        statement.execute("INSERT INTO employees VALUES (" +
            "'EMP" + i + "', 'Employee " + i + "', 'Engineering', 'Active', " +
            "DATE '2024-01-01', TIMESTAMP '2024-01-01 09:30:00')");
      }
    }
  }

  @AfterEach
  void tearDown() throws SQLException {
    connection.close();
  }

  private static SQLSheetData sheet(String sheetName, String sql) {
    SQLSheetData data = new SQLSheetData();
    data.setSheetName(sheetName);
    data.setSqlQuery(sql);
    return data;
  }

  private static ExcelExportConfig.Builder baseConfig() {
    return ExcelExportConfig.builder().documentName("export").targetFolderId(1L);
  }

  // XSSFWorkbook parses the entire OOXML tree into memory upfront, so a sheet
  // returned from an already-closed workbook still has valid XMLBeans objects.
  private static XSSFSheet readBack(byte[] excelBytes, int sheetIndex) throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excelBytes))) {
      return workbook.getSheetAt(sheetIndex);
    }
  }

  @Test
  void testBasicGeneration() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of("STATUS")).build();

    byte[] excelBytes = new ExcelGenerator(config).generate(connection,
        List.of(sheet("Employees", "SELECT employee_id, name, department, status FROM employees ORDER BY employee_id")));

    assertTrue(excelBytes.length > 0);
    XSSFSheet sheet = readBack(excelBytes, 0);
    assertEquals(10, sheet.getLastRowNum());
    assertEquals(4, sheet.getRow(0).getLastCellNum());
  }

  @Test
  void testEditableColumnsApplyGloballyAcrossDifferentlyShapedSheets() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of("STATUS", "TOTAL")).build();

    List<SQLSheetData> sheets = List.of(
        sheet("Employees", "SELECT employee_id, name, status FROM employees ORDER BY employee_id"),
        sheet("Summary", "SELECT department, COUNT(*) AS total FROM employees GROUP BY department"));

    byte[] excelBytes = new ExcelGenerator(config).generate(connection, sheets);

    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excelBytes))) {
      assertEquals(2, workbook.getNumberOfSheets());
      assertEquals("Employees", workbook.getSheetName(0));
      assertEquals("Summary", workbook.getSheetName(1));

      XSSFSheet employeesSheet = workbook.getSheetAt(0);
      for (int r = 1; r <= employeesSheet.getLastRowNum(); r++) {
        Row row = employeesSheet.getRow(r);
        assertTrue(row.getCell(0).getCellStyle().getLocked(), "employee_id must stay locked");
        assertFalse(row.getCell(2).getCellStyle().getLocked(), "status must be unlocked");
      }

      XSSFSheet summarySheet = workbook.getSheetAt(1);
      for (int r = 1; r <= summarySheet.getLastRowNum(); r++) {
        Row row = summarySheet.getRow(r);
        assertTrue(row.getCell(0).getCellStyle().getLocked(), "department must stay locked");
        assertFalse(row.getCell(1).getCellStyle().getLocked(), "total must be unlocked");
      }
    }
  }

  @Test
  void testDefaultHeaderColorsAreYellow() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of("STATUS")).build();

    byte[] excelBytes = new ExcelGenerator(config).generate(connection,
        List.of(sheet("Employees", "SELECT employee_id, status FROM employees ORDER BY employee_id")));

    XSSFSheet sheet = readBack(excelBytes, 0);
    Row header = sheet.getRow(0);
    short nonEditableColor = header.getCell(0).getCellStyle().getFillForegroundColor();
    short editableColor = header.getCell(1).getCellStyle().getFillForegroundColor();

    assertEquals(IndexedColors.YELLOW.getIndex(), nonEditableColor, "non-editable header should default to yellow");
    assertEquals(IndexedColors.YELLOW.getIndex(), editableColor, "editable header should always be yellow");
  }

  @Test
  void testDataCellColorsReflectEditability() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of("STATUS")).build();

    byte[] excelBytes = new ExcelGenerator(config).generate(connection,
        List.of(sheet("Employees", "SELECT employee_id, status FROM employees ORDER BY employee_id")));

    XSSFSheet sheet = readBack(excelBytes, 0);
    Row row = sheet.getRow(1);
    short nonEditableColor = row.getCell(0).getCellStyle().getFillForegroundColor();
    short editableColor = row.getCell(1).getCellStyle().getFillForegroundColor();

    assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), nonEditableColor, "non-editable data cells should be light gray");
    assertEquals(IndexedColors.GREY_50_PERCENT.getIndex(), editableColor, "editable data cells should be dark gray");
    assertNotEquals(nonEditableColor, editableColor);
  }

  @Test
  void testCustomNonEditableHeaderColorOverridesDefault() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of()).nonEditableHeaderColor("#336699").build();

    byte[] excelBytes = new ExcelGenerator(config).generate(connection,
        List.of(sheet("Employees", "SELECT employee_id FROM employees ORDER BY employee_id")));

    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excelBytes))) {
      XSSFColor fillColor = ((XSSFCellStyle) workbook.getSheetAt(0).getRow(0).getCell(0).getCellStyle())
          .getFillForegroundColorColor();
      assertArrayEquals(new byte[] {(byte) 0xFF, (byte) 0x33, (byte) 0x66, (byte) 0x99}, fillColor.getARGB());
    }
  }

  @Test
  void testDateAndDateTimeColumnsUseConfiguredFormat() throws Exception {
    ExcelExportConfig config = baseConfig()
        .editableColumns(List.of())
        .dateFormat("yyyy-mm-dd")
        .dateTimeFormat("yyyy-mm-dd hh:mm")
        .build();

    byte[] excelBytes = new ExcelGenerator(config).generate(connection,
        List.of(sheet("Employees", "SELECT hire_date, last_login FROM employees ORDER BY employee_id")));

    XSSFSheet sheet = readBack(excelBytes, 0);
    Row row = sheet.getRow(1);
    assertEquals("yyyy-mm-dd", row.getCell(0).getCellStyle().getDataFormatString());
    assertEquals("yyyy-mm-dd hh:mm", row.getCell(1).getCellStyle().getDataFormatString());
  }

  @Test
  void testDefaultDateAndDateTimeFormatsAreUsedWhenNotConfigured() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of()).build();

    byte[] excelBytes = new ExcelGenerator(config).generate(connection,
        List.of(sheet("Employees", "SELECT hire_date, last_login FROM employees ORDER BY employee_id")));

    XSSFSheet sheet = readBack(excelBytes, 0);
    Row row = sheet.getRow(1);
    assertEquals("dd-mm-yyyy", row.getCell(0).getCellStyle().getDataFormatString());
    assertEquals("dd-mm-yyyy hh:mm:ss", row.getCell(1).getCellStyle().getDataFormatString());
  }

  @Test
  void testEverySheetIsProtectedWithStructuralEditsAllowed() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of()).build();

    List<SQLSheetData> sheets = List.of(
        sheet("One", "SELECT employee_id FROM employees"),
        sheet("Two", "SELECT department FROM employees"));

    byte[] excelBytes = new ExcelGenerator(config).generate(connection, sheets);

    try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(excelBytes))) {
      for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
        XSSFSheet sheet = workbook.getSheetAt(i);
        assertTrue(sheet.getProtect(), "sheet " + i + " should be protected");

        CTSheetProtection protection = sheet.getCTWorksheet().getSheetProtection();
        assertTrue(protection.getInsertRows());
        assertTrue(protection.getDeleteRows());
        assertTrue(protection.getInsertColumns());
        assertTrue(protection.getDeleteColumns());
      }
    }
  }

  @Test
  void testSqlErrorIsWrappedWithSheetName() {
    ExcelExportConfig config = baseConfig().editableColumns(List.of()).build();

    SQLException exception = assertThrows(SQLException.class, () -> new ExcelGenerator(config).generate(connection,
        List.of(sheet("BadSheet", "SELECT nonexistent_column FROM employees"))));

    assertTrue(exception.getMessage().contains("BadSheet"));
  }

  @Test
  void testEncryptRequiresCorrectPasswordToOpen() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of()).build();
    byte[] plainBytes = new ExcelGenerator(config).generate(connection,
        List.of(sheet("Employees", "SELECT employee_id FROM employees")));

    byte[] encryptedBytes = ExcelGenerator.encrypt(plainBytes, "s3cret");

    assertThrows(Exception.class, () -> WorkbookFactory.create(new ByteArrayInputStream(encryptedBytes)));

    try (POIFSFileSystem fileSystem = new POIFSFileSystem(new ByteArrayInputStream(encryptedBytes))) {
      EncryptionInfo encryptionInfo = new EncryptionInfo(fileSystem);
      Decryptor decryptor = Decryptor.getInstance(encryptionInfo);
      assertTrue(decryptor.verifyPassword("s3cret"));

      try (InputStream decryptedStream = decryptor.getDataStream(fileSystem);
          XSSFWorkbook workbook = new XSSFWorkbook(decryptedStream)) {
        assertEquals(1, workbook.getNumberOfSheets());
      }
    }
  }

  @Test
  void testLargeDataset() throws Exception {
    ExcelExportConfig config = baseConfig().editableColumns(List.of("STATUS")).build();
    String sql = "SELECT CAST(X AS VARCHAR) AS employee_id, 'Employee ' || X AS name, "
        + "'Engineering' AS department, 'Active' AS status FROM SYSTEM_RANGE(1, 50000)";

    long start = System.nanoTime();
    byte[] excelBytes = new ExcelGenerator(config).generate(connection, List.of(sheet("Employees", sql)));
    long elapsedMillis = (System.nanoTime() - start) / 1_000_000;

    assertTrue(elapsedMillis < 60_000, "Generation took too long: " + elapsedMillis + "ms");

    XSSFSheet sheet = readBack(excelBytes, 0);
    assertEquals(50_000, sheet.getLastRowNum());
  }
}
