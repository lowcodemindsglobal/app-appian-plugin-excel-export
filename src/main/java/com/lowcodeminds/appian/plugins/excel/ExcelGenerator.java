package com.lowcodeminds.appian.plugins.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Generates a column-protected .xlsx workbook from row data.
 *
 * Uses SXSSFWorkbook (streaming) so memory stays flat regardless of row count -
 * required to support the 50,000-row target without OutOfMemoryError.
 */
public final class ExcelGenerator {

  private static final Logger LOG = LoggerFactory.getLogger(ExcelGenerator.class);
  private static final int ROW_ACCESS_WINDOW_SIZE = 100;
  private static final String SHEET_NAME = "Export";
  private static final int DEFAULT_COLUMN_WIDTH = 5000;

  private final ColumnProtectionConfig config;

  public ExcelGenerator(ColumnProtectionConfig config) {
    this.config = config;
  }

  public byte[] generate(List<Map<String, Object>> dataRows) {
    SXSSFWorkbook workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE);
    try {
      SXSSFSheet sheet = workbook.createSheet(SHEET_NAME);
      sheet.trackAllColumnsForAutoSizing();

      CellStyle lockedStyle = createLockedStyle(workbook);
      CellStyle unlockedStyle = createUnlockedStyle(workbook);
      CellStyle headerStyle = createHeaderStyle(workbook);

      List<String> columns = config.getAllColumns();
      writeHeaderRow(sheet, columns, headerStyle);
      writeDataRows(sheet, columns, dataRows, lockedStyle, unlockedStyle);

      for (int i = 0; i < columns.size(); i++) {
        sheet.setColumnWidth(i, DEFAULT_COLUMN_WIDTH);
      }

      // Protection must be applied after all cells are written -
      // applying it earlier can interfere with cell writes.
      protectSheet(sheet);

      ByteArrayOutputStream out = new ByteArrayOutputStream();
      workbook.write(out);
      return out.toByteArray();
    } catch (IOException e) {
      LOG.error("Failed to generate Excel workbook", e);
      throw new ExcelGenerationException("Failed to generate Excel workbook", e);
    } finally {
      workbook.dispose();
      try {
        workbook.close();
      } catch (IOException e) {
        LOG.warn("Failed to close workbook", e);
      }
    }
  }

  private CellStyle createLockedStyle(Workbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setLocked(true);
    return style;
  }

  private CellStyle createUnlockedStyle(Workbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setLocked(false);
    return style;
  }

  private CellStyle createHeaderStyle(Workbook workbook) {
    CellStyle style = workbook.createCellStyle();
    style.setLocked(true);
    Font font = workbook.createFont();
    font.setBold(true);
    style.setFont(font);
    style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    return style;
  }

  private void writeHeaderRow(Sheet sheet, List<String> columns, CellStyle headerStyle) {
    Row headerRow = sheet.createRow(0);
    for (int c = 0; c < columns.size(); c++) {
      Cell cell = headerRow.createCell(c);
      cell.setCellValue(columns.get(c));
      cell.setCellStyle(headerStyle);
    }
  }

  private void writeDataRows(Sheet sheet, List<String> columns, List<Map<String, Object>> dataRows,
      CellStyle lockedStyle, CellStyle unlockedStyle) {
    for (int r = 0; r < dataRows.size(); r++) {
      Row row = sheet.createRow(r + 1);
      Map<String, Object> record = dataRows.get(r);
      for (int c = 0; c < columns.size(); c++) {
        String columnName = columns.get(c);
        Cell cell = row.createCell(c);
        setCellValue(cell, record.get(columnName));
        cell.setCellStyle(config.isEditable(columnName) ? unlockedStyle : lockedStyle);
      }
    }
  }

  private void setCellValue(Cell cell, Object value) {
    if (value == null) {
      cell.setBlank();
    } else if (value instanceof Number) {
      cell.setCellValue(((Number) value).doubleValue());
    } else if (value instanceof Boolean) {
      cell.setCellValue((Boolean) value);
    } else if (value instanceof Date) {
      cell.setCellValue((Date) value);
    } else {
      cell.setCellValue(String.valueOf(value));
    }
  }

  /**
   * SXSSFSheet does not publicly expose the underlying XSSFSheet needed to reach
   * CTSheetProtection, so the package-private field "_sh" (verified against
   * Apache POI 5.2.5) is accessed via reflection. This is the documented
   * workaround for blocking row/column insertion and deletion under SXSSF.
   */
  private void protectSheet(SXSSFSheet sheet) {
    String password = config.getProtectionPassword();
    sheet.protectSheet(password != null ? password : "");

    try {
      Field shField = SXSSFSheet.class.getDeclaredField("_sh");
      shField.setAccessible(true);
      XSSFSheet xssfSheet = (XSSFSheet) shField.get(sheet);
      CTSheetProtection protection = xssfSheet.getCTWorksheet().getSheetProtection();
      protection.setInsertRows(true);
      protection.setInsertColumns(true);
      protection.setDeleteRows(true);
      protection.setDeleteColumns(true);
    } catch (ReflectiveOperationException e) {
      throw new ExcelGenerationException("Failed to apply sheet structural protection", e);
    }
  }
}
