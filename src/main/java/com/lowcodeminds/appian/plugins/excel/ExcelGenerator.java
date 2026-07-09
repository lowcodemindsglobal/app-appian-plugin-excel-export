package com.lowcodeminds.appian.plugins.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.security.GeneralSecurityException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.List;

/**
 * Generates a multi-sheet, column-protected .xlsx workbook by running one SQL
 * query per sheet against a shared JDBC connection.
 *
 * Uses SXSSFWorkbook (streaming) so memory stays flat regardless of row count -
 * the row access window below keeps only the most recent rows in memory per
 * sheet, flushing older ones to a temp file automatically.
 */
public final class ExcelGenerator {

  private static final Logger LOG = LoggerFactory.getLogger(ExcelGenerator.class);

  private static final int ROW_ACCESS_WINDOW_SIZE = 100;
  private static final int FETCH_SIZE = 100;
  // A Smart Service node runs synchronously inside a process instance's own
  // execution thread, so a runaway per-sheet query (e.g. a missing join
  // predicate) has no other ceiling on it without this.
  private static final int QUERY_TIMEOUT_SECONDS = 120;

  // POI column widths are in units of 1/256th of a character width, so 5000
  // is roughly 19-20 characters wide - a reasonable default for most columns.
  private static final int DEFAULT_COLUMN_WIDTH = 5000;

  // Default header fill when NoneEditableHeaderColor isn't supplied. Editable
  // column headers always use the fixed color below - there's no separate
  // input to override that side, only the non-editable one.
  private static final short DEFAULT_NON_EDITABLE_HEADER_COLOR = IndexedColors.YELLOW.getIndex();
  private static final short EDITABLE_HEADER_COLOR = IndexedColors.YELLOW.getIndex();

  // Data cell fill colors: non-editable (locked) cells get a light gray
  // background, editable (unlocked) cells get a dark gray background.
  private static final short NON_EDITABLE_DATA_COLOR = IndexedColors.GREY_25_PERCENT.getIndex();
  private static final short EDITABLE_DATA_COLOR = IndexedColors.GREY_50_PERCENT.getIndex();

  private final ExcelExportConfig config;

  public ExcelGenerator(ExcelExportConfig config) {
    this.config = config;
  }

  /**
   * Runs each entry in {@code sheets} as a query against {@code connection} and
   * writes its results into its own sheet, in order.
   */
  public byte[] generate(Connection connection, List<SQLSheetData> sheets) throws SQLException {
    SXSSFWorkbook workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE);
    try {
      CellStyle nonEditableHeaderStyle = createNonEditableHeaderStyle(workbook);
      CellStyle editableHeaderStyle = createHeaderStyle(workbook, EDITABLE_HEADER_COLOR);
      DataStyles dataStyles = new DataStyles(workbook, config.getDateFormat(), config.getDateTimeFormat());

      for (SQLSheetData sheetData : sheets) {
        writeSheet(workbook, connection, sheetData, nonEditableHeaderStyle, editableHeaderStyle, dataStyles);
      }

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

  /**
   * Encrypts an already-generated .xlsx file so a password is required to open it,
   * via Apache POI's standard agile-encryption re-wrap: the plain OOXML package is
   * read back and re-saved into an encrypted POIFS container.
   */
  public static byte[] encrypt(byte[] xlsxBytes, String password) throws IOException, GeneralSecurityException, InvalidFormatException {
    EncryptionInfo encryptionInfo = new EncryptionInfo(EncryptionMode.agile);
    Encryptor encryptor = encryptionInfo.getEncryptor();
    encryptor.confirmPassword(password);

    try (POIFSFileSystem fileSystem = new POIFSFileSystem()) {
      try (OPCPackage opcPackage = OPCPackage.open(new ByteArrayInputStream(xlsxBytes));
          OutputStream encryptedStream = encryptor.getDataStream(fileSystem)) {
        opcPackage.save(encryptedStream);
      }

      ByteArrayOutputStream out = new ByteArrayOutputStream();
      fileSystem.writeFilesystem(out);
      return out.toByteArray();
    }
  }

  private void writeSheet(SXSSFWorkbook workbook, Connection connection, SQLSheetData sheetData,
      CellStyle nonEditableHeaderStyle, CellStyle editableHeaderStyle, DataStyles dataStyles) throws SQLException {
    String sheetName = sheetData.getSheetName();
    LOG.debug("Start processing sheet: {}", sheetName);

    try (PreparedStatement statement = connection.prepareStatement(sheetData.getSqlQuery())) {
      statement.setFetchSize(FETCH_SIZE);
      statement.setQueryTimeout(QUERY_TIMEOUT_SECONDS);
      try (ResultSet resultSet = statement.executeQuery()) {
        ResultSetMetaData metadata = resultSet.getMetaData();
        int columnCount = metadata.getColumnCount();

        String[] columnLabels = new String[columnCount];
        boolean[] editableFlags = new boolean[columnCount];
        for (int i = 0; i < columnCount; i++) {
          columnLabels[i] = metadata.getColumnLabel(i + 1);
          editableFlags[i] = config.isEditable(columnLabels[i]);
        }

        SXSSFSheet sheet = workbook.createSheet(sheetName);
        writeHeaderRow(sheet, columnLabels, editableFlags, nonEditableHeaderStyle, editableHeaderStyle);
        writeDataRows(sheet, resultSet, metadata, columnCount, editableFlags, dataStyles);

        for (int i = 0; i < columnCount; i++) {
          sheet.setColumnWidth(i, DEFAULT_COLUMN_WIDTH);
        }

        // Protection must be applied after all cells are written -
        // applying it earlier can interfere with cell writes.
        protectSheet(sheet);
        LOG.info("Completed processing sheet: {}", sheetName);
      }
    } catch (SQLException e) {
      throw new SQLException("Error processing sheet '" + sheetName + "': " + e.getMessage(), e);
    }
  }

  private void writeHeaderRow(SXSSFSheet sheet, String[] columnLabels, boolean[] editableFlags,
      CellStyle nonEditableHeaderStyle, CellStyle editableHeaderStyle) {
    Row headerRow = sheet.createRow(0);
    for (int c = 0; c < columnLabels.length; c++) {
      Cell cell = headerRow.createCell(c);
      cell.setCellValue(columnLabels[c]);
      cell.setCellStyle(editableFlags[c] ? editableHeaderStyle : nonEditableHeaderStyle);
    }
  }

  private void writeDataRows(SXSSFSheet sheet, ResultSet resultSet, ResultSetMetaData metadata, int columnCount,
      boolean[] editableFlags, DataStyles dataStyles) throws SQLException {
    int rowIndex = 1;
    while (resultSet.next()) {
      Row row = sheet.createRow(rowIndex);
      for (int c = 0; c < columnCount; c++) {
        int col = c + 1;
        Cell cell = row.createCell(c);
        setCellValue(cell, resultSet, metadata.getColumnType(col), col);
        cell.setCellStyle(dataStyles.styleFor(metadata.getColumnType(col), editableFlags[c]));
      }
      rowIndex++;
    }
  }

  private void setCellValue(Cell cell, ResultSet resultSet, int columnType, int col) throws SQLException {
    switch (columnType) {
      case Types.BIT:
      case Types.BOOLEAN:
        cell.setCellValue(resultSet.getBoolean(col) ? "Yes" : "No");
        break;
      case Types.DATE:
        java.sql.Date date = resultSet.getDate(col);
        if (date == null) {
          cell.setBlank();
        } else {
          cell.setCellValue(date);
        }
        break;
      case Types.TIMESTAMP:
        java.sql.Timestamp timestamp = resultSet.getTimestamp(col);
        if (timestamp == null) {
          cell.setBlank();
        } else {
          cell.setCellValue(timestamp);
        }
        break;
      case Types.DOUBLE:
      case Types.FLOAT:
      case Types.DECIMAL:
      case Types.NUMERIC:
      case Types.REAL:
        double doubleValue = resultSet.getDouble(col);
        if (resultSet.wasNull()) {
          cell.setBlank();
        } else {
          cell.setCellValue(doubleValue);
        }
        break;
      case Types.INTEGER:
      case Types.BIGINT:
      case Types.SMALLINT:
      case Types.TINYINT:
        long longValue = resultSet.getLong(col);
        if (resultSet.wasNull()) {
          cell.setBlank();
        } else {
          cell.setCellValue(longValue);
        }
        break;
      default:
        String stringValue = resultSet.getString(col);
        if (stringValue == null) {
          cell.setBlank();
        } else {
          cell.setCellValue(stringValue);
        }
        break;
    }
  }

  private CellStyle createNonEditableHeaderStyle(SXSSFWorkbook workbook) {
    String hex = config.getNonEditableHeaderColor();
    if (hex == null || hex.isBlank()) {
      return createHeaderStyle(workbook, DEFAULT_NON_EDITABLE_HEADER_COLOR);
    }
    try {
      return createHeaderStyle(workbook, parseHexColor(hex));
    } catch (NumberFormatException e) {
      LOG.warn("Invalid NoneEditableHeaderColor '{}', falling back to the default yellow", hex, e);
      return createHeaderStyle(workbook, DEFAULT_NON_EDITABLE_HEADER_COLOR);
    }
  }

  private XSSFColor parseHexColor(String hex) {
    String normalized = hex.startsWith("#") ? hex : "#" + hex;
    return new XSSFColor(Color.decode(normalized), null);
  }

  private CellStyle createHeaderStyle(SXSSFWorkbook workbook, short indexedFillColor) {
    CellStyle style = workbook.createCellStyle();
    style.setFillForegroundColor(indexedFillColor);
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    finishHeaderStyle(workbook, style);
    return style;
  }

  private CellStyle createHeaderStyle(SXSSFWorkbook workbook, XSSFColor fillColor) {
    CellStyle style = workbook.createCellStyle();
    ((XSSFCellStyle) style).setFillForegroundColor(fillColor);
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    finishHeaderStyle(workbook, style);
    return style;
  }

  private void finishHeaderStyle(SXSSFWorkbook workbook, CellStyle style) {
    style.setLocked(true);
    Font font = workbook.createFont();
    font.setBold(true);
    style.setFont(font);
    style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
    style.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
    style.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
    style.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
  }

  /**
   * SXSSFSheet does not publicly expose the underlying XSSFSheet needed to reach
   * CTSheetProtection, so the package-private field "_sh" (verified against
   * Apache POI 5.2.5) is accessed via reflection. This is the documented
   * workaround for blocking row/column insertion and deletion under SXSSF.
   *
   * <p>Protection is always applied with an empty password: ExcelPassword only
   * controls whether the whole file is encrypted for opening (see {@link #encrypt}),
   * not whether cell locks require a password to remove.
   */
  private void protectSheet(SXSSFSheet sheet) {
    sheet.protectSheet("");

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

  /**
   * Caches the 6 locked/unlocked x plain/date/date-time cell style combinations
   * needed per workbook, so repeated rows reuse the same CellStyle instances
   * instead of creating a new one per cell (POI caps the number of distinct
   * styles a workbook can hold).
   */
  private static final class DataStyles {
    private final CellStyle lockedPlain;
    private final CellStyle unlockedPlain;
    private final CellStyle lockedDate;
    private final CellStyle unlockedDate;
    private final CellStyle lockedDateTime;
    private final CellStyle unlockedDateTime;

    DataStyles(SXSSFWorkbook workbook, String dateFormat, String dateTimeFormat) {
      this.lockedPlain = createStyle(workbook, true, null);
      this.unlockedPlain = createStyle(workbook, false, null);
      this.lockedDate = createStyle(workbook, true, dateFormat);
      this.unlockedDate = createStyle(workbook, false, dateFormat);
      this.lockedDateTime = createStyle(workbook, true, dateTimeFormat);
      this.unlockedDateTime = createStyle(workbook, false, dateTimeFormat);
    }

    private static CellStyle createStyle(SXSSFWorkbook workbook, boolean locked, String numberFormat) {
      CellStyle style = workbook.createCellStyle();
      style.setLocked(locked);
      style.setFillForegroundColor(locked ? NON_EDITABLE_DATA_COLOR : EDITABLE_DATA_COLOR);
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      if (numberFormat != null) {
        style.setDataFormat(workbook.createDataFormat().getFormat(numberFormat));
      }
      return style;
    }

    CellStyle styleFor(int columnType, boolean editable) {
      switch (columnType) {
        case Types.DATE:
          return editable ? unlockedDate : lockedDate;
        case Types.TIMESTAMP:
          return editable ? unlockedDateTime : lockedDateTime;
        default:
          return editable ? unlockedPlain : lockedPlain;
      }
    }
  }
}
