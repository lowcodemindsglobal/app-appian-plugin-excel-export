package com.lowcodeminds.appian.plugins.excel;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelExportConfigTest {

  private static ExcelExportConfig.Builder baseConfig() {
    return ExcelExportConfig.builder()
        .documentName("export")
        .targetFolderId(1L);
  }

  @Test
  void testBuilderConstructsConfig() {
    ExcelExportConfig config = baseConfig()
        .editableColumns(List.of("Status"))
        .dateFormat("yyyy-mm-dd")
        .dateTimeFormat("yyyy-mm-dd hh:mm")
        .nonEditableHeaderColor("#336699")
        .excelPassword("secret")
        .build();

    assertEquals(List.of("Status"), config.getEditableColumns());
    assertEquals("export", config.getDocumentName());
    assertEquals(1L, config.getTargetFolderId());
    assertEquals("yyyy-mm-dd", config.getDateFormat());
    assertEquals("yyyy-mm-dd hh:mm", config.getDateTimeFormat());
    assertEquals("#336699", config.getNonEditableHeaderColor());
    assertEquals("secret", config.getExcelPassword());
    assertTrue(config.isPasswordProtected());
  }

  @Test
  void testIsEditableReturnsTrueOnlyForListedColumns() {
    ExcelExportConfig config = baseConfig().editableColumns(List.of("Status")).build();

    assertTrue(config.isEditable("Status"));
    assertFalse(config.isEditable("ID"));
    assertFalse(config.isEditable("AnyOtherColumn"));
  }

  @Test
  void testBlankDateFormatsFallBackToDefaults() {
    ExcelExportConfig config = baseConfig().dateFormat(" ").dateTimeFormat("").build();

    assertEquals("dd-mm-yyyy", config.getDateFormat());
    assertEquals("dd-mm-yyyy hh:mm:ss", config.getDateTimeFormat());
  }

  @Test
  void testBlankPasswordIsNotPasswordProtected() {
    ExcelExportConfig config = baseConfig().excelPassword(" ").build();

    assertFalse(config.isPasswordProtected());
  }

  @Test
  void testUnsetNonEditableHeaderColorIsNull() {
    ExcelExportConfig config = baseConfig().build();

    assertNull(config.getNonEditableHeaderColor());
  }

  @Test
  void testMissingRequiredFieldsThrows() {
    ExcelExportConfig.Builder builder = ExcelExportConfig.builder().documentName("export");

    assertThrows(NullPointerException.class, builder::build);
  }
}
