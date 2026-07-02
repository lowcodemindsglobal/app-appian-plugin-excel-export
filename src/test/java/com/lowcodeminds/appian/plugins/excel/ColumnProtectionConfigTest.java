package com.lowcodeminds.appian.plugins.excel;

import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ColumnProtectionConfigTest {

  private static final List<String> ALL_COLUMNS = Arrays.asList("ID", "Name", "Status");

  @Test
  void testBuilderConstructsConfig() {
    ColumnProtectionConfig config = ColumnProtectionConfig.builder()
        .allColumns(ALL_COLUMNS)
        .editableColumns(List.of("Status"))
        .fileName("export.xlsx")
        .targetFolderId(42L)
        .protectionPassword("secret")
        .build();

    assertEquals(ALL_COLUMNS, config.getAllColumns());
    assertEquals(List.of("Status"), config.getEditableColumns());
    assertEquals("export.xlsx", config.getFileName());
    assertEquals(42L, config.getTargetFolderId());
    assertEquals("secret", config.getProtectionPassword());
  }

  @Test
  void testIsEditableReturnsTrueOnlyForEditableColumns() {
    ColumnProtectionConfig config = ColumnProtectionConfig.builder()
        .allColumns(ALL_COLUMNS)
        .editableColumns(List.of("Status"))
        .fileName("export.xlsx")
        .targetFolderId(1L)
        .build();

    assertTrue(config.isEditable("Status"));
    assertFalse(config.isEditable("ID"));
    assertFalse(config.isEditable("Name"));
  }

  @Test
  void testValidationRejectsUnknownEditableColumn() {
    ColumnProtectionConfig.Builder builder = ColumnProtectionConfig.builder()
        .allColumns(ALL_COLUMNS)
        .editableColumns(List.of("NotAColumn"))
        .fileName("export.xlsx")
        .targetFolderId(1L);

    assertThrows(IllegalArgumentException.class, builder::build);
  }

  @Test
  void testMissingRequiredFieldsThrows() {
    ColumnProtectionConfig.Builder builder = ColumnProtectionConfig.builder()
        .allColumns(ALL_COLUMNS);

    assertThrows(NullPointerException.class, builder::build);
  }
}
