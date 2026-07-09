package com.lowcodeminds.appian.plugins.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * Immutable settings for one Excel export run: which columns stay editable
 * (matched by name across every sheet), where the generated file is saved,
 * and how it's formatted/protected. Built via {@link #builder()} rather than
 * a public constructor so every required field is checked before an instance
 * can exist.
 */
public final class ExcelExportConfig {

  private static final String DEFAULT_DATE_FORMAT = "dd-mm-yyyy";
  private static final String DEFAULT_DATE_TIME_FORMAT = "dd-mm-yyyy hh:mm:ss";

  private final List<String> editableColumns;
  private final String documentName;
  private final Long targetFolderId;
  private final String dateFormat;
  private final String dateTimeFormat;
  private final String nonEditableHeaderColor;
  private final String excelPassword;

  private ExcelExportConfig(Builder builder) {
    this.editableColumns = Collections.unmodifiableList(new ArrayList<>(builder.editableColumns));
    this.documentName = builder.documentName;
    this.targetFolderId = builder.targetFolderId;
    this.dateFormat = isBlank(builder.dateFormat) ? DEFAULT_DATE_FORMAT : builder.dateFormat;
    this.dateTimeFormat = isBlank(builder.dateTimeFormat) ? DEFAULT_DATE_TIME_FORMAT : builder.dateTimeFormat;
    this.nonEditableHeaderColor = builder.nonEditableHeaderColor;
    this.excelPassword = builder.excelPassword;
  }

  private static boolean isBlank(String value) {
    return value == null || value.isBlank();
  }

  public List<String> getEditableColumns() {
    return editableColumns;
  }

  public String getDocumentName() {
    return documentName;
  }

  public Long getTargetFolderId() {
    return targetFolderId;
  }

  public String getDateFormat() {
    return dateFormat;
  }

  public String getDateTimeFormat() {
    return dateTimeFormat;
  }

  /**
   * @return the hex fill color for non-editable column headers, or null if none was
   *     supplied (in which case a default yellow is used - see ExcelGenerator)
   */
  public String getNonEditableHeaderColor() {
    return nonEditableHeaderColor;
  }

  public String getExcelPassword() {
    return excelPassword;
  }

  public boolean isPasswordProtected() {
    return !isBlank(excelPassword);
  }

  /**
   * @return true if this column should be left unlocked (editable) in every
   *     generated sheet; false if it should be locked, including for any
   *     column name that doesn't appear in {@link #getEditableColumns()} at all.
   *     Matched case-insensitively, since {@code ResultSetMetaData} column-label
   *     casing is JDBC-driver/database-dependent (e.g. Oracle uppercases
   *     unquoted identifiers, PostgreSQL lowercases them) and a process designer
   *     configuring this list has no reliable way to predict which casing a
   *     given data source will actually return.
   */
  public boolean isEditable(String columnName) {
    return columnName != null && editableColumns.stream().anyMatch(columnName::equalsIgnoreCase);
  }

  public static Builder builder() {
    return new Builder();
  }

  public static final class Builder {
    private List<String> editableColumns = new ArrayList<>();
    private String documentName;
    private Long targetFolderId;
    private String dateFormat;
    private String dateTimeFormat;
    private String nonEditableHeaderColor;
    private String excelPassword;

    private Builder() {
    }

    public Builder editableColumns(List<String> editableColumns) {
      this.editableColumns = Objects.requireNonNull(editableColumns, "editableColumns must not be null");
      return this;
    }

    public Builder documentName(String documentName) {
      this.documentName = Objects.requireNonNull(documentName, "documentName must not be null");
      return this;
    }

    public Builder targetFolderId(Long targetFolderId) {
      this.targetFolderId = Objects.requireNonNull(targetFolderId, "targetFolderId must not be null");
      return this;
    }

    public Builder dateFormat(String dateFormat) {
      this.dateFormat = dateFormat;
      return this;
    }

    public Builder dateTimeFormat(String dateTimeFormat) {
      this.dateTimeFormat = dateTimeFormat;
      return this;
    }

    public Builder nonEditableHeaderColor(String nonEditableHeaderColor) {
      this.nonEditableHeaderColor = nonEditableHeaderColor;
      return this;
    }

    public Builder excelPassword(String excelPassword) {
      this.excelPassword = excelPassword;
      return this;
    }

    /**
     * @throws NullPointerException if a required field was never set
     */
    public ExcelExportConfig build() {
      Objects.requireNonNull(editableColumns, "editableColumns must not be null");
      Objects.requireNonNull(documentName, "documentName must not be null");
      Objects.requireNonNull(targetFolderId, "targetFolderId must not be null");
      return new ExcelExportConfig(this);
    }
  }
}
