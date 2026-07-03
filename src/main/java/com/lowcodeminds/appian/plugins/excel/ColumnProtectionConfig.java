package com.lowcodeminds.appian.plugins.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * Immutable settings for one Excel export: which columns exist, which of those
 * stay editable, and where the generated file gets saved. Built via
 * {@link #builder()} rather than a public constructor so every required field
 * is checked (including the editable/all-columns cross-check) before an
 * instance can exist.
 */
public final class ColumnProtectionConfig {

  private final List<String> allColumns;
  private final List<String> editableColumns;
  private final String protectionPassword;
  private final String fileName;
  private final Long targetFolderId;

  private ColumnProtectionConfig(Builder builder) {
    this.allColumns = Collections.unmodifiableList(new ArrayList<>(builder.allColumns));
    this.editableColumns = Collections.unmodifiableList(new ArrayList<>(builder.editableColumns));
    this.protectionPassword = builder.protectionPassword;
    this.fileName = builder.fileName;
    this.targetFolderId = builder.targetFolderId;
  }

  public List<String> getAllColumns() {
    return allColumns;
  }

  public List<String> getEditableColumns() {
    return editableColumns;
  }

  public String getProtectionPassword() {
    return protectionPassword;
  }

  public String getFileName() {
    return fileName;
  }

  public Long getTargetFolderId() {
    return targetFolderId;
  }

  /**
   * @return true if this column should be left unlocked (editable) in the
   *     generated sheet; false if it should be locked, including for any
   *     column name not found in {@link #getAllColumns()} at all.
   */
  public boolean isEditable(String columnName) {
    return editableColumns.contains(columnName);
  }

  public static Builder builder() {
    return new Builder();
  }

  public static final class Builder {
    private List<String> allColumns = new ArrayList<>();
    private List<String> editableColumns = new ArrayList<>();
    private String protectionPassword;
    private String fileName;
    private Long targetFolderId;

    private Builder() {
    }

    public Builder allColumns(List<String> allColumns) {
      this.allColumns = Objects.requireNonNull(allColumns, "allColumns must not be null");
      return this;
    }

    public Builder editableColumns(List<String> editableColumns) {
      this.editableColumns = Objects.requireNonNull(editableColumns, "editableColumns must not be null");
      return this;
    }

    public Builder protectionPassword(String protectionPassword) {
      this.protectionPassword = protectionPassword;
      return this;
    }

    public Builder fileName(String fileName) {
      this.fileName = Objects.requireNonNull(fileName, "fileName must not be null");
      return this;
    }

    public Builder targetFolderId(Long targetFolderId) {
      this.targetFolderId = Objects.requireNonNull(targetFolderId, "targetFolderId must not be null");
      return this;
    }

    /**
     * @throws NullPointerException if a required field was never set
     * @throws IllegalArgumentException if editableColumns contains a name
     *     that isn't in allColumns - this is also checked earlier, at design
     *     time, by ExcelExportSmartService.validate(), so reaching this
     *     exception at run() time should be rare in practice.
     */
    public ColumnProtectionConfig build() {
      Objects.requireNonNull(allColumns, "allColumns must not be null");
      Objects.requireNonNull(editableColumns, "editableColumns must not be null");
      Objects.requireNonNull(fileName, "fileName must not be null");
      Objects.requireNonNull(targetFolderId, "targetFolderId must not be null");

      for (String editableColumn : editableColumns) {
        if (!allColumns.contains(editableColumn)) {
          throw new IllegalArgumentException(
              "editableColumns contains a column not present in allColumns: " + editableColumn);
        }
      }

      return new ColumnProtectionConfig(this);
    }
  }
}
