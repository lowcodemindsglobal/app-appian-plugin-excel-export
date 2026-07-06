package com.lowcodeminds.appian.plugins.excel;

import com.appiancorp.suiteapi.common.Name;
import com.appiancorp.suiteapi.common.exceptions.PrivilegeException;
import com.appiancorp.suiteapi.common.exceptions.StorageLimitException;
import com.appiancorp.suiteapi.content.Content;
import com.appiancorp.suiteapi.content.ContentConstants;
import com.appiancorp.suiteapi.content.ContentOutputStream;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.content.exceptions.DuplicateUuidException;
import com.appiancorp.suiteapi.content.exceptions.InsufficientNameUniquenessException;
import com.appiancorp.suiteapi.content.exceptions.InvalidContentException;
import com.appiancorp.suiteapi.knowledge.Document;
import com.appiancorp.suiteapi.knowledge.DocumentDataType;
import com.appiancorp.suiteapi.knowledge.FolderDataType;
import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import com.appiancorp.suiteapi.process.framework.AppianSmartService;
import com.appiancorp.suiteapi.process.framework.Input;
import com.appiancorp.suiteapi.process.framework.MessageContainer;
import com.appiancorp.suiteapi.process.framework.Order;
import com.appiancorp.suiteapi.process.framework.Required;
import com.appiancorp.suiteapi.process.framework.Unattended;
import com.appiancorp.suiteapi.process.palette.AutomationSmartServicesDocumentGeneration;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.naming.Context;
import javax.sql.DataSource;
import java.io.IOException;
import java.sql.Connection;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

/**
 * Smart Service node: runs one SQL query per sheet against a JNDI data source and
 * saves the combined result as a column-protected .xlsx Appian Document.
 *
 * <p>{@link ContentService} and {@link Context} are injected as constructor
 * parameters. The Appian plug-in framework recognizes constructor parameters typed
 * as public {@code *Service} interfaces or {@code javax.naming.Context}, and
 * supplies them automatically - this class does not need, and must not have, any
 * manual Spring/DI wiring for either.
 *
 * <p>{@code @AutomationSmartServicesDocumentGeneration} is one of the SDK's
 * built-in convenience annotations; it is equivalent to writing
 * {@code @PaletteInfo(paletteCategory = "Automation Smart Services", palette = "Document Generation")}
 * but avoids hand-typing palette category strings that have to exactly match
 * values Appian recognizes.
 *
 * <p>{@code @Unattended} tells Appian this node always runs headless and must
 * never be assignable to a person as a task.
 *
 * <p>Failures are surfaced both ways: {@code ErrorMessage} is populated so a
 * calling process can branch on it without a fault handler, and the same
 * failure is also thrown as a {@link SmartServiceException} so a process that
 * doesn't check the output still stops instead of silently continuing with a
 * null {@code NewDocumentCreated}.
 */
@AutomationSmartServicesDocumentGeneration
@Unattended
@Order({"SQLSheetDataList", "jndiName", "editableColumns", "targetFolder", "NewDocumentName", "DateFormat", "DateTimeFormat", "NoneEditableHeaderColor", "ExcelPassword"})
public class ExcelExportSmartService extends AppianSmartService {

  private static final Logger LOG = LoggerFactory.getLogger(ExcelExportSmartService.class);
  private static final String XLSX_EXTENSION = "xlsx";
  private static final int MAX_SHEETS = 20;
  private static final String SQL_SHEET_DATA_LIST_INPUT = "SQLSheetDataList";

  private final ContentService contentService;
  private final Context context;

  // Smart Service inputs - each field is set by its matching @Input setter below.
  private SQLSheetData[] sqlSheetDataList;
  private String jndiName;
  private String[] editableColumns;
  private Long targetFolder;
  private String newDocumentName;
  private String dateFormat;
  private String dateTimeFormat;
  private String nonEditableHeaderColor;
  private String excelPassword;

  // Smart Service outputs - exposed by the getter methods below.
  private Long newDocumentCreated;
  private String errorMessage;

  public ExcelExportSmartService(ContentService contentService, Context context) {
    this.contentService = contentService;
    this.context = context;
  }

  @Input(required = Required.ALWAYS)
  @Name("SQLSheetDataList")
  public void setSqlSheetDataList(SQLSheetData[] sqlSheetDataList) {
    this.sqlSheetDataList = sqlSheetDataList;
  }

  @Input(required = Required.ALWAYS)
  @Name("jndiName")
  public void setJndiName(String jndiName) {
    this.jndiName = jndiName;
  }

  // Global by column name: any column with a matching name, in any sheet, is left
  // unlocked. A name that doesn't match any sheet's columns simply has no effect.
  @Input(required = Required.OPTIONAL)
  @Name("editableColumns")
  public void setEditableColumns(String[] editableColumns) {
    this.editableColumns = editableColumns;
  }

  @Input(required = Required.ALWAYS)
  @Name("targetFolder")
  @FolderDataType
  public void setTargetFolder(Long targetFolder) {
    this.targetFolder = targetFolder;
  }

  @Input(required = Required.ALWAYS)
  @Name("NewDocumentName")
  public void setNewDocumentName(String newDocumentName) {
    this.newDocumentName = newDocumentName;
  }

  @Input(required = Required.OPTIONAL)
  @Name("DateFormat")
  public void setDateFormat(String dateFormat) {
    this.dateFormat = dateFormat;
  }

  @Input(required = Required.OPTIONAL)
  @Name("DateTimeFormat")
  public void setDateTimeFormat(String dateTimeFormat) {
    this.dateTimeFormat = dateTimeFormat;
  }

  // Fill color for non-editable column headers. Editable column headers always use
  // a fixed darker gray instead - there's no input to override that side.
  @Input(required = Required.OPTIONAL)
  @Name("NoneEditableHeaderColor")
  public void setNonEditableHeaderColor(String nonEditableHeaderColor) {
    this.nonEditableHeaderColor = nonEditableHeaderColor;
  }

  // Controls whether the generated file is encrypted for opening (see
  // ExcelGenerator.encrypt). Sheet-level cell protection is always applied
  // regardless of this input, with no password required to remove it.
  @Input(required = Required.OPTIONAL)
  @Name("ExcelPassword")
  public void setExcelPassword(String excelPassword) {
    this.excelPassword = excelPassword;
  }

  @Name("NewDocumentCreated")
  @DocumentDataType
  public Long getNewDocumentCreated() {
    return newDocumentCreated;
  }

  @Name("ErrorMessage")
  public String getErrorMessage() {
    return errorMessage;
  }

  @Override
  public void validate(MessageContainer messages) {
    super.validate(messages);

    if (sqlSheetDataList == null || sqlSheetDataList.length == 0) {
      messages.addError(SQL_SHEET_DATA_LIST_INPUT, "SQL Sheet Data List must not be empty.");
      return;
    }

    if (sqlSheetDataList.length > MAX_SHEETS) {
      messages.addError(SQL_SHEET_DATA_LIST_INPUT, "A maximum of " + MAX_SHEETS + " sheets is allowed.");
    }

    for (SQLSheetData sheetData : sqlSheetDataList) {
      validateSheetData(messages, sheetData);
    }
  }

  private void validateSheetData(MessageContainer messages, SQLSheetData sheetData) {
    if (sheetData.getSheetName() == null || sheetData.getSheetName().isBlank()) {
      messages.addError(SQL_SHEET_DATA_LIST_INPUT, "Sheet Name is missing.");
    }

    String sqlQuery = sheetData.getSqlQuery();
    if (sqlQuery == null || sqlQuery.isBlank()) {
      messages.addError(SQL_SHEET_DATA_LIST_INPUT, "SQL Query is missing.");
      return;
    }
    if (!sqlQuery.trim().toUpperCase().startsWith("SELECT")) {
      messages.addError(SQL_SHEET_DATA_LIST_INPUT, "SQL Query invalid syntax. Only SELECT statements are allowed.");
    }
    if (sqlQuery.trim().endsWith(";")) {
      messages.addError(SQL_SHEET_DATA_LIST_INPUT, "SQL Query must not end with a semicolon.");
    }
  }

  @Override
  public void run() throws SmartServiceException {
    try {
      ExcelExportConfig config = ExcelExportConfig.builder()
          .editableColumns(editableColumns != null ? Arrays.asList(editableColumns) : Collections.emptyList())
          .documentName(newDocumentName)
          .targetFolderId(targetFolder)
          .dateFormat(dateFormat)
          .dateTimeFormat(dateTimeFormat)
          .nonEditableHeaderColor(nonEditableHeaderColor)
          .excelPassword(excelPassword)
          .build();

      byte[] excelBytes = generateExcelBytes(config);

      this.newDocumentCreated = saveAsDocument(excelBytes, config);
      LOG.info("Excel export completed, documentId={}", newDocumentCreated);
    } catch (Exception e) {
      errorMessage = e.getMessage();
      LOG.error("Error generating Excel from SQL", e);
      throw new SmartServiceException(getClass(), e, errorMessage);
    }
  }

  private byte[] generateExcelBytes(ExcelExportConfig config) throws Exception {
    LOG.debug("Looking up JNDI data source: {}", jndiName);
    DataSource dataSource = (DataSource) context.lookup(jndiName); // nosemgrep datasource read from appian-configured JNDI

    byte[] excelBytes;
    try (Connection connection = dataSource.getConnection()) {
      excelBytes = new ExcelGenerator(config).generate(connection, List.of(sqlSheetDataList));
    }

    if (config.isPasswordProtected()) {
      excelBytes = ExcelGenerator.encrypt(excelBytes, config.getExcelPassword());
    }
    return excelBytes;
  }

  /**
   * ContentService.upload()/ContentOutputStream are marked @Deprecated in the SDK, but
   * Document.write(InputStream)/getOutputStream() - the only non-deprecated alternatives -
   * delegate to a private documentHelper field that is only populated by the Appian engine
   * on framework-returned Document instances. Calling them on a plugin-constructed Document
   * NPEs. upload() remains the only working path for plugins to push byte content, so the
   * deprecation warning is intentionally suppressed here.
   *
   * <p>If a document named {@code NewDocumentName.xlsx} already exists directly in the
   * target folder, this overwrites it by uploading a new version onto that existing
   * content id instead of creating a sibling. Otherwise it creates a new document.
   */
  @SuppressWarnings("deprecation")
  private Long saveAsDocument(byte[] excelBytes, ExcelExportConfig config)
      throws InvalidContentException, StorageLimitException, PrivilegeException,
      InsufficientNameUniquenessException, DuplicateUuidException, IOException {
    Long existingDocumentId = findExistingDocumentId(config);

    Document document = new Document();
    document.setName(config.getDocumentName());
    document.setParent(config.getTargetFolderId());
    document.setExtension(XLSX_EXTENSION);

    Integer uploadMode;
    if (existingDocumentId != null) {
      document.setId(existingDocumentId);
      uploadMode = ContentConstants.VERSION_CURRENT;
    } else {
      uploadMode = ContentConstants.UNIQUE_FOR_PARENT;
    }

    Long newDocumentId;
    try (ContentOutputStream contentOutputStream = contentService.upload(document, uploadMode)) {
      contentOutputStream.write(excelBytes);
      newDocumentId = contentOutputStream.getContentId();
    }
    return newDocumentId;
  }

  /**
   * Resolves the content id of an existing {@code NewDocumentName.xlsx} document
   * directly inside the target folder, or null if none exists yet.
   */
  private Long findExistingDocumentId(ExcelExportConfig config) throws InvalidContentException {
    String relativePath = config.getDocumentName() + "." + XLSX_EXTENSION;
    Content existingContent;
    try {
      existingContent = contentService.getByPath(config.getTargetFolderId(), relativePath);
    } catch (InvalidContentException e) {
      return null;
    }
    return existingContent instanceof Document ? existingContent.getId() : null;
  }
}
