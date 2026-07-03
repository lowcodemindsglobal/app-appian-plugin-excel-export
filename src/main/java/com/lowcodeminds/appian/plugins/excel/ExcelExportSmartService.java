package com.lowcodeminds.appian.plugins.excel;

import com.appiancorp.suiteapi.common.Name;
import com.appiancorp.suiteapi.common.exceptions.PrivilegeException;
import com.appiancorp.suiteapi.common.exceptions.StorageLimitException;
import com.appiancorp.suiteapi.content.ContentConstants;
import com.appiancorp.suiteapi.content.ContentOutputStream;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.content.exceptions.DuplicateUuidException;
import com.appiancorp.suiteapi.content.exceptions.InsufficientNameUniquenessException;
import com.appiancorp.suiteapi.content.exceptions.InvalidContentException;
import com.appiancorp.suiteapi.knowledge.Document;
import com.appiancorp.suiteapi.knowledge.FolderDataType;
import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import com.appiancorp.suiteapi.process.framework.AppianSmartService;
import com.appiancorp.suiteapi.process.framework.Input;
import com.appiancorp.suiteapi.process.framework.MessageContainer;
import com.appiancorp.suiteapi.process.framework.Order;
import com.appiancorp.suiteapi.process.framework.Required;
import com.appiancorp.suiteapi.process.framework.Unattended;
import com.appiancorp.suiteapi.process.palette.AutomationSmartServicesDocumentGeneration;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Smart Service node: exports data rows to a column-protected .xlsx file and
 * saves it as an Appian Document.
 *
 * <p>{@link ContentService} is injected as a constructor parameter. The Appian
 * plug-in framework recognizes constructor parameters typed as public
 * {@code *Service} interfaces and supplies them automatically - this class does
 * not need, and must not have, any manual Spring/DI wiring for it.
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
 * <p>{@code dataRowsJson} is a JSON-encoded string, not a native Appian "List
 * of Dictionary" input. An earlier version of this class declared it as
 * {@code Map<String, Object>[]}, which failed to deploy against a real Appian
 * server with "Invalid Type: Unsupported type [Ljava.util.Map;" - Appian's
 * plugin loader has no built-in mapping for a raw Dictionary array, and the
 * SDK exposes no annotation plugin authors can use to declare one (only
 * pre-built convenience annotations like {@code @DocumentDataType} exist, and
 * none covers Dictionary). Accepting the rows as JSON and parsing them with
 * Jackson sidesteps that limitation entirely; the calling process passes the
 * row data through something like {@code a!toJson(rows)}.
 */
@AutomationSmartServicesDocumentGeneration
@Unattended
@Order({"dataRowsJson", "allColumns", "editableColumns", "targetFolderId", "targetFolder", "fileName", "protectionPassword", "documentId"})
public class ExcelExportSmartService extends AppianSmartService {

  private static final Logger LOG = LoggerFactory.getLogger(ExcelExportSmartService.class);
  private static final String XLSX_EXTENSION = "xlsx";
  private static final ObjectMapper JSON = new ObjectMapper();
  private static final TypeReference<List<Map<String, Object>>> DATA_ROWS_TYPE = new TypeReference<>() {
  };

  private final ContentService contentService;

  // Smart Service inputs - each field is set by its matching @Input setter below.
  private String dataRowsJson;
  private String[] allColumns;
  private String[] editableColumns;
  private Long targetFolderId;
  private Long targetFolder;
  private String fileName;
  private String protectionPassword;

  // Smart Service output - exposed by the getDocumentId() getter below.
  private Long documentId;

  public ExcelExportSmartService(ContentService contentService) {
    this.contentService = contentService;
  }

  @Input(required = Required.ALWAYS)
  @Name("dataRowsJson")
  public void setDataRowsJson(String dataRowsJson) {
    this.dataRowsJson = dataRowsJson;
  }

  @Input(required = Required.ALWAYS)
  @Name("allColumns")
  public void setAllColumns(String[] allColumns) {
    this.allColumns = allColumns;
  }

  @Input(required = Required.ALWAYS)
  @Name("editableColumns")
  public void setEditableColumns(String[] editableColumns) {
    this.editableColumns = editableColumns;
  }

  // Two alternate ways to specify the destination folder: a plain numeric ID,
  // or - via @FolderDataType, which tells Appian this Long is a folder
  // reference (system type "CollaborationFolder") - a folder-picker control
  // in the Process Modeler. Both are optional individually, but validate()
  // below requires that at least one of them be supplied; if both are, run()
  // prefers targetFolder since it comes from the picker and can't be typoed.
  @Input(required = Required.OPTIONAL)
  @Name("targetFolderId")
  public void setTargetFolderId(Long targetFolderId) {
    this.targetFolderId = targetFolderId;
  }

  @Input(required = Required.OPTIONAL)
  @Name("targetFolder")
  @FolderDataType
  public void setTargetFolder(Long targetFolder) {
    this.targetFolder = targetFolder;
  }

  @Input(required = Required.ALWAYS)
  @Name("fileName")
  public void setFileName(String fileName) {
    this.fileName = fileName;
  }

  @Input(required = Required.OPTIONAL)
  @Name("protectionPassword")
  public void setProtectionPassword(String protectionPassword) {
    this.protectionPassword = protectionPassword;
  }

  /**
   * Output parameter. A getter with no matching setter is what tells the
   * Appian plug-in framework this is an output rather than an input - do not
   * add a setDocumentId() method, or this stops being treated as an output.
   */
  @Name("documentId")
  public Long getDocumentId() {
    return documentId;
  }

  /**
   * Design-time validation, run before {@link #run()} whether the node is
   * attended or unattended. Catching the "editable column not in all columns"
   * mistake here - instead of only inside run(), which is what
   * ColumnProtectionConfig's own builder already enforces - lets a process
   * designer see the mistake immediately in the modeler instead of only
   * finding out when the process actually executes.
   */
  @Override
  public void validate(MessageContainer messages) {
    super.validate(messages);

    if (targetFolderId == null && targetFolder == null) {
      messages.addError("targetFolder", "Either \"Target Folder ID\" or \"Target Folder\" must be provided.");
    }

    if (allColumns == null || editableColumns == null) {
      // Required-input validation (missing/blank values) is already handled
      // by the framework via @Input(required = Required.ALWAYS); nothing
      // further to check here until both arrays are actually present.
      return;
    }

    List<String> allColumnsList = Arrays.asList(allColumns);
    for (String editableColumn : editableColumns) {
      if (!allColumnsList.contains(editableColumn)) {
        messages.addError("editableColumns", "\"" + editableColumn + "\" is not one of the columns listed in All Columns.");
      }
    }
  }

  @Override
  public void run() throws SmartServiceException {
    // targetFolder (picker) takes precedence over targetFolderId (typed
    // value) when both are supplied; validate() already guarantees at least
    // one of them is present.
    Long effectiveTargetFolderId = targetFolder != null ? targetFolder : targetFolderId;

    ColumnProtectionConfig config;
    try {
      config = ColumnProtectionConfig.builder()
          .allColumns(allColumns != null ? Arrays.asList(allColumns) : Collections.emptyList())
          .editableColumns(editableColumns != null ? Arrays.asList(editableColumns) : Collections.emptyList())
          .protectionPassword(protectionPassword)
          .fileName(fileName)
          .targetFolderId(effectiveTargetFolderId)
          .build();
    } catch (IllegalArgumentException | NullPointerException e) {
      // IllegalArgumentException: editableColumns has a name not present in
      // allColumns (validate() above should normally catch this first).
      // NullPointerException: a Required.ALWAYS input came through null,
      // which should not be possible if Appian enforced the annotation -
      // handled here defensively rather than surfacing a raw NPE to the user.
      LOG.warn("Rejected invalid column configuration for smart service execution", e);
      throw new SmartServiceException(getClass(), e, "Invalid column configuration: " + e.getMessage());
    }

    List<Map<String, Object>> rows;
    try {
      rows = dataRowsJson != null && !dataRowsJson.isBlank()
          ? JSON.readValue(dataRowsJson, DATA_ROWS_TYPE)
          : Collections.emptyList();
    } catch (JsonProcessingException e) {
      LOG.warn("Rejected malformed Data Rows JSON", e);
      throw new SmartServiceException(getClass(), e, "Invalid JSON in Data Rows: " + e.getMessage());
    }

    byte[] excelBytes;
    try {
      excelBytes = new ExcelGenerator(config).generate(rows);
    } catch (ExcelGenerationException e) {
      LOG.error("Excel workbook generation failed", e);
      throw new SmartServiceException(getClass(), e, "Failed to generate the Excel file: " + e.getMessage());
    }

    try {
      this.documentId = saveAsDocument(excelBytes, config);
    } catch (InvalidContentException | StorageLimitException | PrivilegeException
        | InsufficientNameUniquenessException | DuplicateUuidException | IOException e) {
      LOG.error("Saving the generated Excel file as an Appian document failed", e);
      throw new SmartServiceException(getClass(), e,
          "Failed to save the generated Excel file as an Appian document: " + e.getMessage());
    }

    LOG.info("Excel export completed, documentId={}", documentId);
  }

  /**
   * ContentService.upload()/ContentOutputStream are marked @Deprecated in the SDK, but
   * Document.write(InputStream)/getOutputStream() - the only non-deprecated alternatives -
   * delegate to a private documentHelper field that is only populated by the Appian engine
   * on framework-returned Document instances. Calling them on a plugin-constructed Document
   * NPEs. upload() remains the only working path for plugins to push byte content, so the
   * deprecation warning is intentionally suppressed here.
   */
  @SuppressWarnings("deprecation")
  private Long saveAsDocument(byte[] excelBytes, ColumnProtectionConfig config)
      throws InvalidContentException, StorageLimitException, PrivilegeException,
      InsufficientNameUniquenessException, DuplicateUuidException, IOException {
    Document document = new Document();
    document.setName(config.getFileName());
    document.setParent(config.getTargetFolderId());
    document.setExtension(XLSX_EXTENSION);

    // upload() alone both creates the content item and gives back a stream to
    // write its bytes to - it must not be preceded by a separate create()
    // call. An earlier version of this method called create() first (with
    // UNIQUE_FOR_PARENT) and then upload() (with VERSION_CURRENT) on the same
    // Document; upload() runs its own name-uniqueness check independently of
    // create(), so it saw the document create() had just made as a colliding
    // sibling and failed with InsufficientNameUniquenessException every time.
    Long newDocumentId;
    try (ContentOutputStream contentOutputStream = contentService.upload(document, ContentConstants.UNIQUE_FOR_PARENT)) {
      contentOutputStream.write(excelBytes);
      newDocumentId = contentOutputStream.getContentId();
    }
    return newDocumentId;
  }
}
