package com.lowcodeminds.appian.plugins.excel;

import com.appiancorp.suiteapi.common.Name;
import com.appiancorp.suiteapi.content.ContentConstants;
import com.appiancorp.suiteapi.content.ContentOutputStream;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.knowledge.Document;
import com.appiancorp.suiteapi.process.exceptions.SmartServiceException;
import com.appiancorp.suiteapi.process.framework.AppianSmartService;
import com.appiancorp.suiteapi.process.framework.Input;
import com.appiancorp.suiteapi.process.framework.Required;
import com.appiancorp.suiteapi.process.palette.PaletteInfo;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Smart Service node: exports data rows to a column-protected .xlsx file and
 * saves it as an Appian Document.
 *
 * ContentService is supplied via Spring constructor injection, wired in
 * appian-plugin.xml against Appian's "content-service" bean.
 */
@PaletteInfo(paletteCategory = "Data", palette = "Excel Export with Protection")
public class ExcelExportSmartService extends AppianSmartService {

  private static final Logger LOG = LoggerFactory.getLogger(ExcelExportSmartService.class);
  private static final String XLSX_EXTENSION = "xlsx";

  private final ContentService contentService;

  private Map<String, Object>[] dataRows;
  private String[] allColumns;
  private String[] editableColumns;
  private Long targetFolderId;
  private String fileName;
  private String protectionPassword;

  private Long documentId;

  public ExcelExportSmartService(ContentService contentService) {
    this.contentService = contentService;
  }

  @Input(required = Required.ALWAYS)
  @Name("dataRows")
  public void setDataRows(Map<String, Object>[] dataRows) {
    this.dataRows = dataRows;
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

  @Input(required = Required.ALWAYS)
  @Name("targetFolderId")
  public void setTargetFolderId(Long targetFolderId) {
    this.targetFolderId = targetFolderId;
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

  @Name("documentId")
  public Long getDocumentId() {
    return documentId;
  }

  @Override
  public void run() throws SmartServiceException {
    try {
      ColumnProtectionConfig config = ColumnProtectionConfig.builder()
          .allColumns(allColumns != null ? Arrays.asList(allColumns) : Collections.emptyList())
          .editableColumns(editableColumns != null ? Arrays.asList(editableColumns) : Collections.emptyList())
          .protectionPassword(protectionPassword)
          .fileName(fileName)
          .targetFolderId(targetFolderId)
          .build();

      List<Map<String, Object>> rows = dataRows != null ? Arrays.asList(dataRows) : Collections.emptyList();
      byte[] excelBytes = new ExcelGenerator(config).generate(rows);

      this.documentId = saveAsDocument(excelBytes, config);
    } catch (IllegalArgumentException e) {
      throw new SmartServiceException(getClass(), e, "Invalid column configuration: " + e.getMessage());
    } catch (Exception e) {
      LOG.error("Excel export failed", e);
      throw new SmartServiceException(getClass(), e, "Failed to export Excel document: " + e.getMessage());
    }
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
  private Long saveAsDocument(byte[] excelBytes, ColumnProtectionConfig config) throws Exception {
    Document document = new Document();
    document.setName(config.getFileName());
    document.setParent(config.getTargetFolderId());
    document.setExtension(XLSX_EXTENSION);

    Long newDocumentId = contentService.create(document, ContentConstants.UNIQUE_NONE);
    document.setId(newDocumentId);

    try (ContentOutputStream contentOutputStream = contentService.upload(document, ContentConstants.VERSION_CURRENT)) {
      contentOutputStream.write(excelBytes);
    }
    return newDocumentId;
  }
}
