package com.lowcodeminds.appian.plugins.excel;

import java.util.ArrayList;
import java.util.List;

import javax.naming.Context;

import static org.junit.jupiter.api.Assertions.assertTrue;
import org.junit.jupiter.api.Test;

import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.process.framework.MessageContainer;
import com.appiancorp.suiteapi.process.framework.SmartServiceContext;

/**
 * Covers {@link AdvanceExcelExport#validate} - the field-level validation rules
 * flagged as untested in the codebase review.
 *
 * <p>NOT covered here, and not coverable by a plain JUnit test against the
 * public {@code appian-plugin-sdk} artifact: {@link AdvanceExcelExport#run},
 * including its create-vs-overwrite document branch and its exception-wrapping
 * catch block. Both were attempted and both fail with {@code NoClassDefFoundError}
 * for classes that exist only inside a running Appian server, not in the
 * plugin-sdk jar plugins compile/test against:
 * <ul>
 *   <li>{@code new Document()} (used by {@code saveAsDocument()} for every
 *       {@code run()} call, success or failure) requires
 *       {@code com.appiancorp.content.DocumentHelper} at construction time.</li>
 *   <li>{@code SmartServiceException.Builder#build()} (used by {@code run()}'s
 *       catch block) requires {@code com.appiancorp.process.admin.LoadSmartNodeACSchemas}
 *       to resolve this plug-in's message bundle.</li>
 * </ul>
 * Both work correctly when actually deployed to Appian, where these classes are
 * present - this is a test-environment limitation, not a defect in this
 * plug-in's code. Verifying the create/overwrite branch and the exception path
 * requires either a real Appian server or an Appian-provided test harness;
 * faking out these internal, undocumented classes on the test classpath would
 * be brittle stubbing of Appian internals, not a real test.
 */
class AdvanceExcelExportTest {

  private static SQLSheetData sheet(String sheetName, String sql) {
    SQLSheetData data = new SQLSheetData();
    data.setSheetName(sheetName);
    data.setSqlQuery(sql);
    return data;
  }

  // None of AdvanceExcelExport.validate()'s rules read smartServiceContext,
  // contentService, or context, so plain nulls are enough to construct the
  // instance under test - no mocking library needed for this test class.
  private static AdvanceExcelExport newServiceForValidation() {
    ContentService contentService = null;
    Context jndiContext = null;
    SmartServiceContext smartServiceContext = null;
    return new AdvanceExcelExport(smartServiceContext, contentService, jndiContext);
  }

  /** Minimal {@link MessageContainer} fake - just records which bundle keys were added. */
  private static final class RecordingMessageContainer implements MessageContainer {
    private final List<String> errorKeys = new ArrayList<>();

    @Override
    public void addError(String acpKey, String messageKey, Object... messageValues) {
      errorKeys.add(messageKey);
    }

    @Override
    public boolean isEmpty() {
      return errorKeys.isEmpty();
    }

    @Override
    public String getErrorMessage() {
      return String.join(", ", errorKeys);
    }

    boolean hasError(String messageKey) {
      return errorKeys.contains(messageKey);
    }
  }

  @Test
  void testValidateRejectsEmptySheetDataList() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[0]);

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.sqlSheetDataListEmpty"));
  }

  @Test
  void testValidateRejectsMoreThanMaxSheets() {
    SQLSheetData[] sheets = new SQLSheetData[21];
    for (int i = 0; i < sheets.length; i++) {
      sheets[i] = sheet("Sheet" + i, "SELECT 1");
    }
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(sheets);

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.tooManySheets"));
  }

  @Test
  void testValidateRejectsMissingSheetName() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[] {sheet(" ", "SELECT 1")});

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.sheetNameMissing"));
  }

  @Test
  void testValidateRejectsSheetNameOverExcelLimit() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[] {sheet("A".repeat(32), "SELECT 1")});

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.sheetNameTooLong"));
  }

  @Test
  void testValidateRejectsIllegalSheetNameCharacters() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[] {sheet("Bad:Name", "SELECT 1")});

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.sheetNameIllegalCharacters"));
  }

  @Test
  void testValidateRejectsDuplicateSheetNamesCaseInsensitively() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[] {
        sheet("Report", "SELECT 1"),
        sheet("REPORT", "SELECT 2")
    });

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.duplicateSheetName"));
  }

  @Test
  void testValidateRejectsMissingSqlQuery() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[] {sheet("Sheet1", " ")});

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.sqlQueryMissing"));
  }

  @Test
  void testValidateRejectsNonSelectQuery() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[] {sheet("Sheet1", "DELETE FROM t")});

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.sqlQueryNotSelect"));
  }

  @Test
  void testValidateRejectsQueryWithEmbeddedSemicolonNotJustTrailing() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[] {sheet("Sheet1", "SELECT 1; SELECT 2")});

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.hasError("advanceExcelExport.error.sqlQueryContainsSemicolon"));
  }

  @Test
  void testValidateAcceptsFullyValidInput() {
    AdvanceExcelExport service = newServiceForValidation();
    service.setSqlSheetDataList(new SQLSheetData[] {
        sheet("Sheet1", "SELECT 1 AS id"),
        sheet("Sheet2", "SELECT 2 AS id")
    });

    RecordingMessageContainer messages = new RecordingMessageContainer();
    service.validate(messages);

    assertTrue(messages.isEmpty());
  }
}
