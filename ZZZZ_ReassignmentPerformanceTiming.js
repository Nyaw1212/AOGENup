//----------------------------------
// Reassignment report performance timing
// Loaded after Code.js so these functions instrument the live doPost path.
//----------------------------------

function generateReassignmentReport(payload) {
  const startedAt = Date.now();
  const timing = {
    totalMs: 0,
    lockMs: 0,
    normalizeMs: 0,
    validatePayloadMs: 0,
    openSpreadsheetMs: 0,
    loadSheetsMs: 0,
    writeOrderMs: 0,
    writePersonnelMs: 0,
    readOrderMs: 0,
    readPersonnelMs: 0,
    validateSheetMs: 0,
    createDocumentMs: 0,
    saveGeneratedMs: 0,
    releaseLockMs: 0,
    personnel: 0,
  };

  console.log("[PERF][AOGEN Reassignment] START");

  const lock = LockService.getScriptLock();
  const lockStartedAt = Date.now();

  try {
    lock.waitLock(30000);
    timing.lockMs = Date.now() - lockStartedAt;

    const normalizeStartedAt = Date.now();
    const normalizedPayload = normalizeReassignmentPayload_(payload);
    timing.normalizeMs = Date.now() - normalizeStartedAt;
    timing.personnel = Array.isArray(normalizedPayload.personnel)
      ? normalizedPayload.personnel.length
      : 0;

    const validatePayloadStartedAt = Date.now();
    validateReassignmentPayload_(normalizedPayload);
    timing.validatePayloadMs = Date.now() - validatePayloadStartedAt;

    const openStartedAt = Date.now();
    const spreadsheet = getAogenSpreadsheet_();
    timing.openSpreadsheetMs = Date.now() - openStartedAt;

    const sheetsStartedAt = Date.now();
    const sheets = getRequiredSheets_(spreadsheet);
    timing.loadSheetsMs = Date.now() - sheetsStartedAt;

    const writeOrderStartedAt = Date.now();
    writeOrderData_(sheets.orderSheet, normalizedPayload);
    timing.writeOrderMs = Date.now() - writeOrderStartedAt;

    const writePersonnelStartedAt = Date.now();
    writePersonnelData_(sheets.personnelSheet, normalizedPayload.personnel);
    timing.writePersonnelMs = Date.now() - writePersonnelStartedAt;

    const readOrderStartedAt = Date.now();
    const orderData = readOrderData_(sheets.orderSheet);
    timing.readOrderMs = Date.now() - readOrderStartedAt;

    const readPersonnelStartedAt = Date.now();
    const personnel = readPersonnel_(sheets.personnelSheet);
    timing.readPersonnelMs = Date.now() - readPersonnelStartedAt;

    const validateSheetStartedAt = Date.now();
    validateReassignmentData_(orderData, personnel);
    timing.validateSheetMs = Date.now() - validateSheetStartedAt;

    const createStartedAt = Date.now();
    const result = createDocumentFromTemplate_({
      templateDocId: AOGEN_CONFIG.reassignmentTemplateId,
      outputName: `NBP Reassignment ${orderData.orderNumber}`,
      orderData,
      personnel,
      templateType: "REASSIGNMENT",
    });
    timing.createDocumentMs = Date.now() - createStartedAt;

    const saveStartedAt = Date.now();
    saveGeneratedDocument_(sheets.orderSheet, result.url);
    timing.saveGeneratedMs = Date.now() - saveStartedAt;

    return {
      success: true,
      message: "Reassignment Report generated successfully.",
      documentId: result.documentId,
      url: result.url,
      outputName: result.outputName,
      personnelCount: personnel.length,
      timing,
    };
  } catch (error) {
    console.error("generateReassignmentReport:", error);
    return {
      success: false,
      message: getErrorMessage_(error),
      timing,
    };
  } finally {
    const releaseStartedAt = Date.now();
    try {
      lock.releaseLock();
    } catch (ignored) {}
    timing.releaseLockMs = Date.now() - releaseStartedAt;
    timing.totalMs = Date.now() - startedAt;

    console.log(
      "[PERF][AOGEN Reassignment] total=%sms personnel=%s lock=%sms normalize=%sms validatePayload=%sms openSpreadsheet=%sms loadSheets=%sms writeOrder=%sms writePersonnel=%sms readOrder=%sms readPersonnel=%sms validateSheet=%sms createDocument=%sms saveGenerated=%sms releaseLock=%sms",
      timing.totalMs,
      timing.personnel,
      timing.lockMs,
      timing.normalizeMs,
      timing.validatePayloadMs,
      timing.openSpreadsheetMs,
      timing.loadSheetsMs,
      timing.writeOrderMs,
      timing.writePersonnelMs,
      timing.readOrderMs,
      timing.readPersonnelMs,
      timing.validateSheetMs,
      timing.createDocumentMs,
      timing.saveGeneratedMs,
      timing.releaseLockMs
    );
  }
}

function createDocumentFromTemplate_(options) {
  const startedAt = Date.now();
  const timing = {
    totalMs: 0,
    validateMs: 0,
    getTemplateMs: 0,
    copyTemplateMs: 0,
    openDocumentMs: 0,
    getBodyMs: 0,
    replaceDetailsMs: 0,
    replacePersonnelMs: 0,
    saveCloseMs: 0,
    getResultMs: 0,
    personnel: Array.isArray(options && options.personnel)
      ? options.personnel.length
      : 0,
  };

  console.log(
    "[PERF][AOGEN Document] START personnel=%s",
    timing.personnel
  );

  const validateStartedAt = Date.now();
  const templateDocId = clean_(options.templateDocId);
  if (!templateDocId) {
    throw new Error("Template document ID is missing.");
  }
  timing.validateMs = Date.now() - validateStartedAt;

  const getTemplateStartedAt = Date.now();
  const templateFile = DriveApp.getFileById(templateDocId);
  timing.getTemplateMs = Date.now() - getTemplateStartedAt;

  const copyStartedAt = Date.now();
  const outputFile = templateFile.makeCopy(options.outputName);
  timing.copyTemplateMs = Date.now() - copyStartedAt;

  const openDocumentStartedAt = Date.now();
  const outputDocument = DocumentApp.openById(outputFile.getId());
  timing.openDocumentMs = Date.now() - openDocumentStartedAt;

  const getBodyStartedAt = Date.now();
  const body = outputDocument.getBody();
  timing.getBodyMs = Date.now() - getBodyStartedAt;

  const detailsStartedAt = Date.now();
  if (options.templateType === "REASSIGNMENT") {
    replaceReassignmentDetails_(body, options.orderData);
  } else {
    replaceStandardPlaceholders_(body, options.orderData);
  }
  timing.replaceDetailsMs = Date.now() - detailsStartedAt;

  const personnelStartedAt = Date.now();
  replacePersonnelTable_(body, options.personnel);
  timing.replacePersonnelMs = Date.now() - personnelStartedAt;

  const saveStartedAt = Date.now();
  outputDocument.saveAndClose();
  timing.saveCloseMs = Date.now() - saveStartedAt;

  const resultStartedAt = Date.now();
  const result = {
    success: true,
    documentId: outputFile.getId(),
    url: outputFile.getUrl(),
    outputName: options.outputName,
  };
  timing.getResultMs = Date.now() - resultStartedAt;
  timing.totalMs = Date.now() - startedAt;

  console.log(
    "[PERF][AOGEN Document] total=%sms personnel=%s validate=%sms getTemplate=%sms copyTemplate=%sms openDocument=%sms getBody=%sms replaceDetails=%sms replacePersonnel=%sms saveClose=%sms getResult=%sms",
    timing.totalMs,
    timing.personnel,
    timing.validateMs,
    timing.getTemplateMs,
    timing.copyTemplateMs,
    timing.openDocumentMs,
    timing.getBodyMs,
    timing.replaceDetailsMs,
    timing.replacePersonnelMs,
    timing.saveCloseMs,
    timing.getResultMs
  );

  result.timing = timing;
  return result;
}
