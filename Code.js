//----------------------------------
// AOGENup Configuration
//----------------------------------

const AOGEN_CONFIG = Object.freeze({
  spreadsheetId:
    "1jgML6lCg1Ap9RIfRjin7kK4XS89KMxO0cubRtcWkTcU",

  reassignmentTemplateId:
    "1DCCt4ePrFrmOE9jnKpfNk9vORZSwzCGw5i5IyZtIOdM",

  sheetNames: {
    order: "Order",
    personnel: "Personnel",
  },
});


//----------------------------------
// Spreadsheet Menu
//----------------------------------

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu("AO Generator")
    .addItem(
      "Generate Administrative Order",
      "generateAdministrativeOrder"
    )
    .addItem(
      "Generate Reassignment Report",
      "generateReassignmentFromSheet"
    )
    .addToUi();
}


//==================================================
// EXISTING ADMINISTRATIVE ORDER
// Uses the template link stored in Order!B7
//==================================================

function generateAdministrativeOrder() {
  try {
    const spreadsheet =
      getAogenSpreadsheet_();

    const sheets =
      getRequiredSheets_(spreadsheet);

    const orderData =
      readOrderData_(sheets.orderSheet);

    const personnel =
      readPersonnel_(sheets.personnelSheet);

    validateAdministrativeOrder_(
      orderData,
      personnel
    );

    const result =
      createDocumentFromTemplate_({
        templateDocId:
          orderData.templateDocId,

        outputName:
          `NBP AO ${orderData.orderNumber}`,

        orderData,
        personnel,

        templateType:
          "ADMINISTRATIVE_ORDER",
      });

    saveGeneratedDocument_(
      sheets.orderSheet,
      result.url
    );

    SpreadsheetApp
      .getUi()
      .alert(
        "Administrative Order generated successfully.\n\n" +
        result.url
      );

    return result;

  } catch (error) {
    showGenerationError_(
      "Administrative Order",
      error
    );

    throw error;
  }
}


//==================================================
// REASSIGNMENT FROM SHEET
// Uses Order and Personnel tabs
//==================================================

function generateReassignmentFromSheet() {
  try {
    const spreadsheet =
      getAogenSpreadsheet_();

    const sheets =
      getRequiredSheets_(spreadsheet);

    const orderData =
      readOrderData_(sheets.orderSheet);

    const personnel =
      readPersonnel_(sheets.personnelSheet);

    validateReassignmentData_(
      orderData,
      personnel
    );

    const result =
      createDocumentFromTemplate_({
        templateDocId:
          AOGEN_CONFIG.reassignmentTemplateId,

        outputName:
          `NBP Reassignment ${orderData.orderNumber}`,

        orderData,
        personnel,

        templateType:
          "REASSIGNMENT",
      });

    saveGeneratedDocument_(
      sheets.orderSheet,
      result.url
    );

    SpreadsheetApp
      .getUi()
      .alert(
        "Reassignment Report generated successfully.\n\n" +
        result.url
      );

    return result;

  } catch (error) {
    showGenerationError_(
      "Reassignment Report",
      error
    );

    throw error;
  }
}


//==================================================
// REASSIGNMENT PAYLOAD FUNCTION
// Will be called later by Personnel Filter
//==================================================

function generateReassignmentReport(payload) {
  const lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const normalizedPayload =
      normalizeReassignmentPayload_(payload);

    validateReassignmentPayload_(
      normalizedPayload
    );

    const spreadsheet =
      getAogenSpreadsheet_();

    const sheets =
      getRequiredSheets_(spreadsheet);

    writeOrderData_(
      sheets.orderSheet,
      normalizedPayload
    );

    writePersonnelData_(
      sheets.personnelSheet,
      normalizedPayload.personnel
    );

    const orderData =
      readOrderData_(sheets.orderSheet);

    const personnel =
      readPersonnel_(sheets.personnelSheet);

    validateReassignmentData_(
      orderData,
      personnel
    );

    const result =
      createDocumentFromTemplate_({
        templateDocId:
          AOGEN_CONFIG.reassignmentTemplateId,

        outputName:
          `NBP Reassignment ${orderData.orderNumber}`,

        orderData,
        personnel,

        templateType:
          "REASSIGNMENT",
      });

    saveGeneratedDocument_(
      sheets.orderSheet,
      result.url
    );

    return {
      success: true,
      message:
        "Reassignment Report generated successfully.",
      documentId:
        result.documentId,
      url:
        result.url,
      outputName:
        result.outputName,
      personnelCount:
        personnel.length,
    };

  } catch (error) {
    console.error(
      "generateReassignmentReport:",
      error
    );

    return {
      success: false,
      message:
        getErrorMessage_(error),
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {
      // Lock was not acquired.
    }
  }
}


//==================================================
// DOCUMENT GENERATION
//==================================================

function createDocumentFromTemplate_(options) {
  const templateDocId =
    clean_(options.templateDocId);

  if (!templateDocId) {
    throw new Error(
      "Template document ID is missing."
    );
  }

  const templateFile =
    DriveApp.getFileById(
      templateDocId
    );

  const outputFile =
    templateFile.makeCopy(
      options.outputName
    );

  const outputDocument =
    DocumentApp.openById(
      outputFile.getId()
    );

  const body =
    outputDocument.getBody();

  if (
    options.templateType ===
    "REASSIGNMENT"
  ) {
    replaceReassignmentDetails_(
      body,
      options.orderData
    );
  } else {
    replaceStandardPlaceholders_(
      body,
      options.orderData
    );
  }

  replacePersonnelTable_(
    body,
    options.personnel
  );

  outputDocument.saveAndClose();

  return {
    success: true,
    documentId:
      outputFile.getId(),
    url:
      outputFile.getUrl(),
    outputName:
      options.outputName,
  };
}


//==================================================
// REASSIGNMENT TEMPLATE DETAILS
// Placeholder-based template replacement
//==================================================

function replaceReassignmentDetails_(
  body,
  orderData
) {
  // The Reassignment Google Docs template must contain:
  // {{ORDER_DATE}}
  // {{ORDER_NUMBER}}
  // {{SIGNING_OFFICER}}
  // {{SIGNING_POSITION}}
  // {{FILE_REFERENCE}}
  //
  // Text formatting and alignment are inherited from the template.

  replaceStandardPlaceholders_(
    body,
    orderData
  );
}


//==================================================
// PLACEHOLDERS
//==================================================

function replaceStandardPlaceholders_(
  body,
  orderData
) {
  replacePlaceholder_(
    body,
    "{{ORDER_NUMBER}}",
    orderData.orderNumber
  );

  replacePlaceholder_(
    body,
    "{{ORDER_DATE}}",
    orderData.orderDate
  );

  replacePlaceholder_(
    body,
    "{{SIGNING_OFFICER}}",
    orderData.signingOfficer
  );

  replacePlaceholder_(
    body,
    "{{SIGNING_POSITION}}",
    orderData.signingPosition
  );

  replacePlaceholder_(
    body,
    "{{FILE_REFERENCE}}",
    orderData.fileReference || ""
  );
}


function replacePlaceholder_(
  body,
  placeholder,
  replacement
) {
  body.replaceText(
    escapeForRegex_(placeholder),
    String(replacement || "")
  );
}


//==================================================
// PERSONNEL TABLE
//==================================================

function replacePersonnelTable_(
  body,
  personnel
) {
  const tables =
    body.getTables();

  if (!tables.length) {
    throw new Error(
      "No personnel table was found in the document template."
    );
  }

  const table =
    tables[0];

  if (table.getNumRows() < 1) {
    throw new Error(
      "The personnel table does not contain a header row."
    );
  }

  //----------------------------------
  // Keep only the header row
  //----------------------------------

  while (
    table.getNumRows() > 1
  ) {
    table.removeRow(1);
  }

  //----------------------------------
  // Append personnel
  //----------------------------------

  personnel.forEach(person => {
    const row =
      table.appendTableRow();

    const rankAndName =
      [
        person.rank,
        person.fullName,
      ]
        .filter(Boolean)
        .join(" ");

    row.appendTableCell(
      String(person.number)
    );

    row.appendTableCell(
      rankAndName
    );

    row.appendTableCell(
      person.from
    );

    const destination =
      person.notes
        ? `${person.to}\n${person.notes}`
        : person.to;

    row.appendTableCell(
      destination
    );

    formatPersonnelRow_(row);
  });
}


function formatPersonnelRow_(
  row
) {
  for (
    let column = 0;
    column < row.getNumCells();
    column++
  ) {
    const cell =
      row.getCell(column);

    cell.setVerticalAlignment(
      DocumentApp
        .VerticalAlignment
        .CENTER
    );

    cell
      .editAsText()
      .setFontFamily("Arial")
      .setFontSize(11);

    const alignment =
      column === 1
        ? DocumentApp
            .HorizontalAlignment
            .LEFT
        : DocumentApp
            .HorizontalAlignment
            .CENTER;

    for (
      let childIndex = 0;
      childIndex < cell.getNumChildren();
      childIndex++
    ) {
      const child =
        cell.getChild(childIndex);

      if (
        child.getType() ===
        DocumentApp
          .ElementType
          .PARAGRAPH
      ) {
        child
          .asParagraph()
          .setAlignment(
            alignment
          );
      }
    }
  }
}


//==================================================
// READ SHEET DATA
//==================================================

function readOrderData_(sheet) {
  return {
    orderNumber:
      clean_(
        sheet
          .getRange("B2")
          .getDisplayValue()
      ),

    orderDate:
      formatOrderDate_(
        sheet
          .getRange("B3")
          .getValue()
      ),

    signingOfficer:
      clean_(
        sheet
          .getRange("B4")
          .getDisplayValue()
      ),

    signingPosition:
      clean_(
        sheet
          .getRange("B5")
          .getDisplayValue()
      ),

    fileReference:
      clean_(
        sheet
          .getRange("B6")
          .getDisplayValue()
      ),

    templateDocId:
      extractFileId_(
        clean_(
          sheet
            .getRange("B7")
            .getDisplayValue()
        )
      ),
  };
}


function readPersonnel_(sheet) {
  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getDisplayValues();

  return values
    .filter(row =>
      row
        .slice(1, 5)
        .some(value =>
          clean_(value) !== ""
        )
    )
    .map((row, index) => ({
      number:
        index + 1,

      rank:
        clean_(row[1])
          .toUpperCase(),

      fullName:
        clean_(row[2])
          .toUpperCase(),

      from:
        clean_(row[3])
          .toUpperCase(),

      to:
        clean_(row[4])
          .toUpperCase(),

      notes:
        clean_(row[5]),
    }));
}


//==================================================
// WRITE PAYLOAD TO SHEETS
//==================================================

function writeOrderData_(
  sheet,
  payload
) {
  sheet
    .getRange("B2")
    .setValue(
      payload.orderNumber
    );

  sheet
    .getRange("B3")
    .setValue(
      parseDateForSheet_(
        payload.orderDate
      )
    );

  sheet
    .getRange("B3")
    .setNumberFormat(
      "MM/dd/yyyy"
    );

  sheet
    .getRange("B4")
    .setValue(
      payload.signingOfficer
    );

  sheet
    .getRange("B5")
    .setValue(
      payload.signingPosition
    );

  sheet
    .getRange("B6")
    .setValue(
      payload.fileReference
    );

  sheet
    .getRange("B7")
    .setValue(
      "https://docs.google.com/document/d/" +
      AOGEN_CONFIG.reassignmentTemplateId +
      "/edit"
    );

  sheet
    .getRange("B8:B9")
    .clearContent();
}


function writePersonnelData_(
  sheet,
  personnel
) {
  const lastRow =
    sheet.getLastRow();

  if (lastRow > 1) {
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .clearContent();
  }

  if (!personnel.length) {
    return;
  }

  const rows =
    personnel.map(
      (person, index) => [
        index + 1,
        clean_(person.rank)
          .toUpperCase(),
        clean_(person.fullName)
          .toUpperCase(),
        clean_(person.from)
          .toUpperCase(),
        clean_(person.to)
          .toUpperCase(),
        clean_(person.notes),
      ]
    );

  sheet
    .getRange(
      2,
      1,
      rows.length,
      6
    )
    .setValues(rows);
}


//==================================================
// VALIDATION
//==================================================

function validateAdministrativeOrder_(
  orderData,
  personnel
) {
  const missing = [];

  if (!orderData.orderNumber) {
    missing.push("Order Number");
  }

  if (!orderData.orderDate) {
    missing.push("Order Date");
  }

  if (!orderData.signingOfficer) {
    missing.push("Signing Officer");
  }

  if (!orderData.signingPosition) {
    missing.push("Signing Position");
  }

  if (!orderData.templateDocId) {
    missing.push("Template Document");
  }

  throwIfMissingFields_(
    missing,
    "Complete the following fields in the Order tab:"
  );

  validatePersonnel_(personnel);
}


function validateReassignmentData_(
  orderData,
  personnel
) {
  const missing = [];

  if (!orderData.orderNumber) {
    missing.push("Order Number");
  }

  if (!orderData.orderDate) {
    missing.push("Order Date");
  }

  if (!orderData.signingOfficer) {
    missing.push("Signing Officer");
  }

  if (!orderData.signingPosition) {
    missing.push("Signing Position");
  }

  throwIfMissingFields_(
    missing,
    "Complete the following fields in the Order tab:"
  );

  validatePersonnel_(personnel);
}


function validatePersonnel_(
  personnel
) {
  if (!personnel.length) {
    throw new Error(
      "Add at least one personnel entry in the Personnel tab."
    );
  }

  personnel.forEach(
    (person, index) => {
      const missing = [];

      if (!person.rank) {
        missing.push("Rank");
      }

      if (!person.fullName) {
        missing.push(
          "Full Name"
        );
      }

      if (!person.from) {
        missing.push("From");
      }

      if (!person.to) {
        missing.push("To");
      }

      if (missing.length) {
        throw new Error(
          `Personnel row ${index + 2} is missing: ` +
          missing.join(", ")
        );
      }
    }
  );
}


function validateReassignmentPayload_(
  payload
) {
  const missing = [];

  if (!payload.orderNumber) {
    missing.push("Order Number");
  }

  if (!payload.orderDate) {
    missing.push("Order Date");
  }

  if (!payload.signingOfficer) {
    missing.push("Signing Officer");
  }

  if (!payload.signingPosition) {
    missing.push("Signing Position");
  }

  throwIfMissingFields_(
    missing,
    "Missing reassignment fields:"
  );

  validatePersonnel_(
    payload.personnel
  );
}


function throwIfMissingFields_(
  missing,
  heading
) {
  if (!missing.length) {
    return;
  }

  throw new Error(
    heading +
    "\n\n" +
    missing.join("\n")
  );
}


//==================================================
// NORMALIZE PAYLOAD
//==================================================

function normalizeReassignmentPayload_(
  payload
) {
  const source =
    payload &&
    typeof payload === "object"
      ? payload
      : {};

  const personnel =
    Array.isArray(source.personnel)
      ? source.personnel
      : [];

  return {
    orderNumber:
      clean_(source.orderNumber),

    orderDate:
      source.orderDate || "",

    signingOfficer:
      clean_(source.signingOfficer),

    signingPosition:
      clean_(source.signingPosition),

    fileReference:
      clean_(source.fileReference),

    personnel:
      personnel.map(person => ({
        rank:
          clean_(person.rank),

        fullName:
          clean_(person.fullName),

        from:
          clean_(person.from),

        to:
          clean_(person.to),

        notes:
          clean_(person.notes),
      })),
  };
}


//==================================================
// SHEET HELPERS
//==================================================

function getAogenSpreadsheet_() {
  return SpreadsheetApp.openById(
    AOGEN_CONFIG.spreadsheetId
  );
}


function getRequiredSheets_(
  spreadsheet
) {
  const orderSheet =
    spreadsheet.getSheetByName(
      AOGEN_CONFIG.sheetNames.order
    );

  const personnelSheet =
    spreadsheet.getSheetByName(
      AOGEN_CONFIG.sheetNames.personnel
    );

  const missing = [];

  if (!orderSheet) {
    missing.push(
      AOGEN_CONFIG.sheetNames.order
    );
  }

  if (!personnelSheet) {
    missing.push(
      AOGEN_CONFIG.sheetNames.personnel
    );
  }

  if (missing.length) {
    throw new Error(
      "Required tabs were not found:\n\n" +
      missing.join("\n")
    );
  }

  return {
    orderSheet,
    personnelSheet,
  };
}


function saveGeneratedDocument_(
  orderSheet,
  documentUrl
) {
  orderSheet
    .getRange("B8")
    .setValue(
      documentUrl
    );

  orderSheet
    .getRange("B9")
    .setValue(
      new Date()
    );

  orderSheet
    .getRange("B9")
    .setNumberFormat(
      "MM/dd/yyyy hh:mm AM/PM"
    );
}


//==================================================
// DATE HELPERS
//==================================================

function formatOrderDate_(
  value
) {
  if (!value) {
    return "";
  }

  if (
    Object.prototype
      .toString
      .call(value) ===
    "[object Date]"
  ) {
    return Utilities
      .formatDate(
        value,
        Session.getScriptTimeZone(),
        "dd MMMM yyyy"
      )
      .toUpperCase();
  }

  const text =
    clean_(value);

  const parsed =
    new Date(text);

  if (!isNaN(parsed.getTime())) {
    return Utilities
      .formatDate(
        parsed,
        Session.getScriptTimeZone(),
        "dd MMMM yyyy"
      )
      .toUpperCase();
  }

  return text.toUpperCase();
}


function parseDateForSheet_(
  value
) {
  if (
    Object.prototype
      .toString
      .call(value) ===
    "[object Date]"
  ) {
    return value;
  }

  const text =
    clean_(value);

  const parsed =
    new Date(text);

  return !isNaN(parsed.getTime())
    ? parsed
    : text;
}


//==================================================
// GENERAL UTILITIES
//==================================================

function extractFileId_(
  value
) {
  if (!value) {
    return "";
  }

  const match =
    value.match(
      /[-\w]{25,}/
    );

  return match
    ? match[0]
    : value;
}


function clean_(
  value
) {
  return String(
    value === null ||
    value === undefined
      ? ""
      : value
  ).trim();
}


function escapeForRegex_(
  value
) {
  return String(value).replace(
    /[.*+?^${}()|[\]\\]/g,
    "\\$&"
  );
}


function getErrorMessage_(
  error
) {
  if (!error) {
    return "Unknown error.";
  }

  return error.message ||
    String(error);
}


function showGenerationError_(
  reportName,
  error
) {
  SpreadsheetApp
    .getUi()
    .alert(
      `${reportName} generation failed.\n\n` +
      getErrorMessage_(error)
    );
}


//==================================================
// MANUAL TEST
//==================================================

function testGenerateReassignmentFromSheet() {
  const result =
    generateReassignmentFromSheet();

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}
//----------------------------------
// Reassignment Web API
//----------------------------------

function doPost(e) {
  try {
    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {
      throw new Error(
        "No request body was received."
      );
    }

    const request = JSON.parse(
      e.postData.contents
    );

    verifyReassignmentApiKey_(
      request.apiKey
    );

    if (
      String(request.action || "")
        .toUpperCase() !==
      "GENERATE_REASSIGNMENT"
    ) {
      throw new Error(
        "Unsupported request action."
      );
    }

    const result =
      generateReassignmentReport(
        request.payload
      );

    return createReassignmentJsonResponse_(
      result
    );

  } catch (error) {
    return createReassignmentJsonResponse_({
      success: false,
      message:
        error.message ||
        String(error),
    });
  }
}


function verifyReassignmentApiKey_(
  suppliedKey
) {
  const configuredKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "REASSIGNMENT_API_KEY"
      );

  if (!configuredKey) {
    throw new Error(
      "REASSIGNMENT_API_KEY is not configured."
    );
  }

  if (
    String(suppliedKey || "").trim() !==
    configuredKey.trim()
  ) {
    throw new Error(
      "Unauthorized request."
    );
  }
}


function createReassignmentJsonResponse_(
  data
) {
  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}