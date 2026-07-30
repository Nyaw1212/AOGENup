//----------------------------------
// Reassignment List Template
//----------------------------------

const AOGEN_LIST_TEMPLATE_ID =
  "1ZEbQaILGTcfRwLtwM3uQ4yCdP2BR9EjT3EHMKqJohCU";

//----------------------------------
// Add list option to AO Generator menu
//----------------------------------

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu("AO Generator")
    .addItem(
      "Generate Administrative Order",
      "generateAdministrativeOrder"
    )
    .addSeparator()
    .addItem(
      "Generate Reassignment - Table",
      "generateReassignmentFromSheet"
    )
    .addItem(
      "Generate Reassignment - List",
      "generateReassignmentListFromSheet"
    )
    .addToUi();
}

//----------------------------------
// Generate list report from sheet data
//----------------------------------

function generateReassignmentListFromSheet() {
  try {
    const spreadsheet = getAogenSpreadsheet_();
    const sheets = getRequiredSheets_(spreadsheet);
    const orderData = readOrderData_(sheets.orderSheet);
    const personnel = readPersonnel_(sheets.personnelSheet);

    validateReassignmentData_(orderData, personnel);

    const result = createListDocumentFromTemplate_({
      templateDocId: AOGEN_LIST_TEMPLATE_ID,
      outputName:
        `NBP Reassignment ${orderData.orderNumber} - List`,
      orderData,
      personnel,
    });

    saveGeneratedDocument_(sheets.orderSheet, result.url);

    sheets.orderSheet
      .getRange("B7")
      .setValue(
        "https://docs.google.com/document/d/" +
        AOGEN_LIST_TEMPLATE_ID +
        "/edit"
      );

    SpreadsheetApp
      .getUi()
      .alert(
        "Reassignment List generated successfully.\n\n" +
        result.url
      );

    return result;
  } catch (error) {
    showGenerationError_("Reassignment List", error);
    throw error;
  }
}

//----------------------------------
// Copy list template and fill placeholders
//----------------------------------

function createListDocumentFromTemplate_(options) {
  const templateDocId = clean_(options.templateDocId);

  if (!templateDocId) {
    throw new Error(
      "List template document ID is missing."
    );
  }

  const templateFile = DriveApp.getFileById(
    templateDocId
  );

  const outputFile = templateFile.makeCopy(
    options.outputName
  );

  const outputDocument = DocumentApp.openById(
    outputFile.getId()
  );

  const body = outputDocument.getBody();

  replaceReassignmentDetails_(
    body,
    options.orderData
  );

  replacePersonnelList_(
    body,
    options.personnel
  );

  outputDocument.saveAndClose();

  return {
    success: true,
    documentId: outputFile.getId(),
    url: outputFile.getUrl(),
    outputName: options.outputName,
  };
}

//----------------------------------
// Replace {{PERSONNEL_LIST}}
//----------------------------------

function replacePersonnelList_(body, personnel) {
  const match = body.findText(
    "\\{\\{PERSONNEL_LIST\\}\\}"
  );

  if (!match) {
    throw new Error(
      "The list template does not contain {{PERSONNEL_LIST}}."
    );
  }

  const textElement = match.getElement().asText();
  const paragraph = textElement.getParent().asParagraph();
  const parent = paragraph.getParent();
  const insertionIndex = parent.getChildIndex(paragraph);

  parent.removeChild(paragraph);

  personnel.forEach((person, index) => {
    const rankAndName = [
      person.rank,
      person.fullName,
    ]
      .filter(Boolean)
      .join(" ")
      .trim();

    const notes = clean_(person.notes);
    const notesText = notes
      ? ` (${notes})`
      : "";

    const itemText =
      `${rankAndName}, from ${person.from} ` +
      `to ${person.to}${notesText}` +
      getPersonnelListEnding_(
        index,
        personnel.length
      );

    const item = parent.insertListItem(
      insertionIndex + index,
      itemText
    );

    item
      .setGlyphType(
        DocumentApp.GlyphType.LATIN_LOWER
      )
      .setIndentStart(54)
      .setIndentFirstLine(36)
      .setSpacingBefore(0)
      .setSpacingAfter(3)
      .setLineSpacing(1)
      .setAlignment(
        DocumentApp.HorizontalAlignment.JUSTIFY
      );

    item
      .editAsText()
      .setFontFamily("Arial")
      .setFontSize(11);
  });
}

function getPersonnelListEnding_(index, total) {
  if (index === total - 1) {
    return ".";
  }

  if (index === total - 2) {
    return "; and";
  }

  return ";";
}

//----------------------------------
// Manual test
//----------------------------------

function testGenerateReassignmentListFromSheet() {
  const result = generateReassignmentListFromSheet();

  console.log(
    JSON.stringify(result, null, 2)
  );

  return result;
}
