//----------------------------------
// Reassignment List Templates
//----------------------------------

const AOGEN_LIST_TEMPLATE_ID =
  "1ZEbQaILGTcfRwLtwM3uQ4yCdP2BR9EjT3EHMKqJohCU";

const AOGEN_GROUP_TEMPLATE_ID =
  "1mFJFfg1Xyo-qdg50nwgAKo_yCo1be4HfjwVjwuEqXVE";

//----------------------------------
// AO Generator menu
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
      "Generate Reassignment - Numbered List",
      "generateReassignmentListFromSheet"
    )
    .addItem(
      "Generate Reassignment - Grouped Movement",
      "generateReassignmentGroupedFromSheet"
    )
    .addToUi();
}

//----------------------------------
// Generate numbered list report
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
    saveTemplateLink_(sheets.orderSheet, AOGEN_LIST_TEMPLATE_ID);

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
// Generate grouped movement report
//----------------------------------

function generateReassignmentGroupedFromSheet() {
  try {
    const spreadsheet = getAogenSpreadsheet_();
    const sheets = getRequiredSheets_(spreadsheet);
    const orderData = readOrderData_(sheets.orderSheet);
    const personnel = readPersonnel_(sheets.personnelSheet);

    validateReassignmentData_(orderData, personnel);

    const result = createGroupedDocumentFromTemplate_({
      templateDocId: AOGEN_GROUP_TEMPLATE_ID,
      outputName:
        `NBP Reassignment ${orderData.orderNumber} - Grouped`,
      orderData,
      personnel,
    });

    saveGeneratedDocument_(sheets.orderSheet, result.url);
    saveTemplateLink_(sheets.orderSheet, AOGEN_GROUP_TEMPLATE_ID);

    SpreadsheetApp
      .getUi()
      .alert(
        "Grouped Reassignment generated successfully.\n\n" +
        result.url
      );

    return result;
  } catch (error) {
    showGenerationError_("Grouped Reassignment", error);
    throw error;
  }
}

function saveTemplateLink_(orderSheet, templateId) {
  orderSheet
    .getRange("B7")
    .setValue(
      "https://docs.google.com/document/d/" +
      templateId +
      "/edit"
    );
}

//----------------------------------
// Copy numbered-list template
//----------------------------------

function createListDocumentFromTemplate_(options) {
  const templateDocId = clean_(options.templateDocId);

  if (!templateDocId) {
    throw new Error("List template document ID is missing.");
  }

  const templateFile = DriveApp.getFileById(templateDocId);
  const outputFile = templateFile.makeCopy(options.outputName);
  const outputDocument = DocumentApp.openById(outputFile.getId());
  const body = outputDocument.getBody();

  replaceReassignmentDetails_(body, options.orderData);
  replacePersonnelList_(body, options.personnel);

  outputDocument.saveAndClose();

  return {
    success: true,
    documentId: outputFile.getId(),
    url: outputFile.getUrl(),
    outputName: options.outputName,
  };
}

//----------------------------------
// Copy grouped movement template
//----------------------------------

function createGroupedDocumentFromTemplate_(options) {
  const templateDocId = clean_(options.templateDocId);

  if (!templateDocId) {
    throw new Error("Grouped template document ID is missing.");
  }

  const templateFile = DriveApp.getFileById(templateDocId);
  const outputFile = templateFile.makeCopy(options.outputName);
  const outputDocument = DocumentApp.openById(outputFile.getId());
  const body = outputDocument.getBody();

  replaceReassignmentDetails_(body, options.orderData);
  replaceMovementGroups_(body, options.personnel);

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
  const placeholder = findPlaceholderParagraph_(
    body,
    "{{PERSONNEL_LIST}}"
  );

  const parent = placeholder.parent;
  const insertionIndex = placeholder.index;
  parent.removeChild(placeholder.paragraph);

  personnel.forEach((person, index) => {
    const rankAndName = [person.rank, person.fullName]
      .filter(Boolean)
      .join(" ")
      .trim();

    const notes = clean_(person.notes);
    const notesText = notes ? ` (${notes})` : "";

    const itemText =
      `${rankAndName}, from ${person.from} ` +
      `to ${person.to}${notesText}` +
      getPersonnelListEnding_(index, personnel.length);

    insertNumberedItem_(
      parent,
      insertionIndex + index,
      itemText
    );
  });
}

//----------------------------------
// Replace {{MOVEMENT_GROUPS}}
//----------------------------------

function replaceMovementGroups_(body, personnel) {
  const placeholder = findPlaceholderParagraph_(
    body,
    "{{MOVEMENT_GROUPS}}"
  );

  const parent = placeholder.parent;
  let insertionIndex = placeholder.index;
  parent.removeChild(placeholder.paragraph);

  const groups = groupPersonnelByMovement_(personnel);
  let personnelNumber = 1;

  groups.forEach((group, groupIndex) => {
    const heading = parent.insertParagraph(
      insertionIndex++,
      `${group.from} TO ${group.to}`
    );

    heading
      .setAlignment(DocumentApp.HorizontalAlignment.LEFT)
      .setSpacingBefore(groupIndex === 0 ? 0 : 8)
      .setSpacingAfter(4)
      .setLineSpacing(1);

    heading
      .editAsText()
      .setFontFamily("Arial")
      .setFontSize(11)
      .setBold(true);

    group.personnel.forEach((person, personIndex) => {
      const rankAndName = [person.rank, person.fullName]
        .filter(Boolean)
        .join(" ")
        .trim();

      const notes = clean_(person.notes);
      const notesText = notes ? ` (${notes})` : "";
      const isLastOverall =
        groupIndex === groups.length - 1 &&
        personIndex === group.personnel.length - 1;

      const itemText =
        `${rankAndName}${notesText}` +
        (isLastOverall ? "." : ";");

      const item = parent.insertListItem(
        insertionIndex++,
        itemText
      );

      item
        .setGlyphType(DocumentApp.GlyphType.NUMBER)
        .setIndentStart(54)
        .setIndentFirstLine(36)
        .setSpacingBefore(0)
        .setSpacingAfter(2)
        .setLineSpacing(1)
        .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY)
        .setStartNumber(personnelNumber++);

      item
        .editAsText()
        .setFontFamily("Arial")
        .setFontSize(11);
    });
  });
}

function groupPersonnelByMovement_(personnel) {
  const groups = [];
  const byKey = {};

  personnel.forEach(person => {
    const from = clean_(person.from).toUpperCase();
    const to = clean_(person.to).toUpperCase();
    const key = `${from}\u0000${to}`;

    if (!byKey[key]) {
      byKey[key] = {
        from,
        to,
        personnel: [],
      };
      groups.push(byKey[key]);
    }

    byKey[key].personnel.push(person);
  });

  return groups;
}

function findPlaceholderParagraph_(body, placeholder) {
  const match = body.findText(
    escapeForRegex_(placeholder)
  );

  if (!match) {
    throw new Error(
      `The template does not contain ${placeholder}.`
    );
  }

  const textElement = match.getElement().asText();
  const paragraph = textElement.getParent().asParagraph();
  const parent = paragraph.getParent();

  return {
    paragraph,
    parent,
    index: parent.getChildIndex(paragraph),
  };
}

function insertNumberedItem_(parent, index, text) {
  const item = parent.insertListItem(index, text);

  item
    .setGlyphType(DocumentApp.GlyphType.NUMBER)
    .setIndentStart(54)
    .setIndentFirstLine(36)
    .setSpacingBefore(0)
    .setSpacingAfter(3)
    .setLineSpacing(1)
    .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  item
    .editAsText()
    .setFontFamily("Arial")
    .setFontSize(11);

  return item;
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
// Manual tests
//----------------------------------

function testGenerateReassignmentListFromSheet() {
  return generateReassignmentListFromSheet();
}

function testGenerateReassignmentGroupedFromSheet() {
  return generateReassignmentGroupedFromSheet();
}
