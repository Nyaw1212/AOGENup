//----------------------------------
// Reassignment API layout support
//----------------------------------

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("No request body was received.");
    }

    const request = JSON.parse(e.postData.contents);
    verifyReassignmentApiKey_(request.apiKey);

    if (String(request.action || "").toUpperCase() !== "GENERATE_REASSIGNMENT") {
      throw new Error("Unsupported request action.");
    }

    return createReassignmentJsonResponse_(
      generateReassignmentReportWithLayout_(request.payload)
    );
  } catch (error) {
    return createReassignmentJsonResponse_({
      success: false,
      message: error && error.message ? error.message : String(error),
    });
  }
}

function generateReassignmentReportWithLayout_(payload) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const normalizedPayload = normalizeReassignmentPayload_(payload);
    const personnelLayout = normalizePersonnelLayout_(
      payload && payload.personnelLayout
    );

    validateReassignmentPayload_(normalizedPayload);

    const spreadsheet = getAogenSpreadsheet_();
    const sheets = getRequiredSheets_(spreadsheet);

    writeOrderData_(sheets.orderSheet, normalizedPayload);
    writePersonnelData_(sheets.personnelSheet, normalizedPayload.personnel);

    const orderData = readOrderData_(sheets.orderSheet);
    const personnel = readPersonnel_(sheets.personnelSheet);

    validateReassignmentData_(orderData, personnel);

    const generation = getLayoutGenerationOptions_(
      personnelLayout,
      orderData,
      personnel
    );

    const result = generation.create();

    saveGeneratedDocument_(sheets.orderSheet, result.url);
    saveTemplateLink_(sheets.orderSheet, generation.templateId);

    return {
      success: true,
      message: generation.successMessage,
      documentId: result.documentId,
      url: result.url,
      outputName: result.outputName,
      personnelCount: personnel.length,
      personnelLayout,
    };
  } catch (error) {
    console.error("generateReassignmentReportWithLayout_:", error);
    return {
      success: false,
      message: getErrorMessage_(error),
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {
      // Lock was not acquired.
    }
  }
}

function normalizePersonnelLayout_(value) {
  const layout = String(value || "TABLE")
    .trim()
    .toUpperCase();

  if (layout === "NUMBERED_LIST" || layout === "LIST") {
    return "NUMBERED_LIST";
  }

  if (layout === "GROUPED_MOVEMENT" || layout === "GROUPED") {
    return "GROUPED_MOVEMENT";
  }

  return "TABLE";
}

function getLayoutGenerationOptions_(layout, orderData, personnel) {
  if (layout === "NUMBERED_LIST") {
    return {
      templateId: AOGEN_LIST_TEMPLATE_ID,
      successMessage: "Numbered Reassignment List generated successfully.",
      create: function () {
        return createListDocumentFromTemplate_({
          templateDocId: AOGEN_LIST_TEMPLATE_ID,
          outputName: `NBP Reassignment ${orderData.orderNumber} - List`,
          orderData,
          personnel,
        });
      },
    };
  }

  if (layout === "GROUPED_MOVEMENT") {
    return {
      templateId: AOGEN_GROUP_TEMPLATE_ID,
      successMessage: "Grouped Movement Reassignment generated successfully.",
      create: function () {
        return createGroupedDocumentFromTemplate_({
          templateDocId: AOGEN_GROUP_TEMPLATE_ID,
          outputName: `NBP Reassignment ${orderData.orderNumber} - Grouped`,
          orderData,
          personnel,
        });
      },
    };
  }

  return {
    templateId: AOGEN_CONFIG.reassignmentTemplateId,
    successMessage: "Reassignment Report generated successfully.",
    create: function () {
      return createDocumentFromTemplate_({
        templateDocId: AOGEN_CONFIG.reassignmentTemplateId,
        outputName: `NBP Reassignment ${orderData.orderNumber}`,
        orderData,
        personnel,
        templateType: "REASSIGNMENT",
      });
    },
  };
}
