//----------------------------------
// Reassignment validation override
// From and To are optional for report generation.
//----------------------------------

function validateReassignmentPersonnel_(personnel) {
  if (!Array.isArray(personnel) || !personnel.length) {
    throw new Error(
      "Add at least one personnel entry in the Personnel tab."
    );
  }

  personnel.forEach((person, index) => {
    const missing = [];

    if (!person.rank) {
      missing.push("Rank");
    }

    if (!person.fullName) {
      missing.push("Full Name");
    }

    if (missing.length) {
      throw new Error(
        `Personnel row ${index + 2} is missing: ` +
        missing.join(", ")
      );
    }
  });
}

validateReassignmentData_ = function (orderData, personnel) {
  const missing = [];

  if (!orderData.orderNumber) missing.push("Order Number");
  if (!orderData.orderDate) missing.push("Order Date");
  if (!orderData.signingOfficer) missing.push("Signing Officer");
  if (!orderData.signingPosition) missing.push("Signing Position");

  throwIfMissingFields_(
    missing,
    "Complete the following fields in the Order tab:"
  );

  validateReassignmentPersonnel_(personnel);
};

validateReassignmentPayload_ = function (payload) {
  const missing = [];

  if (!payload.orderNumber) missing.push("Order Number");
  if (!payload.orderDate) missing.push("Order Date");
  if (!payload.signingOfficer) missing.push("Signing Officer");
  if (!payload.signingPosition) missing.push("Signing Position");

  throwIfMissingFields_(
    missing,
    "Missing reassignment fields:"
  );

  validateReassignmentPersonnel_(payload.personnel);
};
