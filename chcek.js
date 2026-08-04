function testGenerateReassignmentPayload() {
  const result = generateReassignmentReport({
    orderNumber: "TEST-20260803",
    orderDate: "08/03/2026",
    signingOfficer: "TEST OFFICER",
    signingPosition: "TEST POSITION",
    fileReference: "TEST",
    personnel: [
      {
        rank: "CO1",
        fullName: "ALVIN CHIAO",
        from: "OLD OFFICE",
        to: "NEW OFFICE",
        notes: "Diagnostic test"
      }
    ]
  });

  console.log(JSON.stringify(result, null, 2));
  return result;
}