//----------------------------------
// Grouped Movement Continuous Numbering Fix
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

      const paragraph = parent.insertParagraph(
        insertionIndex++,
        `${personnelNumber}. ${rankAndName}${notesText}` +
          (isLastOverall ? "." : ";")
      );

      paragraph
        .setIndentStart(54)
        .setIndentFirstLine(36)
        .setSpacingBefore(0)
        .setSpacingAfter(2)
        .setLineSpacing(1)
        .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

      paragraph
        .editAsText()
        .setFontFamily("Arial")
        .setFontSize(11);

      personnelNumber++;
    });
  });
}
