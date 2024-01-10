function replaceAndFormatText() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var replacements = [
    { find: 'Goal:', replace: 'Goal:', format: DocumentApp.ParagraphHeading.HEADING3 },
    { find: 'Material Required:', replace: 'Material Required:', format: DocumentApp.ParagraphHeading.HEADING3 },
    { find: 'Procedure:', replace: 'Procedure:', format: DocumentApp.ParagraphHeading.HEADING3 },
    { find: 'Output:', replace: 'Output:', format: DocumentApp.ParagraphHeading.HEADING3 },
  ];

  for (var i = 0; i < body.getNumChildren(); i++) {
    var element = body.getChild(i);

    if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
      var text = element.getText();

      for (var j = 0; j < replacements.length; j++) {
        var replacement = replacements[j];

        var index = text.indexOf(replacement.find);

        if (index !== -1) {
          var beforeText = text.substring(0, index);
          var afterText = text.substring(index + replacement.find.length);

          var paragraph = body.insertParagraph(i + 1, afterText);
          element.setText(beforeText);

          paragraph.editAsText().setHeading(replacement.format);
          paragraph.editAsText().setItalic(0, afterText.length, false); // Set as normal text

          return; // Exit the loop once the replacement is made
        }
      }
    }
  }
}