function replaceAndFormatText() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var wordsToReplace = {
    "Goal:": "I. Goal:",
    "Material Required:": "II. Material Required:",
    "Procedure:": "III. Procedure:",
    "Output:": "IV. Output:"
    // Add more word pairs as needed
  };

  for (var i = 0; i < body.getNumChildren(); i++) {
    var element = body.getChild(i);

    if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
      var text = element.asText().getText();

      for (var word in wordsToReplace) {
        if (text.indexOf(word) !== -1) {
          var newText = text.replace(word, wordsToReplace[word]);
          element.asText().setText(newText);
          element.setHeading(DocumentApp.ParagraphHeading.HEADING3);
          element.setItalic(true);
        }
      }
    }
  }
}
