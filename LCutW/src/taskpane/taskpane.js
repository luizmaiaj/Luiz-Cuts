/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    /*document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("apply-style").onclick = applyStyle;
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    document.getElementById("replace-text").onclick = replaceText;*/
    document.getElementById("highlight-text").onclick = highlightText;
    document.getElementById("format-wit").onclick = formatWIT;
    document.getElementById("format-tickets").onclick = formatTickets;
    document.getElementById("format-numbers").onclick = formatNumbers;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function insertParagraph() {
  Word.run(function (context) {

    var docBody = context.document.body;
    
    docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.", "Start");
    
    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function applyStyle() {
  Word.run(function (context) {

    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;

    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function applyCustomStyle() {
  Word.run(function (context) {

    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";

    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function changeFont() {
  Word.run(function (context) {

    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });

    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertTextIntoRange() {
  Word.run(function (context) {

    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");

    originalRange.load("text");
    return context.sync()
        .then(function() {
          doc.body.insertParagraph("Original range: " + originalRange.text, "End");
        })
        .then(context.sync);
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function insertTextBeforeRange() {
  Word.run(function (context) {

    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");

    originalRange.load("text");
    return context.sync()
       .then(function() {
        doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
       })
       .then(context.sync);

  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function replaceText() {
  Word.run(function (context) {

    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function highlightText() {
  Word.run(function (context) {

    var doc = context.document;
    var originalRange = doc.getSelection();

    originalRange.font.highlightColor = "Yellow";
    
    //originalRange.insertText(Clipboard.text, "End");

    return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function formatWIT() {
  Word.run(function (context) {

    var doc = context.document;
    var originalRange = doc.getSelection();

    originalRange.load("text,hyperlink");

    return context.sync()
       .then(function() {
        let ticket = originalRange.text

        ticket = ticket.replace(/\r?\n|\r/g, ""); // remove any line breaks

        let link = originalRange.hyperlink

        let workItemType = ticket.slice(0, ticket.indexOf(" ")) // get WIT

        let ticketNumber = ticket.slice(workItemType.length + 1, ticket.indexOf(":")) // get ticket number

        let title = ticket.slice(ticket.indexOf(":") + 1, ticket.length) // get title

        originalRange.insertText(workItemType + ":" + title + " ", "Replace");
        originalRange.font.set({
          name: "Montserrat",
          bold: true,
          size: 9
        });

        // TODO: PROJECT
        let linkRange = originalRange.insertText("[PROJECT-" + ticketNumber + "]", "After");
        linkRange.font.set({
          name: "Montserrat",
          bold: false,
          size: 9
        });
        linkRange.hyperlink = link;

        linkRange.insertText(": ", "After");

        console.log("Debug info: " + workItemType + ":" + title + link);

       })
       .then(context.sync);
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function formatNumbers() {
  // TODO: PROJECT
  searchAndFormat('[0-9]{5,6}', '[PROJECT-', ']')
}

function formatTickets() {
  // TODO: PROJECT
  searchAndFormat('[[]PROJECT-[0-9]{5,6}[]]', '', '')
}

function searchAndFormat(regularExpression, prefix, postfix) {
  // Run a batch operation against the Word object model.
  Word.run(function (context) {

    // Queue a command to search the document with a wildcard [[]PROJECT-[0-9]{5,6}[]]

    // search and replace only within selection
    var searchResults = context.document.getSelection().search(regularExpression, {matchWildcards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, "text,font,hyperlink");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
          console.log('Found count: ' + searchResults.items.length);

          // Queue a set of commands to change the font for each found item.
          for (var i = 0; i < searchResults.items.length; i++) {
              searchResults.items[i].font.set({
            name: "Montserrat",
            bold: false,
            size: 9,
            highlightColor: null,
            color: "black"
          });

          ticketNumber = searchResults.items[i].text.match(/\d+/)

          // set the link of the search result by using the id found in the text
          if(prefix.length > 0 && postfix.length > 0) {
            searchResults.items[i].insertText(prefix + searchResults.items[i].text + postfix, "Replace");
          }

          // TODO: PROJECT / ORGANISATION
          searchResults.items[i].hyperlink = "https://dev.azure.com/ORGANISATION/PROJECT/_workitems/edit/" + ticketNumber;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
  })
  .catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
  });
}