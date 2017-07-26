/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      $('#run').click(run);
    });
  };

  function run() {
    
    return OneNote.run(function (context) {

            // Get the collection of pageContent items from the page.
            var pageContents = context.application.getActivePage().contents;

            // Get the first PageContent on the page
            // Assuming its an outline, get the outline's paragraphs.
            var pageContent = pageContents.getItemAt(0);
            var paragraphs = pageContent.outline.paragraphs;
            var firstParagraph = paragraphs.getItemAt(0);

            // Queue a command to load the id and type of the first paragraph
            firstParagraph.load("id,type");

          // do stuff to the text everytime it changes
            return context.sync()
                .then(function () {

                    // Queue commands to insert before and after the first paragraph
                    firstParagraph.insertRichTextAsSibling("Before", "Text Appears Before Paragraph");
                    firstParagraph.insertRichTextAsSibling("After", "Text Appears After Paragraph");
                    
                    // Run the command to insert text contents
                    return context.sync();
                });

            // Run the queued commands, and return a promise to indicate task completion.
          
        })  
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }); 
  }

})();