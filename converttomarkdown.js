/*
Usage: 
Adding this script to your doc: 
- Tools > Script Manager > New
- Select "Blank Project", then paste this code in and save.
Running the script:
- Add-ons > ConverToMarkDown
- Select "Add To Drive" function.
- The document will be converted into a MD file and it will be placed in the folder called "Common"
- You can also view or donwload the converted file from the pop up, once the document is converted.
*/

/**
* Creates a menu entry in the Google Docs UI when the document is opened.
*
* @param {object} e The event parameter for a simple onOpen trigger. To
*     determine which authorization mode (ScriptApp.AuthMode) the trigger is
*     running in, inspect e.authMode.
*/
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
  .addItem('Add To Drive', 'ConvertToMarkdown')
  .addItem('Download MD file', 'downloadMdFile')
  .addToUi();
}

/**
* Runs when the add-on is installed.
*
* @param {object} e The event parameter for a simple onInstall trigger. To
*     determine which authorization mode (ScriptApp.AuthMode) the trigger is
*     running in, inspect e.authMode. (In practice, onInstall triggers always
*     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
*     AuthMode.NONE.)
*/
function onInstall(e) {
  //When the document is first installed, clear out the user preferences set already
  var scriptProperties = PropertiesService.getUserProperties();
  
  //Change this value to change the default output folder. This value can be overridden by user at any time using 'Change Output Folder' menu
  var defaultOutputFolder = "Google Docs to Markdown";
  
  //Setting the intial value of common folder if not present
  //if (typeof scriptProperties.getProperty('DEFAULT_OUTPUT_FOLDER') == 'undefined' || scriptProperties.getProperty('DEFAULT_OUTPUT_FOLDER') == null || scriptProperties.getProperty('DEFAULT_OUTPUT_FOLDER') === "")
  scriptProperties.setProperty('DEFAULT_OUTPUT_FOLDER', defaultOutputFolder); // Changing this value affects the default common folder name
  
  //Setting the intial value of common folder if not present
  //if (typeof scriptProperties.getProperty('COMMON_FOLDER') == 'undefined' || scriptProperties.getProperty('COMMON_FOLDER') == null || scriptProperties.getProperty('COMMON_FOLDER') == "")
  scriptProperties.setProperty('COMMON_FOLDER', defaultOutputFolder); // Changing this value affects the common folder name
  
  onOpen(e);
}

function doGet() {
  ConvertToMarkdown();
}

/* This function converts the google document into mark down and saves it in the common folder as configured by the user.
It will also save any inline images in the same folder where converted MD file is placed */
function ConvertToMarkdown() {
  var text = "";
  var inSrc = false;
  var inClass = false;
  var globalImageCounter = 0;
  var globalListCounters = {};
  // edbacher: added a variable for indent in src <pre> block. Let style sheet do margin.
  var srcIndent = "";
  
  var attachments = [];
  var file;
  var commonFolder;
  var folder;
  var folderName;
  var blob;
  var content_type;
  var suffix;
  var image_name;
  var photo;
  
  try {
    var numChildren = DocumentApp.getActiveDocument().getActiveSection().getNumChildren();
    var folder = getCurrentFolder();
    
    // Walk through all the child elements of the doc.
    for (var i = 0; i < numChildren; i++) {
      var child = DocumentApp.getActiveDocument().getActiveSection().getChild(i);
      var result = processParagraph(i, child, inSrc, globalImageCounter, globalListCounters);
      globalImageCounter += (result && result.images) ? result.images.length : 0;
      if (result !== null) {
        if (result.sourcePretty === "start" && !inSrc) {
          inSrc = true;
          text += "<pre class=\"prettyprint\">\n";
        } else if (result.sourcePretty === "end" && inSrc) {
          inSrc = false;
          text += "</pre>\n\n";
        } else if (result.source === "start" && !inSrc) {
          inSrc = true;
          text += "<pre>\n";
        } else if (result.source === "end" && inSrc) {
          inSrc = false;
          text += "</pre>\n\n";
        } else if (result.inClass === "start" && !inClass) {
          inClass = true;
          text += "<div class=\"" + result.className + "\">\n";
        } else if (result.inClass === "end" && inClass) {
          inClass = false;
          text += "</div>\n\n";
        } else if (inClass) {
          text += result.text + "\n\n";
        } else if (inSrc) {
          text += (srcIndent + escapeHTML(result.text) + "\n");
        } else if (result.text && result.text.length > 0) {
          text += result.text + "\n\n";
        }
        
        if (result.images && result.images.length > 0) {
          for (var j = 0; j < result.images.length; j++) {
            attachments.push({
              "fileName": result.images[j].name,
              "mimeType": result.images[j].type,
              "content": result.images[j].bytes
            });
          }
        }
      } else if (inSrc) { // support empty lines inside source code
        text += '\n';
      }
      
    }
    
    //Checking if MD file with same name is already present inside that folder. If present it will override the older file. If not it will create the new file.
    var filename = DocumentApp.getActiveDocument().getName() + '.md';
    removeExistingFiles(folder, filename);
    file = DriveApp.createFile(filename, text, 'text/plain');
    folder.addFile(file)
    
    //If there are any attachments in the file, it has to be saved in the same directory.
    //Due to this issue [http://code.google.com/p/google-apps-script-issues/issues/detail?id=1239], image files are created using blob. and replaced using Drive API.
    if (attachments.length > 0) {
      for (var iterator = 0; iterator < attachments.length; iterator++) {
        blob = attachments[iterator].content;
        content_type = blob.getContentType()
        suffix = content_type.split("/")[1] //e.g gif/jpg or png
        image_name = "test.png" //Invent a name, the blob seems to need it?
        blob.setName(image_name);
        try {
          image_name = "image_" + iterator + "." + suffix;
          photo = checkIfFileExists(folder, image_name);
          if (photo) {
            photo.setTrashed(true);
            photo = folder.createFile(blob);
            //try {
            //Drive.Files.remove(photo.getId());
            //}
            //catch (e) {
            //DocsList.getFileById(photo.getId()).setTrashed(true);
            //}
          } else {
            photo = folder.createFile(blob);
          }
          photo.rename(image_name);
        } catch (e) {
          throw ("Error in saving attached images : " + e);
        }
        
      }
    }
  } catch (e) {
    var errorMsg = "";
    //While displaying error message, we display the last converted text, so that the users can know, after which line the conversion failed.
    //Check if there is any last converted text. If so take the last sentence from the converted text. If not, just display the error message.
    if (text != null && text.length != 0 && text.trim() !== "") {
      var sentence = text.split(".");
      if (sentence.length > 1) {
        errorMsg = "Error after the line : \"" + sentence[sentence.length - 2] + "\".\n\n" + e;
      } else if (sentence.length == 1) {
        errorMsg = "Error after the line : \"" + sentence[sentence.length - 1] + "\".\n\n" + e;
      } else if (sentence.length == 0) {
        errorMsg = "Error after the text : \"" + text + "\".\n\n" + e;;
      }
    } else {
      errorMsg = e;
    }
    //Showing the error message in alert window.
    DocumentApp.getUi().alert("Error", errorMsg, DocumentApp.getUi().ButtonSet.OK);
  }
  return file;
}


/* This function converts the current open document into an MD file and saves it in the common folder. It will then show you the pop up to download the MD file */
function downloadMdFile() {
  var file = ConvertToMarkdown();
  
  DocumentApp.getUi().showDialog(
    HtmlService
    .createHtmlOutput('<a href="https://docs.google.com/a/fusioncharts.com/uc?export=download&id=' + file.getId() + '">Download the converted MD file</a>')
  .setTitle('Download Link for MD file')
  .setWidth(400 /* pixels */ )
  .setHeight(100 /* pixels */ ));
}

/* This function returns the current folder. */
function getCurrentFolder() {
  // get the active document
  var doc = DocumentApp.getActiveDocument();
  
  // get the parent folders, could be more than one
  var directParents = DriveApp.getFileById(doc.getId()).getParents();
  
  var folderCount = 0;
  while(directParents.hasNext()) {
    folderCount++;
    var currentFolder = directParents.next();
  }
  if (folderCount == 1) {
    return currentFolder;
  }
  
  Logger.log('Found more than one parent folder.');
  return null;
}

/* This function removes previous verions of the file */
function removeExistingFiles(folder, filename) {
  Logger.log(folder.getName());
  Logger.log(filename)
  var files = folder.getFilesByName(filename);
  while (files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);
  }
}

function escapeHTML(text) {
  return text.replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Process each child element (not just paragraphs).
function processParagraph(index, element, inSrc, imageCounter, listCounters) {
  // First, check for things that require no processing.
  if (element.getNumChildren() == 0) {
    return null;
  }
  // Punt on TOC.
  if (element.getType() === DocumentApp.ElementType.TABLE_OF_CONTENTS) {
    return {
      "text": "[[TOC]]"
    };
  }
  
  // Set up for real results.
  var result = {};
  var pOut = "";
  var textElements = [];
  var imagePrefix = "image_";
  
  // Handle Table elements. Pretty simple-minded now, but works for simple tables.
  // Note that Markdown does not process within block-level HTML, so it probably 
  // doesn't make sense to add markup within tables.
  if (element.getType() === DocumentApp.ElementType.TABLE) {
    var nCols = element.getChild(0).getNumCells();
    for (var i = 0; i < element.getNumChildren(); i++) {
      textElements.push("| ");
      // process this row
      for (var j = 0; j < nCols; j++) {
        textElements.push(" " + element.getChild(i).getChild(j).getText() + " |");
      }
      textElements.push("\n");
      // process the delimiter row
      textElements.push("| ");
      // process this row
      for (var j = 0; j < nCols; j++) {
        textElements.push(":---|");
      }
      textElements.push("\n");      
    }
  }
  
  // Process various types (ElementType).
  for (var i = 0; i < element.getNumChildren(); i++) {
    var t = element.getChild(i).getType();
    
    if (t === DocumentApp.ElementType.TABLE_ROW) {
      // do nothing: already handled TABLE_ROW
    } else if (t === DocumentApp.ElementType.TEXT) {
      var txt = element.getChild(i);
      pOut += txt.getText();
      textElements.push(txt);
    } else if (t === DocumentApp.ElementType.INLINE_IMAGE) {
      result.images = result.images || [];
      var contentType = element.getChild(i).getBlob().getContentType();
      var extension = "";
      if (/\/png$/.test(contentType)) {
        extension = ".png";
      } else if (/\/gif$/.test(contentType)) {
        extension = ".gif";
      } else if (/\/jpe?g$/.test(contentType)) {
        extension = ".jpg";
      } else {
        throw "Unsupported image type: " + contentType;
      }
      var name = imagePrefix + imageCounter + extension;
      imageCounter++;
      textElements.push('![image alt text](' + name + ')');
      result.images.push({
        "bytes": element.getChild(i).getBlob(),
        "type": contentType,
        "name": name
      });
    } else if (t === DocumentApp.ElementType.PAGE_BREAK) {
      // ignore
    } else if (t === DocumentApp.ElementType.HORIZONTAL_RULE) {
      textElements.push('* * *\n');
    } else if (t === DocumentApp.ElementType.FOOTNOTE) {
      textElements.push(' (NOTE: ' + element.getChild(i).getFootnoteContents().getText() + ')');
    } else {
      //throw "Paragraph "+index+" of type "+element.getType()+" has an unsupported child: "
      //+t+" "+(element.getChild(i)["getText"] ? element.getChild(i).getText():'')+" index="+result;
      throw "Unsupported format in current file :" + t + " " + (element.getChild(i)["getText"] ? element.getChild(i).getText() : '') + ". Cannot be converted into an MD file";
    }
  }
  
  if (textElements.length == 0) {
    // Isn't result empty now?
    return result;
  }
  
  // evb: Add source pretty too. (And abbreviations: src and srcp.)
  // process source code block:
  if (/^\s*---\s+srcp\s*$/.test(pOut) || /^\s*---\s+source pretty\s*$/.test(pOut)) {
    result.sourcePretty = "start";
  } else if (/^\s*---\s+src\s*$/.test(pOut) || /^\s*---\s+source code\s*$/.test(pOut)) {
    result.source = "start";
  } else if (/^\s*---\s+class\s+([^ ]+)\s*$/.test(pOut)) {
    result.inClass = "start";
    result.className = RegExp.$1;
  } else if (/^\s*---\s*$/.test(pOut)) {
    result.source = "end";
    result.sourcePretty = "end";
    result.inClass = "end";
  } else if (/^\s*---\s+jsperf\s*([^ ]+)\s*$/.test(pOut)) {
    result.text = '<iframe style="width: 100%; height: 340px; overflow: hidden; border: 0;" ' +
      'src="http://www.html5rocks.com/static/jsperfview/embed.html?id=' + RegExp.$1 +
        '"></iframe>';
  } else {
    
    prefix = findPrefix(inSrc, element, listCounters);
    
    var pOut = "";
    for (var i = 0; i < textElements.length; i++) {
      pOut += processTextElement(inSrc, textElements[i]);
    }
    
    // replace Unicode quotation marks
    pOut = pOut.replace('\u201d', '"').replace('\u201c', '"');
    
    result.text = prefix + pOut;
  }
  
  return result;
}

// Add correct prefix to list items.
function findPrefix(inSrc, element, listCounters) {
  var prefix = "";
  if (!inSrc) {
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var paragraphObj = element;
      switch (paragraphObj.getHeading()) {
          // Add a # for each heading level. No break, so we accumulate the right number.
        case DocumentApp.ParagraphHeading.HEADING6:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING5:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING4:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING3:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING2:
          prefix += "#";
        case DocumentApp.ParagraphHeading.HEADING1:
          prefix += "# ";
        default:
      }
    } else if (element.getType() === DocumentApp.ElementType.LIST_ITEM) {
      var listItem = element;
      var nesting = listItem.getNestingLevel()
      for (var i = 0; i < nesting; i++) {
        prefix += "    ";
      }
      var gt = listItem.getGlyphType();
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET || gt === DocumentApp.GlyphType.HOLLOW_BULLET || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        prefix += "* ";
      } else {
        // Ordered list (<ol>):
        var key = listItem.getListId() + '.' + listItem.getNestingLevel();
        var counter = listCounters[key] || 0;
        counter++;
        listCounters[key] = counter;
        prefix += counter + ". ";
      }
    }
  }
  return prefix;
}

function processTextElement(inSrc, txt) {
  if (typeof(txt) === 'string') {
    return txt;
  }
  
  var pOut = txt.getText();
  if (!txt.getTextAttributeIndices) {
    return pOut;
  }
  
  var attrs = txt.getTextAttributeIndices();
  var lastOff = pOut.length;
  
  for (var i = attrs.length - 1; i >= 0; i--) {
    var off = attrs[i];
    var url = txt.getLinkUrl(off);
    var font = txt.getFontFamily(off);
    if (url) { // start of link
      if (i >= 1 && attrs[i - 1] == off - 1 && txt.getLinkUrl(attrs[i - 1]) === url) {
        // detect links that are in multiple pieces because of errors on formatting:
        i -= 1;
        off = attrs[i];
        url = txt.getLinkUrl(off);
      }
      pOut = pOut.substring(0, off) + '[' + pOut.substring(off, lastOff) + '](' + url + ')' + pOut.substring(lastOff);
    } else if (font) {
      if (!inSrc && font === font.COURIER_NEW) {
        while (i >= 1 && txt.getFontFamily(attrs[i - 1]) && txt.getFontFamily(attrs[i - 1]) === font.COURIER_NEW) {
          // detect fonts that are in multiple pieces because of errors on formatting:
          i -= 1;
          off = attrs[i];
        }
        pOut = pOut.substring(0, off) + '`' + pOut.substring(off, lastOff) + '`' + pOut.substring(lastOff);
      }
    }
    if (txt.isBold(off)) {
      var d1 = d2 = "**";
      if (txt.isItalic(off)) {
        // edbacher: changed this to handle bold italic properly.
        d1 = "**_";
        d2 = "_**";
      }
      pOut = pOut.substring(0, off) + d1 + pOut.substring(off, lastOff) + d2 + pOut.substring(lastOff);
    } else if (txt.isItalic(off)) {
      pOut = pOut.substring(0, off) + '*' + pOut.substring(off, lastOff) + '*' + pOut.substring(lastOff);
    }
    lastOff = off;
  }
  return pOut;
}
