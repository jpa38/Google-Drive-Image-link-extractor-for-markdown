/*
---
Title:    Image link extractor
Author:   Jérome PALANCA
Date:     Octobre 2019
Comment:  --
---
####################################################################
Functiun to generate a compatible link with MkDown about image in google drive 's folder.
####################################################################

Clic to **Img_MkDown** -> Run at the top-right

How to use :
1. Create Sheet in the images's folder
1. Load Script in the sheet
2. Reload Sheet
3. Clic to "**Img_MkDown**" -> "Run" in the menu item

Note : Folder / image need to be public. (https://support.google.com/drive/answer/2494822)

Version:
V0.0.1 = First publication

Link to test your links: https://markdownlivepreview.com/
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('**Img_MkDown**')
      .addItem('Run', 'go')
      .addToUi();
}

// Lauching when install
function onInstall() {
  onOpen();
}

// Lauching when click to run
function go() {
  list_img_mkdown();
}

// #### PRNCIPAL FUNCTIUN ####
function list_img_mkdown(){
  
  // Clean Sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  
  // Get sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId =  SpreadsheetApp.getActiveSpreadsheet().getId();
  var spreadsheetFile =  DriveApp.getFileById(spreadsheetId);
  var ui = SpreadsheetApp.getUi();
    
  // Get folder
  var folder = spreadsheetFile.getParents().next()
  var folderId = spreadsheetFile.getParents().next().getId();
  var folderName = spreadsheetFile.getParents().next().getName();
  
  // Debug
  Logger.log("Sheet_ID = " + spreadsheetId)
  Logger.log("Folder_ID = " + folderId)
  Logger.log("Folder_Name = " + folderName)
  
  // Sheet's header
  var date = new Date()
  // First Header
  header = [
    "Last Extraction: ", Date(date.setDate(date.getDate()+5))
    ]
  
  // Second Header
  header_columns = [
      "Name", 
    "Markdown - StandAlone Link",
    "Markdown - Referent Link (To Copy on time at top)",
    "Markdown - Link (need referent Link)",
    "Meta - Size", 
    "Meta - Type", 
    "Type - ID",
    "Preview Google Drive",
    'Permissions - Who ? (need "Anyone_with_link")',
    'Permissions - What ? (recommend : "VIEW")',
    ]
    
  spreadsheet.appendRow(header)
  spreadsheet.appendRow(header_columns)
    
  //Color headers
  var changeRange = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  changeRange.setBackgroundRGB(206, 221, 226);

  // Freeze First Row
  spreadsheet.setFrozenRows(2)
    
  //Resize columns
  i = spreadsheet.getLastColumn()
  while(i>0){
    spreadsheet.autoResizeColumn(i)
    i = i-1
  }
    
  //spreadsheet.appendRow(["Name", "Markdown - StandAlone Link", "Markdown - Referent Link (Copy on top)","Markdown - Link (need referent Link)", "Meta - Size", "Meta", "Type"]);
  
  data = []
  i = 0
  // Get files
  var files = DriveApp.getFolderById(folderId).getFiles()
  
  // Loop for files
  while(files.hasNext()) {
    i=i+1
    
    var f = files.next()
    
    // Get first parameters
    fId = f.getId()
    
    fType = DriveApp.getFileById(fId).getMimeType()
    Logger.log(fType)
    // If isn't image file, go to next file
    if(fType.indexOf("image")==-1){
      continue;
        }
        

    // ******************  Get data from files
    fFullName = f.getName()
    fName = fFullName.split(".")[0]
    fExtension = fFullName.split(".")[1]
    
    fSize = f.getSize()
    fThumb = f.getThumbnail()
    
    fUrl= f.getUrl()
    
    sharingAccess = f.getSharingAccess()
    sharingPermission = f.getSharingPermission()

    Logger.log(fName)
    
    // ****************** Combine data
    fImageUrl = 'img src="https://drive.google.com/uc?id=' + fId + '"'
    
    // First one
    fLinkImageView = 'https://drive.google.com/uc?id=' + fId
    
    // Html
    //shareUrlHtml = '<img src="' + 'fLinkImageView' + '">'
    
    // Markdown standalone
    shareUrlMkDown = '!['+ fName + '](https://drive.google.com/uc?id=' + fId + ')'
    // Markdown reference
    mkRef = '[' + fName + ']: ' + fLinkImageView + ' ' + '"' + fName + '"'
    mkAccesWithRef = '![Alt text]' + '[' + fName + ']'
    
    //  ****************** Insert data
    // Array to insert
    data = [
      fName,
      
      shareUrlMkDown,
      
      mkRef,
      mkAccesWithRef,
      
      fSize,
      
      fType,
      fId,
      fUrl,
      sharingAccess,
      sharingPermission
      ];
    
    Logger.log(data)
    
    // Insert array to sheet
    spreadsheet.appendRow(data);

  }
  
  //Resize first columns
  spreadsheet.autoResizeColumn(1)
  
  // End message
  ui.alert("Extract Done ! \n \n Don't forget to share your folder \n => Public Acces / Viewer recommended \n (https://support.google.com/drive/answer/2494822) \n (https://support.google.com/drive/answer/7166529?co=GENIE.Platform%3DDesktop&hl=en)")
  
  
}

