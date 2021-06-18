function WorkHomeRep(sheetName = '2_Shachiku_202101', templateForm='報告書ひな形') {
  // A-1 Set spreadsheet
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(sheetName);
  var dateRange = sheet.getDataRange().getValues();

  // A-2 Set Template
  var template = spreadSheet.getSheetByName(templateForm);
  var startDate = template.getRange('M9').getValue();
  var endDate = template.getRange('M10').getValue();
  var folderurl = template.getRange('M8').getValue();

  // A-3 Record Lists
  var dutyList = [];
  var dateList = [];
  var beginTimeList = [];
  var endTimeList = [];
  var durationList = [];

  // A-4 Filter by the keyword "在宅勤務"
  for (var index=0; index<dateRange.length; index++){
    if ( Object.prototype.toString.call(dateRange[index][1]) != '[object Date]'){ continue; }
    if ( !dateRange[index][4].match(/在宅勤務/)){ continue; }
    var recDate = dateRange[index][0];
    if ( recDate < startDate){ continue; }
    if ( recDate > endDate){ continue; }
    Logger.log( Utilities.formatString("在宅勤務日 %s を抽出", Utilities.formatDate(dateRange[index][0], "GMT", "YYYY-MM-DD")));
    dutyList.push(dateRange[index][4])
    dateList.push(dateRange[index][0])
    beginTimeList.push(dateRange[index][1]);
    endTimeList.push(dateRange[index][2]);
    durationList.push(endTimeList[index] - beginTimeList[index])
  }

　// A-5 Prepare for page format (2 days per page)
  reportLength = beginTimeList.length;
  var pageLength = Math.ceil(reportLength / 2);
  var pageList = [];

  // A-6 Generage pages
  for (var reportIndex=0; reportIndex<reportLength; reportIndex++){
    var pageIndex = Math.floor(reportIndex / 2);
    var rowIndex = reportIndex - pageIndex* 2;
    var duties = dutyList[reportIndex].split(",");
    duties.splice(duties.indexOf("在宅勤務"),1);
    var dutyNum = duties.length;
    if (rowIndex == 0){
      var pageName = Utilities.formatString("Page%02d", pageIndex);
      Logger.log(Utilities.formatString("%s / %dを生成", pageName, pageLength));
      var newsheet = template.copyTo(spreadSheet);
      let delRange = newsheet.getRange("L1:M32");
      delRange.deleteCells(SpreadsheetApp.Dimension.COLUMNS)
      pageList.push(pageName)
      newsheet.setName(pageName);
      newsheet.getRange("C12").setValue(dateList[reportIndex]);
      newsheet.getRange("C13").setValue(beginTimeList[reportIndex]);
      newsheet.getRange("E13").setValue(endTimeList[reportIndex]);
      for (var dutyIndex=0; dutyIndex<dutyNum; dutyIndex++){
        var dutyString = Utilities.formatString("%d.%s", dutyIndex+1, duties[dutyIndex])
        newsheet.getRange(15+dutyIndex,3,1,1).setValue(dutyString);
      }
    } else {
      newsheet.getRange("C22").setValue(dateList[reportIndex]);
      newsheet.getRange("C23").setValue(beginTimeList[reportIndex]);
      newsheet.getRange("E23").setValue(endTimeList[reportIndex]);
      for (var dutyIndex=0; dutyIndex<dutyNum; dutyIndex++){
        var dutyString = Utilities.formatString("%d.%s", dutyIndex+1, duties[dutyIndex])
        newsheet.getRange(25+dutyIndex,3,1,1).setValue(dutyString);
      }
    }
  }
  // Remove auxiliary cells
  if (rowIndex == 0){
    newsheet.getRange("G23").setValue('');
    newsheet.getRange("I23").setValue('');
  }

  // P-1 Prepare for PDF generation
  var myArray= folderurl.split('/');
  var folderid = myArray[myArray.length-1];
  var folder = DriveApp.getFolderById(folderid);

  var sheet = spreadSheet.getActiveSheet();
  var sheetID = sheet.getSheetId();
  var key = spreadSheet.getId();
  var token = ScriptApp.getOAuthToken();

  // P-2 Hide original and template not to print in PDF
  spreadSheet.getSheetByName(templateForm).hideSheet();
  spreadSheet.getSheetByName(sheetName).hideSheet();

  // P-3 PDF properties
  var url = 'https://docs.google.com/spreadsheets/d/'+ key +'/export?';
  var opts = {
    exportFormat:   'pdf',
    format:         'pdf',
    size:           'A4',
    portrate:       'true',
    fitw:           'true',
    sheetnames:     'false',
    printtitle:     'false',
    pagenumbers:    'false',
    gridlines:      'false',
    fzr:            'false',
    range:          'A1%3AJ32'
  };

  // P-4 Print all pages into a single PDF
  var PDFurl = [];
  for( optName in opts){
    PDFurl.push(optName + '=' + opts[optName]);
  }
  var options = PDFurl.join('&');
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch( url + options, {headers: {'Authorization': 'Bearer ' +  token}});
  var PDFname = Utilities.formatString("WorkHomeReport.%s.pdf", Utilities.formatDate(startDate, "GMT", "YYYY-MM-DD"));
  var blob = response.getBlob().setName(PDFname);
  var newFile = folder.createFile(blob);
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // P-5 Revert original and template
  spreadSheet.getSheetByName(templateForm).showSheet();
  spreadSheet.getSheetByName(sheetName).showSheet();

  // P-6 Remove temporary sheets
  for (var index=0; index<pageList.length; index++){
    let newsheet = spreadSheet.getSheetByName(pageList[index]);
    spreadSheet.deleteSheet(newsheet);
  }
  
}

// C-1 Function to cleanup temporary sheets
function cleanUpSheet(startSheet=0, endSheet=100){
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  for (var index=startSheet; index<endSheet; index++){
    var pageName = Utilities.formatString("Page%02d", index);
    let newsheet = spreadSheet.getSheetByName(pageName);
    spreadSheet.deleteSheet(newsheet);
  }
}

