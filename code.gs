function onInstall(e) {
  onOpen(e);
}

function onOpen(e){

  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Draw Road Map")
    .addItem("Draw Task bars & Milestones", "drawRoadMap")
    .addToUi();
}

// TO DO: Set up an input system to take sourceSSName, sourceSheetName as input
// ASK: Those who generate the sheet:- Can I assume the sheet program is going to use to extract and process, will always have the same format as "YCS Only" of "TCS/YCS"?
// ASK: MATT:- Should I allow the user to Input the range for reading the data required?
function drawRoadMap() {
  var ui = SpreadsheetApp.getUi();
  
  var source = ui.prompt("Name the Sheet to use for the Road Map", ui.ButtonSet.OK);

  if (source.getSelectedButton() == ui.Button.OK) {
    var sourceSheetName = source.getResponseText();
    var destSheetName = sourceSheetName + " in Roadmap Format";
    formatData(sourceSheetName, destSheetName);

    var template = ui.prompt("Name the Template slide to be used in Road Map", ui.ButtonSet.OK);
    if (template.getSelectedButton() == ui.Button.OK) { 
      var templateName = template.getResponseText();
      
      var monthlyBox = ui.prompt("For you Template, what is the length of taskbar for a month", ui.ButtonSet.OK);
      if (monthlyBox.getSelectedButton() == ui.Button.OK) {
        var monthlyBoxLength = parseFloat(monthlyBox.getResponseText());

        var startX = ui.prompt("Enter the X-Coordinate you want to start the Road map", ui.ButtonSet.OK);
        if (startX.getSelectedButton() == ui.Button.OK) { 
          var startPosX = parseFloat(startX.getResponseText());

          var startY = ui.prompt("Enter the Y-Coordinate you want to start the Road map", ui.ButtonSet.OK);
          if (startY.getSelectedButton() == ui.Button.OK) {  
            var startPosY = parseFloat(startY.getResponseText());
            
            if (startY.getSelectedButton() == ui.Button.OK) { generateRoadMap(destSheetName, templateName, monthlyBoxLength, startPosX, startPosY); }
            else { ui.alert("Error!! Can't Proceed!! Re-run and make sure you input values correctly and press OK on every input"); }
          }
          else { ui.alert("Error!! Can't Proceed!! Re-run and make sure you input values correctly and press OK on every input"); }
        }
        else { ui.alert("Error!! Can't Proceed!! Re-run and make sure you input values correctly and press OK on every input"); }
      }
      else { ui.alert("Error!! Can't Proceed!! Re-run and make sure you input values correctly and press OK on every input"); }
    }
    else {ui.alert("Error!! Can't Proceed!! Re-run and make sure you input values correctly and press OK on every input");} 
  }
  else {  ui.alert("Error!! Can't Proceed!! Re-run and make sure you input values correctly and press OK on every input"); }
}


function formatData(sourceSheetName, destSheetName) {

  insertNewSheet (destSheetName);

  // ASSUMPTION: Assuming that the values are in between coulumn B-F respectively, and row 1 is empty. Aslo, data start from row 2
  // Importing the Activity ID, Roadmap Task Name, BCS Activity Name, Start, Finish
  var sourceSheetRange = "B1:E";
  var startRow = 1;
  var startCol = 2;
  var numberOfCol = 4;
  importRange(
    sourceSheetName, //Source Sheet - e.g."Task List"
    sourceSheetRange, // Source Range - e.g. "A2:G"
    startRow,
    startCol,
    numberOfCol,
    destSheetName, //Dest Sheet - e.g. sourceSheetName + " in Roadmap Format"
    "A1", // Dest Range Start - e.g. "B3"
  );
  
  formatTheSheetValue(destSheetName);
  formatTheSheetStyle(destSheetName);
}



function insertNewSheet (newSheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(newSheetName);
  
  if (sheet != null) {
    spreadsheet.deleteSheet(sheet);
  }
  
  var newSheet = spreadsheet.insertSheet();
  newSheet.setName(newSheetName);
}


function importRange(sourceSheetName, sourceSheetRange, startRow, startCol, numberOfCol, destSheetName, destRangeStart){
 
  // Gather the source range values
  var sourceSS = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = sourceSS.getSheetByName(sourceSheetName);
  var sourceRng = sourceSheet.getRange(sourceSheetRange)
  var sourceVals = sourceRng.getValues();
 
  // Get the destiation sheet and cell location.
  var destSS = sourceSS;
  var destSheet = destSS.getSheetByName(destSheetName);
  var destStartRange = destSheet.getRange(destRangeStart);
 
  // Get the full data range to paste from start range.
  var destRange = destSheet.getRange(destStartRange.getRow(), destStartRange.getColumn(), sourceVals.length, sourceVals[0].length);
  
  // Paste in the values.
  destRange.setValues(sourceVals);

  // Update the new sheets' Background colour in respective to the source sheet
  for(var i = 0; i < sourceSheet.getLastRow(); i++) {
    var sRng = sourceSheet.getRange(startRow+i, startCol, 1, numberOfCol);
    var dRng = destSheet.getRange(1+i, 1, 1, numberOfCol);
    dRng.setBackground(sRng.getBackground());
  }

  debugger;
};


function formatTheSheetStyle(destSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(destSheetName);
  var rowRange = sheet.getRange(1,1, 1, sheet.getLastColumn())

  // Bolds and changes font size to highlight the header of the sheet
  rowRange.setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");

  sheet.autoResizeColumns(1, sheet.getLastColumn());
  sheet.setRowHeights(1, sheet.getLastRow(), 40)
}


function formatTheSheetValue(destSheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(destSheetName);

  // Extract the range for row holding "Start" and "Finish" 
  var range = sheet.getRange(2, 2, sheet.getLastRow()-1, 2);

  // Extract the values in the range "range"
  var valueRange = range.getValues();
  
  // Iterate through the rows for the number of records in the sheet
  for(var i = 0; i < sheet.getLastRow()-1; i++) {
    for (var j = 0; j<valueRange[i].length; j++) {
      var valueType = typeof valueRange[i][j];
      var isDate = valueRange[i][j] instanceof Date;

      // Expecting only two different type of values, either "16-Jul-21 A" format or ""(null)
      if (valueType == "string" || !isDate ) {

        // Detects the milestones [P6's simulation], as they have either "Start" or "Finish", but not both 
        // Replicates the same value the empty cell of "Start" or "Finish"
        // Handles ""(null)
        if (valueRange[i][j] === "" || valueRange[i][j] === null || valueRange[i][j] === undefined) {
          if (j+1 == valueRange[i].length) {
            valueRange[i][j] = valueRange[i][j-1];
          }
          else {
            valueRange[i][j] = valueRange[i][j+1];
          }
        }

        // Handles "16-Jul-21 A" format
        else if (valueRange[i][j].split(" ").length == 2){
          var strArr = valueRange[i][j].split(" ");
          var dateStr = strArr[0];
          var dateStrSplit = dateStr.split("-");

          var now = new Date();
          var day = parseInt(dateStrSplit[0]);
          var year = parseInt(dateStrSplit[2]) + parseInt(now.getFullYear()/100)*100;
          var dateExtracted = new Date (dateStrSplit[1] + '/' + day + '/' + year);

          valueRange[i][j] = dateExtracted;
        }

        else {
          valueRange[i][j] = new Date(valueRange[i][j]);
        }
      }
    }
  }

  range.setValues(valueRange);
  
  debugger;
}


function generateRoadMap (sourceSheet, templateName, monthlyBoxLength, startPosX, startPosY) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sourceSheet);

  // Extracts all the start and finsh dates into arrays
  var rangeStartDate = sheet.getRange(2, 2, sheet.getLastRow()-1, 1).getValues();
  var rangeFinishDate = sheet.getRange(2, 3, sheet.getLastRow()-1, 1).getValues();
  var rangeDepth = sheet.getRange(2, 4, sheet.getLastRow()-1, 1).getValues();

  var startYr = rangeStartDate[0][0].getFullYear();
  var endYr = rangeFinishDate[0][0].getFullYear();
  var maxDepth = rangeDepth[0][0];

  // Extracting the min start and max finish date
  for (var i = 1; i < sheet.getLastRow()-1 ; i++) {
    if (rangeStartDate[i][0].getFullYear() < startYr) {
      startYr = rangeStartDate[i][0].getFullYear();
    }
    if (rangeFinishDate[i][0].getFullYear() > endYr) {
      endYr = rangeFinishDate[i][0].getFullYear();
    }
    if (rangeDepth[i][0] > maxDepth) {
      maxDepth = rangeDepth[i][0];
    }
  }

  // The Id of the presentation to copy
  var file = DriveApp.searchFiles("title = '" + templateName + "'").next();
  var templateId = file.getId();

  // Access the template presentation
  var template = SlidesApp.openById(templateId);
  var fileName = template.getName();
  var templateSlides = template.getSlides();

  // Create a new presentation first
  // (note: SlidesApp does not support a way to create a copy)
  var newPresenataionName = new Date().toLocaleString() + " " + fileName;
  //newPresenataionName = Browser.inputBox("Please enter the name for your roadmap: ");
  var newDeck = SlidesApp.create(newPresenataionName);

  // Remove default slides from the presentation with the name newPresenataionName 
  var defaultSlides = newDeck.getSlides();
  defaultSlides.forEach(function(slide) {
    slide.remove();
  });

  // Insert slides from template to the presentation named newPresenataionName
  var index = 0;
  templateSlides.forEach(function(slide) {
    var newSlide = newDeck.insertSlide(index);
    var elements = slide.getPageElements();
    elements.forEach(function(element) {
      newSlide.insertPageElement(element);
    });
    index++;
  });

  if (endYr - startYr < 1) {
    addTasksToMonthlySlide(newPresenataionName, sourceSheet, monthlyBoxLength, startPosX, startPosY, maxDepth);
  } else {
    addTasksToYearlySlide(newPresenataionName, sourceSheet, startYr, monthlyBoxLength, startPosX, startPosY, maxDepth);
  }
  debugger;
}

// Insert Shape takes the values in point(pt), [Convertion: 1in = 72pt].
function addTasksToMonthlySlide(destPresentationName, sourceSheet, monthlyBoxLength, startPosX, startPosY, maxDepth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sourceSheet);
  
  // Assume the height of the grid generate is atleast 4.5 inchs
  var gridHeight = 4.0;
  var height = (gridHeight/(maxDepth*2));

  // Extract the range for row holding "Task Name", "Start", "Finish" and "Vertical Position" 
  var range = sheet.getRange(2, 1, sheet.getLastRow()-1, 4);
  
  // Extract the values in the range "range"
  var valueRange = range.getValues();
  
  var destFile = DriveApp.searchFiles("title = '" + destPresentationName + "'").next();
  var destPresentation = SlidesApp.openById(destFile.getId());
  var destSlide = destPresentation.getSlides()[0];
  var count = 0; 

  var now = new Date();
  var dayInTheMonthNow = new Date(now.getFullYear(), now.getMonth()+1, 0).getDate();
  var whereYouAre = (startPosX+(now.getDate()*(monthlyBoxLength/dayInTheMonthNow)) + ((now.getMonth())*(monthlyBoxLength)));
    
  var verLine = destSlide.insertLine(SlidesApp.LineCategory.STRAIGHT, (whereYouAre*72), ((startPosX-0.10)*72), (whereYouAre*72), ((startPosX+gridHeight+0.10)*72));
  verLine.getLineFill().setSolidFill("#ffab40");

  for (var i=0; i<sheet.getLastRow()-1; i++) {
    var taskName= valueRange[i][0];
    var start = valueRange[i][1];
    var finish = valueRange[i][2];
    var vertPos = valueRange[i][3];

    // Reads the background colour to decide the colour of taskbars  
    range = sheet.getRange(2+i, 2, 1, 1);
    var backgroundColour = range.getBackground();

    var dayInTheMonthStart = new Date(start.getFullYear(), start.getMonth()+1, 0).getDate();
    var dayInTheMonthFinish = new Date(finish.getFullYear(), finish.getMonth()+1, 0).getDate();

    var xCoordStart = (start.getDate()*(monthlyBoxLength/dayInTheMonthStart)) + ((start.getMonth())*monthlyBoxLength);
    var xCoordFinish = (finish.getDate()*(monthlyBoxLength/dayInTheMonthFinish)) + ((finish.getMonth())*monthlyBoxLength);
    var width = xCoordFinish - xCoordStart;
    var yCoord = ((vertPos-1-count)*height*2);

    if (width > 0) {
      var posLeft = ((startPosX+xCoordStart)*72);
      var posTop = ((startPosY+yCoord)*72);
      var rectangle = destSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, posLeft, posTop, (width*72), (height*72));
      rectangle.getFill().setSolidFill(backgroundColour);
      
      var rRectangle = destSlide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, (posLeft-80-5), (posTop-(10/4)), 80, 10);
      var textbox = rRectangle.getText();
      textbox.setText(taskName).getTextStyle().setFontFamily("Arial").setFontSize(4).setBold(true);
    }
    else {
      var decagon = destSlide.insertShape(SlidesApp.ShapeType.DECAGON, ((((startPosX+xCoordStart)*72)-((0.1*72)/2))), ((startPosY-0.1)*72), (0.1*72), (0.1*72));
      decagon.getFill().setSolidFill('#38761d'); // Fills the circle with dark green
      var star = destSlide.insertShape(SlidesApp.ShapeType.STAR_5, ((((startPosX+xCoordStart)*72)-((0.1*72)/2))), ((startPosY-0.1)*72), (0.1*72), (0.1*72))
      star.getFill().setSolidFill('#FFFFFF'); // Fills the circle with white
    }
  }
  
  debugger;
}

// Insert Shape takes the values in point(pt), [Convertion: 1in = 72pt].
function addTasksToYearlySlide(destPresentationName, sourceSheet, startYr, monthlyBoxLength, startPosX, startPosY, maxDepth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sourceSheet);

  // Assume the height of the grid generate is atleast 4.5 inchs
  var gridHeight = 4.0;
  var height = (gridHeight/(maxDepth*2));
  
  // Extract the range for row holding "Task Name", "Start", "Finish" and "Vertical Position" 
  var range = sheet.getRange(2, 1, sheet.getLastRow()-1, 4);
  
  // Extract the values in the range "range"
  var valueRange = range.getValues();
  
  var destFile = DriveApp.searchFiles("title = '" + destPresentationName + "'").next();
  var destPresentation = SlidesApp.openById(destFile.getId());
  var destSlide = destPresentation.getSlides()[0];
  var count = 0; 
  
  var now = new Date();
  var dayInTheMonthNow = new Date(now.getFullYear(), now.getMonth()+1, 0).getDate();
  var whereYouAre = (startPosX+(now.getDate()*(monthlyBoxLength/dayInTheMonthNow)) + ((now.getMonth())*(monthlyBoxLength)) + ((monthlyBoxLength*12)*(now.getFullYear()-startYr)));
    
  var verLine = destSlide.insertLine(SlidesApp.LineCategory.STRAIGHT, (whereYouAre*72), ((startPosX-0.10)*72), (whereYouAre*72), ((startPosX+gridHeight)*72));
  verLine.getLineFill().setSolidFill("#ffab40");

  for (var i=0; i<sheet.getLastRow()-1; i++) {
    var taskName= valueRange[i][0];
    var start = valueRange[i][1];
    var finish = valueRange[i][2];
    var vertPos = valueRange[i][3];

    // Reads the background colour to decide the colour of taskbars  
    range = sheet.getRange(2+i, 2, 1, 1);
    var backgroundColour = range.getBackground();

    var dayInTheMonthStart = new Date(start.getFullYear(), start.getMonth()+1, 0).getDate();
    var dayInTheMonthFinish = new Date(finish.getFullYear(), finish.getMonth()+1, 0).getDate();

    var xCoordStart = (start.getDate()*(monthlyBoxLength/dayInTheMonthStart)) + ((start.getMonth())*(monthlyBoxLength)) + ((monthlyBoxLength*12)*(start.getFullYear()-startYr));
    var xCoordFinish = (finish.getDate()*(monthlyBoxLength/dayInTheMonthFinish)) + ((finish.getMonth())*(monthlyBoxLength)) + ((monthlyBoxLength*12)*(finish.getFullYear()-startYr));
    var width = xCoordFinish - xCoordStart;
    var yCoord = ((vertPos-1-count)*height*2);

    if (width > 0) {
      var posLeft = ((startPosX+xCoordStart)*72);
      var posTop = ((startPosY+yCoord)*72);
      var rectangle = destSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, posLeft, posTop, (width*72), (height*72));
      rectangle.getFill().setSolidFill(backgroundColour);
      
      var rRectangle = destSlide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, (posLeft-80-5), (posTop-((height*72)/4)), 80, (height*72));
      var textbox = rRectangle.getText();
      textbox.setText(taskName).getTextStyle().setFontFamily("Arial").setFontSize(4).setBold(true);
    }
    else {
      var decagon = destSlide.insertShape(SlidesApp.ShapeType.DECAGON, ((((startPosX+xCoordStart)*72)-((0.1*72)/2))), ((startPosY-0.1)*72), (0.1*72), (0.1*72));
      decagon.getFill().setSolidFill('#38761d'); // Fills the circle with dark green
      var star = destSlide.insertShape(SlidesApp.ShapeType.STAR_5, ((((startPosX+xCoordStart)*72)-((0.1*72)/2))), ((startPosY-0.1)*72), (0.1*72), (0.1*72))
      star.getFill().setSolidFill('#FFFFFF'); // Fills the circle with white
    }
  }

  debugger;
}
