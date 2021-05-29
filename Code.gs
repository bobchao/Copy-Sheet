var targetSheetId = 0; //change this when you use the script in another spreadsheet file.
var sourceSheetId = 0; //change this when you use the script in another spreadsheet file.
var sSortingColumn = 16; //I need to sort the source sheet before sync the data so...
var baseTime = new Date("2021-05-04"); //first update. 
var sheetNameTemplate = `本期次序 (下次更新: ${getNextUpdateDateString()})`;
var updateNoticeTemplate = `待辦事項參考順序\n（${Utilities.formatDate(new Date(), "GMT+8", "MM-dd")}更新）`;

function main(){
  var targetSheet = getSheetById(targetSheetId);
  var sourceSheet = getSheetById(sourceSheetId);

  var tRange = targetSheet.getRange(
    targetSheet.getFrozenRows()+1,
    1,
    targetSheet.getMaxRows(),
    targetSheet.getMaxColumns()
  );

  var sRange = sourceSheet.getRange(
    sourceSheet.getFrozenRows()+1,
    1,
    sourceSheet.getMaxRows(),
    sourceSheet.getMaxColumns()
  );

  //prevent others run the script at the same time.
  //not sure if it works :P...
  var lock = LockService.getScriptLock();

  try { //hey this is the first time I use try catch... orz... 
    lock.waitLock(10000);
  } catch (e) {
    Logger.log('Could not obtain lock after 10 seconds.');
  }

  if (lock.hasLock()) {
    //be sure to make the target sheet clean first
    tRange.clear();

    //Sorting before copy
    sourceSheet.sort(sSortingColumn, false);
    sRange.copyTo(tRange);
    
    //remove the empty rows in target sheet (while keep the source untouched.)
    Utilities.sleep(500); //Sleep for a while to make sure the process above did finished.
    var tLastRowAfterCopy = getLastRowSpecial(targetSheet.getRange("A1:A"));
    targetSheet.deleteRows(
      tLastRowAfterCopy+1,
      targetSheet.getMaxRows() - tLastRowAfterCopy
    )

    //change the update time and sheet name of the targetSheet
    targetSheet.getRange("A3").setValue(updateNoticeTemplate);
    targetSheet.setName(sheetNameTemplate);
  }

  lock.releaseLock();
  SpreadsheetApp.flush();
}

function getNextUpdateDateString(){ // update biweekly
  var currentDate = new Date(); //Today
  var nextUpdateDate = new Date();
  var baseDayBiweeklyDiff =
    Math.floor(
      (currentDate.getTime()-baseTime.getTime())
      / (1000 * 3600 * 24)
      % 14); 

  nextUpdateDate.setDate(currentDate.getDate() + (14-baseDayBiweeklyDiff));
  Logger.log(`下次更新在 ${Utilities.formatDate(nextUpdateDate, "GMT+8", "MM-dd")}`);
  return Utilities.formatDate(nextUpdateDate, "GMT+8", "MM-dd");
}

function getSheetById(sid){
  //https://stackoverflow.com/a/26682689
  //why google didn't make this in App Sctipt?
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === sid;}
  )[0];
}

function getActiveSheetId(){
  //https://stackoverflow.com/a/26682689
  var id  = SpreadsheetApp.getActiveSheet().getSheetId();
  Logger.log(id.toString());
  return id;
}

function getLastRowSpecial(range){
  //modified from https://yagisanatode.com/2019/05/11/google-apps-script-get-the-last-row-of-a-data-range-when-other-columns-have-content-like-hidden-formulas-and-check-boxes/
  //the original script check the range from the top to the bottom
  //my script check them reversely to deal with skipped rows issue.
  var rangeValue = range.getValues();
  var rowNum = rangeValue.length;

  for(var row = rangeValue.length-1; row > 0 ; row--){
    if(rangeValue[row][0] !== ""){
      rowNum = row+1;
      break;
    };
  };
  return rowNum;
};
