/*
Copyright © 2020 Katherine Dixon Palevich

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

*/
function onEdit(e) {
  if (e.range.getSheet().getName() == 'Home') {
    if (e.range.getA1Notation() == 'B7') {
      e.range.setValue('working...');
      try {
        createLaunch();
      } catch(e) {
        // Programming error, or maybe sheet already exists.
      }
      e.range.clear();
    }
  }
  else if(e.range.getSheet().getRange('F1').getCell(1,1).getValue() == 'Menu'
         && e.range.getA1Notation() == 'F2'){
    e.range.setValue('working...');
      try {
        saveLaunchData(e.range.getSheet());
      } catch(e) {
        // Programming error, or maybe sheet already exists.
      }
      e.range.clear();
  }
}

function saveLaunchData(sheet){
  var launches = sheet.getParent().getSheetByName('Launches');
  launches.insertColumns(2);
  var sourceRange = sheet.getRange(1,2,25);
  var destinationRange = launches.getRange(1,2,25);
  sourceRange.copyTo(destinationRange);
  var members = '';
  for(var i = 2; i <= 11; i++){
    if(sheet.getRange(i,5).getValue()){
      if(members != ''){
        members += ",";
      }
      members += sheet.getRange(i,4).getValue();
    }
  }
  launches.getRange(27,2).setValue(members);
}

function getLaunchTitle(date) {
  let ops = {
    year: '2-digit',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit'
  };
  return date.toLocaleDateString(undefined, ops);
} 

function createLaunch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName('Launch Template');
  var today = new Date();
  var newSheet = ss.insertSheet(getLaunchTitle(today), 1, {template: templateSheet});
  newSheet.getRange(1,2).setValue(today);
}