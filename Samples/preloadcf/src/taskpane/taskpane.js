/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  // Office.addin.setStartupBehavior(Office.StartupBehavior.none);
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  cleanAllData();
}

async function cleanAllData () {
  let cleanData1 = await insertData(); 
  let cleanData2 = await validateData(cleanData1);
  let finalTable = await makeTable(cleanData2);
  return finalTable;
}

async function insertData() {
  try {
    await Excel.run(async context => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();

      // Define values for the range
      var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
      ["Bridles", 5000, 7000, 6544, 4377],
      ["Saddles", 400, 323, 276, 651],
      ["Boots", 12000, 8766, 8456, 9812],
      ["Hay", 1550, 1088, 692, 853],
      ["Wagons", 225, 600.25, 923, 544],
      ["Horseshoes", 6005, 7634, 4589, 8765]];
  
      // Create the range
      var range = sheet.getRange("A1:E7");
      range.values = values;
  
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
          sheet.getUsedRange().format.autofitColumns();
          sheet.getUsedRange().format.autofitRows();
      }
  
      sheet.activate();
    }) 
  } catch(error) {
    console.error(error);
  }
}

// Delete column B, values for Qtr 1
async function validateData() {
  try {
    await Excel.run(async context => {
    var sheet = context.workbook.worksheets.getActiveWorksheet();;
    var range = sheet.getRange("B1:B7");

    range.delete(Excel.DeleteShiftDirection.left);
   
    return context.sync();

  })
  } catch(error) {
    console.error(error); 
  }
} 

async function makeTable() {
  try {
    await Excel.run(async context => {
    var sheet = context.workbook.worksheets.getActiveWorksheet();;
    var Table = sheet.tables.add("A1:D6", true)
    Table.name = "HorseSuppliesTable"

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
  }

  sheet.activate();

  return context.sync();

  })
  } catch(error) {
    console.error(error); 
  }
}