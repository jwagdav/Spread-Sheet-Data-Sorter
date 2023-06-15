function SortData(){
//references 
const ss = SpreadsheetApp.getActiveSpreadsheet();
const formEntries = ss.getSheetByName('Form Entries');
const allRides = ss.getSheetByName('All Rides');

const formEntriesData = formEntries.getDataRange().getValues();
const allRidesData = allRides.getDataRange().getValues();

const allRidesHeader = allRidesData[0];
const formEntriesHeader = formEntriesData[0];

  //filters allRidesData using references to the 'Import Status' and 'Date' columns 
  //returns an array of data to archive
  //row.some(Boolean) ensures that no empty rows are added to the array
const filteredAllRides = allRidesData.filter(row => {
  const importStatus = row[allRidesHeader.indexOf('Import Status')];
  const toBeArchived = importStatus === 'ALL COMPLETE' || importStatus === 'Cancelled';
  const date = new Date(row[allRidesHeader.indexOf('Date')]);
  const currentDate = new Date();
  const timeDiff = currentDate.getTime() - date.getTime();
  const daysDiff = Math.floor(timeDiff / (24 * 60 * 60 * 1000));
  const is31DaysAgo = daysDiff >= 31;
  return toBeArchived && is31DaysAgo && row.some(Boolean);
});

if(filteredAllRides.length <= 0) {
  Logger.log('No data meets the criteria to be archived');
  return;
}

const entryIDs = filteredAllRides.map(row => {
  if(row.some(Boolean)) {
    return row[allRidesHeader.indexOf('Entry ID')].toString();
  }
  else {
    Logger.log('No data meets the criteria')
    return;
  }
});

const filteredFormEntries = formEntriesData.filter(row => {
  const entryID = row[0].toString();
  return entryIDs.includes(entryID);
});

const filteredFormEntryIDs = filteredFormEntries.map(row => row[formEntriesHeader.indexOf('EntryID')]).toString();
//these logging statements can be used to assess if filtering functions are getting the right data
Logger.log(filteredAllRides.length)
Logger.log(entryIDs.length)
Logger.log(filteredFormEntries.length);

//searches users Drive for a spreadsheet of the current year
const currentYear = new Date().getFullYear();
const fileName = `Archived Data ${currentYear}`;

//there is a lot of repeat code in here that could be later optimized to functions
// it is this way for now because it needs to check if the files and sheets exist in the right context to prevent errors
const files = DriveApp.getFilesByName(fileName);
  if (files.hasNext()) {
    const file = files.next();
    const archivedSpreadsheet = SpreadsheetApp.open(file);
    const archivedFormEntries = archivedSpreadsheet.getSheetByName('Form Entries');
    const archivedAllRides = archivedSpreadsheet.getSheetByName('All Rides');

    const archivedFormEntriesData = archivedFormEntries.getDataRange().getValues();
    const archivedAllRidesData = archivedAllRides.getDataRange().getValues();

    const archivedFormEntriesHeader = archivedFormEntriesData[0];
    const archivedAllRidesHeader = archivedAllRidesData[0];

    if (archivedFormEntries) {
      if (archivedFormEntriesHeader != formEntriesHeader) {
        archivedFormEntries.appendRow(formEntriesHeader);
      }

      if (filteredFormEntries.length >= 1) {
        // Create a 2D array to store all rows to be appended
        const rowsToAppend = [];

        for (let i = filteredFormEntries.length - 1; i >= 0; i--) {
          const row = filteredFormEntries[i];
          rowsToAppend.push(row);
          const rowIndex = formEntriesData.indexOf(row) + 1;
        }

        // Append all rows in one batch operation
        archivedFormEntries.getRange(archivedFormEntries.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

        //this clears the content from the archived rows
        //this is done to prevent the script from timing out 
        //the now empty rows will need to be deleted later
        const deleteRowIndices = filteredFormEntries.map(row => formEntriesData.indexOf(row) + 1);
        const deleteRange = formEntries.getRangeList(deleteRowIndices.map(rowIndex => `${rowIndex}:${rowIndex}`));
        deleteRange.clearContent();
        }

        Logger.log(`${filteredFormEntries.length} rows added to archive and deleted from current sheet`);
      } else {
      archivedFormEntries = getSheetByName('sheet1');
      archivedFormEntries.setName('Form Entries');

      if (archivedFormEntriesHeader != formEntriesHeader) {
        archivedFormEntries.appendRow(formEntriesHeader);
      }

      if (filteredFormEntries.length >= 1) {
        // Create a 2D array to store all rows to be appended
        const rowsToAppend = [];

        for (let i = filteredFormEntries.length - 1; i >= 0; i--) {
          const row = filteredFormEntries[i];
          rowsToAppend.push(row);
          const rowIndex = formEntriesData.indexOf(row) + 1;
        }

        archivedFormEntries.getRange(archivedFormEntries.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

        const deleteRowIndices = filteredFormEntries.map(row => formEntriesData.indexOf(row) + 1);
        const deleteRange = formEntries.getRangeList(deleteRowIndices.map(rowIndex => `${rowIndex}:${rowIndex}`));
        deleteRange.clearContent();
      }

      Logger.log(`${filteredFormEntries.length} rows added to archive from Form Entries sheet`);
    }

    if (archivedAllRides) {
      if (archivedAllRidesHeader != allRidesHeader) {
        archivedAllRides.appendRow(archivedAllRidesHeader);
      }

      if (filteredAllRides.length >= 1) {
  
        const rowsToAppend = [];

        for (let i = filteredAllRides.length - 1; i >= 0; i--) {
          const row = filteredAllRides[i];
          rowsToAppend.push(row);
          const rowIndex = allRidesData.indexOf(row) + 1;
        }

        archivedAllRides.getRange(archivedAllRides.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
        
        const deleteRowIndices = filteredAllRides.map(row => allRidesData.indexOf(row) + 1);
        const deleteRange = allRides.getRangeList(deleteRowIndices.map(rowIndex => `${rowIndex}:${rowIndex}`));
        deleteRange.clearContent();
      }

      Logger.log(`${filteredAllRides.length} rows archived from All Rides`);
    } else {
      const newSheet = archivedSpreadsheet.insertSheet('All Rides');

      if (archivedAllRidesHeader != allRidesHeader) {
        archivedAllRides.appendRow(archivedAllRidesHeader);
      }

      if (filteredAllRides.length >= 1) {
        // Create a 2D array to store all rows to be appended
        const rowsToAppend = [];

        for (let i = filteredAllRides.length - 1; i >= 0; i--) {
          const row = filteredAllRides[i];
          rowsToAppend.push(row);
          const rowIndex = allRidesData.indexOf(row) + 1;
        }
        
        archivedAllRides.getRange(archivedAllRides.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

        const deleteRowIndices = filteredAllRides.map(row => allRidesData.indexOf(row) + 1);
        const deleteRange = allRides.getRangeList(deleteRowIndices.map(rowIndex => `${rowIndex}:${rowIndex}`));
        deleteRange.clearContent();
      }

      Logger.log(`${filteredAllRides.length} rows archived from All Rides`);
    }
  } else {
    Logger.log('Spreadsheet not found');
  }
}
