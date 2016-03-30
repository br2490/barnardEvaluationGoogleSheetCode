/**
 * Script written by Benjamin Rosner - Application Designer, Barnard Library and Academic Information Systems (BLAIS)
 * If there are any questions or the code no longer works please contact me.
 *
 * e: brosner@barnard.edu 
 * p: x49005
 *
 * Version 1 - 2015/OCT/07
 * Version 2 - 2015/OCT/12
 *    - added error checking
 *    - fixed various bugs.
 **/


/**
 * Add a UI element.
 * Called on OPEN. 
 **/
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('BARNARD EVALUATIONS')
        .addItem('Generate Master Sheet', 'populateChildren')
        .addToUi();
}

/**
 * A function to build, gather access and display an array of child sheets.
 *
 * All required params come from the Spreadsheet Sheet by name "Variables"
 * 
 * @customfunction
 **/
function populateChildren() {
    // Get the Spreadsheet this script lives in.
    var ss = SpreadsheetApp.getActiveSpreadsheet(),

        // Get the Sheets (hahaha!)
        variables = ss.getSheetByName("variables"),
        access = ss.getSheetByName("accessSheet"),
        fullImportSheet = ss.getSheetByName("ImportedSheet");
    // Done messing with the Sheets :) //


    // START VARIABLES //
    // Grab CONSTANT variables set in the variables sheet.
    var importConst = variables.getRange("D11").getValue(),
        numberOfDepts = variables.getRange("D2").getValue(),
        ecc = variables.getRange("C30").getValue(),

        // Grab all of our URLs.
        urlRange = variables.getRange(2, 1, numberOfDepts).getValues(),

        // Declare our "master  IMPORTRANGE array formula" prefixed with a SORT to 'remove' (hide) String.EMPTY.
        completedCommand = '=SORT(ARRAYFORMULA({';

    // Clear our access variables Sheet, prepare it for new Sheets.
    access.clear();
    access.appendRow(['This sheet will print the first COURSE ID from each department OR present a #REF! error.']);
    access.appendRow(['Any #REF! error requires you to float over it and click the button to allow variables access.']);
    access.appendRow(['----START----']);

    var urls;
    for (urls in urlRange) {
        var thisUrl = urlRange[urls];
        // Create a resource on the acess Sheet for each item using cell "A4"
        access.appendRow(['=IMPORTRANGE("' + thisUrl + '","A4")']);
        // Concat to completedCommand for inclusion
        completedCommand += 'IMPORTRANGE("' + thisUrl + '","' + importConst + '");';
    }
    // Remove the trailing ";"
    completedCommand = completedCommand.slice(0, -1);

    // Done with access Sheet.
    completedCommand = completedCommand + '}))';
    access.appendRow(['-----END-----']);

    // This should be its own function, and run for all Sheets - less specific code - but w/e.
    if (fullImportSheet) {
        // Out with the old
        ss.deleteSheet(fullImportSheet);
        // In with the new
        ss.insertSheet('ImportedSheet', 3);
        // Get the Sheet again
        fullImportSheet = ss.getSheetByName("ImportedSheet");
        // Format the Sheet, assign the headings, etc.
        fullImportSheet.setFrozenRows(1);
        fullImportSheet.appendRow(['=TRANSPOSE(variables!C14:C26)']);
    }

    fullImportSheet.appendRow(['', 'ecc', completedCommand]);
    resize(fullImportSheet, 16);
}


function resize(sheet, numberOfCols) {
    for (i = 1; i <= numberOfCols; i++) {
        sheet.autoResizeColumn(i);
    }
}