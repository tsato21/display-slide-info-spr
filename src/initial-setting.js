/**
 * Creates and shows a custom menu in the Google Spreadsheet UI.
 * This menu provides various options for managing the spreadsheet and its content.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a custom menu with chaining for a cleaner structure.
  ui.createMenu('Custom Menu')
    .addItem('Update Index & Task Sheets', 'updateIndexAndTaskSheets')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Settings')
        .addItem('Set Necessary Info', 'accessSettingModal')
        .addItem('Delete All Sheets and Pre-Set Info', 'deleteAllExceptFirstAndClearProperties')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Others')
        .addItem('Conduct Authorization', 'showAuthorization')
    )
    .addToUi();
}

/**
 * Performs authorization tasks for Google Spreadsheet and Google Slides services.
 * This function is designed to be triggered by a menu item.
 */
function showAuthorization(){
  SpreadsheetApp;
  SlidesApp;
  ScriptApp;
  PropertiesService;
}

/**
 * Displays a modal dialog with a list of settings for the user to configure.
 * This dialog allows the user to set necessary information for the spreadsheet's operation.
 */
function accessSettingModal() {
    // Create a template from the HTML file
    let htmlTemplate = HtmlService.createTemplateFromFile('show-setting');
    
    let html = htmlTemplate
        .evaluate()
        .setWidth(1200)
        .setHeight(350);
    SpreadsheetApp.getUi().showModalDialog(html, 'Set Necessary Info');
}

/**
 * Retrieves the stored settings for the Google Slide URL and the name of the Index Sheet.
 * @return {Object} An object containing the slide URL and index sheet name.
 */
function getSettings() {
  return {
    slideUrl: SCRIPTPROPERTIES.getProperty(SCRIPT_PROPERTY_KEY_SLIDE_URL),
    indexSheetName: JSON.parse(SCRIPTPROPERTIES.getProperty(SCRIPT_PROPERTY_KEY_INDEX_SHEET) || '{}').name
  };
}

/**
 * Sets the settings for the Google Slide URL and the Index Sheet name.
 * Verifies the existence of the provided Google Slide URL and the Index Sheet.
 * If either is not valid, throws an error.
 * Updates the script properties with the provided values.
 * @param {string} slideUrl - The URL of the Google Slide.
 * @param {string} indexSheetName - The name of the Index Sheet in the Google Spreadsheet.
 * @throws {Error} If the slide URL is not provided, not valid, or does not exist.
 * @throws {Error} If the index sheet name is not provided or the sheet does not exist.
 */
function setSettings(slideUrl, indexSheetName) {
  let sheet;
  if (slideUrl) {
    if(!SlidesApp.openByUrl(slideUrl)){
      throw new Error(`URL not exits`);
    }
  } else {
    throw new Error(`URL not input`);
  }

  if (indexSheetName) {
    sheet = SPREADSHEET.getSheetByName(indexSheetName);
    if (!sheet) {
        throw new Error(`Sheet with the name not exits`);
    }
  } else {
    throw new Error(`Index Sheet Name not input`);
  }
  SCRIPTPROPERTIES.setProperty(SCRIPT_PROPERTY_KEY_SLIDE_URL, slideUrl);
  let sheetId = sheet.getSheetId();
  let spreadsheetUrl = SPREADSHEET.getUrl();
  let spreadSheetId = extractIDFromUrl_(spreadsheetUrl);
  let sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadSheetId}/edit#gid=${sheetId}`;
  let indexSheetData = { name: indexSheetName, url: sheetUrl };
  SCRIPTPROPERTIES.setProperty(SCRIPT_PROPERTY_KEY_INDEX_SHEET, JSON.stringify(indexSheetData));

  Browser.msgBox(`Settings were completed.`);
}

/**
 * Deletes the stored settings for the Google Slide URL and the Index Sheet name.
 */
function deleteSettings() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(SCRIPT_PROPERTY_KEY_SLIDE_URL);
  properties.deleteProperty(SCRIPT_PROPERTY_KEY_INDEX_SHEET);

  // Logging for debugging purposes
  console.log('Settings have been deleted.');
  Browser.msgBox('Settings have been deleted.');
}

/**
 * Extracts the unique identifier (ID) from a given Google Sheets or Google Slides URL.
 * @param {string} url - The URL from which the ID is to be extracted.
 * @return {string|null} The extracted ID or null if not found.
 */
function extractIDFromUrl_(url) {
  var match = /\/d\/([a-zA-Z0-9-_]+)/.exec(url);
  return match ? match[1] : null;
}


/**
 * Deletes all sheets in the spreadsheet except the first one and clears all script properties.
 * Intended for resetting the spreadsheet to its initial state.
 */
function deleteAllExceptFirstAndClearProperties() {
  var sheets = SPREADSHEET.getSheets();
  
  sheets[0].clear();
  // Loop through all sheets except the first sheet and delete sheets
  for (var i =1; i< sheets.length; i++) {
      SPREADSHEET.deleteSheet(sheets[i]);
  }
  
  SCRIPTPROPERTIES.deleteAllProperties();

  Browser.msgBox(`All sheets (excep the first one) were deleted and all pre-set information was reset.`);
}

/*
----If you want to display custom menu only when the onwer of the spreadsheet opens this spreadsheet----
1. Make the two functions, `setUpOnOpenTrigger` and `showCustomMenu_` execusabel.
2. Insert a shape from drawing into the index sheet, assign the function, `setUpOnOpenTrigger` into the shape.
3. Disable the function, `onOpen` since this allows the custome menu available to all users of this spreadsheet.
4. Click the shape to execute the function, `setUpOnOpenTrigger`.
*/
/**
 * Sets up a trigger that automatically adds a custom menu to the spreadsheet UI when the spreadsheet is opened.
 * This function should be executed by clicking a button on the Spreadsheet.
 * The first click will prompt for authorization, and subsequent clicks will set up the trigger.
 * This setup is necessary to ensure proper authorization and functionality of the onOpen trigger.
 */
/*
function setUpOnOpenTrigger() {
  ScriptApp.newTrigger('showCustomMenu_')
    .forSpreadsheet(SPREADSHEET)
    .onOpen()
    .create();
}
*/

/**
 * Creates and shows a custom menu in the Google Spreadsheet UI if the current user is the owner of the spreadsheet.
 * The custom menu includes options to update index sheets and task sheets.
 */
/*
function showCustomMenu_() {  
  // Get the email of the currently active user
  let userEmail = Session.getActiveUser().getEmail();
  
  // Get the email of the owner of the file
  let ownerEmail = DriveApp.getFileById(SPREADSHEET_ID).getOwner().getEmail();
  
  // If the active user is the owner, add the custom menu
  if (userEmail === ownerEmail) {
    var ui = SpreadsheetApp.getUi();
    // Create a custom menu.
    var menu = ui.createMenu('Custom Menu');

    // Add menu items with corresponding function names to be called.
    menu.addItem('Update Index & Task Sheets', 'updateIndexAndTaskSheets')
        .addSeparator();

    // Add a sub-menu.
    var othersMenu = ui.createMenu('Others');
    othersMenu.addItem('Conduct Authorization', 'showAuthorization');
    othersMenu.addItem('Set Necessary Info', 'accessSettingModal');

    // Add the sub-menu to the main menu.
    menu.addSubMenu(othersMenu);

    // Add the menu to the UI.
    menu.addToUi();
  }
}
*/