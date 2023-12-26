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
 * Retrieves the stored URL of the Google Slide from script properties.
 * @return {string} The URL of the Google Slide.
 */
function getSlideUrl() {
  return SCRIPTPROPERTIES.getProperty(SCRIPT_PROPERTY_KEY_SLIDE_URL);
}

/**
 * Retrieves the stored name of the Index Sheet from script properties.
 * @return {string} The name of the Index Sheet.
 */
function getIndexSheetName() {
  return SCRIPTPROPERTIES.getProperty(SCRIPT_PROPERTY_KEY_INDEX_SHEET);
}

/**
 * Sets the URL of the Google Slide in script properties.
 * Verifies the URL before setting and shows a message if the URL is invalid.
 * @param {string} url - The URL of the Google Slide to be set.
 */
function setSlideUrl(url) {
  try {
    // Attempt to open the Google Slide to verify the URL
    SlidesApp.openByUrl(url);
    
    SCRIPTPROPERTIES.setProperty(SCRIPT_PROPERTY_KEY_SLIDE_URL, url);

    console.log('URL set successfully.');

    checkNextStep_(`Slide URL`,`set`);
  } catch (e) {
    Browser.msgBox('Invalid Google Slide URL. Try again.');
    console.log('Error: Invalid Google Slide URL.');
    accessSettingModal();
  }
}

/**
 * Sets the name of the Index Sheet in script properties.
 * Verifies the existence of the sheet before setting and shows a message if the sheet does not exist.
 * @param {string} name - The name of the Index Sheet to be set.
 */
function setIndexSheet(name) {
  try {
    let sheet = SPREADSHEET.getSheetByName(name);
    
    if (sheet) {
      let sheetId = sheet.getSheetId();
      var sheetUrl = "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/edit#gid=" + sheetId;
      let indexSheetData = {
        name: name,
        url: sheetUrl
      };
      // Serialize indexSheetData to a JSON string before storing
      SCRIPTPROPERTIES.setProperty(SCRIPT_PROPERTY_KEY_INDEX_SHEET, JSON.stringify(indexSheetData));
      
      // Assuming checkNextStep_ is a function you've defined elsewhere
      checkNextStep_('Index Sheet Name', 'set');
    } else {
      console.log('Error: Sheet not present.');
      Browser.msgBox('The sheet ' + name + ' does not exist in this Google Spreadsheet. Try again.');
      accessSettingModal(); // Assuming this is a function to reopen the modal
    }
  } catch (e) {
    console.log('Error: Unable to access the spreadsheet.');
    Browser.msgBox('Unable to access the spreadsheet. Try again.');
    accessSettingModal(); // Assuming this is a function to reopen the modal
  }
}

/**
 * Deletes a specified property from the script properties.
 * @param {string} propertyName - The name of the property to be deleted.
 */
function deleteProperty(propertyName) {
  SCRIPTPROPERTIES.deleteProperty(propertyName);
  let setTypeName = toTitleCase_(propertyName);
  checkNextStep_(setTypeName,`deleted`); 
}

/**
 * Checks the next step after a setting has been successfully set or deleted.
 * Prompts the user to conduct another setting if desired.
 * @param {string} setTypeName - The type of setting that was processed.
 * @param {string} executionType - The type of execution performed ('set' or 'deleted').
 */
function checkNextStep_(setTypeName,executionType){
  let checkNextStep = Browser.msgBox(`${setTypeName} was successfully ${executionType}. Do you want to conduct another setting?`,Browser.Buttons.YES_NO);
  if(checkNextStep === 'yes'){
    accessSettingModal();
  }
}

/**
 * Converts a string from snake_case to Title Case.
 * @param {string} str - The string to be converted.
 * @return {string} The converted string in Title Case.
 */
function toTitleCase_(str) {
  return str
    // First, replace underscores with spaces
    .replace(/_/g, ' ')
    // Split the string at each space, then transform to title case
    .split(' ')
    // Convert to title case by capitalizing the first letter of each word
    .map(word => word.charAt(0) + word.slice(1).toLowerCase())
    .join(' ');
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