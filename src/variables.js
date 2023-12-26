const SCRIPT_PROPERTY_KEY_SLIDE_URL = 'SLIDE_URL';
const SCRIPT_PROPERTY_KEY_INDEX_SHEET = 'INDEX_SHEET';
const SCRIPT_PROPERTY_KEY_SAVED_DETAILS = 'SAVED_DETAILS';
const SCRIPTPROPERTIES = PropertiesService.getScriptProperties();
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const SPREADSHEET_ID = SPREADSHEET.getId();

/* Lookup object for script property keys
  This object maps the string identifiers (used in client-side interactions)
  to the actual constant values representing script property keys.
  It's used to dynamically retrieve the correct property key based on a string identifier,
  which helps in efficiently managing the deletion or manipulation of script properties
  without hardcoding multiple if-else conditions or switch cases.
*/