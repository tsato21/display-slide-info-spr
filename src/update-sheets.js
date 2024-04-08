/**
 * Updates the index and task sheets of the spreadsheet. This includes deleting
 * all existing task sheets and recreating them based on the latest data from
 * Google Slides, and then updating the index sheet to reflect these changes.
 */
function updateIndexAndTaskSheets() {
  let startTime = new Date().getTime(); // Record the start time of the script
  let maxExecutionTime = 300000; // Set the maximum execution time to 5 minutes (300000 ms)

  let slideUrl = SCRIPTPROPERTIES.getProperty(SCRIPT_PROPERTY_KEY_SLIDE_URL);
  let slideId = extractIDFromUrl_(slideUrl);
  let presentation = SlidesApp.openById(slideId);
  let slides = presentation.getSlides();

  // Retrieve saved data if it exists, otherwise start fresh
  let savedDetails = SCRIPTPROPERTIES.getProperty(
    SCRIPT_PROPERTY_KEY_SAVED_DETAILS
  );
  let allDetails = savedDetails ? JSON.parse(savedDetails) : [];
  let currentDetails =
    allDetails.length > 0
      ? allDetails.pop()
      : {
          workCategory: null,
          subWorkCategory: null,
          tasks: [],
          lastSlideIndex: -1, // Initialize lastSlideIndex to -1 to denote starting from the beginning
        };

  // Determine the starting index based on the last processed slide
  let startingIndex = currentDetails.lastSlideIndex + 1;
  // console.log(`startingIndex is ${startingIndex}`);

  let patternCategory = /Category:\s*【(.*?)】\s*(.*?)(?=Task:|Summary:|$)/;
  let patternTask = /Task:\s*(.*?)(?=Category:|Summary:|$)/;
  let patternSummary = /Summary:\s*(.*?)(?=Category:|Task:|$)/;

  for (let i = startingIndex; i < slides.length; i++) {
    // Simulate a delay to test the timeout functionality
    // For example, sleep for 10 seconds
    // Utilities.sleep(5000); // Sleep for 5 seconds to test
    // Check the elapsed time
    let currentTime = new Date().getTime();
    // let readableTime = new Date(currentTime).toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', hour12: false });
    // console.log(`Current time is ${readableTime} in Slide ${i + 1}.`);
    if (currentTime - startTime >= maxExecutionTime) {
      // Save the current state and set a trigger if the script is approaching the time limit
      currentDetails.lastSlideIndex = i - 1; // Save the index of the last processed slide
      allDetails.push(currentDetails);
      SCRIPTPROPERTIES.setProperty(
        SCRIPT_PROPERTY_KEY_SAVED_DETAILS,
        JSON.stringify(allDetails)
      );
      // Log timeout and set a trigger for a new execution
      ScriptApp.newTrigger('updateIndexAndTaskSheets')
        .timeBased()
        .after(5000) // Set the trigger to run 5 seconds after the current execution ends
        .create();
      console.log(
        `Time out detected in Slide ${
          i + 1
        }, saving current details and setting a trigger.`
      );
      Browser.msgBox(`Still processing. Please wait a while.`);
      return; // Exit the function to allow the trigger to start a new execution
    }
    let slide = slides[i];
    let shapes = slide.getShapes();
    let entireSlideText = shapes
      .map((shape) => shape.getText().asString().trim())
      .join(' ');
    // console.log(entireSlideText);
    // Check if the entire slide text contains all three keywords
    let matchCategory = entireSlideText.match(patternCategory);
    let matchTask = entireSlideText.match(patternTask);
    let matchSummary = entireSlideText.match(patternSummary);

    if (matchCategory && matchTask && matchSummary) {
      // console.log(`Slide ${i + 1} is subject for extraction.`);
      let slideUrl = presentation.getUrl() + '#slide=id.' + slide.getObjectId();
      let workCategory = matchCategory[1].trim();
      let subWorkCategory = matchCategory[2].trim();
      let slideDetails = extractSlideDetails_(
        workCategory,
        subWorkCategory,
        matchTask[1],
        matchSummary[1],
        slideUrl
      );

      // Check if the combination of work category and sub-category has changed in the current slide
      if (
        slideDetails.workCategory + slideDetails.subWorkCategory !==
        currentDetails.workCategory + currentDetails.subWorkCategory
      ) {
        /*
              If the category or sub-category has changed, it means we've moved to a new category or sub-category.
              Therefore, push the currentDetails (which holds the accumulated data for the previous category or sub-category) to allDetails.
              */
        // Check if currentDetails has any tasks before pushing to allDetails
        if (currentDetails.tasks.length > 0) {
          allDetails.push(currentDetails);
        }

        /*
              Initialize currentDetails for the new category or sub-category.
              This will start accumulating tasks for this new category or sub-category.
              */
        currentDetails = {
          workCategory: slideDetails.workCategory, // Set the new work category
          subWorkCategory: slideDetails.subWorkCategory, // Set the new sub-work category
          tasks: [], // Initialize an empty array for tasks
        };
      }

      /*
            Add the tasks from the current slide to the currentDetails.
            This happens regardless of whether the category has changed or not.
            If the category hasn't changed, we continue adding tasks to the current category.
            If it has changed, we're adding the first task(s) of the new category.
            */
      currentDetails.tasks.push(...slideDetails.tasks);

      // Debugging: Log the currentDetails after each slide is processed
      // console.log(`Slide ${i + 1}:`, JSON.stringify(currentDetails, null, 2));
    } else {
      // console.log(`Slide ${i + 1} is NOT subject for extraction.`);
    }
  }

  // Once we've finished processing all slides, push the final currentDetails to allDetails
  if (currentDetails.workCategory !== null) {
    //Before pushing currentDetails into allDetails, delete all task sheets of the Spreadsheet
    deleteAllTaskSheets_();

    allDetails.push(currentDetails);
    inputSlidesInfoToSheet_(allDetails, SPREADSHEET);
    // console.log(`allDetails are successfully input into Google Sheet`);
    // Clear the saved data after successfully processing all slides
    SCRIPTPROPERTIES.deleteProperty(SCRIPT_PROPERTY_KEY_SAVED_DETAILS);
    // delete triggers to execute `updateIndexAndTaskSheets`
    deleteSpecificTrigger_(`updateIndexAndTaskSheets`);
    updateIndexSheet_();
    SpreadsheetApp.getUi().alert(
      `Index Sheet and Task Sheets have been updated.`
    );
  }
}

/**
 * Extracts slide details such as work category, sub-category, and tasks from provided parameters.
 * This function organizes tasks into categories and sub-categories based on the text extracted from a Google Slide.
 *
 * @param {string} workCategory - The work category extracted from the slide.
 * @param {string} subWorkCategory - The sub-work category extracted from the slide.
 * @param {string} task - The task name extracted from the slide.
 * @param {string} summary - The task summary extracted from the slide.
 * @param {string} slideUrl - The URL of the slide from which the details are extracted.
 * @return {Object} An object containing the organized slide details, including:
 *                  - workCategory: The category of the work.
 *                  - subWorkCategory: The sub-category of the work.
 *                  - tasks: An array of task objects, each with a name, summary, and URL of the slide.
 */
function extractSlideDetails_(
  workCategory,
  subWorkCategory,
  task,
  summary,
  slideUrl
) {
  let slideDetails = {
    workCategory: workCategory,
    subWorkCategory: subWorkCategory,
    tasks: [],
  };

  let currentTask = {
    name: task.trim(),
    summary: summary.trim(),
    url: slideUrl,
  };
  slideDetails.tasks.push(currentTask);

  return slideDetails;
}

/**
 * Deletes all sheets in the spreadsheet except for the index sheet.
 */
function deleteAllTaskSheets_() {
  let sheets = SPREADSHEET.getSheets();
  let indexSh = JSON.parse(
    ScriptProperties.getProperty(SCRIPT_PROPERTY_KEY_INDEX_SHEET)
  );
  if (indexSh) {
    let indexShName = indexSh.name;
    for (sheet of sheets) {
      let sheetName = sheet.getName();
      if (sheetName === indexShName) {
        continue;
      }
      SPREADSHEET.deleteSheet(sheet);
    }
  } else {
    for (i = 1; i < sheets.length; i++) {
      SPREADSHEET.deleteSheet(sheet[i]);
    }
  }
}

/**
 * Inputs the organized slide information into the designated Google Sheet.
 *
 * @param {Array} allDetails - Array of objects containing the task details.
 * @param {Object} spreadsheet - The Google Sheets object to input data into.
 */
function inputSlidesInfoToSheet_(allDetails, SPREADSHEET) {
  // Loop through each detail object and add to the spreadsheet
  for (let detail of allDetails) {
    let sheetName = `${detail.workCategory}: ${detail.subWorkCategory}`;
    // console.log(`detail in inputSlidesInfoToSheet_ is ${JSON.stringify(detail)}`);

    sheet = SPREADSHEET.insertSheet(sheetName, SPREADSHEET.getNumSheets() + 1);

    // Log the details for debugging
    // console.log(`Following details will be input into Google Sheet: ${JSON.stringify(detail, null, 2)}`);

    // Start inserting data from the second row
    let startRow = 2;
    for (let i = 0; i < detail.tasks.length; i++) {
      // Generate a hyperlink formula for each task
      let hyperlinkFormula = `=HYPERLINK("${detail.tasks[i].url}", "${detail.tasks[i].name}")`;
      // Set the values for task name and summary in the sheet
      sheet
        .getRange(startRow + i, 2, 1, 2)
        .setValues([[hyperlinkFormula, detail.tasks[i].summary]]);
    }

    // Define the header range and set the header titles
    let headerRange = sheet.getRange(1, 2, 1, 2);
    // Define the range for the data to be formatted
    let dataRange = sheet.getRange(2, 2, detail.tasks.length, 2);
    // Define the range for text wrapping
    let wrapRange = sheet.getRange('B:C');

    let indexSheetUrl = JSON.parse(
      SCRIPTPROPERTIES.getProperty(SCRIPT_PROPERTY_KEY_INDEX_SHEET)
    ).url;
    // Apply the formatting to the sheet
    setSheetFormat_(
      sheet,
      2,
      [400, 600],
      ['Task', 'Summary'],
      headerRange,
      dataRange,
      null,
      wrapRange,
      null,
      indexSheetUrl
    );
  }
}

/**
 * Deletes all triggers associated with a specific function in the Google Apps Script project.
 *
 * @param {string} functionName - The name of the function for which to delete the triggers.
 */
function deleteSpecificTrigger_(functionName) {
  var allTriggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === functionName) {
      // Delete the trigger if the function name matches
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
}

/**
 * Sets the formatting for the Google Sheet including headers, column widths,
 * and wrap text settings. This function applies specific formatting rules to
 * make the sheet more readable and useful.
 * @param {Object} sheet - The sheet to format.
 * @param {number} startCol - The starting column for setting widths.
 * @param {Array} columnWidths - The widths for the columns.
 * @param {Array} headers - The headers for the sheet.
 * @param {Object} headerRange - The range for the header to format.
 * @param {Object} dataRange - The range for the data to format.
 * @param {Object} blueRange - The range to format with blue font color.
 * @param {Object} wrapRange - The range to set text wrapping.
 * @param {Object} vertiMidRange - The range to vertically align in the middle.
 * @param {string} indexSheetUrl - The URL for the index slide.
 */
function setSheetFormat_(
  sheet,
  startCol,
  columnWidths,
  headers,
  headerRange,
  dataRange,
  blueRange,
  wrapRange,
  vertiMidRange,
  indexSheetUrl
) {
  // Format the header if it exists
  if (headers && headers.length) {
    formatHeader_(headerRange, headers, indexSheetUrl);
  }

  // Set the column widths if the array is provided
  if (columnWidths && columnWidths.length) {
    setColumnsWidth_(sheet, startCol || 1, columnWidths);
  }

  // Format the blue range if provided
  if (blueRange) {
    blueRange.setHorizontalAlignment('center').setFontColor('blue');
  }

  // Apply borders and vertical alignment to header and data ranges
  if (headerRange && dataRange) {
    headerRange
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment('middle');
    dataRange
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment('middle');
  }

  // Enable text wrap for the specified range
  if (wrapRange) {
    wrapRange.setWrap(true);
  }

  // Vertically align the middle range if provided
  if (vertiMidRange) {
    vertiMidRange
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
  }
}

/**
 * Applies formatting to the header of the sheet including bold font, background
 * color, and alignment. Optionally adds a hyperlink to the index slide if an
 * indexSheetUrl is provided.
 * @param {Object} headerRange - The range of the header to format.
 * @param {Array} headers - The header titles.
 * @param {string} indexSheetUrl - The URL for the index slide.
 */
function formatHeader_(headerRange, headers, indexSheetUrl) {
  headerRange
    .setFontWeight('bold')
    .setBackground('#CCCCCC')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setValues([headers]);

  // If an indexSheetUrl is provided, create a hyperlink
  // If an indexSheetUrl is provided, create a hyperlink back to the index slide
  if (indexSheetUrl) {
    const hyperLinkToIndexSlide = `=HYPERLINK("${indexSheetUrl}","Back to Index")`;
    headerRange
      .getSheet()
      .getRange('A1')
      .setValue(hyperLinkToIndexSlide)
      .setBackground('#FFC0CB')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
}

/**
 * Sets the width for the specified columns in the sheet. This function is used
 * to make sure that each column in the sheet has an appropriate width for its content.
 * @param {Object} sheet - The sheet where column widths will be set.
 * @param {number} startCol - The starting column index for setting widths.
 * @param {Array} columnWidths - An array of widths to set for the columns.
 */
function setColumnsWidth_(sheet, startCol, columnWidths) {
  // Loop through each width value and set the column width accordingly
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(startCol + index, width);
  });
}

/**
 * Updates the main index sheet with categories and tasks.
 * It applies formatting and inserts hyperlinks to the respective task sheets.
 * To avoid the 6-minute execution limit, this function is designed to be
 * run incrementally, processing a limited number of categories per execution.
 * @param {string} indexSheetName - The name of the index sheet to update.
 */
function updateIndexSheet_() {
  let indexShName = JSON.parse(
    SCRIPTPROPERTIES.getProperty(SCRIPT_PROPERTY_KEY_INDEX_SHEET)
  ).name;
  let indexSheet = SPREADSHEET.getSheetByName(indexShName);
  indexSheet.clear();
  let lastColNum = indexSheet.getMaxColumns();
  let taskShData = fetchTaskSheetsData_();
  let needColNum = Object.keys(taskShData).length;

  // Check if additional columns are needed
  if (lastColNum < needColNum) {
    // Insert enough columns to meet the requirement
    let columnsToInsert = needColNum - lastColNum;
    indexSheet.insertColumnsAfter(lastColNum, columnsToInsert);
    // console.log(`${columnsToInsert} columns were inserted.`);
  }

  let currentCol = 1;

  Object.keys(taskShData).forEach((category) => {
    let tasks = taskShData[category];
    let updates = [[category]];
    let hyperlinkUpdates = [];

    tasks.forEach((taskInfo) => {
      updates.push([taskInfo.task]);
      hyperlinkUpdates.push([
        '=HYPERLINK("' + taskInfo.url + '","' + taskInfo.task + '")',
      ]);
    });

    let range = indexSheet.getRange(1, currentCol, updates.length);
    range.setValues(updates);
    indexSheet
      .getRange(range.getRow() + 1, currentCol, hyperlinkUpdates.length)
      .setFormulas(hyperlinkUpdates);

    // Formatting
    indexSheet
      .getRange(1, currentCol)
      .setBackground('#D3D3D3')
      .setFontSize(16)
      .setFontWeight('bold')
      .setWrap(true);

    indexSheet
      .getRange(1, currentCol, updates.length)
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    indexSheet.setColumnWidth(currentCol, 150);

    currentCol += 1; // Increment to the next category column
  });

  indexSheet.setTabColor('#FF8C00');
}

/**
 * Fetches task data from all visible sheets and organizes them by category.
 * Sheet names are expected to be formatted as "Category: Task".
 * @returns {Object} An object mapping categories to an array of task objects.
 */
function fetchTaskSheetsData_() {
  let allSheets = SPREADSHEET.getSheets();
  let taskSheetData = {};

  allSheets.forEach((sheet) => {
    if (!sheet.isSheetHidden()) {
      let name = sheet.getName();
      if (name.includes(':')) {
        let [category, task] = name.split(':').map((part) => part.trim());
        let sheetGID = sheet.getSheetId();
        let sheetURL = `${SPREADSHEET.getUrl()}#gid=${sheetGID}`;
        let taskInfo = { task: task, url: sheetURL };

        if (!taskSheetData[category]) {
          taskSheetData[category] = [];
        }
        taskSheetData[category].push(taskInfo);
      }
    }
  });

  return taskSheetData;
}
