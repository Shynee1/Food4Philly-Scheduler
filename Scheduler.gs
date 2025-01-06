// Information for recurring events
const START_DATE = new Date(2025, 0, 18, 12, 0, 0);
const END_DATE = new Date(2026, 0, 1, 12, 0, 0);
const DESCRIPTION = "Created by the Food4Philly Scheduler";
const LOCATION = "900 W Hunting Park Ave, Philadelphia, PA 19140";
// Unique IDs and ranges for accessing Food4Philly directory
const DIRECTORY_ID = "1E62O579akZotUWhNHhoeEPbD0B0EfqnjNsd6PkXKl2U";
const DIRECTORY_CHAPTERS_RANGE = "Chapters!A2:A1000";
// Unique IDs and ranges for accessing elements of Food4Philly scheduler
const DATA_SHEET_ID = 1645079220;
const DROPDOWN_CELL = "D10";
// Rankings of all current chapters
const CHAPTER_RANKINGS = new Map([
  ["Penncrest High School",            2],
  ["The Haverford School",             1],
  ["The Agnes Irwin School",          14],
  ["The Baldwin School",               9],
  ["Upper Dublin High School",         5],
  ["The Episcopal Academy",           12],
  ["Garnet Valley High School",        7],
  ["Harriton High School",             4],
  ["The Shipley School",              11],
  ["Wissahickon High School",          8],
  ["Lower Merion High School",        10],
  ["Downington STEM Academy",          6],
  ["Plymouth Whitemarsh High School", 13],
  ["Academy of Notre Dame de Namur",   3],
  ["Food4TheBay",                     -1],
  ["Food4Pitt",                       -1],
  ["Bowls4Boston",                    -1],
  ["Other",                           -1]
]);

/**
 * Runs when the Google Sheet is open
 * Hides "Data" sheet used to store chapters
 */
function onOpen(e) {
  SpreadsheetApp.getActive().getSheetById(DATA_SHEET_ID).hideSheet();
}

/**
 * Runs when "Submit" button is clicked
 * Reads chapter from dropdown and adds events to calendar
 * Alerts and exits if no chapter is selected
 */
function onSubmitClick() {
  const ui = SpreadsheetApp.getUi();
  const dropdownCell = SpreadsheetApp.getActive().getRange(DIRECTORY_DROPDOWN_CELL);

  // Refresh + regroup every time in case new chapter was added
  const chapters = refreshChapters();
  const groups = createChapterGroups(chapters, CHAPTER_RANKINGS);

  const chapter = dropdownCell.getValue();

  if (chapter == ""){
    SpreadsheetApp.getUi().alert(
      "⚠ Missing Input",
      "Please select a chapter using the dropdown before pressing the 'submit' button",
      ui.ButtonSet.OK
    );
    return;
  }

  if (chapter == "ALL CHAPTERS") 
    for (const group of groups) addCalendarEvents(group[0], groups);
  else
    addCalendarEvents(chapter, groups);

  ui.alert(
    "🎉 Success 🎉",
    `Successfully added Food4Philly events to your Google Calendar`,
    ui.ButtonSet.OK
  );
}

/**
 * Stores up-to-date chapters in hidden "Data" sheet of scheduler
 * 
 * @returns {List} list of up-to-date chapters
 */
function refreshChapters() {
  let chapters = loadChapters(DIRECTORY_ID, DIRECTORY_CHAPTERS_RANGE, CHAPTER_RANKINGS);

  // Add "All Chapters" option to the UI 
  chapters.unshift("ALL CHAPTERS");

  const dataSheet = SpreadsheetApp.getActive().getSheetById(DATA_SHEET_ID);

  // Clear previous content
  const clearRange = dataSheet.getRange(1, 1, 1000, 1);
  clearRange.clearContent();

  const chapterRange = dataSheet.getRange(1, 1, chapters.length, 1)
  // Convert to 2D array for Google Sheets fill
  const chapterValues = chapters.map(chapter => [chapter]);
  chapterRange.setValues(chapterValues);

  // Remove "All Chapters" option for chapter grouping
  chapters.shift();

  return chapters;
}

/**
 * Gathers up-to-date chapters from Food4Philly directory
 * Sorts list from "strongest" to "weakest" chapters
 * 
 * @param {String} directoryID - Unique ID for Food4Philly directory
 * @param {String} range - Unique range for accessing chapters in directory
 * @param {Map} rankings - Initial chapters ranked from strongest to weakest
 * @returns {List} sorted list of up-to-date chapters
 */
function loadChapters(directoryID, range, rankings) {
  const directory = SpreadsheetApp.openById(directoryID);
  const rangeValues = directory.getRange(range).getValues().flat();

  // Exclude unnecessary data
  const chapters = rangeValues.filter((entry) => {
    const score = rankings.get(entry);
    return entry != "" && (score == null || score >= 0);
  });

  // Sort chapters by their rankings
  // Push any unranked chapter to the back
  chapters.sort((a, b) => {
    const scoreA = rankings.get(a) ?? Infinity;
    const scoreB = rankings.get(b) ?? Infinity;
    return scoreA - scoreB;
  });

  return chapters;
}

/**
 * Groups chapters together based on rankings
 * Stronger ranking chapters are paired with weaker ranking chapters
 * New chapters (without rankings) are added to the pre-ranked groups
 * 
 * @param {List} chapters - List of all Food4Philly chapters
 * @param {Map} rankings - Initial chapters ranked from strongest to weakest
 * @returns {List} - 2D Array with groups of chapters
 */
function createChapterGroups(chapters, rankings) {

  let start = 0;
  let end = chapters.length - 1;

  let groups = [];
  let ungroupedStack = [];
  
  while (start <= end){
    const first = chapters[start];
    const last = chapters[end];

    // Group pre-ranked chapters first
    // Pair highest with lowest ranked
    if (rankings.has(first) && rankings.has(last)){
      groups.push([first, last]);
      start++;
      end--;
    }
    // Handle new chapters by storing in separate array
    else if (!rankings.has(first)) {
      ungroupedStack.push(first);
      start++;
    }
    else  {
      ungroupedStack.push(last);
      end--;
    }
  }

  // Append new chapters to pre-ranked groups
  const length = ungroupedStack.length;
  for (let i = 0; i < length; i++) {
    const index = i % groups.length;
    groups[index].push(ungroupedStack.pop());
  }

  return groups;
}

/**
 * Adds recurring calendar event for a specific chapter
 * 
 * @param {string} chapter - Chapter to add events for
 * @param {List} groups - Chapter groups
 */
function addCalendarEvents(chapter, groups){
  const calendar = CalendarApp.getDefaultCalendar();

  // Find which group the chapter belongs to
  const index = findChapterIndex(chapter, groups);

  // Offset start date by group order
  let eventStart = new Date(START_DATE);
  eventStart.setDate(eventStart.getDate() + index * 7);
  // Add two hours for event end time
  let eventEnd = new Date(eventStart);
  eventEnd.setHours(eventEnd.getHours() + 2);

  const title = `Food4Philly Distribution Event: ${groups[index].join(", ")}`;

  // Create recurring Google Calendar event with data
  calendar.createEventSeries(
    title,
    eventStart,
    eventEnd,
    CalendarApp.newRecurrence().addWeeklyRule().interval(groups.length),
    {
      description: DESCRIPTION,
      location: LOCATION 
    },
  );
}

/**
 * Finds index of group for a specific chapter
 * 
 * @param {String} chapter
 * @param {List} groups
 * @returns {int} Index of group or -1 if chapter cannot be found
 */
function findChapterIndex(chapter, groups) {
  for (let i = 0; i < groups.length; i++) {
    if (groups[i].includes(chapter))
      return i;
  }
  return -1;
}

/**
 * Deletes all events added by Food4Philly scheduler
 * Called when "Delete All Events" button is pressed
 */
function onDeleteClick() {
  const ui = SpreadsheetApp.getUi();

  let result = ui.alert(
    "⚠ Event deletion",
    "This action will delete all events created by the Food4Philly Schedule Tool. Are you sure want to continue?",
    ui.ButtonSet.YES_NO
  );

  if (result != ui.Button.YES)
    return;

  const calendar = CalendarApp.getDefaultCalendar();
  // Grab all events containing special description
  const events = calendar.getEvents(START_DATE, END_DATE, {search: DESCRIPTION});
  // Delete the events
  for (event of events)
    event.deleteEvent();

  ui.alert(
    "🎉 Success 🎉",
    "All Food4Philly events successfully deleted.",
    ui.ButtonSet.OK
  );
}

/**
 * Refreshes list of chapters in dropdown
 * Called when "Refresh Chapters" button is pressed
 */ 
function onRefreshClick() {
  const ui = SpreadsheetApp.getUi();
  refreshChapters();
  ui.alert(
    "🎉 Success 🎉",
    "Chapter list successfully refreshed. You are now viewing the most up-to-date list of Food4Philly chapters",
    ui.ButtonSet.OK
  );
}

/**
* Loads HTML instructions for using the scheduler
* Called when "Instructions" button is pressed
*/
function onInstructionsClick() {
  var widget = HtmlService.createHtmlOutputFromFile("Instructions.html");
  SpreadsheetApp.getUi().showModalDialog(widget, "Welcome to the Food4Philly Scheduling Tool!");
}
