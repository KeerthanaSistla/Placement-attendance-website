/*******************************************************
 * PLACEMENT ATTENDANCE SYSTEM
 *
 * SOURCE OF TRUTH:
 *     Attendance Responses
 *
 * MIRROR:
 *     IT1
 *     IT2
 *     IT3
 *
 * IMPORTANT:
 * Submission only writes to Attendance Responses.
 * The IT sheets are updated separately.
 *
 * DYNAMIC EVENTS:
 *     Changes to the Events sheet automatically rebuild
 *     IT1 / IT2 / IT3 so attendance follows the event
 *     even when events are moved or new events are added.
 *******************************************************/


// =====================================================
// CONFIGURATION
// =====================================================

const CONFIG = {

  RESPONSE_SHEET: "Attendance Responses",

  EVENTS_SHEET: "Events",

  SECTION_SHEETS: [
    "IT1",
    "IT2",
    "IT3"
  ],

  DATE_ROW: 3,

  EVENT_ROW: 4,

  HOUR_ROW: 5,

  HEADER_ROW: 6,

  ROLL_COLUMN: 1,

  NAME_COLUMN: 2,

  FIRST_EVENT_COLUMN: 3,

  HOURS_PER_EVENT: 6,

  MARK: "P"

};


// =====================================================
// WEB APP
// =====================================================

function doGet() {

  return HtmlService

    .createHtmlOutputFromFile("index")

    .setTitle(
      "Placement Attendance"
    )

    // IMPORTANT FOR MOBILE RESPONSIVENESS
    // Apps Script ignores a viewport meta tag placed
    // directly inside the HTML file.

    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1"
    )

    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


// =====================================================
// GET EVENTS
// =====================================================

function getEvents() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.EVENTS_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Sheet "Events" was not found.'
    );

  }


  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();


  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {

    return [];

  }


  const values =
    sheet
      .getRange(
        1,
        1,
        lastRow,
        lastColumn
      )
      .getDisplayValues();


  const headers =
    values[0].map(
      normalizeText
    );


  let dateColumn =
    findHeaderColumn(
      headers,
      ["date"]
    );


  let companyColumn =
    findHeaderColumn(
      headers,
      [
        "company",
        "company name"
      ]
    );


  let eventColumn =
    findHeaderColumn(
      headers,
      [
        "event",
        "event name"
      ]
    );


  // ---------------------------------------------------
  // FALLBACK COLUMNS
  // ---------------------------------------------------

  if (dateColumn === -1) {
    dateColumn = 0;
  }

  if (companyColumn === -1) {
    companyColumn = 1;
  }

  if (eventColumn === -1) {
    eventColumn = 2;
  }


  const events = [];


  // ---------------------------------------------------
  // READ EVENTS
  // ---------------------------------------------------

  for (
    let r = 1;
    r < values.length;
    r++
  ) {

    const row =
      values[r];


    const rawDate =
      row[dateColumn];


    const company =
      String(
        row[companyColumn] || ""
      ).trim();


    const event =
      String(
        row[eventColumn] || ""
      ).trim();


    if (
      !rawDate ||
      !company ||
      !event
    ) {

      continue;

    }


    const date =
      normalizeDate(
        rawDate
      );


    if (!date) {

      continue;

    }


    events.push({

      date: date,

      company: company,

      event: event,

      display:
        company +
        " - " +
        event

    });

  }


  return events;

}


// =====================================================
// FIND HEADER
// =====================================================

function findHeaderColumn(
  headers,
  possibleNames
) {

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    for (
      let j = 0;
      j < possibleNames.length;
      j++
    ) {

      if (
        headers[i] ===
        normalizeText(
          possibleNames[j]
        )
      ) {

        return i;

      }

    }

  }


  return -1;

}


// =====================================================
// SUBMIT ATTENDANCE
//
// IMPORTANT:
// This function DOES NOT update IT1/IT2/IT3.
//
// It only:
//   1. validates
//   2. verifies roll number belongs to section
//   3. checks duplicate
//   4. writes one row to Attendance Responses
//
// This keeps submission extremely fast.
// =====================================================

function submitAttendance(
  formData
) {

  if (
    !formData ||
    typeof formData !== "object"
  ) {

    throw new Error(
      "No attendance data received."
    );

  }


  // ---------------------------------------------------
  // READ FORM DATA
  // ---------------------------------------------------

  const fullName =
    String(
      formData.fullName || ""
    ).trim();


  const rollNumber =
    String(
      formData.rollNumber || ""
    ).trim();


  const section =
    normalizeSection(
      formData.section
    );


  const date =
    String(
      formData.date || ""
    ).trim();


  const company =
    String(
      formData.company || ""
    ).trim();


  const eventName =
    String(
      formData.event || ""
    ).trim();


  const proofLink =
    String(
      formData.proofLink || ""
    ).trim();


  let hours =
    formData.hours || [];


  if (
    !Array.isArray(hours)
  ) {

    hours =
      String(hours)
        .split(",")
        .map(function(h) {

          return String(h).trim();

        })
        .filter(Boolean);

  }


  // ---------------------------------------------------
  // VALIDATION
  // ---------------------------------------------------

  if (!fullName) {

    throw new Error(
      "Full Name is required."
    );

  }


  if (!rollNumber) {

    throw new Error(
      "Roll Number is required."
    );

  }


  if (
    ![
      "IT1",
      "IT2",
      "IT3"
    ].includes(section)
  ) {

    throw new Error(
      "Please select a valid section."
    );

  }


  if (!date) {

    throw new Error(
      "Date is required."
    );

  }


  if (!company) {

    throw new Error(
      "Company is required."
    );

  }


  if (!eventName) {

    throw new Error(
      "Event is required."
    );

  }


  if (
    hours.length === 0
  ) {

    throw new Error(
      "Please select at least one class hour."
    );

  }


  if (!proofLink) {

    throw new Error(
      "Proof link is required."
    );

  }


  if (
    !/^https?:\/\//i.test(
      proofLink
    )
  ) {

    throw new Error(
      "Please enter a valid proof link."
    );

  }


  // ---------------------------------------------------
  // SPREADSHEET
  // ---------------------------------------------------

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const responseSheet =
    ss.getSheetByName(
      CONFIG.RESPONSE_SHEET
    );


  if (!responseSheet) {

    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );

  }


  // ===================================================
  // VERIFY ROLL NUMBER BELONGS TO SECTION
  // ===================================================

  const sectionSheet =
    ss.getSheetByName(
      section
    );


  if (!sectionSheet) {

    throw new Error(
      "Section sheet not found: " +
      section
    );

  }


  const studentRow =
    findStudentRow(
      sectionSheet,
      rollNumber
    );


  if (
    studentRow === -1
  ) {

    throw new Error(
      "Roll number " +
      rollNumber +
      " is not registered in " +
      section +
      ". Please check your roll number and section."
    );

  }


  // ===================================================
  // DUPLICATE CHECK
  // ===================================================

  if (
    checkDuplicateSubmission(
      responseSheet,
      rollNumber,
      date,
      company,
      eventName
    )
  ) {

    throw new Error(
      "Attendance for this event has already been submitted for roll number " +
      rollNumber +
      "."
    );

  }


  // ===================================================
  // ADD RESPONSE
  //
  // ONE SPREADSHEET WRITE
  // ===================================================

  const hoursText =
    hours.join(", ");


  const newRow =
    responseSheet.getLastRow() + 1;


  responseSheet
    .getRange(
      newRow,
      1,
      1,
      9
    )
    .setValues([[
      new Date(),
      fullName,
      rollNumber,
      section,
      date,
      company,
      eventName,
      hoursText,
      proofLink
    ]]);


  // ===================================================
  // RETURN IMMEDIATELY
  // ===================================================

  return {

    success: true,

    row: newRow,

    message:
      "Attendance submitted successfully."

  };

}


// =====================================================
// CHECK DUPLICATE
// =====================================================

function checkDuplicateSubmission(
  sheet,
  rollNumber,
  date,
  company,
  eventName
) {

  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return false;

  }


  // ---------------------------------------------------
  // ONLY READ C:G
  //
  // C = Roll
  // D = Section
  // E = Date
  // F = Company
  // G = Event
  // ---------------------------------------------------

  const values =
    sheet
      .getRange(
        2,
        3,
        lastRow - 1,
        5
      )
      .getDisplayValues();


  const targetRoll =
    normalizeRollNumber(
      rollNumber
    );


  const targetDate =
    normalizeDate(
      date
    );


  const targetCompany =
    normalizeText(
      company
    );


  const targetEvent =
    normalizeText(
      eventName
    );


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    const existingRoll =
      normalizeRollNumber(
        values[i][0]
      );


    const existingDate =
      normalizeDate(
        values[i][2]
      );


    const existingCompany =
      normalizeText(
        values[i][3]
      );


    const existingEvent =
      normalizeText(
        values[i][4]
      );


    if (
      existingRoll === targetRoll &&
      existingDate === targetDate &&
      existingCompany === targetCompany &&
      existingEvent === targetEvent
    ) {

      return true;

    }

  }


  return false;

}


// =====================================================
// BACKGROUND SYNC
//
// Called by the HTML after submission succeeds.
//
// This is deliberately separate from submitAttendance()
// so the user does not wait for spreadsheet formatting.
// =====================================================

function syncResponseRow(
  rowNumber
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const responseSheet =
    ss.getSheetByName(
      CONFIG.RESPONSE_SHEET
    );


  if (!responseSheet) {

    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );

  }


  if (
    !rowNumber ||
    rowNumber < 2 ||
    rowNumber >
      responseSheet.getLastRow()
  ) {

    return;

  }


  const row =
    responseSheet
      .getRange(
        rowNumber,
        1,
        1,
        9
      )
      .getDisplayValues()[0];


  const fullName =
    row[1];


  const rollNumber =
    row[2];


  const section =
    row[3];


  const date =
    row[4];


  const company =
    row[5];


  const eventName =
    row[6];


  const hoursText =
    row[7];


  const proofLink =
    row[8];


  if (
    !rollNumber ||
    !section ||
    !date ||
    !company ||
    !eventName ||
    !hoursText
  ) {

    return;

  }


  const hours =
    hoursText
      .split(",")
      .map(function(h) {

        return h.trim();

      })
      .filter(Boolean);


  mirrorAttendance(

    ss,

    rollNumber,

    section,

    date,

    company,

    eventName,

    hours,

    proofLink

  );

}


// =====================================================
// MIRROR ONE ATTENDANCE RECORD
// =====================================================

function mirrorAttendance(
  ss,
  rollNumber,
  section,
  date,
  company,
  eventName,
  hours,
  proofLink
) {

  const sheetName =
    normalizeSection(
      section
    );


  const sheet =
    ss.getSheetByName(
      sheetName
    );


  if (!sheet) {

    throw new Error(
      "Sheet not found: " +
      sheetName
    );

  }


  // ---------------------------------------------------
  // FIND STUDENT
  // ---------------------------------------------------

  const studentRow =
    findStudentRow(
      sheet,
      rollNumber
    );


  if (
    studentRow === -1
  ) {

    throw new Error(
      "Student " +
      rollNumber +
      " was not found in " +
      sheetName
    );

  }


  // ---------------------------------------------------
  // FIND CURRENT EVENT COLUMN
  //
  // IMPORTANT:
  // This searches the CURRENT layout of the sheet.
  // Therefore attendance follows the event even when
  // the Events sheet is rearranged.
  // ---------------------------------------------------

  const eventColumn =
    findEventColumn(
      sheet,
      date,
      company,
      eventName
    );


  if (
    eventColumn === -1
  ) {

    throw new Error(
      "Event not found in " +
      sheetName +
      ": " +
      company +
      " - " +
      eventName
    );

  }


  // ---------------------------------------------------
  // WRITE ATTENDANCE
  // ---------------------------------------------------

  writeAttendanceMarks(

    sheet,

    studentRow,

    eventColumn,

    hours,

    proofLink

  );

}


// =====================================================
// WRITE ATTENDANCE
//
// All 6 hours are written in one batch instead of
// doing setValue/setNote/setRichText repeatedly.
// =====================================================

function writeAttendanceMarks(
  sheet,
  studentRow,
  eventColumn,
  hours,
  proofLink
) {

  const hourMap = {

    "1st Hour": 0,

    "2nd Hour": 1,

    "3rd Hour": 2,

    "4th Hour": 3,

    "5th Hour": 4,

    "6th Hour": 5

  };


  const values = [
    [
      "",
      "",
      "",
      "",
      "",
      ""
    ]
  ];


  const richText = [
    []
  ];


  const notes = [
    [
      "",
      "",
      "",
      "",
      "",
      ""
    ]
  ];


  // ---------------------------------------------------
  // INITIALIZE RICH TEXT
  // ---------------------------------------------------

  for (
    let i = 0;
    i < 6;
    i++
  ) {

    richText[0][i] =
      SpreadsheetApp
        .newRichTextValue()
        .setText("")
        .build();

  }


  // ---------------------------------------------------
  // APPLY MARKS
  // ---------------------------------------------------

  hours.forEach(function(hour) {

    if (
      !Object.prototype.hasOwnProperty.call(
        hourMap,
        hour
      )
    ) {

      return;

    }


    const offset =
      hourMap[hour];


    values[0][offset] =
      CONFIG.MARK;


    richText[0][offset] =
      SpreadsheetApp
        .newRichTextValue()
        .setText(
          CONFIG.MARK
        )
        .setLinkUrl(
          proofLink
        )
        .build();


    notes[0][offset] =
      "Class Hour: " +
      hour +
      "\n\nProof:\n" +
      proofLink;

  });


  // ---------------------------------------------------
  // RANGE
  // ---------------------------------------------------

  const range =
    sheet.getRange(
      studentRow,
      eventColumn,
      1,
      6
    );


  // One write for values

  range.setValues(
    values
  );


  // One write for rich text

  range.setRichTextValues(
    richText
  );


  // One write for notes

  range.setNotes(
    notes
  );


  range.setHorizontalAlignment(
    "center"
  );

}


// =====================================================
// REBUILD EVERYTHING
//
// This is the important function for dynamic Events.
//
// It:
//   1. Clears attendance from IT1/IT2/IT3
//   2. Reads every Attendance Responses row
//   3. Finds the CURRENT event position
//   4. Recreates every P
//
// Therefore if an event moves, its attendance moves
// with it.
// =====================================================

function rebuildAllAttendanceSheets() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const responseSheet =
    ss.getSheetByName(
      CONFIG.RESPONSE_SHEET
    );


  if (!responseSheet) {

    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );

  }


  // ===================================================
  // CLEAR ATTENDANCE AREA
  // ===================================================

  CONFIG.SECTION_SHEETS.forEach(
    function(sheetName) {

      const sheet =
        ss.getSheetByName(
          sheetName
        );


      if (!sheet) {

        return;

      }


      const lastRow =
        sheet.getLastRow();


      const lastColumn =
        sheet.getLastColumn();


      if (
        lastRow <
        CONFIG.HEADER_ROW ||
        lastColumn <
        CONFIG.FIRST_EVENT_COLUMN
      ) {

        return;

      }


      const range =
        sheet.getRange(

          CONFIG.HEADER_ROW,

          CONFIG.FIRST_EVENT_COLUMN,

          lastRow -
            CONFIG.HEADER_ROW +
            1,

          lastColumn -
            CONFIG.FIRST_EVENT_COLUMN +
            1

        );


      // ------------------------------------------------
      // ONLY REMOVE ATTENDANCE DATA
      //
      // Formatting remains untouched.
      // ------------------------------------------------

      range.clearContent();

      range.clearNote();

    }
  );


  // ===================================================
  // READ RESPONSES
  // ===================================================

  const lastRow =
    responseSheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return;

  }


  const responses =
    responseSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        9
      )
      .getDisplayValues();


  // ===================================================
  // MIRROR EVERY RESPONSE
  // ===================================================

  responses.forEach(
    function(row) {

      const rollNumber =
        row[2];


      const section =
        row[3];


      const date =
        row[4];


      const company =
        row[5];


      const eventName =
        row[6];


      const hoursText =
        row[7];


      const proofLink =
        row[8];


      if (
        !rollNumber ||
        !section ||
        !date ||
        !company ||
        !eventName ||
        !hoursText
      ) {

        return;

      }


      const hours =
        hoursText
          .split(",")
          .map(function(h) {

            return h.trim();

          })
          .filter(Boolean);


      try {

        mirrorAttendance(

          ss,

          rollNumber,

          section,

          date,

          company,

          eventName,

          hours,

          proofLink

        );

      } catch (error) {

        console.log(

          "Could not mirror " +
          rollNumber +
          ": " +
          error.message

        );

      }

    }
  );

}


// =====================================================
// MANUAL SYNC
// =====================================================

function manualSync() {

  rebuildAllAttendanceSheets();


  SpreadsheetApp
    .getUi()
    .alert(

      "Attendance Sync",

      "IT1, IT2 and IT3 have been synced from Attendance Responses.",

      SpreadsheetApp
        .getUi()
        .ButtonSet.OK

    );

}


// =====================================================
// AUTO SYNC - EDIT
//
// Handles manual edits in:
//
//   Attendance Responses
//   Events
//
// This is what makes Events dynamic.
// =====================================================

function onAttendanceSheetEdit(e) {

  try {

    if (
      !e ||
      !e.range
    ) {

      return;

    }


    const sheet =
      e.range.getSheet();


    const sheetName =
      sheet.getName();


    // -------------------------------------------------
    // ONLY WATCH THESE TWO SHEETS
    // -------------------------------------------------

    if (
      sheetName !==
        CONFIG.RESPONSE_SHEET &&
      sheetName !==
        CONFIG.EVENTS_SHEET
    ) {

      return;

    }


    rebuildAllAttendanceSheets();

  } catch (error) {

    console.log(

      "Attendance edit sync error: " +
      error.message

    );

  }

}


// =====================================================
// AUTO SYNC - STRUCTURAL CHANGES
//
// Handles changes such as:
//
//   REMOVE_ROW
//   INSERT_ROW
//   REMOVE_COLUMN
//   INSERT_COLUMN
//
// Useful when Events are inserted/deleted/moved
// structurally.
// =====================================================

function onAttendanceChange(e) {

  try {

    if (!e) {

      return;

    }


    const changeTypes = [

      "REMOVE_ROW",

      "INSERT_ROW",

      "REMOVE_COLUMN",

      "INSERT_COLUMN"

    ];


    if (
      changeTypes.includes(
        e.changeType
      )
    ) {

      rebuildAllAttendanceSheets();

    }

  } catch (error) {

    console.log(

      "Attendance change sync error: " +
      error.message

    );

  }

}


// =====================================================
// SETUP TRIGGERS
//
// Run this ONCE manually from Apps Script.
//
// This deletes old versions of the attendance triggers
// and creates the new ones.
// =====================================================

function setupAttendanceTriggers() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  // ===================================================
  // DELETE OLD VERSIONS OF OUR TRIGGERS
  // ===================================================

  ScriptApp
    .getProjectTriggers()
    .forEach(
      function(trigger) {

        const handler =
          trigger.getHandlerFunction();


        if (

          handler ===
            "onResponseEdit" ||

          handler ===
            "onResponseChange" ||

          handler ===
            "onAttendanceSheetEdit" ||

          handler ===
            "onAttendanceChange"

        ) {

          ScriptApp
            .deleteTrigger(
              trigger
            );

        }

      }
    );


  // ===================================================
  // EDIT TRIGGER
  //
  // Handles:
  //   Attendance Responses edits
  //   Events edits
  // ===================================================

  ScriptApp

    .newTrigger(
      "onAttendanceSheetEdit"
    )

    .forSpreadsheet(ss)

    .onEdit()

    .create();


  // ===================================================
  // CHANGE TRIGGER
  //
  // Handles structural changes.
  // ===================================================

  ScriptApp

    .newTrigger(
      "onAttendanceChange"
    )

    .forSpreadsheet(ss)

    .onChange()

    .create();


  // ===================================================
  // CONFIRMATION
  // ===================================================

  SpreadsheetApp
    .getUi()
    .alert(

      "Attendance Sync",

      "Automatic response and event syncing has been enabled.",

      SpreadsheetApp
        .getUi()
        .ButtonSet.OK

    );

}


// =====================================================
// MENU
// =====================================================

function onOpen() {

  SpreadsheetApp
    .getUi()

    .createMenu(
      "🔄 Attendance"
    )

    .addItem(
      "Sync Attendance Now",
      "manualSync"
    )

    .addItem(
      "Setup Auto-Sync",
      "setupAttendanceTriggers"
    )

    .addToUi();

}


// =====================================================
// FIND STUDENT
// =====================================================

function findStudentRow(
  sheet,
  rollNumber
) {

  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <
    CONFIG.HEADER_ROW
  ) {

    return -1;

  }


  const values =
    sheet
      .getRange(

        CONFIG.HEADER_ROW,

        CONFIG.ROLL_COLUMN,

        lastRow -
          CONFIG.HEADER_ROW +
          1,

        1

      )
      .getDisplayValues();


  const target =
    normalizeRollNumber(
      rollNumber
    );


  for (
    let i = 0;
    i < values.length;
    i++
  ) {

    if (

      normalizeRollNumber(
        values[i][0]
      ) === target

    ) {

      return (
        CONFIG.HEADER_ROW +
        i
      );

    }

  }


  return -1;

}


// =====================================================
// FIND EVENT
//
// Searches the CURRENT IT sheet layout.
//
// This is what allows attendance to follow an event
// when the Events sheet is rearranged.
// =====================================================

function findEventColumn(
  sheet,
  formDate,
  company,
  eventName
) {

  const lastColumn =
    sheet.getLastColumn();


  if (
    lastColumn <
    CONFIG.FIRST_EVENT_COLUMN
  ) {

    return -1;

  }


  const numberOfColumns =
    lastColumn -
    CONFIG.FIRST_EVENT_COLUMN +
    1;


  // ---------------------------------------------------
  // DATE ROW
  // ---------------------------------------------------

  const dateValues =
    sheet
      .getRange(

        CONFIG.DATE_ROW,

        CONFIG.FIRST_EVENT_COLUMN,

        1,

        numberOfColumns

      )
      .getDisplayValues()[0];


  // ---------------------------------------------------
  // EVENT ROW
  // ---------------------------------------------------

  const eventValues =
    sheet
      .getRange(

        CONFIG.EVENT_ROW,

        CONFIG.FIRST_EVENT_COLUMN,

        1,

        numberOfColumns

      )
      .getDisplayValues()[0];


  const targetDate =
    normalizeDate(
      formDate
    );


  const targetCompany =
    normalizeText(
      company
    );


  const targetEvent =
    normalizeText(
      eventName
    );


  // ===================================================
  // SEARCH
  // ===================================================

  for (
    let i = 0;
    i < numberOfColumns;
    i++
  ) {

    const sheetDate =
      normalizeDate(
        dateValues[i]
      );


    const sheetEvent =
      normalizeText(
        eventValues[i]
      );


    if (!sheetEvent) {

      continue;

    }


    // -------------------------------------------------
    // DATE MUST MATCH WHEN BOTH DATES EXIST
    // -------------------------------------------------

    if (

      targetDate &&
      sheetDate &&
      targetDate !== sheetDate

    ) {

      continue;

    }


    const companyEvent =
      normalizeText(

        targetCompany +
        " - " +
        targetEvent

      );


    const targetCombined =
      normalizeText(

        company +
        eventName

      );


    // -------------------------------------------------
    // EXACT MATCH
    // -------------------------------------------------

    if (
      sheetEvent ===
      companyEvent
    ) {

      return (
        CONFIG.FIRST_EVENT_COLUMN +
        i
      );

    }


    // -------------------------------------------------
    // COMPANY + EVENT
    // -------------------------------------------------

    if (

      sheetEvent.includes(
        targetCompany
      ) &&

      sheetEvent.includes(
        targetEvent
      )

    ) {

      return (
        CONFIG.FIRST_EVENT_COLUMN +
        i
      );

    }


    // -------------------------------------------------
    // EVENT ONLY
    // -------------------------------------------------

    if (
      sheetEvent ===
      targetEvent
    ) {

      return (
        CONFIG.FIRST_EVENT_COLUMN +
        i
      );

    }


    // -------------------------------------------------
    // COMPANY + EVENT WITHOUT SEPARATOR
    // -------------------------------------------------

    if (
      sheetEvent.includes(
        targetCombined
      )
    ) {

      return (
        CONFIG.FIRST_EVENT_COLUMN +
        i
      );

    }

  }


  return -1;

}


// =====================================================
// NORMALIZE SECTION
// =====================================================

function normalizeSection(
  section
) {

  const value =
    String(
      section || ""
    )
    .trim()
    .toUpperCase()
    .replace(
      /-/g,
      ""
    );


  return value;

}


// =====================================================
// NORMALIZE ROLL
// =====================================================

function normalizeRollNumber(
  value
) {

  return String(
    value || ""
  )
  .trim()
  .replace(
    /\s+/g,
    ""
  )
  .toUpperCase();

}


// =====================================================
// NORMALIZE TEXT
// =====================================================

function normalizeText(
  value
) {

  return String(
    value || ""
  )
  .toLowerCase()
  .trim()
  .replace(
    /[–—−]/g,
    "-"
  )
  .replace(
    /\s+/g,
    " "
  )
  .replace(
    /\s*-\s*/g,
    "-"
  )
  .replace(
    /[.,]/g,
    "");

}


// =====================================================
// NORMALIZE DATE
// =====================================================

function normalizeDate(
  value
) {

  if (!value) {

    return "";

  }


  // ---------------------------------------------------
  // REAL DATE OBJECT
  // ---------------------------------------------------

  if (

    Object.prototype.toString.call(
      value
    ) ===
    "[object Date]"

  ) {

    if (
      isNaN(
        value.getTime()
      )
    ) {

      return "";

    }


    return Utilities
      .formatDate(

        value,

        Session
          .getScriptTimeZone(),

        "yyyy-MM-dd"

      );

  }


  const text =
    String(
      value
    ).trim();


  // ---------------------------------------------------
  // YYYY-MM-DD
  // ---------------------------------------------------

  let match =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})/
    );


  if (match) {

    return (

      match[1] +
      "-" +
      pad(match[2]) +
      "-" +
      pad(match[3])

    );

  }


  // ---------------------------------------------------
  // DD.MM.YYYY
  // ---------------------------------------------------

  match =
    text.match(
      /^(\d{1,2})\.(\d{1,2})\.(\d{4})/
    );


  if (match) {

    return (

      match[3] +
      "-" +
      pad(match[2]) +
      "-" +
      pad(match[1])

    );

  }


  // ---------------------------------------------------
  // DD/MM/YYYY
  // ---------------------------------------------------

  match =
    text.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})/
    );


  if (match) {

    return (

      match[3] +
      "-" +
      pad(match[2]) +
      "-" +
      pad(match[1])

    );

  }


  // ---------------------------------------------------
  // FALLBACK DATE PARSER
  // ---------------------------------------------------

  const parsed =
    new Date(
      text
    );


  if (
    !isNaN(
      parsed.getTime()
    )
  ) {

    return Utilities
      .formatDate(

        parsed,

        Session
          .getScriptTimeZone(),

        "yyyy-MM-dd"

      );

  }


  return "";

}


// =====================================================
// PAD
// =====================================================

function pad(
  value
) {

  return String(
    value
  )
  .padStart(
    2,
    "0"
  );

}