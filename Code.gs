/*******************************************************
 * PLACEMENT ATTENDANCE SYSTEM
 *
 * SOURCE OF TRUTH:
 * ----------------
 * Attendance Responses
 *
 * IT1 / IT2 / IT3 are automatically updated from
 * Attendance Responses.
 *
 * Attendance Responses columns:
 *
 * A = Timestamp
 * B = Full Name
 * C = Roll Number
 * D = Section
 * E = Date
 * F = Company
 * G = Event
 * H = Class Hours Attended
 * I = Proof / Drive Link
 *
 *
 * IT1 / IT2 / IT3 structure:
 *
 * Row 3 = Date
 * Row 4 = Company - Event
 * Row 5 = Hours
 * Row 6 = Student headers
 * Row 7+ = Students
 *
 * Each event occupies 6 columns:
 *
 * 1st Hour | 2nd Hour | 3rd Hour |
 * 4th Hour | 5th Hour | 6th Hour
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

  // IT sheet structure
  DATE_ROW: 3,
  EVENT_ROW: 4,
  HOUR_ROW: 5,
  HEADER_ROW: 6,
  STUDENT_START_ROW: 7,

  ROLL_COLUMN: 1,
  NAME_COLUMN: 2,

  FIRST_EVENT_COLUMN: 3,

  MARK: "P"

};


// =====================================================
// WEB APP
// =====================================================

function doGet() {

  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("Placement Attendance")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


// =====================================================
// GET EVENTS
// Used by index.html dropdown
// =====================================================

function getEvents() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(CONFIG.EVENTS_SHEET);

  if (!sheet) {
    throw new Error(
      'Sheet "Events" was not found.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
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
    values[0].map(normalizeText);


  let dateColumn =
    findHeaderColumn(
      headers,
      ["date"]
    );

  let companyColumn =
    findHeaderColumn(
      headers,
      ["company", "company name"]
    );

  let eventColumn =
    findHeaderColumn(
      headers,
      ["event", "event name"]
    );


  // Fallback positions
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


  for (
    let r = 1;
    r < values.length;
    r++
  ) {

    const row = values[r];

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


    const normalizedDate =
      normalizeDate(rawDate);


    if (!normalizedDate) {
      continue;
    }


    events.push({

      date: normalizedDate,

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
// FIND HEADER COLUMN
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
// =====================================================

function submitAttendance(formData) {

  if (
    !formData ||
    typeof formData !== "object"
  ) {

    throw new Error(
      "No attendance data received."
    );

  }


  const fullName =
    String(
      formData.fullName || ""
    ).trim();

  const rollNumber =
    String(
      formData.rollNumber || ""
    ).trim();

  const section =
    String(
      formData.section || ""
    ).trim();

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


  let classHours =
    formData.hours || [];


  if (!Array.isArray(classHours)) {

    classHours =
      String(classHours)
        .split(",")
        .map(
          hour => hour.trim()
        )
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

  if (!section) {
    throw new Error(
      "Section is required."
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

  if (classHours.length === 0) {
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


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  // ---------------------------------------------------
  // DUPLICATE CHECK
  // ---------------------------------------------------

  const alreadySubmitted =
    checkDuplicateSubmission(
      ss,
      rollNumber,
      date,
      company,
      eventName
    );


  if (alreadySubmitted) {

    throw new Error(
      "Attendance for this event has already been submitted for roll number " +
      rollNumber +
      "."
    );

  }


  // ---------------------------------------------------
  // ADD RESPONSE
  // ---------------------------------------------------

  addAttendanceResponse(
    ss,
    {
      fullName: fullName,
      rollNumber: rollNumber,
      section: section,
      date: date,
      company: company,
      event: eventName,
      hours: classHours,
      proofLink: proofLink
    }
  );


  // ---------------------------------------------------
  // SYNC IT SHEETS
  // ---------------------------------------------------

  syncAttendance();


  return {

    success: true,

    message:
      "Attendance submitted successfully."

  };

}


// =====================================================
// ADD RESPONSE TO ATTENDANCE RESPONSES
// =====================================================

function addAttendanceResponse(
  ss,
  data
) {

  const sheet =
    ss.getSheetByName(
      CONFIG.RESPONSE_SHEET
    );


  if (!sheet) {

    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );

  }


  const hoursText =
    data.hours.join(", ");


  sheet.appendRow([

    new Date(),

    data.fullName,

    data.rollNumber,

    data.section,

    data.date,

    data.company,

    data.event,

    hoursText,

    data.proofLink

  ]);

}


// =====================================================
// ⭐ MAIN SYNC FUNCTION
//
// Attendance Responses → IT1 / IT2 / IT3
// =====================================================

function syncAttendance() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const responseSheet =
    ss.getSheetByName(
      CONFIG.RESPONSE_SHEET
    );


  if (!responseSheet) {

    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );

  }


  // ---------------------------------------------------
  // STEP 1
  // Clear old attendance marks
  //
  // IMPORTANT:
  // We start at row 7.
  //
  // We DO NOT touch:
  // Row 3 = Date
  // Row 4 = Event
  // Row 5 = Hours
  // Row 6 = Headers
  // ---------------------------------------------------

  CONFIG.SECTION_SHEETS.forEach(
    section => {

      const sheet =
        ss.getSheetByName(
          section
        );

      if (!sheet) {
        return;
      }


      const lastRow =
        sheet.getLastRow();

      const lastColumn =
        sheet.getLastColumn();


      if (
        lastRow < CONFIG.STUDENT_START_ROW ||
        lastColumn < CONFIG.FIRST_EVENT_COLUMN
      ) {

        return;

      }


      const numberOfRows =
        lastRow -
        CONFIG.STUDENT_START_ROW +
        1;


      const numberOfColumns =
        lastColumn -
        CONFIG.FIRST_EVENT_COLUMN +
        1;


      sheet
        .getRange(
          CONFIG.STUDENT_START_ROW,
          CONFIG.FIRST_EVENT_COLUMN,
          numberOfRows,
          numberOfColumns
        )
        .clearContent()
        .clearNote();

    }
  );


  // ---------------------------------------------------
  // STEP 2
  // Read Attendance Responses
  // ---------------------------------------------------

  const lastResponseRow =
    responseSheet.getLastRow();


  if (lastResponseRow < 2) {
    return;
  }


  const responses =
    responseSheet
      .getRange(
        2,
        1,
        lastResponseRow - 1,
        9
      )
      .getValues();


  // ---------------------------------------------------
  // STEP 3
  // Process every response
  // ---------------------------------------------------

  responses.forEach(
    row => {

      const fullName =
        String(
          row[1] || ""
        ).trim();

      const roll =
        String(
          row[2] || ""
        ).trim();

      const section =
        String(
          row[3] || ""
        ).trim()
        .toUpperCase();

      const date =
        normalizeDate(
          row[4]
        );

      const company =
        String(
          row[5] || ""
        ).trim();

      const event =
        String(
          row[6] || ""
        ).trim();

      const hours =
        String(
          row[7] || ""
        ).trim();

      const proof =
        String(
          row[8] || ""
        ).trim();


      // Ignore incomplete responses
      if (
        !roll ||
        !section ||
        !date ||
        !company ||
        !event ||
        !hours
      ) {

        return;

      }


      // -------------------------------------------------
      // FIND SECTION SHEET
      // -------------------------------------------------

      const sheet =
        ss.getSheetByName(
          section
        );


      if (!sheet) {
        return;
      }


      // -------------------------------------------------
      // FIND STUDENT
      // -------------------------------------------------

      const studentCount =
        sheet.getLastRow() -
        CONFIG.STUDENT_START_ROW +
        1;


      if (studentCount <= 0) {
        return;
      }


      const students =
        sheet
          .getRange(
            CONFIG.STUDENT_START_ROW,
            CONFIG.ROLL_COLUMN,
            studentCount,
            1
          )
          .getDisplayValues();


      let studentRow =
        -1;


      for (
        let i = 0;
        i < students.length;
        i++
      ) {

        if (
          normalizeRollNumber(
            students[i][0]
          ) ===
          normalizeRollNumber(
            roll
          )
        ) {

          studentRow =
            CONFIG.STUDENT_START_ROW +
            i;

          break;

        }

      }


      if (studentRow === -1) {
        return;
      }


      // -------------------------------------------------
      // FIND EVENT
      // -------------------------------------------------

      const eventCount =
        sheet.getLastColumn() -
        CONFIG.FIRST_EVENT_COLUMN +
        1;


      if (eventCount <= 0) {
        return;
      }


      const dates =
        sheet
          .getRange(
            CONFIG.DATE_ROW,
            CONFIG.FIRST_EVENT_COLUMN,
            1,
            eventCount
          )
          .getDisplayValues()[0];


      const eventNames =
        sheet
          .getRange(
            CONFIG.EVENT_ROW,
            CONFIG.FIRST_EVENT_COLUMN,
            1,
            eventCount
          )
          .getDisplayValues()[0];


      let eventColumn =
        -1;


      const targetCompany =
        normalizeText(
          company
        );

      const targetEvent =
        normalizeText(
          event
        );


      for (
        let i = 0;
        i < eventCount;
        i++
      ) {

        const sheetDate =
          normalizeDate(
            dates[i]
          );


        if (
          sheetDate !== date
        ) {

          continue;

        }


        const sheetEvent =
          normalizeText(
            eventNames[i]
          );


        if (!sheetEvent) {
          continue;
        }


        /*
         * Event header normally looks like:
         *
         * Cognizant NPN Salesforce -
         * Virtual Session
         *
         * We check that both the company and
         * event are present.
         */

        if (
          sheetEvent.includes(
            targetCompany
          ) &&
          sheetEvent.includes(
            targetEvent
          )
        ) {

          eventColumn =
            CONFIG.FIRST_EVENT_COLUMN +
            i;

          break;

        }

      }


      if (eventColumn === -1) {
        return;
      }


      // -------------------------------------------------
      // MARK SELECTED HOURS
      // -------------------------------------------------

      const hourOffsets = {

        "1st Hour": 0,
        "2nd Hour": 1,
        "3rd Hour": 2,
        "4th Hour": 3,
        "5th Hour": 4,
        "6th Hour": 5

      };


      hours
        .split(",")
        .map(
          hour => hour.trim()
        )
        .forEach(
          hour => {

            if (
              !Object.prototype
                .hasOwnProperty
                .call(
                  hourOffsets,
                  hour
                )
            ) {

              return;

            }


            const column =
              eventColumn +
              hourOffsets[hour];


            const cell =
              sheet.getRange(
                studentRow,
                column
              );


            cell.setValue(
              CONFIG.MARK
            );


            // Make P clickable
            if (proof) {

              cell.setRichTextValue(

                SpreadsheetApp
                  .newRichTextValue()
                  .setText(
                    CONFIG.MARK
                  )
                  .setLinkUrl(
                    proof
                  )
                  .build()

              );


              cell.setNote(
                "Proof:\n" +
                proof
              );

            }


            cell.setHorizontalAlignment(
              "center"
            );

          }
        );

    }
  );

}


// =====================================================
// CHECK DUPLICATE RESPONSE
// =====================================================

function checkDuplicateSubmission(
  ss,
  rollNumber,
  date,
  company,
  eventName
) {

  const sheet =
    ss.getSheetByName(
      CONFIG.RESPONSE_SHEET
    );


  if (!sheet) {

    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return false;
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        9
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
        values[i][2]
      );

    const existingDate =
      normalizeDate(
        values[i][4]
      );

    const existingCompany =
      normalizeText(
        values[i][5]
      );

    const existingEvent =
      normalizeText(
        values[i][6]
      );


    if (

      existingRoll ===
      targetRoll &&

      existingDate ===
      targetDate &&

      existingCompany ===
      targetCompany &&

      existingEvent ===
      targetEvent

    ) {

      return true;

    }

  }


  return false;

}


// =====================================================
// NORMALIZE ROLL NUMBER
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


  // Actual Google Sheets Date
  if (
    Object.prototype.toString
      .call(value) ===
    "[object Date]"
  ) {

    if (
      isNaN(
        value.getTime()
      )
    ) {

      return "";

    }


    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  }


  const text =
    String(value).trim();


  // yyyy-mm-dd
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


  // dd.mm.yyyy
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


  // dd/mm/yyyy
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


  // Last attempt
  const parsed =
    new Date(text);


  if (
    !isNaN(
      parsed.getTime()
    )
  ) {

    return Utilities.formatDate(
      parsed,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  }


  return "";

}


// =====================================================
// PAD NUMBER
// =====================================================

function pad(value) {

  return String(
    value
  ).padStart(
    2,
    "0"
  );

}


// =====================================================
// MANUAL SYNC
//
// Useful if you ever want to force a sync.
// =====================================================

function manualSync() {

  syncAttendance();

  SpreadsheetApp
    .getUi()
    .alert(
      "✅ Attendance Synced",
      "IT1, IT2 and IT3 have been updated from Attendance Responses.",
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
      "🔄 Attendance Sync"
    )

    .addItem(
      "Sync Attendance Now",
      "manualSync"
    )

    .addItem(
      "Setup Auto Sync",
      "setupAutoSync"
    )

    .addToUi();

}


// =====================================================
// SETUP AUTO SYNC
//
// Run this ONE TIME manually.
// =====================================================

function setupAutoSync() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  // Remove existing triggers created
  // by this system.

  ScriptApp
    .getProjectTriggers()
    .forEach(
      trigger => {

        const functionName =
          trigger.getHandlerFunction();


        if (
          functionName ===
          "attendanceChangeTrigger"
        ) {

          ScriptApp.deleteTrigger(
            trigger
          );

        }

      }
    );


  // Create ONE installable
  // spreadsheet change trigger.

  ScriptApp
    .newTrigger(
      "attendanceChangeTrigger"
    )
    .forSpreadsheet(ss)
    .onChange()
    .create();


  SpreadsheetApp
    .getUi()
    .alert(

      "✅ Auto Sync Enabled",

      "Attendance Responses is now the source of truth.\n\n" +
      "IT1, IT2 and IT3 will automatically sync when spreadsheet changes occur.",

      SpreadsheetApp
        .getUi()
        .ButtonSet.OK

    );

}


// =====================================================
// AUTO SYNC TRIGGER
// =====================================================

function attendanceChangeTrigger(e) {

  try {

    syncAttendance();

  } catch (error) {

    console.error(
      "Attendance sync error: " +
      error.message
    );

  }

}