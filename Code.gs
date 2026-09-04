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
 *******************************************************/


// =====================================================
// CONFIGURATION
// =====================================================

const CONFIG = {
  RESPONSE_SHEET: "Attendance Responses",
  EVENTS_SHEET: "Events",

  SECTION_SHEETS: ["IT1", "IT2", "IT3"],

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
    .setTitle("Placement Attendance")
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName(CONFIG.EVENTS_SHEET);

  if (!sheet) {
    throw new Error('Sheet "Events" was not found.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  const values = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getDisplayValues();

  const headers = values[0].map(normalizeText);

  let dateColumn = findHeaderColumn(headers, ["date"]);
  let companyColumn = findHeaderColumn(
    headers,
    ["company", "company name"]
  );
  let eventColumn = findHeaderColumn(
    headers,
    ["event", "event name"]
  );

  // Fallback
  if (dateColumn === -1) dateColumn = 0;
  if (companyColumn === -1) companyColumn = 1;
  if (eventColumn === -1) eventColumn = 2;

  const events = [];

  for (let r = 1; r < values.length; r++) {

    const row = values[r];

    const rawDate = row[dateColumn];

    const company = String(
      row[companyColumn] || ""
    ).trim();

    const event = String(
      row[eventColumn] || ""
    ).trim();

    if (!rawDate || !company || !event) {
      continue;
    }

    const date = normalizeDate(rawDate);

    if (!date) {
      continue;
    }

    events.push({
      date: date,
      company: company,
      event: event,
      display: company + " - " + event
    });
  }

  return events;
}


// =====================================================
// FIND HEADER
// =====================================================

function findHeaderColumn(headers, possibleNames) {

  for (let i = 0; i < headers.length; i++) {

    for (let j = 0; j < possibleNames.length; j++) {

      if (
        headers[i] ===
        normalizeText(possibleNames[j])
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
// This function intentionally DOES NOT update IT1/IT2/IT3.
//
// It only:
//   1. validates
//   2. checks duplicate
//   3. writes one row to Attendance Responses
//
// This keeps submission extremely fast.
// =====================================================

function submitAttendance(formData) {

  if (!formData || typeof formData !== "object") {
    throw new Error("No attendance data received.");
  }

  const fullName = String(
    formData.fullName || ""
  ).trim();

  const rollNumber = String(
    formData.rollNumber || ""
  ).trim();

  const section = normalizeSection(
    formData.section
  );

  const date = String(
    formData.date || ""
  ).trim();

  const company = String(
    formData.company || ""
  ).trim();

  const eventName = String(
    formData.event || ""
  ).trim();

  const proofLink = String(
    formData.proofLink || ""
  ).trim();

  let hours = formData.hours || [];

  if (!Array.isArray(hours)) {
    hours = String(hours)
      .split(",")
      .map(function (h) {
        return String(h).trim();
      })
      .filter(Boolean);
  }

  // ---------------------------------------------------
  // VALIDATION
  // ---------------------------------------------------

  if (!fullName) {
    throw new Error("Full Name is required.");
  }

  if (!rollNumber) {
    throw new Error("Roll Number is required.");
  }

  if (!["IT1", "IT2", "IT3"].includes(section)) {
    throw new Error("Please select a valid section.");
  }

  if (!date) {
    throw new Error("Date is required.");
  }

  if (!company) {
    throw new Error("Company is required.");
  }

  if (!eventName) {
    throw new Error("Event is required.");
  }

  if (hours.length === 0) {
    throw new Error(
      "Please select at least one class hour."
    );
  }

  if (!proofLink) {
    throw new Error("Proof link is required.");
  }

  if (!/^https?:\/\//i.test(proofLink)) {
    throw new Error(
      "Please enter a valid proof link."
    );
  }


  // ---------------------------------------------------
  // SPREADSHEET
  // ---------------------------------------------------

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const responseSheet =
    ss.getSheetByName(CONFIG.RESPONSE_SHEET);

  if (!responseSheet) {
    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );
  }

  // ---------------------------------------------------
  // VERIFY ROLL NUMBER BELONGS TO SECTION
  // ---------------------------------------------------

  const sectionSheet =
    ss.getSheetByName(section);

  if (!sectionSheet) {
    throw new Error(
      "Section sheet not found: " + section
    );
  }

  const studentRow =
    findStudentRow(
      sectionSheet,
      rollNumber
    );

  if (studentRow === -1) {
    throw new Error(
      "Roll Number " +
      rollNumber +
      " not in " +
      section +
      "."
    );
  }

  // ---------------------------------------------------
  // DUPLICATE CHECK
  // ---------------------------------------------------

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


  // ---------------------------------------------------
  // ADD RESPONSE
  //
  // ONE spreadsheet write.
  // ---------------------------------------------------

  const hoursText = hours.join(", ");

  const newRow = responseSheet.getLastRow() + 1;

  responseSheet
    .getRange(newRow, 1, 1, 9)
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


  // ---------------------------------------------------
  // RETURN IMMEDIATELY
  // ---------------------------------------------------

  return {
    success: true,
    row: newRow,
    message: "Attendance submitted successfully."
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

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  // Only read columns C:G.
  // C = Roll
  // D = Section
  // E = Date
  // F = Company
  // G = Event

  const values = sheet
    .getRange(
      2,
      3,
      lastRow - 1,
      5
    )
    .getDisplayValues();

  const targetRoll =
    normalizeRollNumber(rollNumber);

  const targetDate =
    normalizeDate(date);

  const targetCompany =
    normalizeText(company);

  const targetEvent =
    normalizeText(eventName);


  for (let i = 0; i < values.length; i++) {

    const existingRoll =
      normalizeRollNumber(values[i][0]);

    const existingDate =
      normalizeDate(values[i][2]);

    const existingCompany =
      normalizeText(values[i][3]);

    const existingEvent =
      normalizeText(values[i][4]);


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

function syncResponseRow(rowNumber) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const responseSheet =
    ss.getSheetByName(CONFIG.RESPONSE_SHEET);

  if (!responseSheet) {
    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );
  }

  if (
    !rowNumber ||
    rowNumber < 2 ||
    rowNumber > responseSheet.getLastRow()
  ) {
    return;
  }


  const row = responseSheet
    .getRange(rowNumber, 1, 1, 9)
    .getDisplayValues()[0];


  const fullName = row[1];
  const rollNumber = row[2];
  const section = row[3];
  const date = row[4];
  const company = row[5];
  const eventName = row[6];
  const hoursText = row[7];
  const proofLink = row[8];


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


  const hours = hoursText
    .split(",")
    .map(function (h) {
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
    normalizeSection(section);

  const sheet =
    ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      "Sheet not found: " + sheetName
    );
  }


  const studentRow =
    findStudentRow(
      sheet,
      rollNumber
    );

  if (studentRow === -1) {
    throw new Error(
      "Student " +
      rollNumber +
      " was not found in " +
      sheetName
    );
  }


  const eventColumn =
    findEventColumn(
      sheet,
      date,
      company,
      eventName
    );

  if (eventColumn === -1) {
    throw new Error(
      "Event not found in " +
      sheetName +
      ": " +
      company +
      " - " +
      eventName
    );
  }


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
    ["", "", "", "", "", ""]
  ];

  const richText = [
    [],
  ];

  const notes = [
    ["", "", "", "", "", ""]
  ];


  for (let i = 0; i < 6; i++) {

    richText[0][i] =
      SpreadsheetApp
        .newRichTextValue()
        .setText("")
        .build();
  }


  hours.forEach(function (hour) {

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
        .setText(CONFIG.MARK)
        .setLinkUrl(proofLink)
        .build();


    notes[0][offset] =
      "Class Hour: " +
      hour +
      "\n\nProof:\n" +
      proofLink;
  });


  const range =
    sheet.getRange(
      studentRow,
      eventColumn,
      1,
      6
    );


  // One write for values
  range.setValues(values);


  // One write for rich text
  range.setRichTextValues(richText);


  // One write for notes
  range.setNotes(notes);


  range.setHorizontalAlignment("center");
}


// =====================================================
// REBUILD EVERYTHING
//
// Used when a response is manually edited/deleted.
//
// This is NOT used during normal form submission.
// =====================================================

function rebuildAllAttendanceSheets() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const responseSheet =
    ss.getSheetByName(CONFIG.RESPONSE_SHEET);

  if (!responseSheet) {
    throw new Error(
      'Sheet "Attendance Responses" was not found.'
    );
  }


  // ---------------------------------------------------
  // CLEAR ATTENDANCE AREA
  // ---------------------------------------------------

  CONFIG.SECTION_SHEETS.forEach(function (sheetName) {

    const sheet =
      ss.getSheetByName(sheetName);

    if (!sheet) {
      return;
    }

    const lastRow =
      sheet.getLastRow();

    const lastColumn =
      sheet.getLastColumn();

    if (
      lastRow < CONFIG.HEADER_ROW ||
      lastColumn < CONFIG.FIRST_EVENT_COLUMN
    ) {
      return;
    }


    const range =
      sheet.getRange(
        CONFIG.HEADER_ROW,
        CONFIG.FIRST_EVENT_COLUMN,
        lastRow - CONFIG.HEADER_ROW + 1,
        lastColumn - CONFIG.FIRST_EVENT_COLUMN + 1
      );


    // Only remove attendance data.
    // Formatting remains untouched.
    range.clearContent();
    range.clearNote();

  });


  // ---------------------------------------------------
  // READ RESPONSES
  // ---------------------------------------------------

  const lastRow =
    responseSheet.getLastRow();

  if (lastRow < 2) {
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


  // ---------------------------------------------------
  // MIRROR EVERY RESPONSE
  // ---------------------------------------------------

  responses.forEach(function (row) {

    const rollNumber = row[2];
    const section = row[3];
    const date = row[4];
    const company = row[5];
    const eventName = row[6];
    const hoursText = row[7];
    const proofLink = row[8];


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
        .map(function (h) {
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

  });
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
      SpreadsheetApp.getUi().ButtonSet.OK
    );
}


// =====================================================
// AUTO SYNC - EDIT
//
// If someone changes a response manually,
// rebuild the attendance sheets.
// =====================================================

function onResponseEdit(e) {

  try {

    if (!e || !e.range) {
      return;
    }

    const sheet =
      e.range.getSheet();

    if (
      sheet.getName() !==
      CONFIG.RESPONSE_SHEET
    ) {
      return;
    }

    rebuildAllAttendanceSheets();

  } catch (error) {

    console.log(
      "Response edit sync error: " +
      error.message
    );

  }
}


// =====================================================
// AUTO SYNC - ROW DELETION
//
// Deleting a row is a CHANGE event rather than a normal
// edit, so this handles deleted response rows.
// =====================================================

function onResponseChange(e) {

  try {

    if (!e) {
      return;
    }

    if (
      e.changeType === "REMOVE_ROW"
    ) {
      rebuildAllAttendanceSheets();
    }

  } catch (error) {

    console.log(
      "Response delete sync error: " +
      error.message
    );

  }
}


// =====================================================
// SETUP TRIGGERS
//
// Run this ONCE manually from Apps Script.
// =====================================================

function setupAttendanceTriggers() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  // Delete old versions of our triggers

  ScriptApp
    .getProjectTriggers()
    .forEach(function (trigger) {

      const handler =
        trigger.getHandlerFunction();

      if (
        handler === "onResponseEdit" ||
        handler === "onResponseChange"
      ) {
        ScriptApp.deleteTrigger(trigger);
      }

    });


  // Trigger for normal edits

  ScriptApp
    .newTrigger("onResponseEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();


  // Trigger for deleted rows

  ScriptApp
    .newTrigger("onResponseChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();


  SpreadsheetApp
    .getUi()
    .alert(
      "Attendance Sync",
      "Automatic response syncing has been enabled.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
}


// =====================================================
// MENU
// =====================================================

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu("🔄 Attendance")
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
        lastRow - CONFIG.HEADER_ROW + 1,
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

      return CONFIG.HEADER_ROW + i;

    }

  }


  return -1;
}


// =====================================================
// FIND EVENT
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


  const dateValues =
    sheet
      .getRange(
        CONFIG.DATE_ROW,
        CONFIG.FIRST_EVENT_COLUMN,
        1,
        numberOfColumns
      )
      .getDisplayValues()[0];


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
    normalizeDate(formDate);

  const targetCompany =
    normalizeText(company);

  const targetEvent =
    normalizeText(eventName);


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


    // Exact match
    if (
      sheetEvent === companyEvent
    ) {
      return (
        CONFIG.FIRST_EVENT_COLUMN + i
      );
    }


    // Company + event
    if (
      sheetEvent.includes(
        targetCompany
      ) &&
      sheetEvent.includes(
        targetEvent
      )
    ) {
      return (
        CONFIG.FIRST_EVENT_COLUMN + i
      );
    }


    // Event only
    if (
      sheetEvent === targetEvent
    ) {
      return (
        CONFIG.FIRST_EVENT_COLUMN + i
      );
    }


    // Company + event without separator
    if (
      sheetEvent.includes(
        targetCombined
      )
    ) {
      return (
        CONFIG.FIRST_EVENT_COLUMN + i
      );
    }

  }


  return -1;
}


// =====================================================
// NORMALIZE SECTION
// =====================================================

function normalizeSection(section) {

  let value =
    String(section || "")
      .trim()
      .toUpperCase()
      .replace(/-/g, "");

  return value;
}


// =====================================================
// NORMALIZE ROLL
// =====================================================

function normalizeRollNumber(value) {

  return String(value || "")
    .trim()
    .replace(/\s+/g, "")
    .toUpperCase();

}


// =====================================================
// NORMALIZE TEXT
// =====================================================

function normalizeText(value) {

  return String(value || "")
    .toLowerCase()
    .trim()
    .replace(/[–—−]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*/g, "-")
    .replace(/[.,]/g, "");

}


// =====================================================
// NORMALIZE DATE
// =====================================================

function normalizeDate(value) {

  if (!value) {
    return "";
  }


  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {

    if (isNaN(value.getTime())) {
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


  const parsed =
    new Date(text);


  if (!isNaN(parsed.getTime())) {

    return Utilities.formatDate(
      parsed,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  }


  return "";
}


// =====================================================
// PAD
// =====================================================

function pad(value) {

  return String(value)
    .padStart(2, "0");

}