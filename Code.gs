/*******************************************************
 * PLACEMENT ATTENDANCE SYSTEM
 *
 * SHEETS
 * -----------------------------------------------------
 * Events
 * IT1
 * IT2
 * IT3
 * Attendance Responses
 *
 *
 * ATTENDANCE SHEET STRUCTURE
 * -----------------------------------------------------
 *
 * Row 3  -> Date
 * Row 4  -> Company - Event
 * Row 5  -> Hour
 * Row 6+ -> Students
 *
 * Column A -> Roll Number
 * Column B -> Full Name
 *
 * Each event occupies 6 consecutive columns:
 *
 *        1st   2nd   3rd   4th   5th   6th
 *        Hour  Hour  Hour  Hour  Hour  Hour
 *
 *
 * Example:
 *
 *        TCS Aptitude Test
 *        1st   2nd   3rd   4th   5th   6th
 *
 * If student selects:
 *
 *        2nd Hour
 *        5th Hour
 *
 * then:
 *
 *        -    P    -    -    P    -
 *
 *******************************************************/


// =====================================================
// CONFIGURATION
// =====================================================

const CONFIG = {

  // Attendance sheet structure
  DATE_ROW: 3,

  EVENT_ROW: 4,

  HOUR_ROW: 5,

  HEADER_ROW: 6,

  // Student columns
  ROLL_COLUMN: 1,

  NAME_COLUMN: 2,

  // First event starts from Column C
  FIRST_EVENT_COLUMN: 3,

  // Every event has 6 hour columns
  HOURS_PER_EVENT: 6,

  // Attendance mark
  MARK: "P",

  // Response sheet
  RESPONSE_SHEET: "Attendance Responses",

  // Events source sheet
  EVENTS_SHEET: "Events"

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
// GET EVENTS FOR DROPDOWN
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


  /*
   * Try to find columns using headers.
   *
   * Supported headers:
   *
   * Date
   * Company
   * Event
   */

  const headers =
    values[0].map(function(value) {

      return normalizeText(value);

    });


  let dateColumn =
    findHeaderColumn(
      headers,
      [
        "date"
      ]
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


  /*
   * If headers are not found,
   * assume:
   *
   * Column A = Date
   * Column B = Company
   * Column C = Event
   */

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


    const normalizedDate =
      normalizeDate(rawDate);


    if (!normalizedDate) {

      continue;

    }


    events.push({

      date:
        normalizedDate,

      company:
        company,

      event:
        event,

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


  // ---------------------------------------------------
  // Get values
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


  // ---------------------------------------------------
  // Class hours
  // ---------------------------------------------------

  let classHours =
    formData.hours || [];


  if (
    !Array.isArray(classHours)
  ) {

    classHours =
      String(classHours)
        .split(",")
        .map(function(hour) {

          return String(hour).trim();

        })
        .filter(Boolean);

  }


  // ---------------------------------------------------
  // Validate
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


  if (
    classHours.length === 0
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


  // ---------------------------------------------------
  // Basic URL validation
  //
  // We intentionally DO NOT use new URL()
  // ---------------------------------------------------

  if (
    !/^https?:\/\//i.test(
      proofLink
    )
  ) {

    throw new Error(
      "Please enter a valid proof link."
    );

  }


  // ===================================================
  // SPREADSHEET
  // ===================================================

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  // ===================================================
  // ATTENDANCE SHEET
  // ===================================================

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
      "Attendance sheet not found: " +
      sheetName
    );

  }


  // ===================================================
  // FIND STUDENT
  // ===================================================

  const studentRow =
    findStudentRow(
      sheet,
      rollNumber
    );


  if (studentRow === -1) {

    throw new Error(
      "Student with roll number " +
      rollNumber +
      " was not found in " +
      sheetName +
      "."
    );

  }


  // ===================================================
  // FIND EVENT
  // ===================================================

  const eventColumn =
    findEventColumn(
      sheet,
      date,
      company,
      eventName
    );


  if (eventColumn === -1) {

    throw new Error(
      "Could not find the event in " +
      sheetName +
      ".\n\n" +
      "Date: " +
      date +
      "\n" +
      "Company: " +
      company +
      "\n" +
      "Event: " +
      eventName
    );

  }


  // ===================================================
  // DUPLICATE CHECK
  // ===================================================

  /*
   * Check whether this student has already
   * submitted the same event.
   *
   * If a previous submission exists, we don't
   * create another response record.
   */

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


  // ===================================================
  // MARK ATTENDANCE
  // ===================================================

  markAttendanceForHours(

    sheet,

    studentRow,

    eventColumn,

    classHours,

    proofLink

  );


  // ===================================================
  // ADD RESPONSE RECORD
  // ===================================================

  addAttendanceResponse(

    ss,

    {

      fullName:
        fullName,

      rollNumber:
        rollNumber,

      section:
        section,

      date:
        date,

      company:
        company,

      event:
        eventName,

      hours:
        classHours,

      proofLink:
        proofLink

    }

  );


  // ===================================================
  // SUCCESS
  // ===================================================

  return {

    success:
      true,

    message:
      "Attendance submitted successfully."

  };

}


// =====================================================
// MARK ATTENDANCE FOR SELECTED HOURS
// =====================================================

function markAttendanceForHours(

  sheet,

  studentRow,

  eventColumn,

  classHours,

  proofLink

) {


  /*
   * The event occupies 6 consecutive columns.
   *
   * 1st Hour -> +0
   * 2nd Hour -> +1
   * 3rd Hour -> +2
   * 4th Hour -> +3
   * 5th Hour -> +4
   * 6th Hour -> +5
   */

  const hourColumnMap = {

    "1st Hour": 0,

    "2nd Hour": 1,

    "3rd Hour": 2,

    "4th Hour": 3,

    "5th Hour": 4,

    "6th Hour": 5

  };


  classHours.forEach(
    function(hour) {


      if (
        !Object.prototype.hasOwnProperty.call(
          hourColumnMap,
          hour
        )
      ) {

        return;

      }


      const offset =
        hourColumnMap[hour];


      const column =
        eventColumn +
        offset;


      const cell =
        sheet.getRange(
          studentRow,
          column
        );


      // ---------------------------------------------
      // Mark P
      // ---------------------------------------------

      cell.setValue(
        CONFIG.MARK
      );


      // ---------------------------------------------
      // Make P clickable
      // ---------------------------------------------

      cell.setRichTextValue(

        SpreadsheetApp
          .newRichTextValue()

          .setText(
            CONFIG.MARK
          )

          .setLinkUrl(
            proofLink
          )

          .build()

      );


      // ---------------------------------------------
      // Add note
      // ---------------------------------------------

      cell.setNote(

        "Class Hour: " +
        hour +

        "\n\n" +

        "Proof:\n" +
        proofLink

      );


      // ---------------------------------------------
      // Alignment
      // ---------------------------------------------

      cell.setHorizontalAlignment(
        "center"
      );


    }
  );

}


// =====================================================
// ADD RECORD TO ATTENDANCE RESPONSES
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


  // ---------------------------------------------------
  // Convert hours to text
  // ---------------------------------------------------

  const hoursText =
    data.hours.join(
      ", "
    );


  // ---------------------------------------------------
  // Append record
  // ---------------------------------------------------

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
// CHECK DUPLICATE SUBMISSION
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


  /*
   * Header only
   */

  if (lastRow < 2) {

    return false;

  }


  /*
   * Expected columns:
   *
   * A Timestamp
   * B Full Name
   * C Roll Number
   * D Section
   * E Date
   * F Company
   * G Event
   * H Hours
   * I Proof
   */

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
// NORMALIZE SECTION
// =====================================================

function normalizeSection(
  section
) {

  let value =
    String(
      section || ""
    )
      .trim()
      .toUpperCase();


  value =
    value.replace(
      /-/g,
      ""
    );


  if (
    value === "IT1"
  ) {

    return "IT1";

  }


  if (
    value === "IT2"
  ) {

    return "IT2";

  }


  if (
    value === "IT3"
  ) {

    return "IT3";

  }


  return value;

}


// =====================================================
// FIND STUDENT ROW
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


  const targetRoll =
    normalizeRollNumber(
      rollNumber
    );


  for (
    let i = 0;
    i < values.length;
    i++
  ) {


    const sheetRoll =
      normalizeRollNumber(
        values[i][0]
      );


    if (
      sheetRoll ===
      targetRoll
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
// FIND EVENT COLUMN
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
  // Read date row
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
  // Read event row
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


  // ---------------------------------------------------
  // Target values
  // ---------------------------------------------------

  const targetCompany =
    normalizeText(
      company
    );


  const targetEvent =
    normalizeText(
      eventName
    );


  const targetDate =
    normalizeDate(
      formDate
    );


  /*
   * Since events occupy 6 columns,
   * only the first column of each block
   * needs to be checked.
   *
   * This also works when the event/date cells
   * are merged across the six columns.
   */

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


    if (
      !sheetEvent
    ) {

      continue;

    }


    // -------------------------------------------------
    // Date must match
    // -------------------------------------------------

    if (

      targetDate &&
      sheetDate &&
      targetDate !==
      sheetDate

    ) {

      continue;

    }


    // -------------------------------------------------
    // Possible event strings
    // -------------------------------------------------

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


    const sheetCombined =
      normalizeText(
        sheetEvent
      );


    // -------------------------------------------------
    // Exact match
    // -------------------------------------------------

    if (
      sheetCombined ===
      companyEvent
    ) {

      return (
        CONFIG.FIRST_EVENT_COLUMN +
        i
      );

    }


    // -------------------------------------------------
    // Flexible company + event
    // -------------------------------------------------

    if (

      sheetCombined.includes(
        targetCompany
      ) &&

      sheetCombined.includes(
        targetEvent
      )

    ) {

      return (
        CONFIG.FIRST_EVENT_COLUMN +
        i
      );

    }


    // -------------------------------------------------
    // Event name only
    // -------------------------------------------------

    if (
      sheetCombined ===
      targetEvent
    ) {

      return (
        CONFIG.FIRST_EVENT_COLUMN +
        i
      );

    }


    // -------------------------------------------------
    // Company + event without separator
    // -------------------------------------------------

    if (
      sheetCombined.includes(
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

    // Normalize dash types
    .replace(
      /[–—−]/g,
      "-"
    )

    // Remove extra spaces
    .replace(
      /\s+/g,
      " "
    )

    // Remove spaces around hyphens
    .replace(
      /\s*-\s*/g,
      "-"
    )

    // Remove punctuation
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
  // Date object
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


    return Utilities.formatDate(

      value,

      Session.getScriptTimeZone(),

      "yyyy-MM-dd"

    );

  }


  const text =
    String(
      value
    ).trim();


  // ---------------------------------------------------
  // yyyy-mm-dd
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
  // dd.mm.yyyy
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
  // dd/mm/yyyy
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
  // JavaScript parser
  // ---------------------------------------------------

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

function pad(
  value
) {

  return String(
    value
  ).padStart(
    2,
    "0"
  );

}
