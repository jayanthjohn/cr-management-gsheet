// List of fixed team members
const teamMembers = [
  "Jayanth", "Sudeshna", "Binit", "Neema", "Manju", "Timothy"
];

// Status dropdown options
const statusOptions = ["Closed", "Cancelled", "Deferred"];

// Fixed Guest List for Calendar Invites
const guestEmails = "jayanthfordhon@gmail.com,jayanthfordhon1@gmail.com"; // Add your team emails here

// 🛠 Setup menu when sheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔧 CR Tools")
    .addItem("➕ Add New CR", "showCRForm")
    .addItem("✏️ Update Existing CR", "showCRUpdateForm")
    .addToUi();
}

// 📋 Show form to create new CR
function showCRForm() {
  const html = HtmlService.createHtmlOutputFromFile("CRForm")
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "➕ Add New CR");
}

// ✏️ Show form to update existing CR
function showCRUpdateForm() {
  const html = HtmlService.createHtmlOutputFromFile("CRUpdate")
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "✏️ Update CR");
}

// 🚀 Provide dropdown data to forms
function getDropdownData() {
  return {
    teamMembers: teamMembers,
    statusOptions: statusOptions
  };
}

// 📚 Get all unclosed CRs (Status empty)
function getUnclosedCRs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const unclosed = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][13]; // Status (14th column)
    if (!status || status.trim() === "") {
      unclosed.push({
        row: i + 1,
        crNumber: data[i][0] // CR/Incident
      });
    }
  }
  return unclosed;
}

// 🛠 Add a new CR
function addCR(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const now = new Date();
  const startTime = new Date(formData.startDateTime.replace("T", " ") + ":00");
  const endTime = new Date(formData.endDateTime.replace("T", " ") + ":00");

  const newRow = [
    formData.crIncident,
    formData.reason,
    formData.raisedBy,
    formData.assignedTo,
    formData.approver,
    formData.implementer,
    formData.validator,
    now,
    startTime,
    endTime,
    "", "", "", "" // Actual Start, Actual End, Total Duration, Status
  ];

  // 🔒 Temporarily unprotect the sheet
  const protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  const editors = protection ? protection.getEditors() : [];

  if (protection) protection.remove();
  sheet.appendRow(newRow);
  if (editors.length) {
    const newProtection = sheet.protect().setDescription('Sheet Lock');
    newProtection.addEditors(editors);
    newProtection.setWarningOnly(false);
  }

  // 📅 Create Calendar Event - 45 minutes before End Time
  const calendar = CalendarApp.getDefaultCalendar();
  const reminderStart = new Date(endTime.getTime() - 45 * 60 * 1000);
  const reminderEnd = new Date(reminderStart.getTime() + 15 * 60 * 1000);

  const event = calendar.createEvent(
    `CR Reminder: ${formData.crIncident}`,
    reminderStart,
    reminderEnd,
    {
      description: `Reminder to prepare for closing CR ${formData.crIncident}.`,
      guests: guestEmails,
      sendInvites: true
    }
  );
  event.addPopupReminder(10);

  // 📧 Schedule Email Reminder - 30 mins before End Time
  ScriptApp.newTrigger("sendEmailReminder")
    .timeBased()
    .at(new Date(endTime.getTime() - 30 * 60 * 1000))
    .create();
}

// 📧 Send Email Reminder (30 mins before End Date)
function sendEmailReminder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const tolerance = 5 * 60 * 1000;

  for (let i = 1; i < data.length; i++) {
    const crNumber = data[i][0];
    const endTime = data[i][9];
    const status = data[i][13];

    if (endTime instanceof Date &&
        Math.abs(now - (endTime.getTime() - 30 * 60 * 1000)) <= tolerance &&
        (!status || status.trim() === "")) {

      MailApp.sendEmail({
        to: "jayanthfordhon@gmail.com,jayanthfordhon1@gmail.com", // Same list as calendar invite or different if needed
        subject: `CR Closing Reminder: ${crNumber}`,
        body: `Hi Team,\n\nThe CR '${crNumber}' is scheduled to close soon at ${endTime.toLocaleString()}.\nPlease be ready to complete closure activities.\n\n- Automated Reminder`
      });
    }
  }
}

// ✏️ Update CR with actual times and status
function updateCR(updateData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = Number(updateData.row);
  const actualStart = new Date(updateData.actualStartDateTime.replace("T", " ") + ":00");
  const actualEnd = new Date(updateData.actualEndDateTime.replace("T", " ") + ":00");
  const duration = (actualEnd - actualStart) / (1000 * 60); // in minutes

  sheet.getRange(row, 11).setValue(actualStart); // Actual Start
  sheet.getRange(row, 12).setValue(actualEnd);   // Actual End
  sheet.getRange(row, 13).setValue(duration);    // Total Duration (minutes)
  sheet.getRange(row, 14).setValue(updateData.status); // Status
}