// List of fixed team members
const teamMembers = [
  "Jayanth", "Sudeshna", "Binit", "Neema", "Manju", "Timothy", "Karthik"
];
// Status dropdown options
const statusOptions = ["Closed", "Cancelled", "Deferred"];
// Fixed Guest List for Calendar Invites
const guestEmails = "jayanthfordhon@gmail.com,jayanthfordhon1@gmail.com";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔧 CR Tools")
    .addItem("➕ Add New CR", "showCRForm")
    .addItem("✏️ Update Existing CR", "showCRUpdateForm")
    .addToUi();
}
function showCRForm() {
  const html = HtmlService.createHtmlOutputFromFile("CRForm")
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "➕ Add New CR");
}
function showCRUpdateForm() {
  const html = HtmlService.createHtmlOutputFromFile("CRUpdate")
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "✏️ Update CR");
}
function getDropdownData() {
  return {
    teamMembers: teamMembers,
    statusOptions: statusOptions
  };
}
function getUnclosedCRs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const unclosed = [];
  for (let i = 1; i < data.length; i++) {
    const status = data[i][13];
    if (!status || status.trim() === "") {
      unclosed.push({
        row: i + 1,
        crNumber: data[i][0]
      });
    }
  }
  return unclosed;
}
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
    "", "", "", ""
  ];
  const protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  const editors = protection ? protection.getEditors() : [];
  if (protection) protection.remove();
  sheet.appendRow(newRow);
  if (editors.length) {
    const newProtection = sheet.protect().setDescription('Sheet Lock');
    newProtection.addEditors(editors);
    newProtection.setWarningOnly(false);
  }
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
  ScriptApp.newTrigger("sendEmailReminder")
    .timeBased()
    .at(new Date(endTime.getTime() - 30 * 60 * 1000))
    .create();
}
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
        to: "jayanthfordhon@gmail.com,jayanthfordhon1@gmail.com",
        subject: `CR Closing Reminder: ${crNumber}`,
        body: `Hi Team,

The CR '${crNumber}' is scheduled to close soon at ${endTime.toLocaleString()}.
Please be ready to complete closure activities.

- Automated Reminder`
      });
    }
  }
}
function updateCR(updateData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = Number(updateData.row);
  const actualStart = new Date(updateData.actualStartDateTime.replace("T", " ") + ":00");
  const actualEnd = new Date(updateData.actualEndDateTime.replace("T", " ") + ":00");
  const duration = (actualEnd - actualStart) / (1000 * 60);
  sheet.getRange(row, 11).setValue(actualStart);
  sheet.getRange(row, 12).setValue(actualEnd);
  sheet.getRange(row, 13).setValue(duration);
  sheet.getRange(row, 14).setValue(updateData.status);
}
