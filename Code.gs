// 🔧 Configurations
const teamMembers = ["Jayanth", "Sudeshna", "Binit", "Neema", "Manju", "Timothy", "Karthik"];
const statusOptions = ["Closed", "Cancelled", "Deferred"];
const guestEmails = "jayanthfordhon@gmail.com,jayanthfordhon1@gmail.com";
const ticketBaseUrl = "https://tickets.mycompany.com/browse/";  // 🔗 Update this if needed

// 📌 Menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔧 CR Tools")
    .addItem("➕ Add New CR", "showCRForm")
    .addItem("✏️ Update Existing CR", "showCRUpdateForm")
    .addItem("🧪 Test Summary Email", "sendDailyCRSummary")
    .addToUi();
}

// 🖼️ Forms
function showCRForm() {
  const html = HtmlService.createHtmlOutputFromFile("CRForm").setWidth(500).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "➕ Add New CR");
}
function showCRUpdateForm() {
  const html = HtmlService.createHtmlOutputFromFile("CRUpdate").setWidth(500).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "✏️ Update CR");
}
function getDropdownData() {
  return { teamMembers: teamMembers, statusOptions: statusOptions };
}

// 📋 Add New CR
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

  const protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  const editors = protection ? protection.getEditors() : [];

  if (protection) protection.remove();
  sheet.appendRow(newRow);
  if (editors.length) {
    const newProtection = sheet.protect().setDescription('Sheet Lock');
    newProtection.addEditors(editors);
    newProtection.setWarningOnly(false);
  }

  // 📅 Calendar Event
  const calendar = CalendarApp.getDefaultCalendar();
  const reminderStart = new Date(endTime.getTime() - 45 * 60 * 1000);
  const reminderEnd = new Date(reminderStart.getTime() + 15 * 60 * 1000);

  calendar.createEvent(
    `CR Reminder: ${formData.crIncident}`,
    reminderStart,
    reminderEnd,
    {
      description: `Reminder to prepare for closing CR ${formData.crIncident}.`,
      guests: guestEmails,
      sendInvites: true
    }
  );

  // 🔔 Trigger email reminder for this CR only
  const lastRow = sheet.getLastRow();
  const trigger = ScriptApp.newTrigger("sendEmailReminderWithContext")
    .timeBased()
    .at(new Date(endTime.getTime() - 30 * 60 * 1000))
    .create();

  const props = PropertiesService.getScriptProperties();
  props.setProperty(trigger.getUniqueId(), lastRow);
}

// 🔔 Email Reminder for Specific CR
function sendEmailReminderWithContext(e) {
  const props = PropertiesService.getScriptProperties();
  const triggerUid = e.triggerUid;
  const row = Number(props.getProperty(triggerUid));

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const crNumber = rowData[0];
  const endTime = rowData[9];
  const status = rowData[13];

  if (endTime instanceof Date && (!status || status.trim() === "")) {
    MailApp.sendEmail({
      to: guestEmails,
      subject: `CR Closing Reminder: ${crNumber}`,
      body: `Hi Team,\n\nThe CR '${crNumber}' is scheduled to close at ${endTime.toLocaleString()}.\nPlease be ready to complete closure activities.\n\n– GSheet Automator`
    });
  }

  // 🔄 Clean up
  const triggers = ScriptApp.getProjectTriggers();
  for (let t of triggers) {
    if (t.getUniqueId() === triggerUid) {
      ScriptApp.deleteTrigger(t);
      break;
    }
  }
  props.deleteProperty(triggerUid);
}

// 📬 Daily Summary (9 AM & 5 PM)
function sendDailyCRSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  SpreadsheetApp.flush();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

  let openCRs = [];

  for (let i = 1; i < data.length; i++) {
    const cr = data[i][0];            // CR/Incident
    const raisedBy = data[i][2];      // Raised By (Column C)
    const validator = data[i][6];     // Validator (Column G)
    const endDate = data[i][9];       // End Date & Time (Column J)
    const status = data[i][13];       // Status (Column N)

    if (cr && endDate instanceof Date && (!status || status.toString().trim() === "")) {
      openCRs.push({
        cr,
        raisedBy,
        validator,
        plannedEnd: Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
        isOverdue: endDate < now
      });
    }
  }

  const recipients = guestEmails;

  if (openCRs.length === 0) {
    MailApp.sendEmail({
      to: recipients,
      subject: "✅ GSheet Automator - No Open CRs 🎉",
      htmlBody: `
        <p style="font-family: Arial;">Hi Team,</p>
        <p style="font-family: Arial; font-size: 15px; color: green;">
          🎉 All CRs are closed! No pending actions as of <b>${today}</b>.
        </p>
        <p style="font-family: Arial;">Keep up the great work! 💪</p>
        <hr><p style="font-size:12px;color:gray;">– GSheet Automator</p>`
    });
    return;
  }

  const tableRows = openCRs.map(cr => {
    const ticketLink = `${ticketBaseUrl}${cr.cr}`;
    return `
      <tr>
        <td style="padding: 8px; border: 1px solid #ccc;">
          <a href="${ticketLink}" target="_blank">${cr.cr}</a>
        </td>
        <td style="padding: 8px; border: 1px solid #ccc;">${cr.raisedBy}</td>
        <td style="padding: 8px; border: 1px solid #ccc;">${cr.validator}</td>
        <td style="padding: 8px; border: 1px solid #ccc;">${cr.plannedEnd}</td>
        <td style="padding: 8px; border: 1px solid #ccc; color: ${cr.isOverdue ? 'red' : 'green'};">
          ${cr.isOverdue ? '⚠️ Overdue' : '✅ On Time'}
        </td>
      </tr>`;
  }).join('');

  const htmlBody = `
    <p style="font-family: Arial;">Hi Team,</p>
    <p style="font-family: Arial;">Here is the list of <b>unclosed CRs</b> as of <b>${today}</b>:</p>
    <table style="border-collapse: collapse; font-family: Arial; font-size: 14px;">
      <tr style="background-color: #f2f2f2;">
        <th style="padding: 8px; border: 1px solid #ccc;">CR Number</th>
        <th style="padding: 8px; border: 1px solid #ccc;">Raised By</th>
        <th style="padding: 8px; border: 1px solid #ccc;">Validator</th>
        <th style="padding: 8px; border: 1px solid #ccc;">Planned End Date</th>
        <th style="padding: 8px; border: 1px solid #ccc;">Status</th>
      </tr>${tableRows}
    </table>
    <br>
    <p style="font-family: Arial;">Please take action as required.</p>
    <hr><p style="font-size:12px;color:gray;">– GSheet Automator</p>
  `;

  MailApp.sendEmail({
    to: recipients,
    subject: `📋 GSheet Automator - Open CR Summary (${today})`,
    htmlBody: htmlBody
  });
}

// ✏️ Update CR Actuals
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