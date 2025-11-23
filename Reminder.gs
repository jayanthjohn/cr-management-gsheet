// 📊 DAILY SUMMARY EMAIL (9 AM & 5 PM)
function sendDailyCRSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 🔄 Force Google Sheets to commit latest edits
  SpreadsheetApp.flush();

  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

  let openCRs = [];

  for (let i = 1; i < data.length; i++) {
    const cr = data[i][0];       // CR/Incident
    const endDate = data[i][9];  // End Date & Time (Column J → index 9)
    const status = data[i][13];  // Status (Column N → index 13)

    // 🧠 Only pick CRs with EMPTY status (open)
    if (cr && endDate instanceof Date && (!status || status.toString().trim() === "")) {
      const overdue = endDate < now;

      openCRs.push({
        cr,
        plannedEnd: Utilities.formatDate(endDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
        isOverdue: overdue
      });
    }
  }

  const recipients = "jayanthfordhon@gmail.com,jayanthfordhon1@gmail.com";

  // 🎉 If no pending CRs → Send Happy Message
  if (openCRs.length === 0) {
    MailApp.sendEmail({
      to: recipients,
      subject: "✅ GSheet Automator - No Open CRs 🎉",
      htmlBody: `
        <p style="font-family: Arial; font-size: 14px;">Hi Team,</p>

        <p style="font-family: Arial; font-size: 15px; color: green;">
          🎉 All CRs are closed! No pending actions as of <b>${today}</b>.
        </p>

        <p style="font-family: Arial;">Great job staying on track! 💪</p>
        <hr>
        <p style="font-size:12px;color:gray;">– GSheet Automator</p>
      `
    });
    return;
  }

  // 🧾 Build CR table for email
  let tableRows = openCRs.map(cr => `
    <tr>
      <td style="padding: 8px; border: 1px solid #ccc;">${cr.cr}</td>
      <td style="padding: 8px; border: 1px solid #ccc;">${cr.plannedEnd}</td>
      <td style="padding: 8px; border: 1px solid #ccc; color: ${cr.isOverdue ? 'red' : 'green'};">
        ${cr.isOverdue ? '⚠️ Overdue' : '✅ On Time'}
      </td>
    </tr>
  `).join("");

  const htmlBody = `
    <p style="font-family: Arial; font-size: 14px;">Hi Team,</p>

    <p style="font-family: Arial;">Below are the <b>Unclosed CRs</b> as of <b>${today}</b>:</p>

    <table style="border-collapse: collapse; font-family: Arial; font-size: 14px;">
      <tr style="background-color: #f0f0f0;">
        <th style="padding: 8px; border: 1px solid #ccc;">CR Number</th>
        <th style="padding: 8px; border: 1px solid #ccc;">Planned End Date</th>
        <th style="padding: 8px; border: 1px solid #ccc;">Status</th>
      </tr>
      ${tableRows}
    </table>

    <br>

    <p style="font-family: Arial;">Please take necessary action.</p>

    <hr>
    <p style="font-size:12px;color:gray;">– GSheet Automator</p>
  `;

  // 📧 Send Summary Email
  MailApp.sendEmail({
    to: recipients,
    subject: `📋 TEammmmm -  CR Summary (${today})`,
    htmlBody: htmlBody
  });
}