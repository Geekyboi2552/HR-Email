// ── CONFIG ── edit these before running ──────────────────────────
const SHEET_NAME   = "Responses";        // your sheet tab name
const SENDER_NAME  = "Katrina Kaif";     // shown as sender
const DAILY_LIMIT  = 200;              // stay under Gmail's 100/day limit
const DELAY_MS     = 2000;            // 2-second gap between emails

// ── EMAIL TEMPLATE ──────────────────────────────────────────────
// {{name}}, {{company}}, {{designation}} are replaced per row
function getEmailBody(name, company, designation) {
  return `Hi ${name},

I'm an MBA student at IIM Jammu working on a benchmarking project with peers on automotive requirements engineering processes (ASPICE SYS 1).

I have four quick questions about how your team works with requirements:

1. How many system requirements does your team derive per customer/stakeholder requirement? (e.g. 1:3, 1:5, 1:10)
2. How long does it roughly take to baseline one stakeholder requirement into a system requirement? (e.g. 2 hours, half a day)
3. What percentage of your baselined requirements typically change after sign-off? (e.g. 10%, 30%)
4. On average, what is the no. of stakeholder requirements that can be converted to system requirements per month? (e.g, 5,10,15,20 etc.)

No published data exists on this, which is why I'm reaching out to practitioners directly.

I'm happy to share the full anonymised findings once compiled. Your answers help fill a gap that no published research has covered till now.

Best regards,
${SENDER_NAME}`;
}

function getSubject(name, company) {
  return `Quick research question — automotive requirements engineering`;
}

// ── SETUP: adds status columns if not present ────────────────────
function setupSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const needed = ["Status", "Sent At", "Error"];
  needed.forEach(col => {
    if (!headers.includes(col)) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col);
    }
  });

  SpreadsheetApp.getUi().alert("✓ Sheet setup complete. Status columns added.");
}

// ── MAIN: sends emails row by row ────────────────────────────────
function sendBulkEmails() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SHEET_NAME);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  // find column indices
  const col = {
    name:    headers.indexOf("Name"),
    company: headers.indexOf("Company"),
    email:   headers.indexOf("Email ID"),
    desg:    headers.indexOf("Designation"),
    status:  headers.indexOf("Status"),
    sentAt:  headers.indexOf("Sent At"),
    error:   headers.indexOf("Error")
  };

  let sent = 0;

  for (let i = 1; i < data.length; i++) {
    if (sent >= DAILY_LIMIT) {
      Logger.log(`Daily limit of ${DAILY_LIMIT} reached. Stopping.`);
      break;
    }

    const row    = data[i];
    const status = row[col.status];

    // skip already sent or marked skip
    if (status === "Sent" || status === "Skip") continue;

    const name    = row[col.name]    || "there";
    const company = row[col.company] || "";
    const email   = row[col.email];
    const desg    = row[col.desg]    || "";

    if (!email || !email.includes("@")) {
      sheet.getRange(i + 1, col.error + 1).setValue("Invalid email");
      sheet.getRange(i + 1, col.status + 1).setValue("Failed");
      continue;
    }

    try {
      const firstName = name.split(" ")[0];

      GmailApp.sendEmail(
        email,
        getSubject(firstName, company),
        getEmailBody(firstName, company, desg),
        { name: SENDER_NAME }
      );

      sheet.getRange(i + 1, col.status + 1).setValue("Sent");
      sheet.getRange(i + 1, col.sentAt + 1).setValue(new Date().toLocaleString());
      sent++;

      Logger.log(`Sent to ${email} (${sent}/${DAILY_LIMIT})`);
      Utilities.sleep(DELAY_MS);

    } catch (e) {
      sheet.getRange(i + 1, col.error + 1).setValue(e.message);
      sheet.getRange(i + 1, col.status + 1).setValue("Failed");
      Logger.log(`Failed for ${email}: ${e.message}`);
    }
  }

  SpreadsheetApp.getUi().alert(`Done. ${sent} emails sent.`);
}

// ── OPTIONAL: attach a Word doc from Drive ───────────────────────
// Replace DRIVE_FILE_ID with the ID from your Google Drive URL
// Then replace GmailApp.sendEmail(...) with this version:
/*
function sendWithAttachment(email, subject, body, senderName) {
  const fileId = "YOUR_DRIVE_FILE_ID_HERE";
  const file   = DriveApp.getFileById(fileId);
  GmailApp.sendEmail(email, subject, body, {
    name:        senderName,
    attachments: [file.getAs(MimeType.MICROSOFT_WORD)]
  });
}
*/
