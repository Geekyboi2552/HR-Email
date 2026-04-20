// ════════════════════════════════════════════════════════════
//  RECRUITING MAIL MERGE — FULL SCRIPT
//  Config sheet  →  Drive Folder (Google Doc email body, auto-fetched)
//              →  Drive Folder (CV PDF, auto-fetched)
//              →  Gmail (personalised emails + CV attached)
// ════════════════════════════════════════════════════════════


// ── SECTION 1: Read Config sheet ─────────────────────────────
function getConfig() {
  const configSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Config-enter your details here");

  if (!configSheet) {
    throw new Error("No sheet named 'Config-enter your details here' found! Please check the sheet name.");
  }

  const rows   = configSheet.getDataRange().getValues();
  const config = {};

  rows.forEach(row => {
    if (row[0] && row[1]) {
      config[row[0].trim()] = row[1].toString().trim();
    }
  });

  const required = [
    "YOUR_NAME", "YOUR_PHONE", "YOUR_LINKEDIN",
    "YOUR_COLLEGE", "YOUR_BATCH",
    "Email_Subject", "Email_Doc_Folder_ID", "CV_Folder_ID"
  ];

  const missing = required.filter(k => !config[k]);
  if (missing.length > 0) {
    throw new Error(`Missing in Config sheet: ${missing.join(", ")} — Please fill these in Column B.`);
  }

  return config;
}


// ── SECTION 2: Replace {{placeholders}} in a string ──────────
function fillTemplate(template, values) {
  let result = template;
  for (const key in values) {
    result = result.split(`{{${key}}}`).join(values[key]);
  }
  return result;
}


// ── SECTION 3: Helper — wrap text in HTML formatting tags ────
function wrapWithTags(text, tag) {
  if (!tag) return text;
  let result = text;
  if (tag.includes("u")) result = `<u>${result}</u>`;
  if (tag.includes("i")) result = `<i>${result}</i>`;
  if (tag.includes("b")) result = `<b>${result}</b>`;
  return result;
}


// ── SECTION 4: Extract Google Doc body with formatting ────────
//  Reads the Doc element tree directly — paragraph by paragraph,
//  text run by text run — and reconstructs bold/italic/underline
//  as HTML tags. No export API used.
function getDocBody(folderId) {
  const folder  = DriveApp.getFolderById(folderId);
  const files   = folder.getFilesByType(MimeType.GOOGLE_DOCS);
  const docList = [];

  while (files.hasNext()) docList.push(files.next());

  if (docList.length === 0) {
    throw new Error("No Google Doc found in your Email_Doc_Folder. Please place your email template Doc there.");
  }

  docList.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  const doc  = docList[0];
  Logger.log(`Email Doc selected: ${doc.getName()} (last modified: ${doc.getLastUpdated()})`);

  const body     = DocumentApp.openById(doc.getId()).getBody();
  const numItems = body.getNumChildren();
  let html       = "";

  for (let i = 0; i < numItems; i++) {
    const child = body.getChild(i);

    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para     = child.asParagraph();
      const numTexts = para.getNumChildren();
      let paraHtml   = "";

      for (let j = 0; j < numTexts; j++) {
        const elem = para.getChild(j);

        if (elem.getType() === DocumentApp.ElementType.TEXT) {
          const textElem = elem.asText();
          const text     = textElem.getText();

          if (text.length === 0) continue;

          // Walk character by character to handle mixed formatting runs
          let currentTag  = "";
          let currentText = "";

          for (let k = 0; k < text.length; k++) {
            const isBold      = textElem.isBold(k);
            const isItalic    = textElem.isItalic(k);
            const isUnderline = textElem.isUnderline(k);

            // Build tag key for this character's formatting combo
            let tag = "";
            if (isBold)      tag += "b";
            if (isItalic)    tag += "i";
            if (isUnderline) tag += "u";

            if (tag !== currentTag) {
              // Formatting changed — flush previous run
              if (currentText) {
                paraHtml += wrapWithTags(currentText, currentTag);
              }
              currentTag  = tag;
              currentText = text[k];
            } else {
              currentText += text[k];
            }
          }

          // Flush last run
          if (currentText) {
            paraHtml += wrapWithTags(currentText, currentTag);
          }
        }
      }

      // Wrap paragraph — empty paragraph becomes a line break
      html += paraHtml.length > 0 ? `<p>${paraHtml}</p>` : `<br>`;
    }
  }

  return html;
}


// ── SECTION 5: Auto-fetch latest CV PDF from Drive folder ─────
function getCVBlob(folderId) {
  const folder  = DriveApp.getFolderById(folderId);
  const files   = folder.getFilesByType(MimeType.PDF);
  const pdfList = [];

  while (files.hasNext()) pdfList.push(files.next());

  if (pdfList.length === 0) {
    throw new Error("No PDF found in your CV folder. Please upload your CV (PDF) there.");
  }

  pdfList.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  const cv = pdfList[0];

  Logger.log(`CV selected: ${cv.getName()} (last modified: ${cv.getLastUpdated()})`);
  return cv.getBlob().setName(cv.getName());
}


// ── SECTION 6: Main — send all emails ────────────────────────
function sendRecruitingEmails() {

  // Step 1: Load config
  const cfg          = getConfig();
  const templateBody = getDocBody(cfg["Email_Doc_Folder_ID"]);
  const cvBlob       = getCVBlob(cfg["CV_Folder_ID"]);

  // Step 2: Load data sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  if (!sheet) throw new Error("No sheet named 'test' found! Please check the sheet name.");
  const data  = sheet.getDataRange().getValues();

  // Column indexes — Sr_no(0) | HR_name(1) | Company(2) | email(3)
  // LinkedIn(4) | Contact_No(5) | Personal_Email(6) | Location(7) | Status(8)
  const COL_NAME     = 1;
  const COL_COMPANY  = 2;
  const COL_EMAIL    = 3;
  const COL_LINKEDIN = 4;
  const COL_CONTACT  = 5;
  const COL_PERSONAL = 6;
  const COL_LOCATION = 7;
  const COL_STATUS   = 8;

  let sentCount = 0, skippedCount = 0, failedCount = 0;

  // Step 3: Loop through each row
  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const email  = row[COL_EMAIL];
    const status = row[COL_STATUS];

    // Skip empty rows or already-sent rows
    if (!email || status === "Sent" || status === "sent") {
      skippedCount++;
      continue;
    }

    // Build values object — Config keys + this row's sheet data
    const values = {
      ...cfg,
      HR_name:        row[COL_NAME],
      company:        row[COL_COMPANY],
      email:          row[COL_EMAIL],
      LinkedIn:       row[COL_LINKEDIN],
      Contact_No:     row[COL_CONTACT],
      Personal_Email: row[COL_PERSONAL],
      Location:       row[COL_LOCATION],
    };

    const subject = fillTemplate(cfg["Email_Subject"], values);
    const body    = fillTemplate(templateBody, values);

    try {
      GmailApp.sendEmail(email, subject, "", {
        htmlBody:    body,
        attachments: [cvBlob],
        name:        cfg["YOUR_NAME"]
      });

      sheet.getRange(i + 1, COL_STATUS + 1).setValue("Sent");
      sheet.getRange(i + 1, COL_STATUS + 1).setBackground("#b7e1cd"); // green
      sentCount++;
      Utilities.sleep(2000); // 2s delay to avoid spam flags

    } catch (e) {
      sheet.getRange(i + 1, COL_STATUS + 1).setValue("Failed");
      sheet.getRange(i + 1, COL_STATUS + 1).setBackground("#f4c7c3"); // red
      Logger.log(`Failed for ${email}: ${e}`);
      failedCount++;
    }
  }

  // Step 4: Log summary (visible in Apps Script → Execution Log)
  Logger.log(`✅ Done! Sent: ${sentCount} | Skipped: ${skippedCount} | Failed: ${failedCount}`);
}


// ── SECTION 7: Custom menu in Google Sheet ───────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📧 Email Recruiter")
    .addItem("Send Emails to All", "sendRecruitingEmails")
    .addToUi();
}
