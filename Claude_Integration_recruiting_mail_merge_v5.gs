// ════════════════════════════════════════════════════════════
//  RECRUITING MAIL MERGE — FULL SCRIPT (AI-POWERED) v6
//  Config sheet  →  Drive Folder (Google Doc email structure)
//              →  Drive Folder (CV PDF — attached + text extracted)
//              →  Claude AI (generates personalised email per HR)
//              →  Gmail (sends personalised email + CV attached)
// ════════════════════════════════════════════════════════════
//
//  CONFIG SHEET ("Config-enter your details here") KEYS:
//  ──────────────────────────────────────────────────────
//  YOUR_NAME           → Ankush Kumar Bawa
//  YOUR_PHONE          → 76966-67410
//  YOUR_LINKEDIN       → your-linkedin-url
//  YOUR_COLLEGE        → IIM Jammu
//  YOUR_BATCH          → 2025-2027
//  Email_Subject       → MBA Student from {{YOUR_COLLEGE}} – Exploring Opportunities at {{company}}
//  Email_Doc_Folder_ID → (Google Drive folder ID containing your email template Doc)
//  CV_Folder_ID        → (Google Drive folder ID containing your CV PDF)
//  ANTHROPIC_API_KEY   → (your API key from console.anthropic.com)
//
//  HR DATA SHEET ("test") COLUMNS:
//  ────────────────────────────────
//  Sr_no(0) | HR_name(1) | Company(2) | email(3) | LinkedIn(4)
//  Contact_No(5) | Personal_Email(6) | Location(7) | Domain(8)
//  Roles(9) | Designation(10) | Status(11)
//
//  BEFORE RUNNING:
//  ───────────────
//  1. Enable Drive API v2 → Apps Script editor → Services → Drive API → v2 → Add
//  2. Fill all Config sheet keys in Column B
//  3. Place CV PDF in CV_Folder
//  4. Place email template Google Doc in Email_Doc_Folder
//  5. Run via 📧 Email Recruiter → Send Emails to All
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
    "Email_Subject", "Email_Doc_Folder_ID", "CV_Folder_ID",
    "ANTHROPIC_API_KEY"
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

          let currentTag  = "";
          let currentText = "";

          for (let k = 0; k < text.length; k++) {
            const isBold      = textElem.isBold(k);
            const isItalic    = textElem.isItalic(k);
            const isUnderline = textElem.isUnderline(k);

            let tag = "";
            if (isBold)      tag += "b";
            if (isItalic)    tag += "i";
            if (isUnderline) tag += "u";

            if (tag !== currentTag) {
              if (currentText) paraHtml += wrapWithTags(currentText, currentTag);
              currentTag  = tag;
              currentText = text[k];
            } else {
              currentText += text[k];
            }
          }

          if (currentText) paraHtml += wrapWithTags(currentText, currentTag);
        }
      }

      html += paraHtml.length > 0 ? `<p>${paraHtml}</p>` : `<br>`;
    }
  }

  return html;
}


// ── SECTION 5: Auto-fetch latest CV PDF blob (for attachment) ─
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


// ── SECTION 6: Extract CV text from PDF via Drive OCR ─────────
//  Temporarily converts the PDF to a Google Doc to extract text,
//  then immediately deletes the temp Doc.
//  Requires: Drive API v2 enabled in Services.
function extractCVText(folderId) {
  const folder  = DriveApp.getFolderById(folderId);
  const files   = folder.getFilesByType(MimeType.PDF);
  const pdfList = [];

  while (files.hasNext()) pdfList.push(files.next());

  if (pdfList.length === 0) {
    throw new Error("No PDF found in your CV folder for text extraction.");
  }

  pdfList.sort((a, b) => b.getLastUpdated() - a.getLastUpdated());
  const pdf = pdfList[0];

  Logger.log(`Extracting text from CV: ${pdf.getName()}`);

  // Convert PDF → temporary Google Doc using Drive OCR
  const tempDoc = Drive.Files.insert(
    { title: "temp_cv_ocr", mimeType: MimeType.GOOGLE_DOCS },
    pdf.getBlob()
  );

  // Extract plain text from the temp Doc
  const text = DocumentApp.openById(tempDoc.id).getBody().getText();

  // Delete temp Doc immediately
  DriveApp.getFileById(tempDoc.id).setTrashed(true);

  Logger.log(`CV text extracted (${text.length} characters)`);
  return text;
}


// ── SECTION 7: Generate personalised email via Claude AI ──────
function generatePersonalisedEmail(hrData, cvText, cfg) {
  const prompt = `
You are writing a high-conversion cold recruiting email on behalf of ${cfg.YOUR_NAME}, an MBA student at ${cfg.YOUR_COLLEGE} (Batch ${cfg.YOUR_BATCH}).

Recipient details:
- HR Name: ${hrData.name}
- Designation: ${hrData.designation}
- Company: ${hrData.company}
- Domain: ${hrData.domain}
- Roles they are currently hiring for: ${hrData.roles}
- Location: ${hrData.location}

Sender's CV (full text):
${cvText}

Write a highly personalised, professional cold email that follows these STRICT guidelines:

1. Opens with a warm, natural greeting addressing the HR by name
2. Subtly acknowledges their seniority and designation — tailor the tone accordingly:
   - For senior designations (MD, CEO, Director, Zonal Head, Executive Director): tone should be more formal, respectful, and strategic — focus on business impact and leadership potential
   - For mid-level designations (HR Executive, Talent Acquisition, Regional HR): tone can be slightly warmer and more direct — focus on role fit and skills
   - For technical/specialised hiring (Tech & Leadership Talent, Medical Hiring): highlight only the most relevant domain-specific experience from the CV
3. Mentions the company and its domain organically (avoid generic or forced praise)
4. Clearly aligns the sender's profile with the roles they are hiring for
5. Extracts and highlights ONLY the most relevant skills, experiences, and achievements from the CV (do not list everything)
6. Demonstrates value in a crisp, specific way (avoid vague claims like "hardworking" or "passionate")
7. Keeps the tone confident, respectful, and non-desperate
8. Includes a clear but soft call-to-action (e.g., requesting guidance or a short conversation — NOT directly asking for a job)
9. Keeps the email concise (150–180 words ideal, 200 words max)
10. Uses short paragraphs for readability
11. Avoids clichés, fluff, and mass-email tone — it should feel genuinely written for this specific person
12. Bullet Points Rule:
    - Use bullet points ONLY if they improve clarity and readability
    - Maximum 2–3 bullet points
    - Each point must highlight a strong, relevant achievement or skill
    - Do NOT replicate the full CV or create a long list
    - The email must still feel conversational, not like a resume dump

End with a professional sign-off including:
- Name
- College and batch
- Phone number
- LinkedIn profile

Return ONLY the email body as clean HTML using <p>, <b>, and <br> tags.
Do NOT include:
- Subject line
- Any explanation or meta text
- Placeholder text (everything should feel real and filled)
`;

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    headers: {
      "x-api-key":         cfg.ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "content-type":      "application/json"
    },
    payload: JSON.stringify({
      model:      "claude-haiku-4-5-20251001",
      max_tokens: 1024,
      messages:   [{ role: "user", content: prompt }]
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error(`Claude API error: ${result.error.message}`);
  }

  return result.content[0].text;
}


// ── SECTION 8: Main — send all emails ────────────────────────
function sendRecruitingEmails() {

  // Step 1: Load config
  const cfg = getConfig();

  // Step 2: Fetch CV blob (for attachment) and CV text (for AI prompt)
  // Both are fetched once before the loop — not per row
  const cvBlob = getCVBlob(cfg["CV_Folder_ID"]);
  const cvText = extractCVText(cfg["CV_Folder_ID"]);

  // Step 3: Load HR data sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  if (!sheet) throw new Error("No sheet named 'test' found! Please check the sheet name.");
  const data = sheet.getDataRange().getValues();

  // Column indexes
  // Sr_no(0) | HR_name(1) | Company(2) | email(3) | LinkedIn(4)
  // Contact_No(5) | Personal_Email(6) | Location(7) | Domain(8)
  // Roles(9) | Designation(10) | Status(11)
  const COL_NAME        = 1;
  const COL_COMPANY     = 2;
  const COL_EMAIL       = 3;
  const COL_LINKEDIN    = 4;
  const COL_CONTACT     = 5;
  const COL_PERSONAL    = 6;
  const COL_LOCATION    = 7;
  const COL_DOMAIN      = 8;
  const COL_ROLES       = 9;
  const COL_DESIGNATION = 10;
  const COL_STATUS      = 11;

  let sentCount = 0, skippedCount = 0, failedCount = 0;

  // Step 4: Loop through each row
  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const email  = row[COL_EMAIL];
    const status = row[COL_STATUS];

    // Skip empty rows or already-sent rows
    if (!email || status === "Sent" || status === "sent") {
      skippedCount++;
      continue;
    }

    // Build values object for subject line template
    const values = {
      ...cfg,
      HR_name:        row[COL_NAME],
      company:        row[COL_COMPANY],
      email:          row[COL_EMAIL],
      LinkedIn:       row[COL_LINKEDIN],
      Contact_No:     row[COL_CONTACT],
      Personal_Email: row[COL_PERSONAL],
      Location:       row[COL_LOCATION],
      domain:         row[COL_DOMAIN],
      roles:          row[COL_ROLES],
      designation:    row[COL_DESIGNATION],
    };

    const subject = fillTemplate(cfg["Email_Subject"], values);

    try {
      // Generate AI-personalised email body for this specific HR
      const body = generatePersonalisedEmail({
        name:        row[COL_NAME],
        company:     row[COL_COMPANY],
        domain:      row[COL_DOMAIN],
        roles:       row[COL_ROLES],
        location:    row[COL_LOCATION],
        designation: row[COL_DESIGNATION]
      }, cvText, cfg);

      GmailApp.sendEmail(email, subject, "", {
        htmlBody:    body,
        attachments: [cvBlob],
        name:        cfg["YOUR_NAME"]
      });

      sheet.getRange(i + 1, COL_STATUS + 1).setValue("Sent");
      sheet.getRange(i + 1, COL_STATUS + 1).setBackground("#b7e1cd"); // green
      sentCount++;

      // Delay between emails — avoids spam flags + API rate limits
      Utilities.sleep(3000);

    } catch (e) {
      sheet.getRange(i + 1, COL_STATUS + 1).setValue("Failed");
      sheet.getRange(i + 1, COL_STATUS + 1).setBackground("#f4c7c3"); // red
      Logger.log(`Failed for ${email}: ${e}`);
      failedCount++;
    }
  }

  // Step 5: Log summary (visible in Apps Script → Execution Log)
  Logger.log(`✅ Done! Sent: ${sentCount} | Skipped: ${skippedCount} | Failed: ${failedCount}`);
}


// ── SECTION 9: Custom menu in Google Sheet ───────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📧 Email Recruiter")
    .addItem("Send Emails to All", "sendRecruitingEmails")
    .addToUi();
}
