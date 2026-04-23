// ════════════════════════════════════════════════════════════
//  RECRUITING MAIL MERGE — FULL SCRIPT (AI-POWERED) v7
//  Config sheet  →  Drive Folder (Google Doc email structure)
//              →  Drive Folder (CV PDF — attached + text extracted)
//              →  Claude AI (enriches company domain + roles)
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
//  Roles(9) | Designation(10) | Role_Location(11) | Status(12)
//
//  NOTE: Role_Location (col 11, column L) = city/cities where the company is hiring.
//        Location     (col 7,  column H)  = the HR contact's own city (existing field).
//        Status       (col 12, column M)  = email send status (always last column).
//        Add "Role_Location" header in column L and move "Status" to column M in your sheet.
//
//  BEFORE RUNNING:
//  ───────────────
//  1. Enable Drive API v2 → Apps Script editor → Services → Drive API → v2 → Add
//  2. Fill all Config sheet keys in Column B
//  3. Place CV PDF in CV_Folder
//  4. Place email template Google Doc in Email_Doc_Folder
//  5. In the "test" sheet: insert "Role_Location" in col L, move "Status" to col M
//  6. Run enrichment first: 📧 Email Recruiter → Enrich Missing Domain & Roles
//  7. Then send emails:     📧 Email Recruiter → Send Emails to All
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
- Location of HR contact: ${hrData.location}
- Location where roles are based: ${hrData.roleLocation}

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


// ════════════════════════════════════════════════════════════
// ── SECTION 8 (NEW): Enrich Domain & Roles via Claude AI ────
//
//  Loops through the "test" sheet and for every row where
//  Domain (col 8) or Roles (col 9) is blank, calls Claude
//  with web_search to fetch:
//    • The company's primary operating domain/industry
//    • Up to 6 current or typical open roles
//
//  Results are written directly back into the sheet.
//  Rows already filled are skipped automatically.
//  A 2-second delay between API calls avoids rate limits.
// ════════════════════════════════════════════════════════════
function enrichDomainAndRoles() {

  const cfg = getConfig();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("test");
  if (!sheet) throw new Error("No sheet named 'test' found!");

  const data = sheet.getDataRange().getValues();

  // Column indexes (0-based for array, +1 for sheet range)
  const COL_COMPANY       = 2;
  const COL_DOMAIN        = 8;
  const COL_ROLES         = 9;
  const COL_ROLE_LOCATION = 11;  // column L
  const COL_STATUS        = 12;  // column M — always last

  let enrichedCount = 0;
  let skippedCount  = 0;
  let failedCount   = 0;

  for (let i = 1; i < data.length; i++) {
    const row          = data[i];
    const company      = (row[COL_COMPANY]       || "").toString().trim();
    const domain       = (row[COL_DOMAIN]        || "").toString().trim();
    const roles        = (row[COL_ROLES]         || "").toString().trim();
    const roleLocation = (row[COL_ROLE_LOCATION] || "").toString().trim();

    // Skip if row is empty
    if (!company) {
      skippedCount++;
      continue;
    }

    // Skip only if ALL three fields are already filled
    if (domain && roles && roleLocation) {
      Logger.log(`Row ${i + 1} (${company}): domain + roles + role location all present — skipping.`);
      skippedCount++;
      continue;
    }

    Logger.log(`Row ${i + 1}: Enriching "${company}"…`);

    try {
      const result = fetchDomainAndRoles(company, cfg.ANTHROPIC_API_KEY);

      // Write each field back ONLY if currently blank — never overwrite existing data
      if (!domain && result.domain) {
        sheet.getRange(i + 1, COL_DOMAIN + 1).setValue(result.domain);
        sheet.getRange(i + 1, COL_DOMAIN + 1).setBackground("#cfe2f3");
      }

      if (!roles && result.roles && result.roles.length > 0) {
        sheet.getRange(i + 1, COL_ROLES + 1).setValue(result.roles.join(", "));
        sheet.getRange(i + 1, COL_ROLES + 1).setBackground("#cfe2f3");
      }

      if (!roleLocation && result.roleLocation) {
        sheet.getRange(i + 1, COL_ROLE_LOCATION + 1).setValue(result.roleLocation);
        sheet.getRange(i + 1, COL_ROLE_LOCATION + 1).setBackground("#cfe2f3");
      }

      enrichedCount++;
      Logger.log(`  ✓ Domain: ${result.domain} | Roles: ${result.roles.join(", ")} | Role Location: ${result.roleLocation}`);

    } catch (e) {
      Logger.log(`  ✗ Failed for ${company}: ${e.message}`);
      failedCount++;
    }

    // Pause between API calls
    Utilities.sleep(2000);
  }

  const summary = `✅ Enrichment complete!\nEnriched: ${enrichedCount} | Skipped: ${skippedCount} | Failed: ${failedCount}`;
  Logger.log(summary);
  SpreadsheetApp.getUi().alert(summary);
}


// ── SECTION 8a: Single-company API call (Claude + web search) ─
//
//  Makes one Claude API call with the web_search tool enabled.
//  Returns { domain: string, roles: string[], roleLocation: string }
//  Throws on API error or JSON parse failure.
function fetchDomainAndRoles(companyName, apiKey) {

  const prompt = `You are a company intelligence assistant with access to web search.

Search the web and find the following for the company "${companyName}":
1. The primary industry/domain it operates in (2-4 words max, e.g. "Fintech Payments", "Quick Commerce", "B2B SaaS", "Edtech", "Healthtech", "FMCG", "IT Consulting")
2. Up to 6 specific job roles this company is currently hiring for (based on recent job postings, LinkedIn, or their careers page)
3. The location(s) where these roles are based (e.g. "Bengaluru", "Mumbai, Delhi", "Pan India", "Remote"). If multiple cities, list them comma-separated. If unknown, return "India".

Return ONLY a valid JSON object in this exact format — no markdown, no explanation, no preamble:
{"domain":"<domain>","roles":["<role1>","<role2>","<role3>"],"roleLocation":"<city or cities>"}`;

  const payload = {
    model:      "claude-haiku-4-5-20251001",
    max_tokens: 512,
    tools: [
      {
        type: "web_search_20250305",
        name: "web_search"
      }
    ],
    messages: [{ role: "user", content: prompt }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method:             "post",
    headers: {
      "x-api-key":         apiKey,
      "anthropic-version": "2023-06-01",
      "content-type":      "application/json"
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const raw    = response.getContentText();
  const result = JSON.parse(raw);

  if (result.error) {
    throw new Error(`Claude API error: ${result.error.message}`);
  }

  // Extract the text block from the response (web_search may add tool_use blocks)
  let text = "";
  for (const block of (result.content || [])) {
    if (block.type === "text") text += block.text;
  }

  text = text.trim();

  // Strip markdown code fences if Claude wrapped the JSON anyway
  text = text.replace(/```json/gi, "").replace(/```/g, "").trim();

  // Find the JSON object boundaries robustly
  const start = text.indexOf("{");
  const end   = text.lastIndexOf("}");

  if (start === -1 || end === -1) {
    throw new Error(`No JSON found in Claude response: ${text.substring(0, 200)}`);
  }

  const parsed = JSON.parse(text.slice(start, end + 1));

  return {
    domain:       (parsed.domain       || "").trim(),
    roles:        Array.isArray(parsed.roles) ? parsed.roles.map(r => r.trim()) : [],
    roleLocation: (parsed.roleLocation || "India").trim()
  };
}


// ── SECTION 9: Main — send all emails ────────────────────────
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
  // Roles(9) | Designation(10) | Role_Location(11) | Status(12)
  const COL_NAME          = 1;
  const COL_COMPANY       = 2;
  const COL_EMAIL         = 3;
  const COL_LINKEDIN      = 4;
  const COL_CONTACT       = 5;
  const COL_PERSONAL      = 6;
  const COL_LOCATION      = 7;
  const COL_DOMAIN        = 8;
  const COL_ROLES         = 9;
  const COL_DESIGNATION   = 10;
  const COL_ROLE_LOCATION = 11;  // column L
  const COL_STATUS        = 12;  // column M — always last

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
      roleLocation:   row[COL_ROLE_LOCATION],
    };

    const subject = fillTemplate(cfg["Email_Subject"], values);

    try {
      // Generate AI-personalised email body for this specific HR
      const body = generatePersonalisedEmail({
        name:         row[COL_NAME],
        company:      row[COL_COMPANY],
        domain:       row[COL_DOMAIN],
        roles:        row[COL_ROLES],
        location:     row[COL_LOCATION],
        roleLocation: row[COL_ROLE_LOCATION],
        designation:  row[COL_DESIGNATION]
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


// ── SECTION 10: Custom menu in Google Sheet ──────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📧 Email Recruiter")
    .addItem("① Enrich Missing Domain & Roles", "enrichDomainAndRoles")
    .addSeparator()
    .addItem("② Send Emails to All",            "sendRecruitingEmails")
    .addToUi();
}
