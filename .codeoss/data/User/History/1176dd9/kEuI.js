/**
CONFIGURATION
*/
// Moved configuration to PropertiesService for better management
// Keys to be set in Script Properties:
// MISTRAL_API_KEY, TEMPLATE_DOC_ID_DE, TEMPLATE_DOC_ID_EN, TEMPLATE_DOC_ID_WEB3, PDF_FOLDER_ID

const CONFIG = {
  MISTRAL_MODEL: "mistral-medium-latest",
  TIMEZONE: "Europe/Berlin",
  MY_NAME: "Rey Chancahuaña",
  ALLOWED_PLATFORMS: ["LinkedIn", "Cryptojobslist", "Indeed", "Cryptocurrencyjobs", "Web3career", "Stepstone"],
  PLATFORM_DOMAINS: {
    "linkedin.com": "LinkedIn",
    "cryptojobslist.com": "Cryptojobslist",
    "indeed.com": "Indeed",
    "cryptocurrencyjobs.co": "Cryptocurrencyjobs",
    "web3career.com": "Web3career",
    "stepstone.com": "Stepstone",
    "bybit.com": "Own website",
  }
};

// ─── FIX ISSUE 2 ─────────────────────────────────────────────────────────────
// Central helper: replaces any "remote" or "home office" variant with "Sassnitz"
function normalizeLocation(city) {
  if (!city) return "";
  const cleaned = city.trim();
  if (/^(remote|home\s*office|homeoffice|home-office)$/i.test(cleaned)) {
    return "Sassnitz";
  }
  return cleaned;
}
// ─────────────────────────────────────────────────────────────────────────────

function getScriptProperty(key) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    return value || null;
  } catch (e) {
    return null;
  }
}

function getTemplateDE()  { return getScriptProperty('TEMPLATE_DOC_ID_DE'); }
function getTemplateEN()  { return getScriptProperty('TEMPLATE_DOC_ID_EN'); }
function getTemplateWeb3(){ return getScriptProperty('TEMPLATE_DOC_ID_WEB3'); }
function getPdfFolder()   { return getScriptProperty('PDF_FOLDER_ID'); }
function getMistralKey()  { return getScriptProperty('MISTRAL_API_KEY'); }


/*
AUTOMATIC STATUS TRACKER
*/
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const col   = range.getColumn();
  const row   = range.getRow();
  const newValue = range.getValue();

  if ((col === 5 || col === 6 || (col >= 10 && col <= 18)) && row > 1) {
    if (col === 6) {
      if (newValue === "" || newValue === null) {
        sheet.getRange(row, 9, 1, 10).clearContent();
      } else {
        updateRowStatusLogic(sheet, row, newValue);
        if (newValue === "Rejected") {
          const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "dd.MM.yyyy");
          sheet.getRange(row, 20).setValue(`${dateStr} 🙅🏽‍♂️`);
        } else {
          sheet.getRange(row, 20).clearContent();
        }
      }
    }
    SpreadsheetApp.flush();
    updateSankeyData();
    updateGeoData();
    SpreadsheetApp.flush();
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🤖 AI Recruitment')
    .addItem('Open AI Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Scan Gmail for Applications', 'processGmailApplications')
    .addItem('Scan Gmail for Rejections', 'processRejectionEmails')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('AI Recruitment Suite').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}


/*
CORE PROCESSOR for manual input via Sidebar
*/
function mainJobProcessor(jdInput, cvType) {
  try {
    const sheet = getOrCreateMonthlyTab();
    const isDe = cvType.includes("DE");

    let templateDocId;
    if (cvType === "DE Web2 Marketing Manager") {
      templateDocId = getTemplateDE();
    } else if (cvType === "EN Web2 Marketing Manager") {
      templateDocId = getTemplateEN();
    } else if (cvType === "Web3 Marketing Manager") {
      templateDocId = getTemplateWeb3();
    } else {
      throw new Error(`Unknown CV type selected: ${cvType}`);
    }
    const templateText = DocumentApp.openById(templateDocId).getBody().getText();

    const cleanedJD = jdInput.replace(/[^\x20-\x7E\n]/g, '').substring(0, 5000);

    const signOff = isDe ? "Mit freundlichen Grüßen" : "Best regards";
    const availabilityText = isDe
      ? "Ich bin bereit umzuziehen (falls erforderlich) und stehe kurzfristig mit einer Kündigungsfrist von einer Woche zur Verfügung."
      : "I am fully open to relocation if required and am available to start within a one-week notice period.";

    const prompt = `You are an expert career coach. Output format MUST be exactly: MATCH | COMPANY | POSITION | PLATFORM | CITY | SMART_LOC | SALARY | LETTER

RULES FOR OUTPUT FIELDS:
- MATCH: one of: 🚀 Web3, M4, M3, M2, M1, M0
- COMPANY: company name only
- POSITION: job title only
- PLATFORM: one of [LinkedIn, Cryptojobslist, Indeed, Cryptocurrencyjobs, Web3career, Stepstone] or "Own website"
- CITY: city name only
- SMART_LOC: "Remote", "Hybrid", or "On-site"
- SALARY: If salary range found in JD, return it as-is. If not found but city is known, estimate a realistic market rate for the role and city (e.g. "~€40,000–50,000"). If unknown, return empty.
- LETTER: the full cover letter text (plain text only, zero markdown, zero asterisks)

CURRENT YEAR: 2026.
LANGUAGE: Write the entire cover letter in ${isDe ? "German" : "English"}. Do not mix languages under any circumstance.
CV PROFILE: ${templateText}
JOB DESCRIPTION: ${cleanedJD}

COVER LETTER RULES:
1. Start DIRECTLY with the salutation (e.g. "Sehr geehrte..." or "Dear...").
2. Write in a natural, confident, human tone. Do NOT mirror or parrot phrases from the JD. Draw from the CV to tell a story.
3. Highlight 2-3 specific achievements from the CV that are most relevant. Use concrete numbers.
4. Show genuine interest in the company mission, not just the role requirements.
5. Keep it to 4 paragraphs maximum.
6. Do NOT mention salary, availability, notice period, relocation, or any closing signature.
7. End with the final body paragraph only, no sign-off, no name.
8. PLAIN TEXT ONLY. No markdown, no asterisks, no bold, no bullet points, no dashes used stylistically.
9. Preserve hyphens only in compound words (e.g. "SaaS-Produkt", "e-mail").`;

    const response = callMistralResilient(prompt);
    let parts = response.split("|").map(p => p.trim());

    while (parts.length < 8) {
      parts.splice(parts.length - 1, 0, "");
    }

    let [match, co, pos, plat, city, smartLoc, salary] = parts.slice(0, 7);
    let letterText = parts.slice(7).join(" ");

    letterText = letterText.replace(/```[\s\S]*?```/g, '').trim();
    letterText = letterText.replace(/`/g, '').trim();
    letterText = letterText.replace(/\*\*(.*?)\*\*/g, '$1');
    letterText = letterText.replace(/\*(.*?)\*/g, '$1');

    const closingRegex = /([\n\s]*(Mit freundlichen Grüßen|Best regards)[,]?[\s\S]*)$/i;
    letterText = letterText.replace(closingRegex, '').trim();

    const availabilityDE = /Ich bin bereit umzuziehen[\s\S]*?Verfügung\.?/gi;
    const availabilityEN = /I am fully open to relocation[\s\S]*?notice period\.?/gi;
    const startDE = /Mein Startdatum[\s\S]*?möglich\.?/gi;
    const startEN = /I (can|am able to) start[\s\S]*?notice[\s\S]*?\./gi;
    letterText = letterText.replace(availabilityDE, '').trim();
    letterText = letterText.replace(availabilityEN, '').trim();
    letterText = letterText.replace(startDE, '').trim();
    letterText = letterText.replace(startEN, '').trim();

    letterText = letterText.replace(/—/g, ',');
    letterText = letterText.replace(/–/g, ',');
    letterText = letterText.replace(/\n{3,}/g, '\n\n');

    letterText = `${letterText}\n\n${availabilityText}\n\n${signOff}\n\n${CONFIG.MY_NAME}`;

    let detectedPlatform = plat || "Own website";

    if (jdInput.startsWith("http") && jdInput.includes(".")) {
      try {
        const url = new URL(jdInput);
        const hostname = url.hostname.replace(/^www\./, '');
        if (CONFIG.PLATFORM_DOMAINS[hostname]) {
          detectedPlatform = CONFIG.PLATFORM_DOMAINS[hostname];
        }
      } catch (e) {
        Logger.log(`URL parsing failed for ${jdInput}: ${e.message}`);
      }
    }

    if (detectedPlatform === "Own website" && !CONFIG.ALLOWED_PLATFORMS.includes(plat)) {
      if (!Object.values(CONFIG.PLATFORM_DOMAINS).includes(plat)) {
        detectedPlatform = "Own website";
      } else {
        detectedPlatform = plat;
      }
    } else if (!CONFIG.ALLOWED_PLATFORMS.includes(detectedPlatform) && detectedPlatform !== "Own website") {
      detectedPlatform = "Own website";
    }

    const companyName = co ? co.trim() : "Unknown";
    const position    = pos ? pos.trim() : "Unknown";

    // ─── FIX ISSUE 2: normalize location ─────────────────────────────────────
    const location = normalizeLocation(city);
    // ─────────────────────────────────────────────────────────────────────────

    let cleanedSalary = "";
    if (salary) {
      salary = salary.trim();
      const salaryMatch = salary.match(/([\$£€¥]?\s?(\d{1,3}(?:[.,]\d{3})*|\d+)(?:[.,]\d{2})?)/);
      if (salaryMatch && salaryMatch[1]) {
        cleanedSalary = salaryMatch[1];
      } else {
        cleanedSalary = salary;
        Logger.log(`Potential unparsable salary: "${salary}" for ${companyName}`);
      }
    }

    const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "dd.MM.yyyy");
    let notes = [];
    if (smartLoc && smartLoc !== "") notes.push(`Work: ${smartLoc}`);
    if (cleanedSalary) notes.push(`Salary: ${cleanedSalary}`);
    const finalNotes = notes.join(" | ");

    const rowData = [[match || "M0", companyName, position, detectedPlatform, location, "Applied", dateStr, "", finalNotes]];

    const targetRow = findNextEmptyRow(sheet);
    sheet.getRange(targetRow, 1, 1, 9).setValues(rowData);
    updateRowStatusLogic(sheet, targetRow, "Applied");

    const statusPathFormula = `=JOIN(""; IF(J${targetRow}>0;"📩";""); IF(K${targetRow}>0;"0️⃣";""); IF(L${targetRow}>0;"1️⃣";""); IF(M${targetRow}>0;"2️⃣";""); IF(N${targetRow}>0;"3️⃣";""); IF(O${targetRow}>0;"4️⃣";""); IF(P${targetRow}>0;"🎉";""); IF(Q${targetRow}>0;"⚪";""); IF(R${targetRow}>0;"🛑";""))`;
    sheet.getRange(targetRow, 19).setFormula(statusPathFormula);

    const prefix = isDe ? "Anschreiben Rey" : "Cover letter Rey";
    const tempDocTitle = `${prefix} - ${companyName}`;
    savePdfGhostFree(letterText, companyName, isDe, tempDocTitle);

    // ─── FIX ISSUE 3: flush before AND after, then fresh reads in update fns ─
    SpreadsheetApp.flush();
    Utilities.sleep(1500);
    updateSankeyData();
    updateGeoData();
    SpreadsheetApp.flush();
    // ─────────────────────────────────────────────────────────────────────────

    return `Success: ${companyName} registered as ${match}!`;

  } catch (e) {
    Logger.log(`Error in mainJobProcessor: ${e.toString()}\nStack: ${e.stack}`);
    return e.message ? `Error: ${e.message}` : "An unexpected error occurred. Please check the script logs.";
  }
}


/*
Gmail Application Scanner
*/
function processGmailApplications() {
  const sheet   = getOrCreateMonthlyTab();
  const timezone = CONFIG.TIMEZONE;

  try {
    const lastRow = sheet.getLastRow();
    const existingJobs = lastRow > 1
      ? sheet.getRange("B2:C" + lastRow).getValues()
          .map(row => `${row[0].trim()}-${row[1].trim()}`)
          .filter(e => e !== "-")
      : [];

    const since = new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
    const sinceFormatted = Utilities.formatDate(since, timezone, "yyyy/MM/dd");
    const searchString = `in:inbox after:${sinceFormatted} subject:"Ihre Bewerbung wurde an" subject:"gesendet" -label:LinkedIn-Processed`;

    const threads = GmailApp.search(searchString, 0, 50);

    if (threads.length === 0) {
      SpreadsheetApp.getUi().alert("Keine neuen Bewerbungs-E-Mails in den letzten 24 Stunden gefunden.");
      return;
    }

    let registered = 0;

    for (const thread of threads) {
      const message = thread.getMessages()[0];
      const body    = message.getPlainBody();
      const lines   = body.split('\n').map(l => l.trim()).filter(l => l.length > 0);

      const subject      = message.getSubject();
      const companyMatch = subject.match(/Ihre Bewerbung wurde an (.+?) gesendet/i);
      if (!companyMatch) continue;
      const companyName = companyMatch[1].trim();

      const confirmLineIndex = lines.findIndex(l => l.includes("Ihre Bewerbung wurde an") && l.includes("gesendet"));
      if (confirmLineIndex === -1) continue;

      const jobTitle = lines[confirmLineIndex + 1] || "Unknown Position";

      let city     = "";
      let smartLoc = "";

      const cityLine = lines[confirmLineIndex + 3] || "";
      if (cityLine && !cityLine.startsWith("http") && !cityLine.startsWith("Beworben") && !cityLine.startsWith("Jobangebot")) {
        const cityMatch = cityLine.match(/^(.+?)(?:\s*\((.+?)\))?$/);
        if (cityMatch) {
          city = cityMatch[1].trim()
            .replace(/,.*$/, '')
            .replace(/\s+und\s+Umgebung$/i, '')
            .replace(/\s+and\s+surroundings$/i, '')
            .trim();

          const countryToCity = { "deutschland": "Sassnitz", "germany": "Sassnitz" };
          if (countryToCity[city.toLowerCase()]) city = countryToCity[city.toLowerCase()];

          const workTypeRaw = (cityMatch[2] || "").toLowerCase();
          if      (workTypeRaw.includes("hybrid"))   smartLoc = "Hybrid";
          else if (workTypeRaw.includes("remote"))   smartLoc = "Remote";
          else if (workTypeRaw.includes("vor ort"))  smartLoc = "On-site";
        }
      }

      // ─── FIX ISSUE 2: normalize location ───────────────────────────────────
      city = normalizeLocation(city);
      // ───────────────────────────────────────────────────────────────────────

      const entryKey = `${companyName}-${normalizeJobTitle(jobTitle)}`;
      const normalizedExisting = existingJobs.map(e => {
        const dashIndex     = e.indexOf('-');
        const existingCompany = e.substring(0, dashIndex);
        const existingTitle   = e.substring(dashIndex + 1);
        return `${existingCompany}-${normalizeJobTitle(existingTitle)}`;
      });
      if (normalizedExisting.includes(entryKey)) continue;

      const appDate = Utilities.formatDate(message.getDate(), timezone, "dd.MM.yyyy");

      let salary = "";
      const salaryMatch = body.match(/([\$£€¥]\s?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?|\d{1,3}(?:[.,]\d{3})+\s?(?:EUR|USD|GBP|€|\$|£))/i);
      if (salaryMatch) salary = salaryMatch[0].trim();

      let noteParts = [];
      if (smartLoc) noteParts.push(`Work: ${smartLoc}`);
      if (salary)   noteParts.push(`Salary: ${salary}`);
      const notes = noteParts.length > 0 ? `🤖 ${noteParts.join(" | ")}` : "🤖";

      const targetRow = findNextEmptyRow(sheet);
      sheet.getRange(targetRow, 1, 1, 9).setValues([[
        "M0", companyName, jobTitle, "LinkedIn", city, "Applied", appDate, "", notes
      ]]);
      updateRowStatusLogic(sheet, targetRow, "Applied");

      const statusPathFormula = `=JOIN(""; IF(J${targetRow}>0;"📩";""); IF(K${targetRow}>0;"0️⃣";""); IF(L${targetRow}>0;"1️⃣";""); IF(M${targetRow}>0;"2️⃣";""); IF(N${targetRow}>0;"3️⃣";""); IF(O${targetRow}>0;"4️⃣";""); IF(P${targetRow}>0;"🎉";""); IF(Q${targetRow}>0;"⚪";""); IF(R${targetRow}>0;"🛑";""))`;
      sheet.getRange(targetRow, 19).setFormula(statusPathFormula);

      existingJobs.push(`${companyName}-${normalizeJobTitle(jobTitle)}`);
      registered++;
      thread.addLabel(GmailApp.createLabel('LinkedIn-Processed'));
    }

    if (registered > 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(1500);
      updateSankeyData();
      updateGeoData();
      SpreadsheetApp.flush();
    }

    SpreadsheetApp.getUi().alert(
      registered > 0
        ? `${registered} neue Bewerbung(en) erfolgreich registriert.`
        : "Keine neuen Bewerbungen registriert. Alle gefundenen E-Mails waren bereits vorhanden."
    );

  } catch (e) {
    Logger.log(`Error in processGmailApplications: ${e.toString()}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert("Fehler beim Verarbeiten der Gmail-Bewerbungen. Bitte Logs prüfen.");
  }
}

function normalizeJobTitle(title) {
  return title
    .replace(/\s*\(m\/w\/d\)|\(f\/m\/d\)|\(w\/m\/d\)|\(all genders\)|\(gn\)/gi, '')
    .trim()
    .toLowerCase();
}


/*
PDF Generation
*/
function savePdfGhostFree(letterText, company, isDe, tempDocTitle) {
  const prefix      = isDe ? "Anschreiben Rey" : "Cover letter Rey";
  const pdfFileName = `${prefix} - ${company}.pdf`;
  const pdfFolderId = getPdfFolder();

  if (!pdfFolderId) throw new Error("PDF_FOLDER_ID is not set in Script Properties.");

  const tempDoc = DocumentApp.create(tempDocTitle);
  const tempId  = tempDoc.getId();
  const body    = tempDoc.getBody();

  let cleanText = letterText.trim().replace(/\n{3,}/g, '\n\n');
  body.setText(cleanText);
  const textObj = body.editAsText();
  textObj.setFontSize(11);
  textObj.setFontFamily("Arial");
  tempDoc.saveAndClose();

  const tempFile = DriveApp.getFileById(tempId);
  const pdfBlob  = tempFile.getAs('application/pdf').setName(pdfFileName);

  try {
    DriveApp.getFolderById(pdfFolderId).createFile(pdfBlob);
    Logger.log(`PDF created: ${pdfFileName}`);
  } catch (e) {
    Logger.log(`Error creating PDF: ${e.toString()}`);
    throw new Error("Failed to save PDF to Google Drive. Check PDF_FOLDER_ID.");
  } finally {
    try { tempFile.setTrashed(true); } catch (e) { Logger.log(`Error trashing temp doc: ${e.toString()}`); }
  }
}


/*
Binary Status Columns
*/
function updateRowStatusLogic(sheet, row, newStatus) {
  const statusColumnMap = {
    "Applied": 10, "HR Interview": 11, "1st Interview": 12, "2nd Interview": 13,
    "3rd Interview": 14, "4th Interview": 15, "Offer": 16, "Ignored": 17, "Rejected": 18
  };
  const currentRowValues     = sheet.getRange(row, 10, 1, 9).getValues()[0];
  const targetColIndexInRow  = statusColumnMap[newStatus] - 10;
  if (targetColIndexInRow >= 0 && targetColIndexInRow < currentRowValues.length) {
    if (currentRowValues[targetColIndexInRow] !== 1) currentRowValues[targetColIndexInRow] = 1;
  }
  sheet.getRange(row, 10, 1, 9).setValues([currentRowValues]);
}


/*
Mistral API Caller
*/
function callMistralResilient(prompt) {
  const url     = "https://api.mistral.ai/v1/chat/completions";
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": `Bearer ${getMistralKey()}` },
    payload: JSON.stringify({
      model: CONFIG.MISTRAL_MODEL,
      messages: [
        { role: "system", content: "You are a professional writing tool. Current date is Feb 2026. Never use 2024/2025 in text. Output exactly 8 fields with pipes (|)." },
        { role: "user", content: prompt }
      ],
      temperature: 0.1
    }),
    muteHttpExceptions: true
  };

  for (let i = 0; i < 3; i++) {
    try {
      const res          = UrlFetchApp.fetch(url, options);
      const responseCode = res.getResponseCode();
      const responseBody = res.getContentText();

      if (responseCode === 200) {
        const json    = JSON.parse(responseBody);
        let content   = json.choices[0].message.content.trim();
        if (content.includes("|")) {
          const matchIndex = content.search(/🚀|M[0-4]/);
          if (matchIndex !== -1) {
            content = content.substring(matchIndex).trim();
            const currentPipes  = content.split('|').length - 1;
            const requiredPipes = 7;
            for (let p = 0; p < requiredPipes - currentPipes; p++) content += " |";
            return content;
          }
        }
        Logger.log(`Mistral unexpected format (attempt ${i+1}): ${content}`);
      } else {
        Logger.log(`Mistral API Error (attempt ${i+1}): Code ${responseCode}, Body: ${responseBody}`);
      }
    } catch (error) {
      Logger.log(`Network error (attempt ${i+1}): ${error.message}`);
    }
    Utilities.sleep(2000);
  }
  throw new Error("Mistral API Failure after multiple attempts.");
}


/*
Sheet Helpers
*/
function findNextEmptyRow(sheet) {
  const values = sheet.getRange("B2:B" + sheet.getMaxRows()).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0] || values[i][0] === "") return i + 2;
  }
  return sheet.getMaxRows() + 1;
}

function getOrCreateMonthlyTab() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const month = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "MMM yyyy").replace('.', '');
  let sheet   = ss.getSheetByName(month);

  if (!sheet) {
    sheet = ss.insertSheet(month);
    const headers = [
      "Match Level", "Companies", "Position", "Application Platform", "Location",
      "Status", "Application Date", "Days Posted", "Notes",
      "Applied", "HR Interview", "1st Interview", "2nd Interview", "3rd Interview",
      "4th Interview", "Offer", "Ignored", "Rejected",
      "Status Path", "Email Rejection",
      "Source", "Target", "Count"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#4285f4").setFontColor("white").setFontWeight("bold");

    const emojiFormula = '=JOIN(""; IF(J2>0;"📩";""); IF(K2>0;"0️⃣";""); IF(L2>0;"1️⃣";""); IF(M2>0;"2️⃣";""); IF(N2>0;"3️⃣";""); IF(O2>0;"4️⃣";""); IF(P2>0;"🎉";""); IF(Q2>0;"⚪";""); IF(R2>0;"🛑";""))';
    sheet.getRange("S2:S").setFormula(emojiFormula);

    const lastHeaderCol = headers.length;
    sheet.getRange(2, lastHeaderCol + 1).setFormula('=UNIQUE(E2:E)');
    sheet.getRange(2, lastHeaderCol + 2).setFormula('=ARRAYFORMULA(IF(Y2:Y=""; ""; COUNTIF(E$2:E; Y2:Y)))');

    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
  }
  return sheet;
}

function debugEmailBody() {
  const timezone = CONFIG.TIMEZONE;
  const since    = new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(since, timezone, "yyyy/MM/dd");
  const threads  = GmailApp.search(`in:inbox after:${sinceFormatted} subject:"Ihre Bewerbung wurde an" subject:"gesendet"`, 0, 3);

  if (threads.length === 0) { Logger.log("No threads found."); return; }

  const message = threads[0].getMessages()[0];
  const lines   = message.getPlainBody().split('\n').map(l => l.trim()).filter(l => l.length > 0);
  lines.forEach((line, i) => {
    const charCodes = [...line].map(c => `${c}(U+${c.charCodeAt(0).toString(16).toUpperCase().padStart(4,'0')})`).join(' ');
    Logger.log(`Line ${i}: "${line}"`);
    Logger.log(`Codes: ${charCodes}`);
  });
}


/*
Rejection Email Scanner
*/
function processRejectionEmails() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheets   = ss.getSheets();
  const timezone = CONFIG.TIMEZONE;

  const COMPANY_COL = 1;
  const STATUS_COL  = 5;

  let pendingCompanies = [];
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName !== "Sankey_Data" && sheetName !== "Geo_Data" && /\d{4}/.test(sheetName)) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const status      = data[i][STATUS_COL] ? data[i][STATUS_COL].toString().trim().toLowerCase() : "";
        const companyName = data[i][COMPANY_COL] ? data[i][COMPANY_COL].toString().trim() : "";
        if (companyName && status !== "rejected" && status !== "ignored" && status !== "i withdrew") {
          pendingCompanies.push({ sheet, rowIndex: i + 1, name: companyName });
        }
      }
    }
  });

  if (pendingCompanies.length === 0) return "No pending applications to check.";

  const labelName = 'bot-rejections-processed';
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);

  const since = new Date(new Date().getTime() - 14 * 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(since, timezone, "yyyy/MM/dd");

  const rejectionKeywordsEN = [
  '"unfortunately"',
  '"not moving forward"',
  '"not to move forward"',
  '"other candidates"',
  '"careful consideration"',
  '"careful review"',
  '"we have decided"',
  '"we\'ve decided"',
  '"we will not be moving"',
  '"decided not to move"'
];
  const rejectionKeywordsDE = [
  '"leider"',
  '"absage"',
  '"andere Kandidaten"',
  '"nicht weiterverfolgen"',
  '"keinen positiven Bescheid"',
  '"nicht berücksichtigen"',
  '"haben wir uns entschieden"',
  '"Auswahlprozess"',
  '"Auseinandersetzung"',
  '"Auswahlverfahren"',
  '"nicht weiter im"',
  '"entschieden, Ihre Bewerbung"',
  '"Bewerbende entschieden"',
  '"anderweitig besetzt"',
  '"keine Weiterverfolgung"'
];

  const allKeywords = [...rejectionKeywordsEN, ...rejectionKeywordsDE].join(' OR ');

  // ─── FIX ISSUE 1: scan BOTH unlabeled and already-labeled threads ─────────
  // We run two searches and merge results to avoid missing emails that were
  // labeled in a previous run but never successfully matched to a company.
  const queryFresh   = `after:${sinceFormatted} (${allKeywords}) -label:${labelName}`;
  const queryLabeled = `after:${sinceFormatted} (${allKeywords})  label:${labelName}`;
  const threadsFresh   = GmailApp.search(queryFresh,   0, 30);
  const threadsLabeled = GmailApp.search(queryLabeled, 0, 30);

  // Merge, deduplicate by thread ID
  const seen    = new Set();
  const threads = [];
  for (const t of [...threadsFresh, ...threadsLabeled]) {
    if (!seen.has(t.getId())) { seen.add(t.getId()); threads.push(t); }
  }
  // ─────────────────────────────────────────────────────────────────────────

  let rejectionsFound = 0;
  const dateStr         = Utilities.formatDate(new Date(), timezone, "dd.MM.yyyy");
  const botRejectionMark = `${dateStr} 🤖`;

  
  
  for (const thread of threads) {
    const messages      = thread.getMessages();
    const latestMessage = messages[messages.length - 1];
    const body          = latestMessage.getPlainBody();
    const sender        = latestMessage.getFrom();
    const subject       = latestMessage.getSubject();
    const contentToSearch = sender + " " + subject + " \n" + body;

    

    let matchedCompany = null;

    for (let c = 0; c < pendingCompanies.length; c++) {
      const cName    = pendingCompanies[c].name;
      const coreName = cName
        .split(" (")[0]
        .replace(/\s+(GmbH|AG|SE|KG|OHG|UG|Ltd|Inc|Corp|LLC|SAS|BV|NV|AB)\.?$/i, '')
        .trim();
      const regexPattern = coreName.length <= 4
        ? `(?<![a-zA-ZäöüÄÖÜß])${escapeRegExp_(coreName)}(?![a-zA-ZäöüÄÖÜß])`
        : escapeRegExp_(coreName);
      const regex = new RegExp(regexPattern, "i");
      
      
    }
    
    

    for (let c = 0; c < pendingCompanies.length; c++) {
      const cName    = pendingCompanies[c].name;
      const coreName = cName
  .split(" (")[0]
  .replace(/\s+(GmbH|AG|SE|KG|OHG|UG|Ltd|Inc|Corp|LLC|SAS|BV|NV|AB)\.?$/i, '')
  .trim();

// For very short names (3 chars or fewer like "STI"), require word boundaries
// using a lookahead/lookbehind that handles German characters.
// This prevents "STI" matching inside words like "bestimmt" or "Einstieg".
const regexPattern = coreName.length <= 4
  ? `(?<![a-zA-ZäöüÄÖÜß])${escapeRegExp_(coreName)}(?![a-zA-ZäöüÄÖÜß])`
  : escapeRegExp_(coreName);
const regex = new RegExp(regexPattern, "i");
    }

    if (matchedCompany) {
      matchedCompany.sheet.getRange(matchedCompany.rowIndex, 6).setValue("Rejected");
      matchedCompany.sheet.getRange(matchedCompany.rowIndex, 20).setValue(botRejectionMark);
      updateRowStatusLogic(matchedCompany.sheet, matchedCompany.rowIndex, "Rejected");
      rejectionsFound++;
      thread.addLabel(label);
      pendingCompanies = pendingCompanies.filter(pc => pc !== matchedCompany);
    }
    // ─── FIX ISSUE 1: removed the else-branch that labeled non-matches ───────
    // Non-matching threads are now left unlabeled so future scans can retry them.
    // ─────────────────────────────────────────────────────────────────────────
  }

  if (rejectionsFound > 0) {
    SpreadsheetApp.flush();
    Utilities.sleep(1500);
    updateSankeyData();
    updateGeoData();
    SpreadsheetApp.flush();
    return `Scan complete. Logged ${rejectionsFound} new rejection(s).`;
  }
  return "Scan complete. No new rejections found.";
}

function escapeRegExp_(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


/*
Sankey Data
*/
function updateSankeyData() {
  // ─── FIX ISSUE 3: get a guaranteed-fresh spreadsheet reference ────────────
  const ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  // ─────────────────────────────────────────────────────────────────────────
  const sheets      = ss.getSheets();
  const transitions = {};

  const stages = [
    "Applied", "HR Interview", "1st Interview", "2nd Interview",
    "3rd Interview", "4th Interview", "Offer", "Ignored", "Rejected"
  ];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().replace('.', '');
    if (sheetName.includes("2026")) {
      // ─── FIX ISSUE 3: use getDataRange() to avoid stale getLastRow() ──────
      const dataRange = sheet.getDataRange();
      const lastRow   = dataRange.getNumRows();
      if (lastRow < 2) return;
      const data = dataRange.getValues();
      // ───────────────────────────────────────────────────────────────────────
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[1]) continue;

        let path = ["Applied"];
        for (let j = 1; j < stages.length; j++) {
          if (row[9 + j] === 1) path.push(stages[j]);
        }

        if (path.length > 1) {
          for (let p = 0; p < path.length - 1; p++) {
            const key = `${sheetName}|${path[p]}|${path[p+1]}`;
            transitions[key] = (transitions[key] || 0) + 1;
          }
        } else {
          const key = `${sheetName}|Applied|Pending Response`;
          transitions[key] = (transitions[key] || 0) + 1;
        }
      }
    }
  });

  let sankeySheet = ss.getSheetByName("Sankey_Data");
  if (!sankeySheet) sankeySheet = ss.insertSheet("Sankey_Data");
  sankeySheet.clear();
  sankeySheet.getRange(1, 1, 1, 4).setValues([["Month", "Source", "Target", "Count"]]);

  const output = Object.keys(transitions).map(k => {
    const parts = k.split("|");
    return [parts[0], parts[1], parts[2], transitions[k]];
  });
  if (output.length > 0) {
  const range = sankeySheet.getRange(2, 1, output.length, 4);
  range.setNumberFormat('@STRING@');
  range.setValues(output);
}
}


/*
Geo Data
*/
function updateGeoData() {
  const ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const sheets    = ss.getSheets();
  const geoCounts = {};

  // Build country lookup once, outside the loop
  const countryToCities = {
    "Spain":                ["madrid", "barcelona"],
    "United Kingdom":       ["london"],
    "Czech Republic":       ["praha", "prague"],
    "Cyprus":               ["cyprus", "limasol", "limassol"],
    "Malta":                ["malta"],
    "Latvia":               ["riga"],
    "Estonia":              ["tallinn"],
    "France":               ["paris", "kamrach"],
    "United Arab Emirates": ["dubai"],
    "Portugal":             ["lisbon", "lissabon", "lisboan"],
    "Norway":               ["oslo"],
    "Ireland":              ["dublin"],
    "Austria":              ["vienna", "wien"],
    "Denmark":              ["kopenhagen", "copenhagen"],
    "Belgium":              ["brussels", "brussel", "brüssel", "brüssels"],
    "Netherlands":          ["amsterdam"],
    "Switzerland":          ["zurich", "zürich"],
    "Luxembourg":           ["luxembourg", "luxemburg"],
  };

  // Build reverse lookup once
  const cityToCountry = {};
  Object.entries(countryToCities).forEach(([country, cities]) => {
    cities.forEach(city => { cityToCountry[city] = country; });
  });

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().replace('.', '');
    if (sheetName.includes("2026")) {
      const dataRange = sheet.getDataRange();
      if (dataRange.getNumRows() < 2) return;
      const data = dataRange.getValues();
      for (let i = 1; i < data.length; i++) {
        let city = data[i][4];
        if (city && city.trim() !== "" && city !== "Remote") {
          let cleanCity = city.split('(')[0].trim();
          const matchedCountry = Object.keys(cityToCountry).find(c =>
            cleanCity.toLowerCase().includes(c)
          );
          const country = matchedCountry ? cityToCountry[matchedCountry] : "Germany";
          const key = `${sheetName}|${cleanCity}|${country}`;
          geoCounts[key] = (geoCounts[key] || 0) + 1;
        }
      }
    }
  });

  let geoSheet = ss.getSheetByName("Geo_Data");
  if (!geoSheet) {
    geoSheet = ss.insertSheet("Geo_Data");
  } else {
    geoSheet.clearContents();
    geoSheet.clearFormats();
  }

  const output = Object.keys(geoCounts).map(key => {
    const [month, city, country] = key.split("|");
    return [month, city, country, geoCounts[key]];
  });

  const allRows = [["Month", "City", "Country", "Application Count"], ...output];
  const range = geoSheet.getRange(1, 1, allRows.length, 4);
  range.setNumberFormat('@STRING@');
  range.setValues(allRows);
}
function debugSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => {
    const original = sheet.getName();
    const cleaned = original.replace('.', '');
    Logger.log(`Original: "${original}" → Cleaned: "${cleaned}"`);
  });
}
function debugDataTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const geoSheet = ss.getSheetByName("Geo_Data");
  const geoData = geoSheet.getRange(1, 1, 10, 4).getValues();
  Logger.log("=== GEO_DATA first 10 rows ===");
  geoData.forEach((row, i) => Logger.log(`Row ${i}: "${row[0]}" | "${row[1]}" | "${row[2]}" | "${row[3]}"`));

  const sankeySheet = ss.getSheetByName("Sankey_Data");
  const sankeyData = sankeySheet.getRange(1, 1, 10, 4).getValues();
  Logger.log("=== SANKEY_DATA first 10 rows ===");
  sankeyData.forEach((row, i) => Logger.log(`Row ${i}: "${row[0]}" | "${row[1]}" | "${row[2]}" | "${row[3]}"`));
}
function debugSankeyCycles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sankey_Data");
  const data = sheet.getDataRange().getValues();
  Logger.log("=== CHECKING FOR CYCLES ===");
  for (let i = 1; i < data.length; i++) {
    const source = data[i][1];
    const target = data[i][2];
    if (source === target) {
      Logger.log(`Row ${i+1}: SOURCE "${source}" equals TARGET "${target}" — CYCLE!`);
    }
  }
  Logger.log("=== DONE ===");
}