/**
CONFIGURATION
*/
// Moved configuration to PropertiesService for better management
// Keys to be set in Script Properties:
// MISTRAL_API_KEY, TEMPLATE_DOC_ID_DE, TEMPLATE_DOC_ID_EN, TEMPLATE_DOC_ID_WEB3, PDF_FOLDER_ID

const CONFIG = {
  // Using mistral-medium-latest as requested, but be mindful of its potential issues.
  MISTRAL_MODEL: "mistral-medium-latest", // Set to your preferred model
  TIMEZONE: "Europe/Berlin",
  MY_NAME: "Rey Chancahuaña",
  ALLOWED_PLATFORMS: ["LinkedIn", "Cryptojobslist", "Indeed", "Cryptocurrencyjobs", "Web3career", "Stepstone"],
  // Domain to Platform mapping for better identification
  PLATFORM_DOMAINS: {
    "linkedin.com": "LinkedIn",
    "cryptojobslist.com": "Cryptojobslist",
    "indeed.com": "Indeed",
    "cryptocurrencyjobs.co": "Cryptocurrencyjobs", // Example, adjust if needed
    "web3career.com": "Web3career",
    "stepstone.com": "Stepstone",
    "bybit.com": "Own website", // Example: Bybit might not be directly detectable as a platform link
    // Add more known domains if necessary
  }
};

// Helper function to get script properties safely
function getScriptProperty(key) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    // If it's missing, don't crash the script, just return null
    return value || null;
  } catch (e) {
    return null;
  }
}

// Fetching critical IDs from Script Properties
function getTemplateDE() { return getScriptProperty('TEMPLATE_DOC_ID_DE'); }
function getTemplateEN() { return getScriptProperty('TEMPLATE_DOC_ID_EN'); }
function getTemplateWeb3() { return getScriptProperty('TEMPLATE_DOC_ID_WEB3'); }
function getPdfFolder() { return getScriptProperty('PDF_FOLDER_ID'); }
function getMistralKey() { return getScriptProperty('MISTRAL_API_KEY'); } 


/*
AUTOMATIC STATUS TRACKER
*/
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const col = range.getColumn();
  const row = range.getRow();
  const newValue = range.getValue();

  // Check if the edited cell is in the Status column (F) or the Binary columns (J through R)
  // Column 6 is F. Columns 10-18 are J through R.
  if ((col === 6 || (col >= 10 && col <= 18)) && row > 1) {
    
    // If it was the Status column (F), handle the normal logic
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

    // CRITICAL INTEGRATION: Update the Sankey data tab whenever ANY status changes
    updateSankeyData();
    //CRITICAL INTEGRATION: Update Geo Data tab whenever ANY status changes
    updateGeoData();
  }
}

/**
Creates the custom menu in the Google Sheet UI.
*/
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🤖 AI Recruitment')
    .addItem('Open AI Sidebar', 'showSidebar')
    .addSeparator() // Add a separator for visual distinction
    .addItem('Scan Gmail for Applications', 'processGmailApplications') // New menu item
    .addItem('Scan Gmail for Rejections', 'processRejectionEmails') // <--- NEW LINE ADDED
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
    const isDe = cvType.includes("DE"); // Used for PDF naming prefix

    // --- Determine correct CV template Doc ID based on selection ---
    let templateDocId;
    let templateText;
    if (cvType === "DE Web2 Marketing Manager") {
      templateDocId = getTemplateDE();
    } else if (cvType === "EN Web2 Marketing Manager") {
      templateDocId = getTemplateEN();
    } else if (cvType === "Web3 Marketing Manager") {
      templateDocId = getTemplateWeb3();
    } else {
      throw new Error(`Unknown CV type selected: ${cvType}`);
    }
    templateText = DocumentApp.openById(templateDocId).getBody().getText();

    // --- Clean Job Description ---
    const cleanedJD = jdInput.replace(/[^\x20-\x7E\n]/g, '').substring(0, 5000);

    // --- Construct Prompt for Mistral ---
    const signOff = isDe ? "Mit freundlichen Grüßen" : "Best regards";
    const availabilityText = isDe
      ? "Ich bin bereit umzuziehen (falls erforderlich) und stehe kurzfristig mit einer Kündigungsfrist von einer Woche zur Verfügung."
      : "I am fully open to relocation if required and am available to start within a one-week notice period.";

    // --- REVISED PROMPT for Hyphens and Closing ---
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
    // --- END REVISED PROMPT ---

    // --- Call AI ---
    const response = callMistralResilient(prompt);
    let parts = response.split("|").map(p => p.trim());

    // --- Data Extraction and Validation ---
    // Defensive Padding: Ensure we have at least 8 parts (index 7 for letter text)
    while (parts.length < 8) {
      parts.splice(parts.length - 1, 0, ""); // Insert empty string before the last element
    }

    let [match, co, pos, plat, city, smartLoc, salary] = parts.slice(0, 7); // Extract first 7 parts
    let letterText = parts.slice(7).join(" "); // The rest is the letter content
     // --- Clean the letter text ---
    // Remove markdown code fences
    letterText = letterText.replace(/```[\s\S]*?```/g, '').trim();
    letterText = letterText.replace(/`/g, '').trim();
    letterText = letterText.replace(/\*\*(.*?)\*\*/g, '$1'); // Remove bold markdown
    letterText = letterText.replace(/\*(.*?)\*/g, '$1');     // Remove italic markdown

    // Strip ANY closing signature block the AI may have appended anyway
    const closingRegex = /([\n\s]*(Mit freundlichen Grüßen|Best regards)[,]?[\s\S]*)$/i;
    letterText = letterText.replace(closingRegex, '').trim();

    // Strip stray availability/start-date sentences from the body
    const availabilityDE = /Ich bin bereit umzuziehen[\s\S]*?Verfügung\.?/gi;
    const availabilityEN = /I am fully open to relocation[\s\S]*?notice period\.?/gi;
    const startDE = /Mein Startdatum[\s\S]*?möglich\.?/gi;
    const startEN = /I (can|am able to) start[\s\S]*?notice[\s\S]*?\./gi;
    letterText = letterText.replace(availabilityDE, '').trim();
    letterText = letterText.replace(availabilityEN, '').trim();
    letterText = letterText.replace(startDE, '').trim();
    letterText = letterText.replace(startEN, '').trim();

    // Replace em/en dashes
    letterText = letterText.replace(/—/g, ',');
    letterText = letterText.replace(/–/g, ',');

    // Consolidate excessive newlines
    letterText = letterText.replace(/\n{3,}/g, '\n\n');

    // Append the ONE canonical closing
    letterText = `${letterText}\n\n${availabilityText}\n\n${signOff}\n\n${CONFIG.MY_NAME}`;

    // --- Platform Identification Refinement ---
    let detectedPlatform = plat || "Own website"; // Default if AI missed it

    // 1. Check if the provided JD input was a URL and try to match domain
    if (jdInput.startsWith("http") && jdInput.includes(".")) {
      try {
        const url = new URL(jdInput);
        const hostname = url.hostname.replace(/^www\./, ''); // Remove www. for matching
        if (CONFIG.PLATFORM_DOMAINS[hostname]) {
          detectedPlatform = CONFIG.PLATFORM_DOMAINS[hostname];
        }
      } catch (e) {
        Logger.log(`URL parsing failed for ${jdInput}: ${e.message}`);
      }
    }

    // 2. Validate against allowed list (if not determined by URL)
    if (detectedPlatform === "Own website" && !CONFIG.ALLOWED_PLATFORMS.includes(plat)) {
       if (!Object.values(CONFIG.PLATFORM_DOMAINS).includes(plat)) {
         detectedPlatform = "Own website";
       } else {
         detectedPlatform = plat;
       }
    } else if (!CONFIG.ALLOWED_PLATFORMS.includes(detectedPlatform) && detectedPlatform !== "Own website") {
        detectedPlatform = "Own website";
    }


    // --- Data Cleaning & Validation (Improved) ---
    const companyName = co ? co.trim() : "Unknown";
    const position = pos ? pos.trim() : "Unknown";
    const location = city ? city.trim() : "";

    let cleanedSalary = "";
    if (salary) {
      salary = salary.trim();
      const salaryMatch = salary.match(/([\$£€¥]?\s?(\d{1,3}(?:[.,]\d{3})*|\d+)(?:[.,]\d{2})?)/);
      if (salaryMatch && salaryMatch[1]) {
        cleanedSalary = salaryMatch[1];
      } else {
         cleanedSalary = salary;
         Logger.log(`Potential unparsable salary detected: "${salary}" for company ${companyName}`);
      }
    }

    // --- Prepare Row Data ---
    const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "dd.MM.yyyy");
    let notes = [];
    if (smartLoc && smartLoc !== "") notes.push(`Work: ${smartLoc}`);
    if (cleanedSalary) notes.push(`Salary: ${cleanedSalary}`);
    const finalNotes = notes.join(" | ");

    const rowData = [
      [match || "M0", companyName, position, detectedPlatform, location, "Applied", dateStr, "", finalNotes]
    ];

    // --- Write to Sheet ---
    const targetRow = findNextEmptyRow(sheet);
    sheet.getRange(targetRow, 1, 1, 9).setValues(rowData); // Write to columns A-I
    updateRowStatusLogic(sheet, targetRow, "Applied"); // Sets column J to 1

    // --- Apply Status Path Formula to Column S ---
    const statusPathFormula = `=JOIN(""; IF(J${targetRow}>0;"📩";""); IF(K${targetRow}>0;"0️⃣";""); IF(L${targetRow}>0;"1️⃣";""); IF(M${targetRow}>0;"2️⃣";""); IF(N${targetRow}>0;"3️⃣";""); IF(O${targetRow}>0;"4️⃣";""); IF(P${targetRow}>0;"🎉";""); IF(Q${targetRow}>0;"⚪";""); IF(R${targetRow}>0;"🛑";""))`;
    sheet.getRange(targetRow, 19).setFormula(statusPathFormula);

    // --- Save PDF ---
    const prefix = isDe ? "Anschreiben Rey" : "Cover letter Rey";
    const tempDocTitle = `${prefix} - ${companyName}`;
    savePdfGhostFree(letterText, companyName, isDe, tempDocTitle);

    return `Success: ${companyName} registered as ${match}!`;

  } catch (e) {
    Logger.log(`Error in mainJobProcessor: ${e.toString()}\nStack: ${e.stack}`);
    let userMessage = "An unexpected error occurred. Please check the script logs for details.";
    if (e.message) {
      userMessage = `Error: ${e.message}`;
    }
    return userMessage;
  }
}


/**
NEW FUNCTION: Scans Gmail for job applications and registers them in the sheet.
*/
function processGmailApplications() {
  const sheet = getOrCreateMonthlyTab();
  const timezone = CONFIG.TIMEZONE;

  try {
    // Get existing entries for duplicate check
    const lastRow = sheet.getLastRow();
    const existingJobs = lastRow > 1
      ? sheet.getRange("B2:C" + lastRow).getValues()
          .map(row => `${row[0].trim()}-${row[1].trim()}`)
          .filter(e => e !== "-")
      : [];

    // Build search query for last 24 hours
    const since = new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
    const sinceFormatted = Utilities.formatDate(since, timezone, "yyyy/MM/dd");
    const searchString = `in:inbox after:${sinceFormatted} subject:"Ihre Bewerbung wurde an" subject:"gesendet" -label:LinkedIn-Processed`;

    const threads = GmailApp.search(searchString, 0, 50);

    if (threads.length === 0) {
      SpreadsheetApp.getUi().alert("Keine neuen Bewerbungs-E-Mails in den letzten 24 Stunden gefunden.");
      return;
    }

    Logger.log(`Found ${threads.length} threads.`);
    let registered = 0;

    for (const thread of threads) {
      const message = thread.getMessages()[0];
      const body = message.getPlainBody();

      // Split into non-empty lines
      const lines = body.split('\n').map(l => l.trim()).filter(l => l.length > 0);

      // Extract company from subject
      const subject = message.getSubject();
      const companyMatch = subject.match(/Ihre Bewerbung wurde an (.+?) gesendet/i);
      if (!companyMatch) {
        Logger.log(`Could not extract company from subject: "${subject}". Skipping.`);
        continue;
      }
      const companyName = companyMatch[1].trim();

      // Find the confirmation line in the body
      const confirmLineIndex = lines.findIndex(l => l.includes("Ihre Bewerbung wurde an") && l.includes("gesendet"));
      if (confirmLineIndex === -1) {
        Logger.log(`Could not find confirmation line in body for: "${companyName}". Skipping.`);
        continue;
      }

      // Job title is the next line after the confirmation line
      const jobTitle = lines[confirmLineIndex + 1] || "Unknown Position";

      // Location line: "Company · City (WorkType)"
      let city = "";
      let smartLoc = "";

      const cityLine = lines[confirmLineIndex + 3] || "";
      if (cityLine && !cityLine.startsWith("http") && !cityLine.startsWith("Beworben") && !cityLine.startsWith("Jobangebot")) {
        const cityMatch = cityLine.match(/^(.+?)(?:\s*\((.+?)\))?$/);
        if (cityMatch) {
          city = cityMatch[1].trim()
          .replace(/,.*$/, '')           // Remove everything after a comma
          .replace(/\s+und\s+Umgebung$/i, '')  // Remove "und Umgebung"
          .replace(/\s+and\s+surroundings$/i, '') // English equivalent
          .trim();

        // Replace country-level locations with the actual city
        const countryToCity = {
          "deutschland": "Sassnitz",
          "germany": "Sassnitz",
          // Add more mappings here if needed in the future
        };
        if (countryToCity[city.toLowerCase()]) {
          city = countryToCity[city.toLowerCase()];
        }
          const workTypeRaw = (cityMatch[2] || "").toLowerCase();
          if (workTypeRaw.includes("hybrid")) smartLoc = "Hybrid";
          else if (workTypeRaw.includes("remote")) smartLoc = "Remote";
          else if (workTypeRaw.includes("vor ort")) smartLoc = "On-site";
        }
      }

      // Duplicate check
      const entryKey = `${companyName}-${normalizeJobTitle(jobTitle)}`;
      const normalizedExisting = existingJobs.map(e => {
        const dashIndex = e.indexOf('-');
        const existingCompany = e.substring(0, dashIndex);
        const existingTitle = e.substring(dashIndex + 1);
        return `${existingCompany}-${normalizeJobTitle(existingTitle)}`;
      });
      if (normalizedExisting.includes(entryKey)) {
        Logger.log(`Duplicate: "${entryKey}". Skipping.`);
        continue;
      }

      // Application date from email
      const appDate = Utilities.formatDate(message.getDate(), timezone, "dd.MM.yyyy");

      // Try to extract salary from the email body
      let salary = "";
      const salaryMatch = body.match(/([\$£€¥]\s?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?|\d{1,3}(?:[.,]\d{3})+\s?(?:EUR|USD|GBP|€|\$|£))/i);
      if (salaryMatch) salary = salaryMatch[0].trim();

      // Build notes
      let noteParts = [];
      if (smartLoc) noteParts.push(`Work: ${smartLoc}`);
      if (salary) noteParts.push(`Salary: ${salary}`);
      const notes = noteParts.length > 0 ? `🤖 ${noteParts.join(" | ")}` : "🤖";

      // Write to sheet
      const targetRow = findNextEmptyRow(sheet);
      sheet.getRange(targetRow, 1, 1, 9).setValues([[
        "M0", companyName, jobTitle, "LinkedIn", city, "Applied", appDate, "", notes
      ]]);
      updateRowStatusLogic(sheet, targetRow, "Applied");

      // Apply Status Path Formula to Column S
      const statusPathFormula = `=JOIN(""; IF(J${targetRow}>0;"📩";""); IF(K${targetRow}>0;"0️⃣";""); IF(L${targetRow}>0;"1️⃣";""); IF(M${targetRow}>0;"2️⃣";""); IF(N${targetRow}>0;"3️⃣";""); IF(O${targetRow}>0;"4️⃣";""); IF(P${targetRow}>0;"🎉";""); IF(Q${targetRow}>0;"⚪";""); IF(R${targetRow}>0;"🛑";""))`;
      sheet.getRange(targetRow, 19).setFormula(statusPathFormula);

      existingJobs.push(`${companyName}-${normalizeJobTitle(jobTitle)}`);
      registered++;
      thread.addLabel(GmailApp.createLabel('LinkedIn-Processed'));
      Logger.log(`Registered: ${companyName} - ${jobTitle} (${city}, ${smartLoc})`);
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

/**
FIXED PDF GENERATION for manual submissions
*/
function savePdfGhostFree(letterText, company, isDe, tempDocTitle) {
  const prefix = isDe ? "Anschreiben Rey" : "Cover letter Rey";
  const pdfFileName = `${prefix} - ${company}.pdf`;

  const pdfFolderId = getPdfFolder();
  // Ensure the PDF folder ID is set
  if (!pdfFolderId) {
    throw new Error("PDF_FOLDER_ID is not set in Script Properties. Cannot save PDF.");
  }

  const tempDoc = DocumentApp.create(tempDocTitle);
  const tempId = tempDoc.getId();
  const body = tempDoc.getBody();

  let cleanText = letterText.trim();
  cleanText = cleanText.replace(/\n{3,}/g, '\n\n');

  body.setText(cleanText);
  const textObj = body.editAsText();
  textObj.setFontSize(11);
  textObj.setFontFamily("Arial");
  tempDoc.saveAndClose();

  const tempFile = DriveApp.getFileById(tempId);
  const pdfBlob = tempFile.getAs('application/pdf').setName(pdfFileName);

  try {
    DriveApp.getFolderById(pdfFolderId).createFile(pdfBlob);
    Logger.log(`PDF created successfully: ${pdfFileName}`);
  } catch (e) {
    Logger.log(`Error creating PDF file in Drive folder ${pdfFolderId}: ${e.toString()}`);
    throw new Error(`Failed to save PDF to Google Drive. Ensure PDF_FOLDER_ID is correct and the folder exists.`);
  } finally {
    // Clean up the temporary document
    try {
      tempFile.setTrashed(true);
    } catch (e) {
      Logger.log(`Error trashing temporary document ${tempId}: ${e.toString()}`);
    }
  }
}


// Function to UPDATE the binary status columns (J-R) WITHOUT CLEARING PREVIOUS ONES
function updateRowStatusLogic(sheet, row, newStatus) {
  const statusColumnMap = {
    "Applied": 10, "HR Interview": 11, "1st Interview": 12, "2nd Interview": 13,
    "3rd Interview": 14, "4th Interview": 15, "Offer": 16, "Ignored": 17, "Rejected": 18
  };

  const currentRowValues = sheet.getRange(row, 10, 1, 9).getValues()[0]; // Get values for row J to R

  const targetColIndexInRow = statusColumnMap[newStatus] - 10; // Adjust index for the 0-based array (10th col is index 0)

  if (targetColIndexInRow >= 0 && targetColIndexInRow < currentRowValues.length) {
    // Only set to 1 if it's not already 1, to preserve history if needed, though for 'Applied' it's usually a fresh start.
    if (currentRowValues[targetColIndexInRow] !== 1) {
      currentRowValues[targetColIndexInRow] = 1;
    }
  }
  sheet.getRange(row, 10, 1, 9).setValues([currentRowValues]);
}


/*
MISTRAL API CALLER (for manual submissions only)
*/
function callMistralResilient(prompt) {
  const url = "https://api.mistral.ai/v1/chat/completions";
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
      const res = UrlFetchApp.fetch(url, options);
      const responseCode = res.getResponseCode();
      const responseBody = res.getContentText();

      if (responseCode === 200) {
        const json = JSON.parse(responseBody);
        let content = json.choices[0].message.content.trim();

        if (content.includes("|")) {
          const matchIndex = content.search(/🚀|M[0-4]/);
          if (matchIndex !== -1) {
             content = content.substring(matchIndex).trim();
             const currentPipes = content.split('|').length - 1;
             const requiredPipes = 7; // Expecting 7 pipes for 8 fields
             for (let p = 0; p < requiredPipes - currentPipes; p++) {
               content += " |";
             }
             return content;
          }
        }
        Logger.log(`Mistral returned unexpected format (attempt ${i+1}): ${content}`);

      } else {
        Logger.log(`Mistral API Error (attempt ${i+1}): Code ${responseCode}, Body: ${responseBody}`);
      }
    } catch (error) {
      Logger.log(`Network or fetch error (attempt ${i+1}): ${error.message}`);
    }
    Utilities.sleep(2000);
  }
  throw new Error("Mistral API Failure after multiple attempts.");
}

// Finds the next empty row in column B (Companies)
function findNextEmptyRow(sheet) {
  const dataRange = sheet.getRange("B2:B" + sheet.getMaxRows());
  const values = dataRange.getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0] || values[i][0] === "") {
      return i + 2; // Return row number (1-based index + 2 for header)
    }
  }
  return sheet.getMaxRows() + 1; // If sheet is full, return the next row
}

// Gets the current month's sheet or creates it
function getOrCreateMonthlyTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const month = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "MMM yyyy");
  let sheet = ss.getSheetByName(month);

  if (!sheet) {
    sheet = ss.insertSheet(month);
    // Headers are defined to match the expected columns A-I for rowData in mainJobProcessor
    const headers = [
      "Match Level", "Companies", "Position", "Application Platform", "Location",
      "Status", "Application Date", "Days Posted", "Notes", // Columns A-I
      "Applied", "HR Interview", "1st Interview", "2nd Interview", "3rd Interview", // Columns J-N (Binary Status)
      "4th Interview", "Offer", "Ignored", "Rejected", // Columns O-R (Binary Status)
      "Status Path", "Email Rejection", // Columns S-T
      "Source", "Target", "Count" // Columns U-W for Sankey
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#4285f4").setFontColor("white").setFontWeight("bold");

    // Formula for Status Path (Column S)
    const emojiFormula = '=JOIN(""; IF(J2>0;"📩";""); IF(K2>0;"0️⃣";""); IF(L2>0;"1️⃣";""); IF(M2>0;"2️⃣";""); IF(N2>0;"3️⃣";""); IF(O2>0;"4️⃣";""); IF(P2>0;"🎉";""); IF(Q2>0;"⚪";""); IF(R2>0;"🛑";""))';
    sheet.getRange("S2:S").setFormula(emojiFormula);

    // Formulas for Sankey (Columns Y, Z for current setup, if headers end before U-W)
    // Note: If headers end at W, these formulas would start in X, Y. Adjust if needed.
    // Using placeholder Y and Z for unique/count formulas as per original code, assuming these are added AFTER 'Count'
    const geomapUniqueFormula = '=UNIQUE(E2:E)'; // Assuming unique locations in Column E
    const geomapCountFormula = '=ARRAYFORMULA(IF(Y2:Y=""; ""; COUNTIF(E$2:E; Y2:Y)))'; // Count occurrences of unique locations
    const lastHeaderCol = headers.length;
    sheet.getRange(2, lastHeaderCol + 1).setFormula(geomapUniqueFormula); // Column Y
    sheet.getRange(2, lastHeaderCol + 2).setFormula(geomapCountFormula); // Column Z


    sheet.setFrozenRows(1);
    // Auto-resize columns based on header content
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
  return sheet;
}

function debugEmailBody() {
  const timezone = CONFIG.TIMEZONE;
  const since = new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(since, timezone, "yyyy/MM/dd");
  const threads = GmailApp.search(`in:inbox after:${sinceFormatted} subject:"Ihre Bewerbung wurde an" subject:"gesendet"`, 0, 3);
  
  if (threads.length === 0) {
    Logger.log("No threads found.");
    return;
  }

  const message = threads[0].getMessages()[0];
  const body = message.getPlainBody();
  const lines = body.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  lines.forEach((line, i) => {
    // Log each character's Unicode code point so we can see exactly what symbols are used
    const charCodes = [...line].map(c => `${c}(U+${c.charCodeAt(0).toString(16).toUpperCase().padStart(4,'0')})`).join(' ');
    Logger.log(`Line ${i}: "${line}"`);
    Logger.log(`Codes: ${charCodes}`);
  });
}

/**
 * NEW FUNCTION: Scans Gmail for rejection emails, matches them to pending companies,
 * and updates the Status and Date columns.
 */
function processRejectionEmails() {
  const sheet = getOrCreateMonthlyTab();
  const data = sheet.getDataRange().getValues();
  const timezone = CONFIG.TIMEZONE;
  
  // Column Indices (0-based for arrays)
  const COMPANY_COL = 1; // Column B (Companies)
  const STATUS_COL = 5;  // Column F (Status)
  
  // 1. Get pending applications from the current month's sheet
  let pendingCompanies =[];
  for (let i = 1; i < data.length; i++) { // Skip header row
    let status = data[i][STATUS_COL] ? data[i][STATUS_COL].toString().trim().toLowerCase() : "";
    let companyName = data[i][COMPANY_COL] ? data[i][COMPANY_COL].toString().trim() : "";
    
    // Only look for companies where status is not already finalized
    if (companyName && status !== "rejected" && status !== "ignored" && status !== "i withdrew") {
      pendingCompanies.push({
        rowIndex: i + 1, // +1 because sheet rows are 1-based
        name: companyName
      });
    }
  }

  if (pendingCompanies.length === 0) {
    return "No pending applications to check.";
  }

  // 2. Setup Gmail Label
  const labelName = 'bot-rejections-processed';
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }

  // 3. Search Gmail for potential rejections
  // We look back 14 days to keep it fast, but ensuring we don't miss anything.
  const since = new Date(new Date().getTime() - 14 * 24 * 60 * 60 * 1000); 
  const sinceFormatted = Utilities.formatDate(since, timezone, "yyyy/MM/dd");
  
  // NEW — Bilingual, covers real-world German rejection phrasing
const rejectionKeywordsEN = [
  '"unfortunately"',
  '"not moving forward"',
  '"other candidates"',
  '"careful consideration"',
  '"we have decided"',
  '"we will not be moving"'
];

const rejectionKeywordsDE = [
  '"leider"',                        // "unfortunately" — catches iVentureGroup & freenet
  '"absage"',                        // direct rejection noun
  '"andere Kandidaten"',             // "other candidates"
  '"für andere Kandidat"',           // catches "für andere Kandidat*innen" (gendered form)
  '"nicht weiterverfolgen"',         // "not moving forward"
  '"keinen positiven Bescheid"',     // "no positive news" — catches iVentureGroup exactly
  '"eingehender Prüfung"',           // "careful consideration" (German formal)
  '"sorgfältiger Prüfung"',          // alternative formal phrasing — catches freenet exactly
  '"nicht berücksichtigen"',         // "cannot consider your application"
  '"haben wir uns entschieden"'      // "we have decided" (German)
];

const allKeywords = [...rejectionKeywordsEN, ...rejectionKeywordsDE].join(' OR ');
const query = `after:${sinceFormatted} (${allKeywords}) -label:${labelName}`;
  const threads = GmailApp.search(query, 0, 30);
  
  let rejectionsFound = 0;
  const dateStr = Utilities.formatDate(new Date(), timezone, "dd.MM.yyyy");
  const botRejectionMark = `${dateStr} 🤖`;

  // 4. Process emails and match to companies
  for (const thread of threads) {
    const messages = thread.getMessages();
    const latestMessage = messages[messages.length - 1];
    const body = latestMessage.getPlainBody();
    const sender = latestMessage.getFrom();
    const subject = latestMessage.getSubject();
    
    // We search across sender, subject, and body for the company name
    const contentToSearch = sender + " " + subject + " \n" + body;
    let matchedCompany = null;

    for (let c = 0; c < pendingCompanies.length; c++) {
  const cName = pendingCompanies[c].name;

  // Remove legal suffixes to get the core brand name
  // e.g. "freenet AG" → "freenet", "iVentureGroup GmbH" → "iVentureGroup"
  const coreName = cName
    .split(" (")[0]
    .replace(/\s+(GmbH|AG|SE|KG|OHG|UG|Ltd|Inc|Corp|LLC|SAS|BV|NV|AB)\.?$/i, '')
    .trim();

  // Use a looser match: just check if the core name appears anywhere (case-insensitive)
  // This handles camelCase names like "iVentureGroup" that break \b word boundaries
  const regex = new RegExp(escapeRegExp_(coreName), "i");

  if (regex.test(contentToSearch)) {
    matchedCompany = pendingCompanies[c];
    break;
  }
}

    if (matchedCompany) {
      // 5. Update the Sheet
      sheet.getRange(matchedCompany.rowIndex, 6).setValue("Rejected"); // Column F = Status
      sheet.getRange(matchedCompany.rowIndex, 20).setValue(botRejectionMark); // Column T = Email Rejection
      
      // 6. Programmatically trigger the binary columns update
      updateRowStatusLogic(sheet, matchedCompany.rowIndex, "Rejected");
      
      rejectionsFound++;
      thread.addLabel(label);
      
      // Remove company from the pending list so we don't process it twice in the same run
      pendingCompanies = pendingCompanies.filter(pc => pc.rowIndex !== matchedCompany.rowIndex);
    } else {
      // Even if it didn't match, label it so the bot skips it next time
      thread.addLabel(label);
    }
  }

  if (rejectionsFound > 0) {
    return `Scan complete. Logged ${rejectionsFound} new rejection(s).`;
  } else {
    return "Scan complete. No new rejections found.";
  }
}

/**
 * Helper function: Escapes special characters in company names 
 * so they don't break the Regex search.
 */
function escapeRegExp_(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Scans all monthly tabs from 2026 onwards and prepares a Source/Target table.
 */
function updateSankeyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const transitions = {}; // Key format: "Month|Source|Target"

  const stages = [
    "Applied", "HR Interview", "1st Interview", "2nd Interview", 
    "3rd Interview", "4th Interview", "Offer", "Ignored", "Rejected"
  ];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // Only process tabs ending in 2026
    if (sheetName.includes("2026")) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[1]) continue; // Skip empty rows (Company name in Col B)

        let path = ["Applied"];
        for (let j = 1; j < stages.length; j++) {
          if (row[9 + j] === 1) { path.push(stages[j]); }
        }

        if (path.length > 1) {
          for (let p = 0; p < path.length - 1; p++) {
            // Include sheetName (Month) in the key
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
  
  // New Header with "Month"
  sankeySheet.getRange(1, 1, 1, 4).setValues([["Month", "Source", "Target", "Count"]]);
  
  const output = Object.keys(transitions).map(k => {
    const parts = k.split("|"); // [Month, Source, Target]
    return [parts[0], parts[1], parts[2], transitions[k]];
  });

  if (output.length > 0) {
    sankeySheet.getRange(2, 1, output.length, 4).setValues(output);
  }
}
/**
 * Generates city-based data for a Heatmap in Looker Studio.
 */
function updateGeoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const geoCounts = {}; // Key format: "Month|City|Country"

  const international = [
    "Madrid", "Barcelona", "London", "Praha", "Prague", "Cyprus", "Limasol", "Limassol", 
    "Malta", "Riga", "Tallinn", "Paris", "Dubai", "Lisbon", "Lissabon", "Oslo", "Dublin", 
    "Vienna", "Wien", "Kopenhagen", "Copenhagen", "Kamrach", "Brussels", "Brussel", "Brüssel", 
    "Amsterdam", "Zurich", "Zürich", "Luxembourg", "Luxemburg"
  ];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.includes("2026")) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        let city = data[i][4]; // Column E: Location
        if (city && city.trim() !== "" && city !== "Remote") {
          let cleanCity = city.split('(')[0].trim();
          
          let country = "Germany";
          const isInternational = international.some(intl => 
            cleanCity.toLowerCase().includes(intl.toLowerCase())
          );
          
          if (isInternational || cleanCity.toLowerCase().includes("international")) {
            country = "International";
          }

          // ADDED SHEETNAME (MONTH) TO THE KEY
          const key = `${sheetName}|${cleanCity}|${country}`;
          geoCounts[key] = (geoCounts[key] || 0) + 1;
        }
      }
    }
  });

  let geoSheet = ss.getSheetByName("Geo_Data");
  if (!geoSheet) geoSheet = ss.insertSheet("Geo_Data");
  geoSheet.clear();
  
  // NEW HEADERS: Month, City, Country, Count
  geoSheet.getRange(1,1,1,4).setValues([["Month", "City", "Country", "Application Count"]]);
  
  const output = Object.keys(geoCounts).map(key => {
    const [month, city, country] = key.split("|");
    return [month, city, country, geoCounts[key]];
  });

  if (output.length > 0) {
    geoSheet.getRange(2, 1, output.length, 4).setValues(output);
  }
}