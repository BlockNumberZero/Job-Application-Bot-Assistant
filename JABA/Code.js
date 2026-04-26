/**
CONFIGURATION
*/
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

// Central helper: replaces any "remote" or "home office" variant with "Sassnitz"
function normalizeLocation(city) {
  if (!city) return "";
  const cleaned = city.trim();
  if (/^(remote|home\s*office|homeoffice|home-office)$/i.test(cleaned)) {
    return "Sassnitz";
  }
  return cleaned;
}

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
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🤖 AI Recruitment')
    .addItem('Open AI Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Scan Gmail for Applications', 'processGmailApplications')
    .addItem('Scan Gmail for Rejections', 'processRejectionEmails')
    .addSeparator()
    .addItem('🔄 Refresh Dashboard Data', 'refreshAllData')
    .addToUi();
}

function refreshAllData() {
  const ui = SpreadsheetApp.getUi();
  try {
    updateSankeyData();
    updateGeoData();
    ui.alert('Success: Dashboard data updated.');
  } catch (e) {
    ui.alert('Error updating data: ' + e.toString());
  }
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AI Recruitment Suite')
    .setWidth(720)
    .setHeight(820);
  SpreadsheetApp.getUi().showModelessDialog(html, 'AI Recruitment Suite');
}


/* ============================================================
   NEW: SKILLS MATCH MAKER (SMM) — Step 1
   Called from sidebar when user clicks a DE/EN/Web3 SMM button.
   Returns a JSON string with 8 skills, scores, and match level.
   ============================================================ */
function analyzeSkillsMatch(jdInput, cvType) {
  try {
    // Load the correct CV template
    let templateDocId;
    if (cvType === "DE Web2 Marketing Manager") {
      templateDocId = getTemplateDE();
    } else if (cvType === "EN Web2 Marketing Manager") {
      templateDocId = getTemplateEN();
    } else if (cvType === "Web3 Marketing Manager") {
      templateDocId = getTemplateWeb3();
    } else {
      throw new Error(`Unknown CV type: ${cvType}`);
    }

    const templateText = DocumentApp.openById(templateDocId).getBody().getText();
    const cleanedJD    = jdInput.replace(/[^\x20-\x7E\n]/g, '').substring(0, 4000);
    const cleanedCV    = templateText.replace(/[^\x20-\x7E\n]/g, '').substring(0, 3000);

    const prompt = `You are a precise recruitment skills analyst. Analyze the job description and CV profile below.

JOB DESCRIPTION:
${cleanedJD}

CV PROFILE:
${cleanedCV}

TASK:
1. Identify the TOP 8 most important skills from the JD, ranked by the JD's own priority (not by CV match).
2. For each skill, score the CV evidence from 0 to 5:
   - 0: Skill not mentioned in CV at all
   - 1: Briefly mentioned, no supporting evidence
   - 2: Mentioned with some context or description
   - 3: Clear experience with a concrete example
   - 4: Strong experience with partial metrics or results
   - 5: Expert-level evidence with quantified achievements (e.g. "+20% revenue")
3. Classify the JD's importance for that skill: "Crucial" (must-have), "Necessary" (important), or "Optional" (nice-to-have).
4. Extract the CV evidence (max 15 words) or write "Not found in CV" if score is 0.
5. Write a practical gap tip (max 12 words) explaining how to improve the score for that skill.

SCORING RULES:
- Base scores ONLY on concrete CV evidence — do NOT award points for keywords alone.
- A score of 0 is valid and expected when a skill is genuinely absent from the CV.
- The sum of all 8 scores is the total_score (max 40).

MATCH LEVEL (based on total_score):
- M0: 0–10
- M1: 11–20
- M2: 21–29
- M3: 30–35
- M4: 36–40

RESPOND WITH ONLY THIS JSON OBJECT (no markdown, no code fences, no explanation):
{
  "skills": [
    {
      "name": "Skill name",
      "score": 0,
      "importance": "Crucial",
      "evidence": "Brief CV quote or Not found in CV",
      "gap_tip": "Practical advice to improve this score"
    }
  ],
  "total_score": 0,
  "match_level": "M0"
}`;

    const url = "https://api.mistral.ai/v1/chat/completions";
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { "Authorization": `Bearer ${getMistralKey()}` },
      payload: JSON.stringify({
        model: CONFIG.MISTRAL_MODEL,
        messages: [
          {
            role: "system",
            content: "You are a precise recruitment skills analyst. Always respond with valid JSON only. No markdown, no code blocks, no explanation whatsoever."
          },
          { role: "user", content: prompt }
        ],
        temperature: 0.1,
        max_tokens: 2500
      }),
      muteHttpExceptions: true
    };

    const response     = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      throw new Error(`Mistral API error ${responseCode}: ${response.getContentText().substring(0, 200)}`);
    }

    const json    = JSON.parse(response.getContentText());
    let   content = json.choices[0].message.content.trim()
                        .replace(/```json|```/g, '').trim();

    const parsed = JSON.parse(content);

    // Server-side validation: recalculate total and match level to prevent AI drift
    if (parsed.skills && parsed.skills.length > 0) {
      parsed.total_score = parsed.skills.reduce((sum, s) => sum + (Number(s.score) || 0), 0);

      if      (parsed.total_score <= 10) parsed.match_level = "M0";
      else if (parsed.total_score <= 20) parsed.match_level = "M1";
      else if (parsed.total_score <= 29) parsed.match_level = "M2";
      else if (parsed.total_score <= 35) parsed.match_level = "M3";
      else                               parsed.match_level = "M4";
    }

    Logger.log(`SMM Analysis complete — ${cvType} | Score: ${parsed.total_score}/40 | Level: ${parsed.match_level}`);
    return JSON.stringify(parsed);

  } catch (e) {
    Logger.log(`Error in analyzeSkillsMatch: ${e.toString()}\nStack: ${e.stack}`);
    return JSON.stringify({ error: e.message || "Unknown error in SMM analysis." });
  }
}


/* ============================================================
   NEW: SMM_Raw_Data sheet — create or retrieve
   ============================================================ */
function getOrCreateSmmRawDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("SMM_Raw_Data");
  if (!sheet) {
    sheet = ss.insertSheet("SMM_Raw_Data");
    const headers = [
      "UID", "Date", "Company", "Position", "CV_Type",
      "Skill_Rank", "Skill_Name", "Match_Score", "JD_Importance",
      "CV_Evidence", "Gap_Tip", "Interview_Reached"
    ];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setBackground("#4285f4")
         .setFontColor("white")
         .setFontWeight("bold");
    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
    Logger.log("SMM_Raw_Data sheet created.");
  }
  return sheet;
}


/* ============================================================
   NEW: Write 8 skill rows to SMM_Raw_Data
   ============================================================ */
function writeSmmRawData(uid, dateStr, company, position, cvType, smmData) {
  try {
    const sheet  = getOrCreateSmmRawDataSheet();
    const skills = smmData.skills || [];

    const rows = skills.map((s, i) => [
      uid,
      dateStr,
      company,
      position,
      cvType,
      i + 1,                     // Skill_Rank (1–8, ordered by JD priority)
      s.name        || "",
      s.score       || 0,
      s.importance  || "",
      s.evidence    || "",
      s.gap_tip     || "",
      false                      // Interview_Reached — updated later by onEdit hook (Phase 2)
    ]);

    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, 12).setValues(rows);
      Logger.log(`SMM_Raw_Data: wrote ${rows.length} rows for UID ${uid} (${company})`);
    }
  } catch (e) {
    Logger.log(`Error in writeSmmRawData: ${e.toString()}`);
    // Non-fatal: do not throw — main registration should still succeed
  }
}


/*
CORE PROCESSOR — Step 2 (Generate & Register)
Now accepts optional smmDataJson from the sidebar (pre-calculated SMM result).
When smmDataJson is provided, match level comes from SMM score (accurate).
When not provided (fallback), Mistral estimates it as before.
*/
function mainJobProcessor(jdInput, cvType, smmDataJson) {
  try {
    const sheet = getOrCreateMonthlyTab();
    const isDe  = cvType.includes("DE");

    // ── Parse pre-calculated SMM data (if provided by sidebar) ──
    let smmData       = null;
    let smmMatchLevel = null;
    if (smmDataJson) {
      try {
        smmData       = JSON.parse(smmDataJson);
        smmMatchLevel = smmData.match_level || null;
      } catch (e) {
        Logger.log(`Could not parse smmDataJson: ${e.message}`);
      }
    }

    // ── Load CV template ──
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

    // ── Mistral prompt — Step 2 focuses ONLY on cover letter + metadata ──
    // MATCH is still in the output format so parsing stays consistent,
    // but we override it with the SMM value when available.
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

    // ── Platform detection (unchanged) ──
    let detectedPlatform = plat || "Own website";
    if (jdInput.startsWith("http") && jdInput.includes(".")) {
      try {
        const url      = new URL(jdInput);
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
    const location    = normalizeLocation(city);

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

    // ── Determine final match level ──
    // Priority: SMM-calculated level > Mistral's guess
    let finalMatch;
    if (smmMatchLevel) {
      // When Web3 CV + high match, use the Web3 emoji indicator
      if (cvType === "Web3 Marketing Manager" && (smmMatchLevel === "M3" || smmMatchLevel === "M4")) {
        finalMatch = "🚀 Web3";
      } else {
        finalMatch = smmMatchLevel;
      }
      Logger.log(`Using SMM match level: ${finalMatch} (score: ${smmData ? smmData.total_score : 'n/a'})`);
    } else {
      // Fallback: use Mistral's MATCH output (legacy behaviour, no SMM run)
      finalMatch = match || "M0";
      Logger.log(`Using Mistral match level (no SMM data): ${finalMatch}`);
    }

    const dateStr  = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "dd.MM.yyyy");
    let   notes    = [];
    if (smartLoc && smartLoc !== "") notes.push(`Work: ${smartLoc}`);
    if (cleanedSalary)               notes.push(`Salary: ${cleanedSalary}`);
    // Append SMM score to notes for quick reference in the sheet
    if (smmData && smmData.total_score !== undefined) {
      notes.push(`SMM: ${smmData.total_score}/40`);
    }
    const finalNotes = notes.join(" | ");

    const rowData   = [[finalMatch, companyName, position, detectedPlatform, location, "Applied", dateStr, "", finalNotes]];
    const targetRow = findNextEmptyRow(sheet);
    sheet.getRange(targetRow, 1, 1, 9).setValues(rowData);
    updateRowStatusLogic(sheet, targetRow, "Applied");

    const statusPathFormula = `=JOIN(""; IF(J${targetRow}>0;"📩";""); IF(K${targetRow}>0;"0️⃣";""); IF(L${targetRow}>0;"1️⃣";""); IF(M${targetRow}>0;"2️⃣";""); IF(N${targetRow}>0;"3️⃣";""); IF(O${targetRow}>0;"4️⃣";""); IF(P${targetRow}>0;"🎉";""); IF(Q${targetRow}>0;"⚪";""); IF(R${targetRow}>0;"🛑";""))`;
    sheet.getRange(targetRow, 19).setFormula(statusPathFormula);

    // ── Generate cover letter PDF ──
    const prefix       = isDe ? "Anschreiben Rey" : "Cover letter Rey";
    const tempDocTitle = `${prefix} - ${companyName}`;
    savePdfGhostFree(letterText, companyName, isDe, tempDocTitle);

    // ── Write SMM skills to SMM_Raw_Data (if SMM was run) ──
    if (smmData && smmData.skills && smmData.skills.length > 0) {
      const uid = Utilities.getUuid();
      writeSmmRawData(uid, dateStr, companyName, position, cvType, smmData);
    }

    // ── Refresh dashboard data ──
    SpreadsheetApp.flush();
    Utilities.sleep(1500);
    updateSankeyData();
    updateGeoData();
    SpreadsheetApp.flush();

    const scoreNote = smmData ? ` (SMM ${smmData.total_score}/40 · ${finalMatch})` : "";
    return `✅ Success: ${companyName} registered${scoreNote}!`;

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

    const since          = new Date(new Date().getTime() - 24 * 60 * 60 * 1000);
    const sinceFormatted = Utilities.formatDate(since, timezone, "yyyy/MM/dd");
    const searchString   = `in:inbox after:${sinceFormatted} subject:"Ihre Bewerbung wurde an" subject:"gesendet" -label:LinkedIn-Processed`;
    const threads        = GmailApp.search(searchString, 0, 50);

    if (threads.length === 0) {
      SpreadsheetApp.getUi().alert("Keine neuen Bewerbungs-E-Mails in den letzten 24 Stunden gefunden.");
      return;
    }

    let registered = 0;

    for (const thread of threads) {
      const message = thread.getMessages()[0];
      const body    = message.getPlainBody();
      const lines   = body.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      const subject = message.getSubject();

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
          if      (workTypeRaw.includes("hybrid"))  smartLoc = "Hybrid";
          else if (workTypeRaw.includes("remote"))  smartLoc = "Remote";
          else if (workTypeRaw.includes("vor ort")) smartLoc = "On-site";
        }
      }

      city = normalizeLocation(city);

      const entryKey = `${companyName}-${normalizeJobTitle(jobTitle)}`;
      const normalizedExisting = existingJobs.map(e => {
        const dashIndex       = e.indexOf('-');
        const existingCompany = e.substring(0, dashIndex);
        const existingTitle   = e.substring(dashIndex + 1);
        return `${existingCompany}-${normalizeJobTitle(existingTitle)}`;
      });
      if (normalizedExisting.includes(entryKey)) continue;

      const appDate = Utilities.formatDate(message.getDate(), timezone, "dd.MM.yyyy");

      let salary      = "";
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
  const currentRowValues    = sheet.getRange(row, 10, 1, 9).getValues()[0];
  const targetColIndexInRow = statusColumnMap[newStatus] - 10;
  if (targetColIndexInRow >= 0 && targetColIndexInRow < currentRowValues.length) {
    if (currentRowValues[targetColIndexInRow] !== 1) currentRowValues[targetColIndexInRow] = 1;
  }
  sheet.getRange(row, 10, 1, 9).setValues([currentRowValues]);
}


/*
Mistral API Caller (resilient, for cover letter + metadata)
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
        { role: "system", content: "You are a professional writing tool. Current date is 2026. Never use 2024/2025 in text. Output exactly 8 fields with pipes (|)." },
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
        const json  = JSON.parse(responseBody);
        let content = json.choices[0].message.content.trim();
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
Pre-filter: quickly exclude obvious non-job emails
*/
function isLikelyJobEmail(sender, subject) {
  const senderLower  = sender.toLowerCase();
  const subjectLower = subject.toLowerCase();

  const nonJobDomains = [
    'amazon', 'ebay', 'paypal', 'netflix', 'spotify', 'apple.com',
    'facebook', 'instagram', 'twitter', 'tiktok', 'youtube', 'reddit',
    'quora', 'medium.com', 'booking.com', 'airbnb', 'expedia', 'trivago',
    'dhl', 'fedex', 'ups', 'hermes-europe', 'dpd', 'gls-group',
    'sparkasse', 'commerzbank', 'ing.de', 'deutsche-bank', 'comdirect',
    'christ.de', 'zalando', 'otto.de', 'aboutyou', 'hm.com',
    'lieferando', 'deliveroo', 'uber', 'mjam'
  ];
  const nonJobSubjects = [
    'newsletter', 'angebot', '% rabatt', 'sale', 'discount',
    'bestellung', 'order confirmation', 'rechnung', 'invoice',
    'digest', 'quora', 'sparangebot', 'nur heute', 'flash sale',
    'deine lieferung', 'your delivery', 'tracking', 'versandbestatigung'
  ];

  if (nonJobDomains.some(d => senderLower.includes(d)))  return false;
  if (nonJobSubjects.some(s => subjectLower.includes(s))) return false;
  return true;
}


/*
Mistral API Caller for rejection detection
*/
function callGeminiForRejection(sender, subject, body, pendingCompanyNames) {
  const mistralKey = getMistralKey();
  if (!mistralKey) {
    Logger.log("MISTRAL_API_KEY not set in Script Properties.");
    return null;
  }

  const url           = "https://api.mistral.ai/v1/chat/completions";
  const truncatedBody = body.substring(0, 1500);

  const prompt = `You are analyzing a job application email.

EMAIL:
From: ${sender}
Subject: ${subject}
Body: ${truncatedBody}

Answer these two questions:
1. Is this a job application rejection email? (The company is saying they will NOT proceed with this candidate)
2. If yes, what is the exact company name that sent this rejection?

IMPORTANT RULES:
- Acknowledgment emails ("we received your application") are NOT rejections
- Interview invitations are NOT rejections
- Job alert emails from Indeed/LinkedIn are NOT rejections
- Only return true if the company explicitly declines to proceed

Respond ONLY with valid JSON:
{"isRejection": true, "companyName": "Exact company name from the email"}
or
{"isRejection": false, "companyName": null}`;

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": `Bearer ${mistralKey}` },
    payload: JSON.stringify({
      model: CONFIG.MISTRAL_MODEL,
      messages: [
        { role: "system", content: "You are a precise email classifier. Always respond with valid JSON only. No markdown, no explanation." },
        { role: "user", content: prompt }
      ],
      temperature: 0.0,
      max_tokens: 100
    }),
    muteHttpExceptions: true
  };

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const response     = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 429) {
        const waitMs = attempt * 5000; // 5s, 10s, 15s
        Logger.log(`Mistral 429 rate limit (attempt ${attempt}) — waiting ${waitMs/1000}s before retry.`);
        Utilities.sleep(waitMs);
        continue;
      }

      if (responseCode !== 200) {
        Logger.log(`Mistral rejection API error ${responseCode}: ${response.getContentText()}`);
        return null;
      }
    const json   = JSON.parse(response.getContentText());
    const text   = json.choices[0].message.content.trim().replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(text);
    if (parsed.companyName === "null" || parsed.companyName === "") {
      parsed.companyName = null;
    }
    return parsed;
  } catch (e) {
      Logger.log(`Mistral rejection call failed (attempt ${attempt}): ${e.message}`);
      if (attempt < 3) Utilities.sleep(3000);
    }
  }
  Logger.log('Mistral rejection: all 3 attempts failed.');
  return null;
}


/*
Fuzzy company name matcher
*/
function findBestCompanyMatch(geminiName, pendingCompanies) {
  if (!geminiName) return null;

  function normalize(name) {
    return name
      .toLowerCase()
      .replace(/\s+(gmbh\s*&\s*co\.?\s*kg|gmbh|ag|se|kg|ohg|ug|ltd|inc|corp|llc|sas|bv|nv|ab)\.?/gi, '')
      .replace(/[&.,\-]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  const normalizedGemini = normalize(geminiName);

  let match = pendingCompanies.find(p => normalize(p.name) === normalizedGemini);
  if (match) return match;

  match = pendingCompanies.find(p => {
    const normalizedSheet = normalize(p.name);
    return normalizedSheet.includes(normalizedGemini) || normalizedGemini.includes(normalizedSheet);
  });
  if (match) return match;

  const geminiWords = normalizedGemini.split(' ').filter(w => w.length > 3);
  match = pendingCompanies.find(p => {
    const sheetWords = normalize(p.name).split(' ').filter(w => w.length > 3);
    const overlap    = geminiWords.filter(w => sheetWords.includes(w));
    return overlap.length >= 2;
  });
  if (match) return match;

  if (geminiWords.length > 0) {
    const firstWord  = geminiWords[0];
    const candidates = pendingCompanies.filter(p => {
      const sheetWords = normalize(p.name).split(' ').filter(w => w.length > 3);
      return sheetWords[0] === firstWord;
    });
    if (candidates.length === 1) return candidates[0];
  }

  return null;
}


/*
Rejection Email Scanner — Mistral powered
*/
function processRejectionEmails() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheets   = ss.getSheets();
  const timezone = CONFIG.TIMEZONE;
  const props    = PropertiesService.getScriptProperties();

  const ts           = props.getProperty('LAST_REJECTION_SCAN');
  const testSince    = ts ? new Date(ts) : new Date(new Date().getTime() - 14 * 24 * 60 * 60 * 1000);
  const testFormatted = Utilities.formatDate(testSince, timezone, "yyyy/MM/dd");
  Logger.log(`Rejection scan — searching after: ${testFormatted}`);

  const COMPANY_COL = 1;
  const STATUS_COL  = 5;

  let pendingCompanies = [];
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName !== "Sankey_Data" && sheetName !== "Geo_Data" && sheetName !== "SMM_Raw_Data" && /\d{4}/.test(sheetName)) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const status      = data[i][STATUS_COL] ? data[i][STATUS_COL].toString().trim().toLowerCase() : "";
        const companyName = data[i][COMPANY_COL] ? data[i][COMPANY_COL].toString().trim() : "";
        if (companyName &&
            !companyName.startsWith("http") &&
            status !== "rejected" &&
            status !== "ignored" &&
            status !== "i withdrew") {
          pendingCompanies.push({ sheet, rowIndex: i + 1, name: companyName });
        }
      }
    }
  });

  if (pendingCompanies.length === 0) return "No pending applications to check.";

  const lastScanStr = props.getProperty('LAST_REJECTION_SCAN');
  let sinceDate;
  if (lastScanStr) {
    sinceDate = new Date(lastScanStr);
  } else {
    sinceDate = new Date(new Date().getTime() - 14 * 24 * 60 * 60 * 1000);
    Logger.log("First run — scanning last 14 days.");
  }
  const sinceFormatted = Utilities.formatDate(sinceDate, timezone, "yyyy/MM/dd");

  const labelName = 'bot-rejections-processed';
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);

  const queryFresh   = `after:${sinceFormatted} -label:${labelName}`;
  const queryLabeled = `after:${sinceFormatted} label:${labelName}`;
  const freshThreads   = GmailApp.search(queryFresh,   0, 100);
  const labeledThreads = GmailApp.search(queryLabeled, 0, 50);

  const seen = new Set();
  const allThreads = [];
  for (const t of [...freshThreads, ...labeledThreads]) {
    if (!seen.has(t.getId())) { seen.add(t.getId()); allThreads.push(t); }
  }

  if (allThreads.length === 0) {
    props.setProperty('LAST_REJECTION_SCAN', new Date().toISOString());
    return "Scan complete. No new emails to process.";
  }

  const jobThreads = allThreads.filter(thread => {
    const msg = thread.getMessages()[thread.getMessages().length - 1];
    return isLikelyJobEmail(msg.getFrom(), msg.getSubject());
  });

  Logger.log(`Total threads: ${allThreads.length} | After pre-filter: ${jobThreads.length}`);

  const MAX_PER_RUN     = 150;
  const threadsToProcess = jobThreads.slice(0, MAX_PER_RUN);

  let rejectionsFound = 0;
  const dateStr       = Utilities.formatDate(new Date(), timezone, "dd.MM.yyyy");
  const botMark       = `${dateStr} 🤖`;

  for (const thread of threadsToProcess) {
    const messages      = thread.getMessages();
    const latestMessage = messages[messages.length - 1];
    let   body          = latestMessage.getPlainBody();
    if (body.includes('Nachricht gekürzt') || body.includes('message has been truncated')) {
      body = latestMessage.getRawContent();
    }
    const sender  = latestMessage.getFrom();
    const subject = latestMessage.getSubject();

    Logger.log(`Processing: "${subject}" from "${sender}"`);
    const result = callGeminiForRejection(sender, subject, body, []);

    if (!result) { Utilities.sleep(300); continue; }

    Logger.log(`Mistral result: isRejection=${result.isRejection}, company="${result.companyName}"`);

    if (result.isRejection && result.companyName) {
      const matchedEntry = findBestCompanyMatch(result.companyName, pendingCompanies);
      if (matchedEntry) {
        matchedEntry.sheet.getRange(matchedEntry.rowIndex, 6).setValue("Rejected");
        matchedEntry.sheet.getRange(matchedEntry.rowIndex, 20).setValue(botMark);
        updateRowStatusLogic(matchedEntry.sheet, matchedEntry.rowIndex, "Rejected");
        rejectionsFound++;
        thread.addLabel(label);
        Logger.log(`✓ Rejection registered: "${matchedEntry.name}"`);
        pendingCompanies = pendingCompanies.filter(p => p !== matchedEntry);
      } else {
        Logger.log(`⚠ Rejection detected but "${result.companyName}" has no sheet match — labeled for review.`);
        thread.addLabel(label);
      }
    }
    Utilities.sleep(2200);
  }

  props.setProperty('LAST_REJECTION_SCAN', new Date().toISOString());

  if (rejectionsFound > 0) {
    SpreadsheetApp.flush();
    Utilities.sleep(1500);
    updateSankeyData();
    updateGeoData();
    SpreadsheetApp.flush();
  }

  const remaining = jobThreads.length - threadsToProcess.length;
  if (remaining > 0) {
    return `Scan complete. Logged ${rejectionsFound} new rejection(s). ${remaining} more emails pending — run again to continue.`;
  }
  return `Scan complete. Logged ${rejectionsFound} new rejection(s).`;
}


/*
Sankey Data
*/
function updateSankeyData() {
  const ss          = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const sheets      = ss.getSheets();
  const transitions = {};

  const stages = [
    "Applied", "HR Interview", "1st Interview", "2nd Interview",
    "3rd Interview", "4th Interview", "Offer", "Ignored", "Rejected"
  ];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().replace('.', '');
    if (sheetName.includes("2026")) {
      const dataRange = sheet.getDataRange();
      const lastRow   = dataRange.getNumRows();
      if (lastRow < 2) return;
      const data = dataRange.getValues();
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
  const ss        = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const sheets    = ss.getSheets();
  const geoCounts = {};

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
    "Finland":              ["helsinki"],
  };

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
          const key     = `${sheetName}|${cleanCity}|${country}`;
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

  const output  = Object.keys(geoCounts).map(key => {
    const [month, city, country] = key.split("|");
    return [month, city, country, geoCounts[key]];
  });
  const allRows = [["Month", "City", "Country", "Application Count"], ...output];
  const range   = geoSheet.getRange(1, 1, allRows.length, 4);
  range.setNumberFormat('@STRING@');
  range.setValues(allRows);
}


// ── Utility / Debug Functions ──────────────────────────────────────────────

function checkScanTimestamp() {
  const props = PropertiesService.getScriptProperties();
  const ts    = props.getProperty('LAST_REJECTION_SCAN');
  Logger.log(`LAST_REJECTION_SCAN = "${ts}"`);
}

function resetScanTimestamp() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('LAST_REJECTION_SCAN');
  Logger.log("Timestamp cleared. Next scan will go back 14 days.");
}

function testGeminiDirectly() {
  const result = callGeminiForRejection(
    "wattfox-jobs@m.personio.de",
    "Deine Bewerbung als AI & Marketing Automation Manager bei WattFox",
    "Hallo Rey, noch einmal vielen Dank. Wir haben uns dieses Mal allerdings für Bewerbende entschieden, deren Erfahrungen besser zur ausgeschriebenen Rolle gepasst haben.",
    ["WattFox", "Rhenus Warehousing Solutions Global GmbH & Co. KG", "STI GmbH"]
  );
  Logger.log(`Gemini test result: ${JSON.stringify(result)}`);
}

function testSpecificRejections() {
  const pendingNames = ["Town & Country Kundenservice GmbH", "TV NORD Systems GmbH & Co. KG",
                        "WeWork", "Rhenus Warehousing Solutions Global GmbH & Co. KG",
                        "Lidl Dienstleistung", "Columbia Road", "koenig.solutions"];
  const testEmails = [
    { sender: "towncountry-jobs@m.personio.de", subject: "TC - Vielen Dank für Ihre Bewerbung als Spezialist CRM", body: "Nach eingehender Prüfung Ihrer Unterlagen müssen wir Ihnen leider mitteilen, dass wir Ihre Bewerbung für diese Position nicht weiterverfolgen können." },
    { sender: "bbidarnariman@tuevnordgroup.recruitmail.com", subject: "Feedback zu deiner Bewerbung als PowerPlatform Manager:in", body: "Leider müssen wir dir mitteilen, dass wir deine Bewerbung im Auswahlverfahren nicht weiter berücksichtigen können." },
    { sender: "wework@myworkday.com", subject: "Thank You, from WeWork", body: "We recently filled the role you originally applied to, so we won't be moving forward with your candidacy for this position." },
    { sender: "rhe@myworkday.com", subject: "Deine Bewerbung für Junior Innovation Manager", body: "Nach sorgfältiger Überlegung bedauern wir, dir mitteilen zu müssen, dass wir deine Bewerbung zu diesem Zeitpunkt nicht weiter verfolgen werden." },
    { sender: "noreply@lidl.com", subject: "Deine Bewerbung als Junior Automation & AI Specialist", body: "Andere Bewerbungen entsprechen dem Stellenprofil jedoch noch etwas besser. Wir bedauern, den Bewerbungsprozess deshalb an dieser Stelle beenden zu müssen." },
    { sender: "hannu.saarinen@columbiaroad.teamtailor-mail.com", subject: "Your application to Columbia Road", body: "we will not be moving forward with your application for now." },
    { sender: "noreply@indeed.com", subject: "Neuigkeiten zu Ihrer Bewerbung von koenig.solutions", body: "Leider konnte Ihre Bewerbung dieses Mal nicht berücksichtigt werden." }
  ];
  testEmails.forEach(email => {
    const result = callGeminiForRejection(email.sender, email.subject, email.body, pendingNames);
    Logger.log(`"${email.subject}" → ${JSON.stringify(result)}`);
  });
}

function checkLabeledEmails() {
  const threads = GmailApp.search('label:bot-rejections-processed', 0, 50);
  Logger.log(`Total labeled threads: ${threads.length}`);
  threads.forEach(t => {
    const msg = t.getMessages()[t.getMessages().length - 1];
    Logger.log(`"${msg.getSubject()}" from "${msg.getFrom()}"`);
  });
}