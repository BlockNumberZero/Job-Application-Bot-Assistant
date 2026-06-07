/**
CONFIGURATION
*/
const CONFIG = {
  MISTRAL_MODEL: "mistral-medium-latest",
  TIMEZONE: "Europe/Berlin",
  MY_NAME: "Rey Chancahuaña",
  ALLOWED_PLATFORMS: [
  "LinkedIn", "Cryptojobslist", "Indeed", "Cryptocurrencyjobs",
  "Web3career", "Stepstone", "Arbeitsagentur", "Remotive",
  "Jobicy", "Himalayas", "WeWorkRemotely", "EuroJobs"
],
PLATFORM_DOMAINS: {
  "linkedin.com":           "LinkedIn",
  "cryptojobslist.com":     "Cryptojobslist",
  "indeed.com":             "Indeed",
  "cryptocurrencyjobs.co":  "Cryptocurrencyjobs",
  "web3career.com":         "Web3career",
  "stepstone.com":          "Stepstone",
  "stepstone.de":           "Stepstone",
  "arbeitsagentur.de":      "Arbeitsagentur",
  "jobboerse.arbeitsagentur.de": "Arbeitsagentur",
  "remotive.com":           "Remotive",
  "jobicy.com":             "Jobicy",
  "himalayas.app":          "Himalayas",
  "weworkremotely.com":     "WeWorkRemotely",
  "eurojobs.com":           "EuroJobs",
  "bybit.com":              "Own website"
},
};

const PROMPT_VERSIONS = {
  SMM:           'v1.5',  // language_note extracted separately, not scored
  REJECTION:     'v1.2',
  COVER_LETTER:  'v1.3'
};

/**
 * Attempts to parse potentially malformed JSON from Mistral/Groq.
 * Handles: trailing content, truncation at end, truncation mid-object,
 * and extra text before the JSON object.
 */
function repairAndParseSmm(raw) {
  if (!raw) throw new Error('Empty SMM response');
  let content = raw.replace(/```json|```/g, '').trim();

  // Strip any leading text before the first {
  const firstBrace = content.indexOf('{');
  if (firstBrace > 0) content = content.substring(firstBrace);

  // Try direct parse first
  try { return JSON.parse(content); } catch(e) {}

  // Truncated at end — find last complete skill object
  if (!content.endsWith('}')) {
    const lastClose = content.lastIndexOf('},');
    if (lastClose > 0) {
      const repaired = content.substring(0, lastClose + 1) +
        '\n  ],\n  "total_score": 0,\n  "match_level": "M0"\n}';
      try {
        Logger.log('SMM JSON: repaired end-truncation');
        return JSON.parse(repaired);
      } catch(e2) {}
    }
  }

  // Mid-object corruption — try to extract valid skills array
  const skillsMatch = content.match(/"skills"\s*:\s*(\[[\s\S]*?\])/);
  if (skillsMatch) {
    try {
      const skills = JSON.parse(skillsMatch[1]);
      Logger.log(`SMM JSON: extracted ${skills.length} skills from corrupted response`);
      return { skills, total_score: 0, match_level: 'M0' };
    } catch(e3) {}
  }

  throw new Error(`Could not repair SMM JSON. Preview: ${content.substring(0, 100)}`);
}

function getSmmCacheKey(jdText, cvType, cvText) {
  const input = cvType + '|' +
                jdText.replace(/\s+/g, ' ').substring(0, 3000) + '|' +
                (cvText || '').replace(/\s+/g, ' ').substring(0, 500);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
  return 'smm2_' + digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

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
function getGroqKey()     { return getScriptProperty('GROQ_API_KEY'); }
function getTavilyKey()  { return getScriptProperty('TAVILY_API_KEY'); }

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
      // ── Phase 2: flag Interview_Reached in SMM_Raw_Data ──
      const interviewStatuses = ["HR Interview", "1st Interview"];
      if (interviewStatuses.includes(newValue)) {
        const company = sheet.getRange(row, 2).getValue().toString().trim();
        const appDate = sheet.getRange(row, 7).getValue();
        flagSmmInterviewReached(company, appDate);
      }
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
    .addItem('Open AI Sidebar',                         'showSidebar')
    .addSeparator()
    .addItem('Scan Gmail for Applications',             'processGmailApplications')
    .addItem('Scan Gmail for Rejections',               'processRejectionEmails')
    .addSeparator()
    .addItem('📧 Process Indeed Job Alerts',            'processIndeedAlertEmails_Phase1')
    .addItem('📧 Process BA Job Alerts',                'processArbeitsagenturAlertEmails_Phase1')
    .addSeparator()
    .addItem('🔄 Refresh Dashboard Data',               'refreshAllData')
    .addItem('🧠 Refresh SMM Categories (min. 20 apps)', 'batchRefreshMasterCategories')
    .addSeparator()
    .addItem('🔍 Run Daily Job Search (Phase 1)',       'runJobSearchPhase1')
    .addSeparator()
    .addItem('🗑️ Clear Job Search Cache (run once)',   'clearAllJobCache')
    .addItem('🔔 Setup Daily Notification (15:00)',     'createDailyNotificationTrigger')
    .addSeparator()
    .addItem('🔍 Audit Cache Quality',                  'auditCacheQuality')
    .addItem('🎨 Format Job Cache Colors',              'applyJobCacheFormatting')
    .addToUi();

  checkCombinedOpenNotifications();
}

/**
 * Shows a combined startup popup for:
 *  1. Unread M2+ jobs in Alert_Results (marks them read after display)
 *  2. Pending notification text stored by sendDailyNotificationSummary
 * Called by onOpen(). Defined at top level so it can be tested independently.
 */
function checkCombinedOpenNotifications() {
  const messages = [];
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Alert_Results');
    if (sheet && sheet.getLastRow() >= 2) {
      const data   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
      const unread = [];
      data.forEach(function(row, i) {
        const n = parseInt((row[7] || '').toString().replace(/\D/g, '')) || 0;
        if (n >= 2 && !row[11]) {
          unread.push({
            rowIndex: i + 2,
            company:  row[2],
            title:    row[3],
            level:    row[7],
            score:    row[6],
            source:   row[1]
          });
        }
      });
      if (unread.length > 0) {
        const lines = unread.map(function(j) {
          return (j.source === 'BA' ? '🏛' : '🔔') + ' ' +
                 j.company + ' — ' + j.title +
                 ' (' + j.level + ': ' + j.score + '/40)';
        }).join('\n');
        messages.push(
          '🚨 ' + unread.length + ' new M2+ Alert Job' + (unread.length > 1 ? 's' : '') +
          ':\n' + lines +
          '\n\nSee Alert_Results tab for links.'
        );
        unread.forEach(function(j) {
          sheet.getRange(j.rowIndex, 12).setValue(true);
        });
        SpreadsheetApp.flush();
      }
    }
  } catch(e) {
    Logger.log('combinedNotif alerts: ' + e.message);
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const text  = props.getProperty('PENDING_NOTIFICATION_TEXT');
    if (text) {
      messages.push(text + '\n\nOpen Job_Search_Cache → mark "fit" → register via sidebar.');
      props.deleteProperty('PENDING_NOTIFICATION_TEXT');
    }
  } catch(e) {
    Logger.log('combinedNotif pending: ' + e.message);
  }
  if (messages.length === 0) return;
  SpreadsheetApp.getUi().alert(
    '📋 JABA Updates',
    messages.join('\n\n──────\n\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function refreshAllData() {
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();
  try {
    updateSankeyData();
       updateGeoData();
       updateInterviewGeoData();
       updateWeeklyData();
       if (ui) ui.alert('Success: Dashboard data updated (including weekly trend).');
  } catch (e) {
    Logger.log('refreshAllData error: ' + e.toString());
    if (ui) ui.alert('Error updating data: ' + e.toString());
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
function analyzeSkillsMatch(jdInput, cvType, runStart, timeBudgetMs) {
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
    const cleanedJD = focusJdContent(
         jdInput.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
       ).substring(0, 6000);

// Guard: JD too short for meaningful SMM analysis
if (cleanedJD.length < 300) {
  Logger.log(`SMM skipped — JD too short: ${cleanedJD.length} chars`);
  return JSON.stringify({ error: 'JD too short for SMM analysis (< 300 chars). Try fetching the full job page.' });
}
        // Check cache before calling Mistral
    const cacheKey = getSmmCacheKey(cleanedJD, cvType, templateText);
    const scriptCache = CacheService.getScriptCache();
    const cached      = scriptCache.get(cacheKey);
    if (cached) {
      Logger.log(`SMM cache hit — skipping Mistral call`);
      return cached;
    }
    const cleanedCV    = templateText.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').substring(0, 3000);

    // Load existing Master_Category values to keep new ones consistent
    const existingCategories = getExistingMasterCategories();
    const categoryContext = existingCategories.length > 0
      ? `\nEXISTING MASTER CATEGORIES (reuse these before creating new ones):\n${existingCategories.join(', ')}\n`
      : '';

    const prompt = `You are a precise recruitment skills analyst. Analyze the job description and CV profile below.

JOB DESCRIPTION:
${cleanedJD}

CV PROFILE:
${cleanedCV}
${categoryContext}

TASK:
1. Identify the TOP 8 most important skills from the JD, ranked strictly by the JD's own emphasis (most repeated / most described first).
   2. For each skill score the CV from 0 to 5 using ONLY these exact criteria — no interpretation allowed:
   - 0: The exact skill or a direct synonym does NOT appear anywhere in the CV.
   - 1: The skill word appears once with no context (e.g. listed in a tools section only).
   - 2: The skill is described in 1 sentence but with no measurable outcome.
   - 3: The skill is described with a specific project or campaign as context.
   - 4: The skill is described with a specific project AND at least one number (%, €, users, etc.).
   - 5: The skill is described with a specific project AND two or more quantified outcomes.
   RULE: When unsure between two scores, always choose the LOWER one. This is mandatory, not optional.
   RULE: Score 0 means the skill is explicitly absent from the CV. Give 0 even if a tangentially related skill exists. "Sort of implied" = 0. Never promote 0 to 1 through reasoning or inference.
   RULE: To justify score 1, the exact skill keyword or a direct synonym must appear explicitly in the CV — not a related concept, not an adjacent skill, not implied by a project description.
   RULE: To justify score 2 over 1, there must be a complete sentence about the skill — not a noun in a list, not a tool name in a skills section.
   RULE: To justify score 3 over 2, a specific named project, campaign, or employer context must be mentioned alongside the skill in the same sentence or paragraph.
   RULE: To justify score 4 over 3, at least one concrete number (%, €, users, months, team size) must appear in direct connection with the skill evidence.
   RULE: To justify score 5 over 4, two or more distinct quantified outcomes must appear.
   Applying these rules strictly prevents score inflation. Be conservative, not generous.
   FUNCTIONAL EQUIVALENCE — EXCEPTION TO THE SCORE-0 RULE:
   Apply ONLY when the exact skill and all direct synonyms are completely absent from the CV.

   A. TOOL EQUIVALENCE: Tools within the same group below are interchangeable. If the JD requires a tool absent from the CV but the CV contains a different tool from the same group, do NOT score 0. Apply: 1 equivalent tool in CV → score minimum 1. 2 or more equivalents in CV → score minimum 2. Then also apply depth rules (3–5) normally to the best-described equivalent tool, and use the HIGHER of the count-based or depth-based result.
   GROUPS: [BI & Visualisation: Power BI · Looker Studio · Google Data Studio · Tableau · Qlik] [CRM: Salesforce · HubSpot CRM · Pipedrive · Zoho CRM · Microsoft Dynamics] [Marketing Automation: Marketo · Pardot · HubSpot Marketing · Salesforce Marketing Cloud · ActiveCampaign · Klaviyo · Brevo] [Email Marketing: Mailchimp · Klaviyo · Brevo · HubSpot · Campaign Monitor] [Paid Advertising: Google Ads · Meta Ads Manager · LinkedIn Ads · TikTok Ads · Microsoft Ads] [SEO Tools: SEMrush · Ahrefs · Moz · Google Search Console] [Social Tools: Hootsuite · Buffer · Sprout Social · Later]

   B. SKILL ADJACENCY: The pairs below share enough core outcomes that outcome evidence of one justifies a score of 1 for the other. This rule converts 0 → 1 ONLY — never higher. Require specific named outcomes in the CV (not just job titles) before applying.
   Copywriting ↔ Content Writing or Social Media Content: only if CV shows CTA writing, conversion copy, ad copy, or persuasive writing outcomes.
   Paid Search ↔ Performance Marketing or Paid Social: only if CV shows ROAS, CPC, conversion rate, or campaign budget outcomes.
   Community Management ↔ Social Media Management: only if CV shows community growth metrics or engagement rate outcomes.
   Email Marketing ↔ CRM or Marketing Automation: only if CV shows open rate, CTR, subscriber count, or campaign automation outcomes.
   Data Analysis/Reporting ↔ BI Tools or Web Analytics: only if CV shows data-driven decisions, performance dashboards, or reporting outcomes.
3. Language requirements — do NOT include any language as one of the 8 skills.
   Instead, analyse the JD and populate language_note with exactly one of:
   - German C2 / Muttersprache / native speaker / verhandlungssicher / muttersprachlich:
     { "text": "German C2 (native) required — candidate holds C1, requirement not met.", "status": "unmet" }
   - German B1 / B2 / C1 / fließend / sehr gute Deutschkenntnisse / gute Deutschkenntnisse / fluent German:
     { "text": "German C1 required — requirement met.", "status": "met" }
   - English required (any phrasing, any level):
     { "text": "English required — requirement met.", "status": "met" }
   - Third language required (Spanish, French, Dutch, Italian, Polish, etc.):
     { "text": "[Language] [level if stated] required — verify this requirement manually.", "status": "other" }
   - No language requirement anywhere in the JD:
     { "text": "No specific language requirement stated.", "status": "none" }
   Priority if multiple apply: "unmet" > "other" > "met" > "none".
4. Classify JD importance: "Crucial" = role cannot be done without it. "Necessary" = strongly preferred. "Optional" = mentioned once or as a plus.
5. Evidence: copy max 10 words verbatim from the CV, or write exactly "Not found in CV".
6. Gap tip: max 10 words, start with a verb (e.g. "Add", "Quantify", "Include").
7. Job location: extract the city and country where this role is physically based.
    - Remote only, no office mentioned → { "city": "Remote", "country": null }
    - City unclear but country stated  → { "city": null, "country": "Germany" }
    - Completely unknown               → { "city": null, "country": null }
   
   SCORING RULES:
- Scores are based ONLY on what is written in the CV — not on assumptions about the candidate.
- Do NOT give partial credit for related skills. Score the specific skill requested.
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
      "gap_tip": "Practical advice to improve this score",
      "master_category": "Overarching skill cluster (e.g. 'CRM & Automation', 'Data Analytics', 'Content Marketing', 'Social Media', 'SEO & Performance', 'Web3 & Blockchain', 'Leadership & Strategy', 'Community Management')"
    }
  ],
  "total_score": 0,
  "match_level": "M0",
  "language_note": {
       "text": "German C2 (native) required — candidate holds C1, requirement not met.",
       "status": "unmet"
     },
     "job_location": {
       "city": "Berlin",
       "country": "Germany"
     }
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
        temperature: 0,
        max_tokens: 3000
      }),
      muteHttpExceptions: true
    };

    // REPLACE this block inside analyzeSkillsMatch (the retry loop + response handling):

let response, responseCode;
const RETRY_DELAYS = [20000, 45000, 90000]; // 20s, 45s, 90s
for (let attempt = 1; attempt <= 3; attempt++) {
  response     = UrlFetchApp.fetch(url, options);
  responseCode = response.getResponseCode();
  if (responseCode === 200) break;
  if (responseCode === 429 || responseCode === 503) {
    const wait = RETRY_DELAYS[attempt - 1];
    // ── Time-budget guard (automated runs only) ──────────────────────────
    if (runStart && timeBudgetMs) {
      const remaining = timeBudgetMs - (Date.now() - runStart);
      if (remaining < wait + 45000) { // need wait + 45s safety buffer
        Logger.log(`⏱ SMM 429 — only ${Math.round(remaining/1000)}s left in budget, aborting retry → Groq`);
        break; // fall through to Groq fallback below
      }
    }
    // ────────────────────────────────────────────────────────────────────
    Logger.log(`Mistral SMM ${responseCode} (attempt ${attempt}) — waiting ${wait/1000}s`);
    Utilities.sleep(wait);
  } else {
    throw new Error(`Mistral API error ${responseCode}: ${response.getContentText().substring(0, 200)}`);
  }
}

// If Mistral is still failing after 3 attempts, try Groq as fallback
if (responseCode !== 200) {
  Logger.log(`Mistral SMM failed after 3 attempts (code ${responseCode}) — trying Groq fallback`);
  const groqRaw = callGroqApi(
    "You are a precise recruitment skills analyst. Always respond with valid JSON only. No markdown, no code blocks, no explanation whatsoever.",
    prompt,
    "llama-3.3-70b-versatile",
    3000
  );
  if (!groqRaw) throw new Error(`Both Mistral and Groq failed for SMM analysis.`);
  let content = groqRaw.replace(/```json|```/g, '').trim();
  const parsed = repairAndParseSmm(content);
  if (!parsed.language_note || typeof parsed.language_note !== 'object') {
    parsed.language_note = { text: 'Language requirement not analyzed.', status: 'none' };
  }
  if (!parsed.language_note.text)   parsed.language_note.text   = 'No language data.';
  if (!parsed.job_location || typeof parsed.job_location !== 'object') parsed.job_location = { city: null, country: null };
  // recalculate totals
  if (parsed.skills && parsed.skills.length > 0) {
    parsed.total_score = parsed.skills.reduce((sum, s) => sum + (Number(s.score) || 0), 0);
    if      (parsed.total_score <= 10) parsed.match_level = "M0";
    else if (parsed.total_score <= 20) parsed.match_level = "M1";
    else if (parsed.total_score <= 29) parsed.match_level = "M2";
    else if (parsed.total_score <= 35) parsed.match_level = "M3";
    else                               parsed.match_level = "M4";
  }
  Logger.log(`Groq fallback SMM — Score: ${parsed.total_score}/40 | Level: ${parsed.match_level}`);
  return JSON.stringify(parsed);
}

    const json    = JSON.parse(response.getContentText());
    let content = json.choices[0].message.content.trim()
                  .replace(/```json|```/g, '').trim();
const parsed = repairAndParseSmm(content);
    if (!parsed.language_note || typeof parsed.language_note !== 'object') {
      parsed.language_note = { text: 'Language requirement not analyzed.', status: 'none' };
    }
    if (!parsed.language_note.text)   parsed.language_note.text   = 'No language data.';
    if (!parsed.job_location || typeof parsed.job_location !== 'object') parsed.job_location = { city: null, country: null };
 
    // Server-side validation: recalculate total and match level to prevent AI drift
    if (parsed.skills && parsed.skills.length > 0) {
      parsed.total_score = parsed.skills.reduce((sum, s) => sum + (Number(s.score) || 0), 0);

      if      (parsed.total_score <= 10) parsed.match_level = "M0";
      else if (parsed.total_score <= 20) parsed.match_level = "M1";
      else if (parsed.total_score <= 29) parsed.match_level = "M2";
      else if (parsed.total_score <= 35) parsed.match_level = "M3";
      else                               parsed.match_level = "M4";
    }

    // ── Location fallback: if JD had no location, search by company name ──
    // companyName is extracted from jdInput only when it was passed as a URL;
    // for plain JD text we don't have a reliable company name here, so
    // location_source stays "jd" and the sidebar/processor enriches it later.
    const hasJdLocation = (parsed.job_location &&
      (parsed.job_location.city || parsed.job_location.country));
    parsed.location_source = hasJdLocation ? 'jd' : 'unknown';

    Logger.log(`SMM Analysis complete — ${cvType} | Score: ${parsed.total_score}/40 | Level: ${parsed.match_level} | Location: ${hasJdLocation ? JSON.stringify(parsed.job_location) : 'none (will enrich later)'}`);
    const resultStr = JSON.stringify(parsed);
scriptCache.put(cacheKey, resultStr, 43200); // cache for 12 hours
return resultStr;

  } catch (e) {
    Logger.log(`Error in analyzeSkillsMatch: ${e.toString()}\nStack: ${e.stack}`);
    return JSON.stringify({ error: e.message || "Unknown error in SMM analysis." });
  }
}

function clearSmmCache() {
  // GAS ScriptCache does not support listing or bulk-clearing all keys.
  // Cache entries expire automatically after 12 hours — no manual clear needed.
  Logger.log('SMM cache expires automatically after 12h. No action required.');
}

function getPhase2Status() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty('PHASE2_STATUS');
    if (!raw) return { running: false };
    const status = JSON.parse(raw);
    // Auto-expire after 12 min — protects against crashes without cleanup
    if (status.startedAt && (Date.now() - status.startedAt) > 720000) {
      props.deleteProperty('PHASE2_STATUS');
      return { running: false };
    }
    // Add ETA: ~45 seconds average per SMM job
    const AVG_SMM_SECONDS = 45;
    const pending          = status.pending || 0;
    const etaMinutes       = Math.max(1, Math.ceil((pending * AVG_SMM_SECONDS) / 60));
    return { ...status, etaMinutes };
  } catch(e) {
    return { running: false };
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
      "CV_Evidence", "Gap_Tip", "Interview_Reached", "Master_Category",
      "Prompt_Version"             // ← new
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
   INDEED ALERT PROCESSOR — keyword filter config
   ============================================================ */
const RELEVANT_KEYWORDS = [
  // English
  'marketing', 'crm', 'growth', 'content', 'social media', 'web3', 'blockchain',
  'community', 'automation', 'digital', 'campaign', 'brand', 'branding', 'seo', 'sea',
  'performance', 'email marketing', 'influencer', 'analytics', 'communications',
  'copywriter', 'storytelling', 'acquisition', 'retention', 'engagement', 'e-commerce',
  // ← new English additions
  'lifecycle', 'martech', 'growth marketer', 'demand generation', 'demand gen',
  'b2b marketing', 'b2c marketing', 'marketing operations', 'marketing automation',
  'digital marketing manager', 'online marketing',
  // German
  'wachstum', 'inhalt', 'gemeinschaft', 'automatisierung', 'kampagne', 'marke',
  'leistung', 'kommunikation', 'öffentlichkeitsarbeit', 'digitalmarketing',
  'onlinemarketing', 'online-marketing', 'markenführung', 'reichweite',
  // ← new German additions
  'wachstumsmarketing', 'kundenbindung', 'lifecycle-marketing'
];

const SKIP_KEYWORDS = [
  // English
  'sales', 'engineer', 'software developer', 'software engineer', 'lawyer',
  'accountant', 'nurse', 'driver', 'warehouse', 'key account', 'key-account',
  'recruiter', 'finance controller', 'electrician', 'plumber', 'mechanic',
  'chef', 'cook', 'cleaner', 'internship', 'intern ', 'fellowship', 'security guard',
  // ← new
  'customer success',
  // German
  'vertrieb', 'verkauf', 'ingenieur', 'softwareentwickler', 'entwickler',
  'rechtsanwalt', 'buchhalter', 'steuerberater', 'krankenschwester', 'pfleger',
  'fahrer', 'lagerarbeiter', 'lagermitarbeiter', 'schlüsselkunde', 'key account',
  'personalvermittler', 'werkstudent', 'werkstudentin', 'praktikum', 'praktikant',
  'praktikantin', 'pflichtpraktikum', 'elektriker', 'klempner', 'mechaniker',
  'koch', 'reinigungskraft'
];

// Domains where fetching always fails — skip immediately
const BLOCKED_FETCH_DOMAINS = [
  'linkedin.com','indeed.com','glassdoor.com','monster.com','jobware.de'
];

// Extracts full text from a known URL via Tavily Extract endpoint
function tavilyExtract(url) {
  const key = getTavilyKey();
  if (!key) { Logger.log('TAVILY_API_KEY not set'); return null; }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      api_key: key,
      urls: [url],
      extract_depth: 'basic'
    }),
    muteHttpExceptions: true
  };

  try {
    const res  = UrlFetchApp.fetch('https://api.tavily.com/extract', options);
    const code = res.getResponseCode();
    if (code !== 200) {
      Logger.log(`Tavily Extract error ${code}: ${res.getContentText().substring(0, 200)}`);
      return null;
    }
    const data   = JSON.parse(res.getContentText());
    const result = data.results && data.results[0];
    if (!result || !result.raw_content) return null;
    const text = result.raw_content.substring(0, 10000);
    Logger.log(`Tavily Extract OK: ${url} (${text.length} chars)`);
    return text;
  } catch (e) {
    Logger.log(`Tavily Extract exception: ${e.message}`);
    return null;
  }
}

// Searches for a job by company + title, then extracts the best result
function tavilySearch(company, jobTitle) {
  const key = getTavilyKey();
  if (!key) { Logger.log('TAVILY_API_KEY not set'); return null; }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      api_key: key,
      query: `${company} ${jobTitle} job`,
      search_depth: 'basic',
      max_results: 5,
      include_raw_content: false,
      exclude_domains: ['linkedin.com', 'xing.com', 'glassdoor.com', 'kununu.com']
    }),
    muteHttpExceptions: true
  };

  try {
    const res  = UrlFetchApp.fetch('https://api.tavily.com/search', options);
    const code = res.getResponseCode();
    if (code !== 200) {
      Logger.log(`Tavily Search error ${code}: ${res.getContentText().substring(0, 200)}`);
      return null;
    }
    const data    = JSON.parse(res.getContentText());
    const results = data.results || [];
    Logger.log(`Tavily Search: ${results.length} results for "${company} — ${jobTitle}"`);

    for (const r of results) {
      const text = tavilyExtract(r.url);
      if (text && text.length > 300) return text;
      Utilities.sleep(500);
    }
    return null; // no partial fallback — partial JDs are useless for SMM
  } catch (e) {
    Logger.log(`Tavily Search exception: ${e.message}`);
    return null;
  }
}

/**
 * Searches Tavily for a job, but validates each result is actually
 * about the right company/role before returning it.
 */
function tavilySearchValidated(company, jobTitle) {
  const key = getTavilyKey();
  if (!key) return null;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      api_key: key,
      query: `"${company}" ${jobTitle} Stelle`,
      search_depth: 'basic',
      max_results: 5,
      include_raw_content: false,
      exclude_domains: ['linkedin.com', 'xing.com', 'glassdoor.com', 'kununu.com', 'navvis.com']
    }),
    muteHttpExceptions: true
  };

  try {
    const res  = UrlFetchApp.fetch('https://api.tavily.com/search', options);
    const code = res.getResponseCode();
    if (code !== 200) return null;

    const data    = JSON.parse(res.getContentText());
    const results = data.results || [];

    for (const r of results) {
      if (!isJobDetailPage(r.url)) {
        Logger.log(`  Tavily result rejected (category page URL): ${r.url}`);
        Utilities.sleep(400);
        continue;
      }
      const text = tavilyExtract(r.url);
      if (!text || !looksLikeJobContent(text)) { Utilities.sleep(400); continue; }
      if (!isJdRelevantToJob(text, company, jobTitle)) {
        Logger.log(`  Tavily result rejected (irrelevant): ${r.url}`);
        Utilities.sleep(400);
        continue;
      }
      Logger.log(`  Tavily search validated OK: ${r.url}`);
      return { text, url: r.url };
    }
    return null;

  } catch (e) {
    Logger.log(`tavilySearchValidated exception: ${e.message}`);
    return null;
  }
}

/* ── Helpers ─────────────────────────────────────────────── */

function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

function isRelevantJobTitle(title) {
  if (!title) return false;
  const lower = title.toLowerCase();
  // Skip takes priority
  if (SKIP_KEYWORDS.some(k => lower.includes(k.toLowerCase()))) return false;
  return RELEVANT_KEYWORDS.some(k => lower.includes(k.toLowerCase()));
}

function detectCvTypeFromText(text) {
  if (!text) return 'EN Web2 Marketing Manager';
  const lower = text.toLowerCase();
  const web3Keys = ['web3','blockchain','crypto','defi','nft','token','dao',
                    'smart contract','solidity','decentralized'];
  if (web3Keys.some(k => lower.includes(k))) return 'Web3 Marketing Manager';
  const deKeys = ['(m/w/d)','(w/m/d)','stellenanzeige','karriere','gehalt',
                  'vollzeit','teilzeit','berufserfahrung','bewerbung'];
  if (deKeys.some(k => lower.includes(k))) return 'DE Web2 Marketing Manager';
  return 'EN Web2 Marketing Manager';
}

function stripHtmlToText(html) {
  // ── Adzuna-specific structural noise removal ──────────────────────────────
  // Remove the country-selector modal (always at the top of Adzuna pages)
  html = html.replace(/<[^>]*id=["'][^"']*country[^"']*["'][^>]*>[\s\S]*?<\/(?:div|section|form)>/gi, ' ');
  // Remove the "Ähnliche Jobs" / "Similar Jobs" carousel section at the bottom
  // This section starts with a heading containing "hnliche Jobs" and runs to end of main content
  html = html.replace(/<[^>]*>[\s]*[ÄA]hnliche Jobs[\s\S]*$/i, ' ');
  html = html.replace(/<[^>]*>[\s]*Similar Jobs[\s\S]*$/i, ' ');
  // Remove "Jetzt ähnliche Jobs per E-Mail erhalten" block and everything after it
  html = html.replace(/Jetzt\s+[äa]hnliche\s+Jobs[\s\S]*$/i, ' ');
  // Remove "Job-E-Mail bestellen" subscription block
  html = html.replace(/Job-E-Mail\s+bestellen[\s\S]*$/i, ' ');
  // ── Standard structural removal ───────────────────────────────────────────
  return html
    // Remove entire semantic sections that are never JD content
    .replace(/<nav[\s\S]*?<\/nav>/gi, ' ')
    .replace(/<header[\s\S]*?<\/header>/gi, ' ')
    .replace(/<footer[\s\S]*?<\/footer>/gi, ' ')
    .replace(/<aside[\s\S]*?<\/aside>/gi, ' ')
    // Remove scripts and styles
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    // Strip remaining tags
    .replace(/<[^>]+>/g, ' ')
    // Decode entities
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ')
    .replace(/&#\d+;/g, ' ')
    // Clean whitespace
    .replace(/[ \t]{3,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}


function extractJobsFromAlertEmail(emailBody) {
  const systemPrompt = `You are a precise data extractor. Respond with valid JSON only. No markdown.`;
  const userPrompt   = `Extract job listings from this Indeed alert email.

EMAIL BODY:
${emailBody.substring(0, 2000)}

Return ONLY a JSON array of job title and company. No URLs. Empty array if none found:
[{"title": "Job Title", "company": "Company Name"}]`;

  const raw = callGroqApi(systemPrompt, userPrompt, 'llama-3.1-8b-instant', 400);
  if (!raw) return [];
  try {
    return JSON.parse(raw.replace(/```json|```/g, '').trim());
  } catch(e) {
    Logger.log(`extractJobsFromAlertEmail parse error: ${e.message}`);
    return [];
  }
}

/**
 * Extracts ALL job listings from an Indeed alert email HTML body.
 * Uses href jk= pattern — reliable regardless of email length or batch size.
 * Replaces Groq-based extractJobsFromAlertEmail.
 */
function extractJobListingsFromHtml(htmlBody) {
  if (!htmlBody) return [];
  const jobs = [];
  const seen = new Set();

  // Normalise HTML entities before parsing
  const html = htmlBody
    .replace(/&amp;/g, '&').replace(/&#38;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&#\d+;/g, ' ').replace(/&nbsp;/g, ' ');

  // Each Indeed job link: <a href="...jk=ID...">Job Title</a>
  const pattern = /href="[^"]*jk=([a-zA-Z0-9]+)[^"]*"[^>]*>([\s\S]{4,150}?)<\/a>/gi;
  let match;

  while ((match = pattern.exec(html)) !== null) {
    const jk       = match[1];
    const rawTitle = match[2].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();

    if (seen.has(jk)) continue;
    if (!rawTitle || rawTitle.length < 4) continue;
    // Skip action/navigation links
    if (/^(bewerb|apply|view|mehr|weiter|vollständig|see |anzeigen|abmeld|unsubscrib|alle jobs|job alert)/i
        .test(rawTitle)) continue;
    if (/^https?:|^\d+$/.test(rawTitle)) continue;

    seen.add(jk);

    // Extract company from the HTML that immediately follows the job title link
    const afterLink = html.substring(match.index + match[0].length,
                                     match.index + match[0].length + 600);
    const company   = extractCompanyFromHtmlContext(afterLink) || 'Unknown';

    jobs.push({
      title:   rawTitle,
      company: company,
      url:     `https://de.indeed.com/viewjob?jk=${jk}`,
      jk:      jk
    });
  }

  Logger.log(`  HTML extraction: ${jobs.length} job(s) found in email`);
  return jobs;
}

/**
 * Extracts the most likely company name from the HTML immediately after a job title.
 */
function extractCompanyFromHtmlContext(contextHtml) {
  const chunks = contextHtml
    .replace(/<[^>]+>/g, '\n')
    .split('\n')
    .map(s => s.trim())
    .filter(s => s.length > 2 && s.length < 80)
    .filter(s => !/^https?:\/\//.test(s))
    .filter(s => !/^\d/.test(s))
    .filter(s => !/^(Vollzeit|Teilzeit|Minijob|Remote|Hybrid|Vor Ort|Homeoffice|Berlin|Hamburg|München|Frankfurt|Köln|Stuttgart|Düsseldorf|Dresden|Leipzig|Deutschland|Germany|Anzeige|Gesponsert|vor \d|seit \d|\+\d)/i
                 .test(s));
  return chunks[0] || null;
}

function buildAlertLabel(smmResult) {
  const score    = smmResult.total_score || 0;
  const level    = parseInt((smmResult.match_level || 'M0').replace(/\D/g, '')) || 0;
  const skills   = smmResult.skills || [];
  const base     = `JABA Alert/M${level}-${score}`;

  if (level === 0) return base; // M0: no indicators

  function indicator(importance, emoji) {
    const group = skills.filter(s => s.importance === importance);
    if (group.length === 0) return '-';
    return group.every(s => (s.score || 0) >= 1) ? emoji : '0';
  }

  const c = indicator('Crucial',   '🟢');
  const n = indicator('Necessary', '🟡');
  const o = indicator('Optional',  '🔵');

  return `${base} ${c}${n}${o}`;
}


/* ============================================================
   DEPRECATED: processIndeedAlertEmails
   Replaced by processIndeedAlertEmails_Phase1 + runPhase2.
   Kept as dead code reference only — do not run or trigger.
   ============================================================ */
function _DEPRECATED_processIndeedAlertEmails() {
  const props    = PropertiesService.getScriptProperties();
  const timezone = CONFIG.TIMEZONE;
  const runStart = Date.now();
  const ALERT_TIME_BUDGET_MS = 280000; // 4.67 minutes
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();

  // Determine scan window
  const lastScanStr    = props.getProperty('LAST_ALERT_SCAN');
  const sinceDate      = lastScanStr
    ? new Date(lastScanStr)
    : new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000);
  const queryDate      = new Date(sinceDate.getTime() - 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(queryDate, timezone, 'yyyy/MM/dd');
  Logger.log(`Indeed alert scan — searching since: ${sinceFormatted}`);

  // Gmail labels
  // Score labels are built dynamically by buildAlertLabel()
  // Only static labels needed upfront:
  const labelInternship = getOrCreateLabel('JABA Alert/🎓 Internship');

  // Search for Indeed job alert emails only — exclude application/rejection threads
  const query = [
    'from:indeed.com',
    `after:${sinceFormatted}`,
    '-label:JABA Alert/Processed',
    '-subject:Bewerbung',
    '-subject:"Neuigkeiten zu Ihrer Bewerbung"',
    '-subject:"Your application"',
    '-subject:indeedapply',
    '-subject:"Heben Sie sich"'
  ].join(' ');

  const threads = GmailApp.search(query, 0, 30);
  Logger.log(`Found ${threads.length} alert thread(s) to process.`);

  if (threads.length === 0) {
    props.setProperty('LAST_ALERT_SCAN', new Date().toISOString());
    if (ui) ui.alert('No new Indeed alert emails found since last scan.');
    return;
  }

  let totalReviewed = 0;
  let totalLow      = 0;
  let totalSkipped  = 0;
  const summaryLines = [];
  const MAX_JOBS_PER_RUN = 10; // stay within 6-min GAS limit
  let jobsProcessed = 0;

  for (const thread of threads) {
    if (jobsProcessed >= MAX_JOBS_PER_RUN) {
      Logger.log(`MAX_JOBS_PER_RUN (${MAX_JOBS_PER_RUN}) reached — run again for remaining emails.`);
      break;
    }

    const message   = thread.getMessages()[thread.getMessages().length - 1];
    const body      = message.getPlainBody();
    const htmlBody  = message.getBody();
    const subject   = message.getSubject();
    Logger.log(`\nEmail: "${subject}"`);

    // Pre-filter by subject line — avoids Groq call for obviously irrelevant emails
    if (!isRelevantJobTitle(subject)) {
  Logger.log(`Subject pre-filtered: "${subject}"`);
  const isInternship = /werkstudent|praktikum|pflichtpraktikum|internship/i.test(subject);
  if (isInternship) {
    thread.addLabel(labelInternship);
  }
  thread.addLabel(getOrCreateLabel('JABA Alert/Processed'));
  continue;
}

    const allJobs      = extractJobListingsFromHtml(htmlBody);
const relevantJobs = allJobs.filter(j => isRelevantJobTitle(j.title));
Logger.log(`Jobs in email: ${allJobs.length} | After keyword filter: ${relevantJobs.length}`);

    if (relevantJobs.length === 0) {
  thread.addLabel(getOrCreateLabel('JABA Alert/Processed'));
  continue;
}

    const threadResults = [];

    for (const job of relevantJobs) {
      if (jobsProcessed >= MAX_JOBS_PER_RUN) break;
      // Time budget check
      if (Date.now() - runStart > ALERT_TIME_BUDGET_MS - 60000) {
        Logger.log(`⏱ Alert scan: time budget nearly exhausted — stopping gracefully`);
        break;
      }

      Logger.log(`→ "${job.title}" at "${job.company}"`);

// ── JD Fetch: company page → web search → skip (no email body fallback) ──
let jdText   = null;
let jdSource = 'unknown';

// Tier 1: Fetch the Indeed job page directly via Tavily
Logger.log(`  Tier 1: fetching ${job.url}`);
const t1 = tavilyExtractAdvanced(job.url);
if (t1 && looksLikeJobContent(t1) &&
    isJdRelevantToJob(t1, job.company, job.title) &&
    isCompleteJobDescription(t1)) {
  jdText  = t1;
  jdSource = 'indeed_direct';
  incrementTavilyCounter(2);
} else if (t1) {
  Logger.log(`  Tier 1 result discarded (relevance or truncation check failed)`);
}

// Tier 2: Targeted web search — quoted company name for precision
if (!jdText) {
  Logger.log(`  Tier 2: searching web for "${job.company}" + "${job.title}"`);
  const t2 = tavilySearchValidated(job.company, job.title);
  if (t2 && isCompleteJobDescription(t2)) {
    jdText  = t2;
    jdSource = 'web_search';
  }
}

// No JD found — skip cleanly, no SMM, no fake score
if (!jdText) {
  Logger.log(`  ⚠ No valid complete JD found — skipping (no fake score)`);
  totalSkipped++;
  threadResults.push({
    title:   job.title,
    company: job.company,
    skipped: true,
    reason:  'no_jd'
  });
  jobsProcessed++;
  continue;
}

      // Step 2: detect CV profile
      const cvType = detectCvTypeFromText(job.title + ' ' + jdText.substring(0, 500));

      // Step 3: run SMM analysis
      let smmResult;
      try {
        const smmRaw = analyzeSkillsMatch(jdText, cvType, runStart, ALERT_TIME_BUDGET_MS);
        smmResult    = JSON.parse(smmRaw);
        if (smmResult.error) throw new Error(smmResult.error);
      } catch(e) {
        Logger.log(`  ✗ SMM error: ${e.message}`);
        totalSkipped++;
        threadResults.push({ title: job.title, company: job.company, skipped: true });
        jobsProcessed++;
        Utilities.sleep(3000);
        continue;
      }

      const score      = smmResult.total_score || 0;
      const matchLevel = smmResult.match_level  || 'M0';
      const skills     = smmResult.skills        || [];
      const levelNum   = parseInt(matchLevel.replace(/\D/g, '')) || 0;

      // Step 4: apply qualification rule
      const crucialSkills  = skills.filter(s => s.importance === 'Crucial');
      const zeroCrucial    = crucialSkills.length === 0;
      const allCrucialPass = !zeroCrucial && crucialSkills.every(s => (s.score || 0) >= 1);
      const qualifies      = levelNum >= 1 && allCrucialPass;

      Logger.log(`  Score: ${score}/40 | ${matchLevel} | Crucial: ${crucialSkills.length} | Pass: ${allCrucialPass} | Qualifies: ${qualifies}`);

      if (zeroCrucial) {
        // Edge case: zero Crucial skills → manual review
        totalLow++;
        threadResults.push({ title: job.title, company: job.company, smmResult: smmResult, score, matchLevel, qualifies: false, reason: 'no-crucial' });
        summaryLines.push(`⚠ MANUAL REVIEW (no Crucial skills): ${job.company} — ${job.title} — ${score}/40`);
      } else if (qualifies) {
        totalReviewed++;
        threadResults.push({ title: job.title, company: job.company, smmResult: smmResult, score, matchLevel, qualifies: true });
        summaryLines.push(`✅ ${job.company} — ${job.title} — ${score}/40 (${matchLevel})`);
      } else {
        totalLow++;
        threadResults.push({ title: job.title, company: job.company, smmResult: smmResult, score, matchLevel, qualifies: false, reason: 'below-threshold' });
        summaryLines.push(`⬇ ${job.company} — ${job.title} — ${score}/40 (${matchLevel}) — below threshold`);
      }

      jobsProcessed++;
      Utilities.sleep(2500); // respect Mistral rate limit
    }

    const analyzed  = threadResults.filter(r => r.smmResult);
const noJd      = threadResults.filter(r => r.skipped && r.reason === 'no_jd');
const otherSkip = threadResults.filter(r => r.skipped && r.reason !== 'no_jd');

if (analyzed.length > 0) {
  // Pick the highest scoring job in this thread
  const best = analyzed.reduce((a, b) =>
    (a.smmResult.total_score || 0) >= (b.smmResult.total_score || 0) ? a : b
  );
  const scoreLabel = buildAlertLabel(best.smmResult);
  thread.addLabel(getOrCreateLabel(scoreLabel));
} else if (noJd.length > 0 && analyzed.length === 0) {
  thread.addLabel(getOrCreateLabel('JABA Alert/⏭ No JD Found'));
} else if (otherSkip.length > 0 && analyzed.length === 0) {
  thread.addLabel(getOrCreateLabel('JABA Alert/⏭ Skipped'));
}

thread.addLabel(getOrCreateLabel('JABA Alert/Processed'));

    Utilities.sleep(500);
  }

  props.setProperty('LAST_ALERT_SCAN', new Date().toISOString());

  const summary = [
    '📧 Indeed Alert Scan Complete',
    '─────────────────────────────',
    `✅ Qualifying jobs: ${totalReviewed}`,
    `⬇  Below threshold: ${totalLow}`,
    `⚠  Skipped (fetch failed): ${totalSkipped}`,
    '',
    ...summaryLines,
    '',
    `Jobs processed this run: ${jobsProcessed}/${MAX_JOBS_PER_RUN}`,
    jobsProcessed >= MAX_JOBS_PER_RUN ? 'Run again to continue with remaining emails.' : ''
  ].join('\n');

  Logger.log(summary);
  if (ui) ui.alert(summary);
}

/* ============================================================
   INDEED ALERT PHASE 1
   Fetches and validates JDs, writes to Pending_SMM.
   SMM analysis handled by universal runPhase2.
   Replaces processIndeedAlertEmails (old function kept as dead code).
   ============================================================ */
function processIndeedAlertEmails_Phase1() {
  const props    = PropertiesService.getScriptProperties();
  const timezone = CONFIG.TIMEZONE;
  const runStart = Date.now();
  const TIME_BUDGET_MS = 280000;
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();

  cleanAlertResults();

  const lastScanStr    = props.getProperty('LAST_ALERT_SCAN');
  const sinceDate      = lastScanStr
    ? new Date(lastScanStr)
    : new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000);
  const queryDate      = new Date(sinceDate.getTime() - 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(queryDate, timezone, 'yyyy/MM/dd');
  Logger.log(`Indeed Phase 1 — since: ${sinceFormatted}`);

  const labelInternship = getOrCreateLabel('JABA Alert/🎓 Internship');
  const labelQueued     = getOrCreateLabel('JABA Alert/⏳ Queued');

  const query = [
    'from:indeed.com',
    `after:${sinceFormatted}`,
    '-label:JABA Alert/Processed',
    '-subject:Bewerbung',
    '-subject:"Neuigkeiten zu Ihrer Bewerbung"',
    '-subject:"Your application"',
    '-subject:indeedapply',
    '-subject:"Heben Sie sich"'
  ].join(' ');

  const threads = GmailApp.search(query, 0, 30);
  Logger.log(`Found ${threads.length} Indeed alert thread(s)`);

  if (threads.length === 0) {
    props.setProperty('LAST_ALERT_SCAN', new Date().toISOString());
    if (ui) ui.alert('No new Indeed alert emails found since last scan.');
    return;
  }

  const pendingSheet   = getOrCreatePendingSmmSheet();
  const dateStr        = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
  const MAX_JD_FETCHES = 12;
  let jdFetches    = 0;
  let totalQueued  = 0;
  let totalSkipped = 0;
  let totalCached  = 0;
  let totalApplied = 0;

  // Load applied jobs once — avoids per-job sheet reads and prevents
  // queuing a job that's already been registered in a monthly tab.
  const appliedSetIndeed = buildAppliedJobsSet();

  for (const thread of threads) {
    if (Date.now() - runStart > TIME_BUDGET_MS - 30000) {
      Logger.log('⏱ Indeed Phase 1 time budget — stopping');
      break;
    }

    const threadId = thread.getId();
    const message  = thread.getMessages()[thread.getMessages().length - 1];
    const htmlBody = message.getBody();
    const subject  = message.getSubject();
    Logger.log(`\nIndeed thread: "${subject}"`);

    const allJobs      = extractJobListingsFromHtml(htmlBody);
    const relevantJobs = allJobs.filter(j => isRelevantJobTitle(j.title));
    Logger.log(`  Jobs: ${allJobs.length} total | ${relevantJobs.length} relevant`);

    if (relevantJobs.length === 0) {
      const hasInternship = allJobs.some(j =>
        /werkstudent|praktikum|pflichtpraktikum|internship/i.test(j.title)
      );
      if (hasInternship) thread.addLabel(labelInternship);
      thread.addLabel(getOrCreateLabel('JABA Alert/Processed'));
      continue;
    }

    let threadQueued         = 0;
    let threadFullyProcessed = true;

    for (const job of relevantJobs) {
      if (isJobInCache(job.company, job.title)) {
        Logger.log(`    [cache] "${job.title}" @ ${job.company} — skip`);
        totalCached++;
        continue;
      }

      if (jdFetches >= MAX_JD_FETCHES) {
        Logger.log(`  MAX_JD_FETCHES (${MAX_JD_FETCHES}) reached — thread not fully processed`);
        threadFullyProcessed = false;
        break;
      }

      if (Date.now() - runStart > TIME_BUDGET_MS - 30000) {
        threadFullyProcessed = false;
        break;
      }

      Logger.log(`  → "${job.title}" at "${job.company}"`);
      jdFetches++;

      let jdText    = null;
      let fetchedUrl = job.url;

      // Tier 1: Indeed job page via Tavily advanced extract
      Logger.log(`    Tier 1: ${job.url}`);
      const t1 = tavilyExtractAdvanced(job.url);
      if (t1 && looksLikeJobContent(t1) && isCompleteJobDescription(t1) &&
          isJdRelevantToJob(t1, job.company, job.title, false)) {
        jdText = t1;
        Logger.log(`    Tier 1 OK (${t1.length} chars)`);
      } else if (t1) {
        Logger.log(`    Tier 1 discarded (relevance or completeness failed)`);
      }

      // Tier 2: Tavily web search — now returns {text, url}
      if (!jdText) {
        Logger.log(`    Tier 2: web search`);
        const t2 = tavilySearchValidated(job.company, job.title);
        if (t2 && isCompleteJobDescription(t2.text) &&
            isJdRelevantToJob(t2.text, job.company, job.title, false)) {
          jdText     = t2.text;
          fetchedUrl = t2.url;
          Logger.log(`    Tier 2 OK (${t2.text.length} chars)`);
        } else if (t2) {
          Logger.log(`    Tier 2 discarded`);
        }
      }

      if (!jdText) {
        Logger.log(`    ✗ No valid JD — writing skipped to Alert_Results`);
        const skipDate = Utilities.formatDate(new Date(), timezone, 'dd.MM.yyyy HH:mm');
        writeToAlertResults(skipDate, 'Indeed', job.company, job.title,
                            job.url, '', 0, 'Skip', 'No valid JD found', '', '', threadId);
        addJobToCache(job, { match_level: 'SKIP', total_score: 0 }, '', 'no_jd', job.url, 'Indeed', '');
        totalSkipped++;
        Utilities.sleep(500);
        continue;
      }

      // ── Already applied check (post-JD-fetch, consistent with job search) ──
      const indeedApplyKey = job.company.toLowerCase().trim() + '||' +
                             normalizeJobTitle(job.title).toLowerCase().trim();
      if (appliedSetIndeed.has(indeedApplyKey)) {
        Logger.log(`    [already_applied] "${job.title}" @ ${job.company} — skipping`);
        const skipDate = Utilities.formatDate(new Date(), timezone, 'dd.MM.yyyy HH:mm');
        writeToAlertResults(skipDate, 'Indeed', job.company, job.title,
                            job.url, fetchedUrl, 0, 'already_applied',
                            'Already applied', '', '', threadId);
        addJobToCache(job, { match_level: 'already_applied', total_score: 0 },
                      '', 'already_applied', fetchedUrl, 'Indeed', '');
        totalApplied++;
        Utilities.sleep(300);
        continue;
      }
      // ─────────────────────────────────────────────────────────────────────

      // Write to Pending_SMM — Phase 2 handles SMM scoring
      pendingSheet.appendRow([
        dateStr, job.company, job.title, job.url,
        '', '',                           // City, Country (not in email)
        jdText.substring(0, 10000),       // Description — pre-fetched and validated
        'TRUE',                           // DescriptionFull
        scoreJobTitleQuality(job.title),  // QualityScore
        job.id || '',                     // ID
        '', '', '', '',                   // Status, CvType, FetchSource, SmmData
        'Indeed',                         // Source
        threadId,                         // Thread_ID
        'TRUE',                           // JD_Fetched
        fetchedUrl                        // Source_URL — actual JD source (may differ from job.url)
      ]);
      addJobToCache(job, { match_level: 'QUEUED', total_score: 0 }, '', 'queued', fetchedUrl, 'Indeed', '');
      totalQueued++;
      threadQueued++;
      Logger.log(`    ✓ Queued for SMM`);

      Utilities.sleep(1000);
    }

    // Mark thread as Processed only if every relevant job was handled
    // Unfinished threads stay unlabeled so the next trigger finds them again
    if (threadFullyProcessed) {
      thread.addLabel(getOrCreateLabel('JABA Alert/Processed'));
      if (threadQueued > 0) thread.addLabel(labelQueued);
    } else {
      Logger.log(`  Thread not fully processed — will retry on next trigger run`);
    }

    Utilities.sleep(300);
  }

  props.setProperty('LAST_ALERT_SCAN', new Date().toISOString());
  SpreadsheetApp.flush();
  Logger.log(`Indeed Phase 1: ${totalQueued} queued | ${totalSkipped} skipped | ${totalCached} cached | ${totalApplied} already_applied`);

  if (totalQueued > 0) {
    deletePhase2Triggers();
    ScriptApp.newTrigger('runPhase2').timeBased().after(PHASE2_DELAY_MS).create();
    Logger.log(`Phase 2 trigger created — runs in ${PHASE2_DELAY_MS / 60000} min`);
  }

  const summary = [
    '📧 Indeed Alert Phase 1 Complete',
    '─────────────────────────────',
    `✅ Queued for SMM: ${totalQueued}`,
    `⚠  Skipped (no valid JD): ${totalSkipped}`,
    `⏭  Already cached: ${totalCached}`,
    `✅ Already applied (skipped): ${totalApplied}`,
    totalQueued > 0 ? `\nSMM analysis runs automatically in ${PHASE2_DELAY_MS / 60000} minutes.` : ''
  ].filter(Boolean).join('\n');

  if (ui) ui.alert(summary);
}

// Debug helper — resets the alert scan window to 7 days ago
function resetAlertScanTimestamp() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_ALERT_SCAN');
  Logger.log('Alert scan timestamp cleared. Next run will scan last 7 days.');
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
      i + 1,
      s.name            || "",
      s.score           || 0,
      s.importance      || "",
      s.evidence        || "",
      s.gap_tip         || "",
      false,                     // Interview_Reached
      s.master_category || "",    // Master_Category
      PROMPT_VERSIONS.SMM          // ← new column 14
    ]);

    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, 14).setValues(rows);
      Logger.log(`SMM_Raw_Data: wrote ${rows.length} rows for UID ${uid} (${company})`);
    }
  } catch (e) {
    Logger.log(`Error in writeSmmRawData: ${e.toString()}`);
    // Non-fatal: do not throw — main registration should still succeed
  }
}
/* ============================================================
   Phase 2: flag Interview_Reached = TRUE in SMM_Raw_Data
   Triggered by onEdit when status changes to HR/1st Interview.
   Matches by Company name — no UID column needed in monthly sheet.
   ============================================================ */
function flagSmmInterviewReached(company, appDate) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SMM_Raw_Data");
    if (!sheet || sheet.getLastRow() < 2) return;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    let flagged = 0;

    data.forEach((row, i) => {
      const smmCompany = row[2] ? row[2].toString().trim() : "";
      if (smmCompany.toLowerCase() === company.toLowerCase()) {
        sheet.getRange(i + 2, 12).setValue(true); // column 12 = Interview_Reached
        flagged++;
      }
    });

    if (flagged > 0) {
      SpreadsheetApp.flush();
      Logger.log(`Interview_Reached flagged for "${company}": ${flagged} rows updated.`);
    }
  } catch (e) {
    Logger.log(`Error in flagSmmInterviewReached: ${e.toString()}`);
  }
}
/* ============================================================
   Helper: get existing Master_Category values from SMM_Raw_Data
   Used to keep new categorisations consistent with past ones.
   ============================================================ */
function getExistingMasterCategories() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SMM_Raw_Data");
    if (!sheet || sheet.getLastRow() < 2) return [];
    // Master_Category is column 13
    const values = sheet.getRange(2, 13, sheet.getLastRow() - 1, 1).getValues();
    const unique  = [...new Set(values.flat().filter(v => v && v.toString().trim() !== ''))];
    return unique.slice(0, 40); // cap context size
  } catch (e) {
    return [];
  }
}


/* ============================================================
   Groq API caller — reusable helper
   model: e.g. "llama-3.1-8b-instant" or "llama-3.3-70b-versatile"
   ============================================================ */
function callGroqApi(systemPrompt, userPrompt, model, maxTokens) {
  const groqKey = getGroqKey();
  if (!groqKey) throw new Error("GROQ_API_KEY not set in Script Properties.");

  const url     = "https://api.groq.com/openai/v1/chat/completions";
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": `Bearer ${groqKey}` },
    payload: JSON.stringify({
      model: model || "llama-3.3-70b-versatile",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user",   content: userPrompt   }
      ],
      temperature: 0.0,
      max_tokens:  maxTokens || 200
    }),
    muteHttpExceptions: true
  };

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const res  = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      if (code === 429) {
        Logger.log(`Groq 429 (attempt ${attempt}) — waiting ${attempt * 4}s`);
        Utilities.sleep(attempt * 4000);
        continue;
      }
      if (code !== 200) {
        Logger.log(`Groq API error ${code}: ${res.getContentText().substring(0, 200)}`);
        return null;
      }
      const json = JSON.parse(res.getContentText());
      return json.choices[0].message.content.trim().replace(/```json|```/g, '').trim();
    } catch (e) {
      Logger.log(`Groq call failed (attempt ${attempt}): ${e.message}`);
      if (attempt < 3) Utilities.sleep(3000);
    }
  }
  return null;
}


/* ============================================================
   Batch Refresh Master Categories
   Re-clusters all SMM_Raw_Data rows using Groq llama-3.3-70b.
   Only runs if >= 20 applications are registered.
   ============================================================ */
function batchRefreshMasterCategories() {
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();

  // Guard: require at least 20 SMM applications (160 rows = 20 × 8 skills)
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SMM_Raw_Data");
  if (!sheet || sheet.getLastRow() < 161) {
    ui.alert("Not enough data yet. Run the refresh once you have at least 20 applications registered via SMM.");
    return;
  }

  const lastRow = sheet.getLastRow();
  const data    = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  // Collect all unique skill + importance + current category combos
  const skillMap = {}; // key: skill_name|importance → {rows: [...rowIndices], currentCategory}
  data.forEach((row, i) => {
    const skillName  = row[6]  ? row[6].toString().trim()  : "";
    const importance = row[8]  ? row[8].toString().trim()  : "";
    const category   = row[12] ? row[12].toString().trim() : "";
    if (!skillName) return;
    const key = `${skillName}|${importance}`;
    if (!skillMap[key]) skillMap[key] = { skillName, importance, currentCategory: category, rowIndices: [] };
    skillMap[key].rowIndices.push(i + 2); // +2 for 1-indexed + header
  });

  const uniqueSkills      = Object.values(skillMap);
  const existingCategories = [...new Set(uniqueSkills.map(s => s.currentCategory).filter(Boolean))];

  Logger.log(`Batch refresh: ${uniqueSkills.length} unique skills, ${existingCategories.length} existing categories`);

  // Build prompt — send all unique skills in one call for efficiency
  const skillList = uniqueSkills.map((s, i) =>
    `${i + 1}. Skill: "${s.skillName}" | Importance: "${s.importance}" | Current category: "${s.currentCategory}"`
  ).join('\n');

  const systemPrompt = `You are a precise skills categorisation engine. Always respond with valid JSON only. No markdown, no explanation.`;

  const userPrompt = `You are consolidating skill categories for a job application analytics dashboard.

EXISTING CATEGORIES (use these first, create new ones only when truly necessary):
${existingCategories.join(', ')}

CONSTRAINT: Maximum 20 unique categories per JD_Importance level (Crucial / Necessary / Optional).
Merge semantically similar categories (e.g. "CRM Tools" and "CRM Platforms" → "CRM & Automation").

SKILLS TO CATEGORISE:
${skillList}

Respond ONLY with a JSON array. One object per skill, same order as input:
[{"index": 1, "master_category": "Category Name"}, ...]`;

  const raw = callGroqApi(systemPrompt, userPrompt, "llama-3.3-70b-versatile", 3000);
  if (!raw) {
    ui.alert("Groq API call failed. Check logs and try again.");
    return;
  }

  let assignments;
  try {
    assignments = JSON.parse(raw);
  } catch (e) {
    Logger.log(`Batch refresh parse error: ${e.message}\nRaw: ${raw}`);
    ui.alert("Could not parse Groq response. Check logs.");
    return;
  }

  // Write updated categories back — only column 13
  let updated = 0;
  assignments.forEach(a => {
    const idx  = a.index - 1; // back to 0-based
    const skill = uniqueSkills[idx];
    if (!skill || !a.master_category) return;
    skill.rowIndices.forEach(rowNum => {
      sheet.getRange(rowNum, 13).setValue(a.master_category);
      updated++;
    });
  });

  SpreadsheetApp.flush();
  Logger.log(`Batch refresh complete: ${updated} rows updated.`);
  ui.alert(`✅ Done. ${updated} rows updated across ${assignments.length} unique skills.`);
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

    const cleanedJD = jdInput.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').substring(0, 6000);

    const signOff = isDe ? "Mit freundlichen Grüßen" : "Best regards";
    const availabilityText = isDe
  ? "Ich bin bereit umzuziehen (falls erforderlich) und stehe kurzfristig mit einer Kündigungsfrist von zwei Wochen zur Verfügung."
  : "I am fully open to relocation if required and am available to start within a two-week notice period.";

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

    letterText = letterText.replace(/\s*—\s*/g, ', ');  // em-dash → comma
    letterText = letterText.replace(/\s*–\s*/g, ', ');  // en-dash → comma
    letterText = letterText.replace(/ - /g, ', ');       // connector hyphen → comma (compound words untouched)
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

    // ── Location resolution ───────────────────────────────────────────────
    // Priority 1: location from JD (via SMM job_location field)
    // Priority 2: location from company name lookup (Germany only for sheet)
    // Priority 3: Mistral's CITY field from the pipe-delimited output
    let resolvedCity = city ? city.trim() : '';

    if (smmData) {
      // Enrich SMM data with company location if JD had none
      enrichSmmLocation(smmData, companyName);

      const jl = smmData.job_location;
      if (jl && (jl.city || jl.country)) {
        if (smmData.location_source === 'jd') {
          // JD location is authoritative — always use it
          resolvedCity = jl.city || resolvedCity;
          Logger.log(`Location from JD: "${resolvedCity}"`);
        } else if (smmData.location_source === 'company_lookup') {
          // Company lookup — only use for sheet if German location
          if (jl.city && isGermanLocation(jl.city)) {
            resolvedCity = jl.city;
            Logger.log(`Location from company lookup (German): "${resolvedCity}"`);
          } else if (jl.city) {
            Logger.log(`Location from company lookup (non-German, map only): "${jl.city}" — not written to sheet`);
            // resolvedCity stays as Mistral's city guess (may be empty)
          }
        }
      }
    }

    const location = normalizeLocation(resolvedCity);

    let cleanedSalary = "";
    if (salary) {
      salary = salary.trim();
      const salaryMatch = salary.match(/([\$£€¥]?\s?(\d{1,3}(?:[.,]\d{3})*|\d+)(?:[.,]\d{2})?)/);
      if (salaryMatch && salaryMatch[1]) {
        cleanedSalary = salaryMatch[1];
      } else {
  // Only keep if it contains at least one digit — discard pure text hallucinations
  cleanedSalary = /\d/.test(salary) ? salary : '';
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

    // ── German language requirement (column H) ────────────────────────────
    // Maps SMM language_note to a short label for analysis of C2 rejections.
    // EN requirements are intentionally ignored — Rey holds C2 English.
    // Blank = no German requirement stated, or only English required.
    // ── Derive column H language requirement from SMM language_note ──
    let langReq = '';
    if (smmData && smmData.language_note) {
      const ln = smmData.language_note;
      if (ln.status === 'unmet') {
        langReq = 'C2 required';
      } else if (ln.status === 'met' && ln.text && ln.text.toLowerCase().includes('c1')) {
        langReq = 'C1 required';
      } else if (ln.status === 'other') {
        langReq = 'Other';
      }
      // status === 'met' for English, or status === 'none' → stays blank intentionally
    }

    const rowData   = [[finalMatch, companyName, position, detectedPlatform, location, "Applied", dateStr, langReq, finalNotes]];
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
    updateInterviewGeoData();
    SpreadsheetApp.flush();

    try { markCacheRowStatus(companyName, position, 'registered'); } catch(e) {
      Logger.log(`markCacheRowStatus non-fatal: ${e.message}`);
    }
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const month = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "MMM yyyy").replace(/\./g, '');

  // If tab already exists, return it directly
  const existing = ss.getSheetByName(month);
  if (existing) return existing;

  // ── Find the most recent monthly tab by parsing tab names as dates ──────────
  const MONTH_NAMES = {
    jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
    jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
  };

  let templateSheet = null;
  let latestDate    = null;

  ss.getSheets().forEach(function(s) {
    const parts = s.getName().trim().split(/\s+/);
    if (parts.length !== 2) return;
    const mon = MONTH_NAMES[parts[0].toLowerCase().slice(0, 3)];
    const yr  = parseInt(parts[1], 10);
    if (mon === undefined || isNaN(yr)) return;
    const d = new Date(yr, mon, 1);
    if (!latestDate || d > latestDate) {
      latestDate    = d;
      templateSheet = s;
    }
  });

  // ── Create a fresh new sheet ─────────────────────────────────────────────────
  const sheet = ss.insertSheet(month);
  SpreadsheetApp.flush();

  if (templateSheet) {
    // ── Read header row values and formatting from the template ──────────────
    // We only copy row 1 (headers) and the dropdown validations from row 2.
    // We do NOT use copyTo() — it fails in web app context.

    const templateHeaderRange = templateSheet.getRange(1, 1, 1, 20);

    // Copy header values (A1:T1)
    const headerValues = templateHeaderRange.getValues();
    sheet.getRange(1, 1, 1, 20).setValues(headerValues);

    // Copy header background colors
    const bgColors = templateHeaderRange.getBackgrounds();
    sheet.getRange(1, 1, 1, 20).setBackgrounds(bgColors);

    // Copy header font colors
    const fontColors = templateHeaderRange.getFontColors();
    sheet.getRange(1, 1, 1, 20).setFontColors(fontColors);

    // Copy header font weights
    const fontWeights = templateHeaderRange.getFontWeights();
    sheet.getRange(1, 1, 1, 20).setFontWeights(fontWeights);

    // Copy header font sizes
    const fontSizes = templateHeaderRange.getFontSizes();
    sheet.getRange(1, 1, 1, 20).setFontSizes(fontSizes);

    // Copy header font families
    const fontFamilies = templateHeaderRange.getFontFamilies();
    sheet.getRange(1, 1, 1, 20).setFontFamilies(fontFamilies);

    // Copy column widths for A–T (columns 1–20)
    for (let col = 1; col <= 20; col++) {
      const width = templateSheet.getColumnWidth(col);
      sheet.setColumnWidth(col, width);
    }

    // Copy frozen rows setting
    const frozenRows = templateSheet.getFrozenRows();
    if (frozenRows > 0) sheet.setFrozenRows(frozenRows);

    // ── Re-apply dropdown validations from template row 2 ───────────────────
    const colsWithDropdowns = [1,2,3,4,5,6,7,9,10,11,12,13,14,15,16,17,18,20];
    colsWithDropdowns.forEach(function(col) {
      const validation = templateSheet.getRange(2, col).getDataValidation();
      if (!validation) return;
      sheet.getRange(2, col, 1000, 1).setDataValidation(validation);
    });

    // ── Column H: Language Requirement dropdown ───────────────────────────
    // Applied fresh here because the template predates this column.
    // Uses requireValueInList (same style as cols A, D, F) — no inline arrow.
    const langReqValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['C2 required', 'C1 required', 'Other'], true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, 8, 1000, 1).setDataValidation(langReqValidation);
    // ─────────────────────────────────────────────────────────────────────

  } else {
    // ── Fallback: no monthly tab found — build from scratch ─────────────────
    const headers = [
      "Match Level", "Companies", "Position", "Application Platform", "Location",
      "Status", "Application Date", "Language Req.", "Notes",
      "Applied", "HR Interview", "1st Interview", "2nd Interview", "3rd Interview",
      "4th Interview", "Offer", "Ignored", "Rejected",
      "Status Path", "Email Rejection"
    ];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setBackground("#4285f4")
         .setFontColor("white")
         .setFontWeight("bold");
    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);

    // ── Column H: Language Requirement dropdown ───────────────────────────
    const langReqValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['C2 required', 'C1 required', 'Other'], true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, 8, 1000, 1).setDataValidation(langReqValidation);
    // ─────────────────────────────────────────────────────────────────────
  }

  // ── Apply =JOIN formula to column S (rows 2 onward) ─────────────────────────
  const emojiFormula = '=JOIN(""; IF(J2>0;"📩";""); IF(K2>0;"0️⃣";""); IF(L2>0;"1️⃣";""); IF(M2>0;"2️⃣";""); IF(N2>0;"3️⃣";""); IF(O2>0;"4️⃣";""); IF(P2>0;"🎉";""); IF(Q2>0;"⚪";""); IF(R2>0;"🛑";""))';
  sheet.getRange("S2:S").setFormula(emojiFormula);

  SpreadsheetApp.flush();
  return sheet;
}


/**
 * Copies data validations from templateSheet row 2 (columns A–T, excluding S)
 * and applies them to all data rows in the new sheet.
 * This is needed because clearDataValidations() removes the dropdown rules
 * that were brought over by copyTo().
 */
function _copyValidationsFromTemplate(templateSheet, newSheet, lastRow) {
  // Columns to restore validations for: A-R (1-18) and T (20). Skip S (19).
  const colsToRestore = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,20];

  colsToRestore.forEach(function(col) {
    const templateCell = templateSheet.getRange(2, col);
    const validation   = templateCell.getDataValidation();
    if (!validation) return; // no dropdown on this column — skip

    // Apply to the full data range in the new sheet
    newSheet.getRange(2, col, lastRow - 1, 1).setDataValidation(validation);
  });
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
    'christ.de', 'zalando', 'otto.de', 'aboutyou', 'hm.com', 'notion.so',
    'match.indeed.com', 'jobalert.indeed.com', 'hackernoon.com',
    'lieferando', 'deliveroo', 'uber', 'mjam'
  ];
  const nonJobSubjects = [
    'newsletter', 'angebot', '% rabatt', 'sale', 'discount', 'Kontoprüfcode',
    'bestellung', 'order confirmation', 'rechnung', 'invoice',
    'digest', 'quora', 'sparangebot', 'nur heute', 'flash sale',
    'deine lieferung', 'your delivery', 'tracking', 'versandbestatigung',
    'jaba job report', 'jaba daily report'
  ];

  if (nonJobDomains.some(d => senderLower.includes(d)))  return false;
  if (nonJobSubjects.some(s => subjectLower.includes(s))) return false;
  return true;
}


/**
 * Tier 1: deterministic rule-based rejection detection.
 * Returns 'rejection', 'not_rejection', or 'uncertain'.
 */
function classifyRejectionByRules(subject, body) {
  const combined = (subject + ' ' + body).toLowerCase().substring(0, 3000);

  // Hard rejection signals — any one of these = confirmed rejection
  const rejectionPhrases = [
    // ── German formal (Sie/Ihnen) — original phrases ──────────────────────
    'leider müssen wir ihnen mitteilen',
    'leider können wir ihre bewerbung',
    'leider können wir ihnen',
    'leider müssen wir ihnen',
    'nicht weiterverfolgen',
    'anderweitig besetzt',
    'haben uns für andere kandidaten',
    'haben uns für andere bewerber',
    'entschieden wir uns für',
    'kein passendes profil',
    'entspricht nicht dem gesuchten profil',
    'nicht dem anforderungsprofil',
    'bedauern wir ihnen mitteilen',
    'bedauern, ihnen mitteilen',
    'absage für ihre bewerbung',
    'ihre bewerbung war leider nicht erfolgreich',
    'bewerbung nicht berücksichtigen',
    'im auswahlverfahren nicht',
    'an dieser stelle beenden',
    'bewerbungsprozess beenden',
    'leider keine möglichkeit',
    'nach reiflicher überlegung',
    'nach sorgfältiger prüfung',
    'nach eingehender prüfung',
    'anderen kandidaten den vorzug',
    'anderen bewerbern den vorzug',

    // ── German informal (du/dir/deine) — added variants ───────────────────
    'leider müssen wir dir mitteilen',
    'leider können wir deine bewerbung',
    'leider können wir dir',
    'bedauern wir dir mitteilen',
    'bedauern, dir mitteilen',
    'absage für deine bewerbung',
    'deine bewerbung war leider nicht erfolgreich',

    // ── German — new patterns missing entirely ────────────────────────────
    'inzwischen besetzt',                        // "position filled in the meantime"
    'nicht berücksichtigt werden konnte',        // very common soft rejection
    'nicht berücksichtigt werden konnten',
    'nicht in die engere auswahl',               // not shortlisted
    'leider nicht in die engere auswahl',
    'uns leider nicht für dich entschieden',
    'uns leider nicht für sie entschieden',
    'müssen wir dir leider absagen',
    'müssen wir ihnen leider absagen',
    'leider absagen müssen',
    'haben wir uns gegen deine bewerbung',
    'haben wir uns gegen ihre bewerbung',

    // ── English — original phrases ────────────────────────────────────────
    'not moving forward with your application',
    'not moving forward with your candidacy',
    'will not be moving forward',
    'unable to move your application forward',
    'not selected for',
    'decided to move forward with other',
    'went with another candidate',
    'position has been filled',
    'no longer considering your application',
    'we won\'t be proceeding',
    'unfortunately we will not',
    'unfortunately, we will not',
    'unfortunately we\'re unable to',
    'unfortunately, we\'re unable to',
    'regret to inform you',

    // ── English — new patterns missing entirely ───────────────────────────
    'we have decided not to move forward',
    'we have decided to pursue other',
    'will not be proceeding with your',
    'your application was unsuccessful',
    'we are unable to offer you',
    'the role has been filled',
    'we have filled the position',
    'chosen to move forward with another candidate',
    'not shortlisted',
    'other applicants meet our requirements',
    'We are truly sorry not to be able to give you positive feedback',
    'decided not to proceed with your',
  ];

  // Hard non-rejection signals — any one of these = definitely not rejection
  const nonRejectionPhrases = [
    'einladung zum vorstellungsgespräch',
    'einladung zum gespräch',
    'wir möchten sie zu einem gespräch einladen',
    'wir möchten dich zu einem gespräch einladen',
    'bewerbungsgespräch vereinbaren',
    'we would like to invite you',
    "we'd like to schedule",
    'interview invitation',
    'bitte reichen sie',         // document request
    'bitte schicken sie uns',
    'unterlagen nachzureichen',
    'fehlen uns noch',
    'fehlende unterlagen',
    'arbeitszeugnisse',
    'wir haben ihre bewerbung erhalten',
    'wir haben deine bewerbung erhalten',
    'your application has been received',
    'bewerbungseingang',
    'eingangsbestätigung',
    'bestätige deine identität',
    'verifizierungscode',
    'code:',                     // identity verification codes
    'wir prüfen ihre unterlagen',
    'wir prüfen deine unterlagen',
    'werden uns bei ihnen melden',
    'werden uns bei dir melden',
    'we will be in touch',
    'we\'ll be in touch',
    'assessment',
    'online-test',
    'online test einladung',
    'talentpool',
    'talent pool',
    'talent-pool'
  ];

  if (nonRejectionPhrases.some(p => combined.includes(p))) return 'not_rejection';
  if (rejectionPhrases.some(p => combined.includes(p))) return 'rejection';
  return 'uncertain';
}

/**
 * Main rejection classifier: rules first, AI only for uncertain cases.
 * Uses llama-3.3-70b (more accurate than 3.1-8b, still free on Groq).
 */
function classifyRejectionEmail(sender, subject, body) {
  const truncatedBody = body.substring(0, 1500);
  const ruleResult    = classifyRejectionByRules(subject, truncatedBody);

  if (ruleResult === 'not_rejection') {
    Logger.log(`  Rules: NOT rejection`);
    return { isRejection: false, companyName: null };
  }

  if (ruleResult === 'rejection') {
    // Rules confirmed rejection — still need company name, use AI for extraction only
    Logger.log(`  Rules: REJECTION confirmed — extracting company name`);
    const prompt = `This email is a job application rejection. Extract only the company name that sent it.
From: ${sender}
Subject: ${subject}
Body: ${truncatedBody.substring(0, 500)}
Respond ONLY with JSON: {"companyName": "Company Name"} or {"companyName": null} if unclear.`;

    const raw = callGroqApi(
      "Extract the company name from this rejection email. JSON only.",
      prompt,
      "llama-3.1-8b-instant",
      30
    );
    try {
      const parsed = JSON.parse(raw || '{}');
      return { isRejection: true, companyName: parsed.companyName || null };
    } catch(e) {
      return { isRejection: true, companyName: null };
    }
  }

  // Uncertain: use AI with better model for full classification
  Logger.log(`  Rules: uncertain — calling llama-3.3-70b`);
  const prompt = `Classify this job application email.
From: ${sender}
Subject: ${subject}
Body: ${truncatedBody}

Rules:
- Rejection = company explicitly declines to proceed in this email body
- Document requests ("please send us missing documents") = NOT rejection
- Acknowledgments ("we received your application") = NOT rejection
- Interview invitations = NOT rejection
- Verification codes = NOT rejection

Respond ONLY with JSON:
{"isRejection": true, "companyName": "Company Name"}
or
{"isRejection": false, "companyName": null}`;

  const raw = callGroqApi(
    "Classify job emails. JSON only.",
    prompt,
    "llama-3.3-70b-versatile",
    60
  );
  try {
    const parsed = JSON.parse(raw || '{}');
    if (parsed.companyName === 'null') parsed.companyName = null;
    return parsed;
  } catch(e) {
    Logger.log(`  Classification parse failed: ${e.message}`);
    return null;
  }
}


/*
Fuzzy company name matcher
*/
function findBestCompanyMatch(geminiName, pendingCompanies) {
  if (!geminiName) return null;

  function normalize(name) {
    return name
      .toLowerCase()
      .replace(/['`'']/g, '')                    // ← (removes apostrophes)
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

/**
 * Pre-filter: returns true if the email is clearly an acknowledgment,
 * confirmation, or identity verification — never a rejection.
 * Called BEFORE the AI to avoid false positives on small models.
 */
function isAcknowledgmentEmail(subject, body) {
  const combined = (subject + ' ' + body).toLowerCase();

  // Hard signals — these patterns are NEVER rejections, skip immediately
  const hardSignals = [
    'bestätige deine identität',
    'bestätige dein profil',
    'bestätige deine e-mail',
    'verifizierungscode',
    'verification code',
    'bestätigungscode',
    'confirm your identity',
    'eingangsbestätigung',
    'hiermit übersende ich',
    'hiermit bewerbe ich mich',
    'ich bewerbe mich hiermit',
    'anbei sende ich',
    'anbei übersende ich',
    'please find attached my',
    'i am writing to apply',
    'i would like to apply',
    'bewerbungseingang',
    'fehlen uns noch folgende',
'fehlende unterlagen',
'unterlagen nachzureichen',
'unterlagen nachreichen',
'nachzureichen als antwort',
'bitte reichen sie',
'bitte senden sie uns',
'bitte schicken sie',
'folgende dokument',
'arbeitszeugnisse',
'missing document',
'please provide the following',
'please send us the following',
'additional documents required',
'wir benötigen noch',
'wir bitten dich um',
'bitten wir dich, uns',
    'bewerbung eingegangen',
    'wir haben deine bewerbung erhalten',
    'wir haben ihre bewerbung erhalten',
    'your application has been received',
    'we have received your application'
  ];
  if (hardSignals.some(s => combined.includes(s))) return true;

  // Soft signals: "we're reviewing" language combined with NO rejection markers
  const acknowledgmentPhrases = [
    'wir freuen uns über dein interesse',
    'wir freuen uns über ihr interesse',
    'freuen uns, dass du teil',
    'freuen uns, dass sie teil',
    'danke für deine bewerbung',
    'danke für ihre bewerbung',
    'lieben dank für deine bewerbung',
    'deine unterlagen werden geprüft',
    'unterlagen werden im ersten schritt',
    'sorgfältig prüfen',
    'für eine rückmeldung noch etwas zeit',
    'wir setzen uns so bald wie möglich',
    'wir melden uns bei dir',
    'wir melden uns bei ihnen',
    'thank you for your application',
    'thank you for applying',
    'we will be in touch'
  ];

  const rejectionMarkers = [
    'leider', 'bedauern', 'nicht weiterverfolgen', 'anderweitig besetzt',
    'absage', 'entschieden uns für andere', 'andere bewerber', 'nicht berücksichtigen',
    'unable to', 'not moving forward', 'will not be moving', 'unfortunately',
    'no longer considering', 'decided not to'
  ];

  const hasAcknowledgment = acknowledgmentPhrases.some(s => combined.includes(s));
  const hasRejection      = rejectionMarkers.some(s => combined.includes(s));

  // Acknowledgment phrase present + zero rejection language = safe to skip
  return hasAcknowledgment && !hasRejection;
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
    const SKIP_REJECTION_SHEETS = new Set([
           "Sankey_Data","Geo_Data","SMM_Raw_Data","Interview_Geo_Data",
           "Job_Search_Cache","Pending_SMM","Alert_Results","M2_Notifications",
           "Weekly_Data"
         ]);
      if (!SKIP_REJECTION_SHEETS.has(sheetName) && /\d{4}/.test(sheetName)) {
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
  const queryDate      = new Date(sinceDate.getTime() - 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(queryDate, timezone, "yyyy/MM/dd");

  const labelName = 'bot-rejections-processed';
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);

  const queryFresh = `after:${sinceFormatted} -label:${labelName} -subject:"JABA Job Report" -in:sent`;
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

  const MAX_PER_RUN     = 60;
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
    // Fix 2 — Skip interview and meeting invitation emails before calling AI
    const MEETING_SIGNALS = [
      'bewerbungsgespräch', 'vorstellungsgespräch',
      'einladung zum gespräch', 'gesprächseinladung',
      'interview einladung', 'einladung zum interview',
      'telefoninterview', 'phone interview', 'video interview',
      'we would like to invite', "we'd like to invite",
      'calendar invite', 'meeting invitation', 'besprechungseinladung'
    ];
    if (MEETING_SIGNALS.some(sig => subject.toLowerCase().includes(sig))) {
      Logger.log(`⏭ Interview/meeting invitation — skipped: "${subject}"`);
      Utilities.sleep(300);
      continue;
    }

    // Pre-filter: skip acknowledgment and identity verification emails
    if (isAcknowledgmentEmail(subject, body.substring(0, 1500))) {
      Logger.log(`⏭ Acknowledgment/confirmation email — skipped: "${subject}"`);
      Utilities.sleep(300);
      continue;
    }

    const result = classifyRejectionEmail(sender, subject, body);

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
    updateInterviewGeoData();
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
    const sheetName = sheet.getName().replace(/\./g, '');
    if (/[A-Za-z]+ \d{4}/.test(sheetName)) {
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
    const sheetName = sheet.getName().replace(/\./g, '');
    if (/[A-Za-z]+ \d{4}/.test(sheetName)) {
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



function updateInterviewGeoData() {
  const ss     = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const sheets = ss.getSheets();

  const INTERVIEW_STAGES = [
    { dataIdx: 10, name: 'HR Interview',  stageNum: 1 },
    { dataIdx: 11, name: '1st Interview', stageNum: 2 },
    { dataIdx: 12, name: '2nd Interview', stageNum: 3 },
    { dataIdx: 13, name: '3rd Interview', stageNum: 4 },
    { dataIdx: 14, name: '4th Interview', stageNum: 5 },
    { dataIdx: 15, name: 'Offer',         stageNum: 6 },
  ];

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

  // key: "month|city|country" → { maxStageNum, maxStageName, count }
  const cityMap = {};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName().replace(/\./g, '');
    if (!/[A-Za-z]+ \d{4}/.test(sheetName)) return;

    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() < 2) return;
    const data = dataRange.getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1]) continue;

      let maxStageNum  = 0;
      let maxStageName = '';
      for (const stage of INTERVIEW_STAGES) {
        if (row[stage.dataIdx] === 1) {
          maxStageNum  = stage.stageNum;
          maxStageName = stage.name;
        }
      }
      if (maxStageNum === 0) continue; // no interview stage reached

      const rawCity = row[4];
      if (!rawCity || rawCity.trim() === '') continue;
      const cleanCity = rawCity.split('(')[0].trim();

      const matchedCountry = Object.keys(cityToCountry).find(c =>
        cleanCity.toLowerCase().includes(c)
      );
      const country = matchedCountry ? cityToCountry[matchedCountry] : 'Germany';

      const key = `${sheetName}|${cleanCity}|${country}`;
      if (!cityMap[key]) {
        cityMap[key] = { maxStageNum: 0, maxStageName: '', count: 0 };
      }
      cityMap[key].count++;
      if (maxStageNum > cityMap[key].maxStageNum) {
        cityMap[key].maxStageNum  = maxStageNum;
        cityMap[key].maxStageName = maxStageName;
      }
    }
  });

  let geoSheet = ss.getSheetByName('Interview_Geo_Data');
  if (!geoSheet) {
    geoSheet = ss.insertSheet('Interview_Geo_Data');
  } else {
    geoSheet.clearContents();
    geoSheet.clearFormats();
  }

  const output  = Object.keys(cityMap).map(key => {
    const [month, city, country] = key.split('|');
    const d = cityMap[key];
    return [month, city, country, d.maxStageName, d.maxStageNum, d.count];
  });

  const allRows = [['Month', 'City', 'Country', 'Max_Stage', 'Stage_Num', 'Count'], ...output];
  const range   = geoSheet.getRange(1, 1, allRows.length, 6);
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
function backfillInterviewReached() {
  // Add the company names of your two interview applications here
  const interviewCompanies = [
    "Trailblazer Summits",  // replace with exact name from your sheet
    "Hays AG"   // replace with exact name from your sheet
  ];
  interviewCompanies.forEach(company => flagSmmInterviewReached(company, null));
  Logger.log("Backfill complete.");
}

/* ============================================================
   AUTOMATED DAILY JOB SEARCH — Complete Section (Adzuna)
   Paste this entire block at the end of Code.js, replacing
   everything from this comment to the end of the file.
   ============================================================ */


// ── Search config ─────────────────────────────────────────────────────────────

const JOB_SEARCH_KEYWORDS = [
  'marketing manager',
  'marketing automation',
  'online marketing',
  'digital marketing',
  'crm manager',
  'email marketing',
  'community manager',
  'campaign manager',
  'web3 marketing',
  'AI marketing',
  'automation manager,'
];

const JOB_SEARCH_EXCLUDE_TERMS = [
  // English
  'senior', 'sr.', 'lead ', 'head of', 'director',
  'vp ', 'vice president', 'sales', 'vertrieb', 'ausbildung',
  'verkauf', 'key account', 'leiter', 'leitung', 'cmo', 'praktikum',
  'praktikant', 'werkstudent', 'werkstudentin',
  'internship', 'intern ',
  'product marketing', 'field marketing',
  // ← new: non-marketing roles confirmed M0
  'customer success', 'account executive', 'revops'
];

const DIRECT_FETCH_BLOCKED = [
  'linkedin.com', 'xing.com', 'glassdoor.com',
  'monster.com', 'stepstone.de'
];


// ── Tavily credit counter ─────────────────────────────────────────────────────

function incrementTavilyCounter(credits) {
  if (!credits || credits <= 0) return;
  const props = PropertiesService.getScriptProperties();
  const now   = new Date();

  const resetStr  = props.getProperty('TAVILY_COUNTER_RESET');
  const resetDate = resetStr ? new Date(resetStr) : null;
  if (!resetDate ||
      now.getMonth()    !== resetDate.getMonth() ||
      now.getFullYear() !== resetDate.getFullYear()) {
    props.setProperty('TAVILY_CREDITS_MONTH', '0');
    props.setProperty('TAVILY_COUNTER_RESET', now.toISOString());
  }

  const current  = parseInt(props.getProperty('TAVILY_CREDITS_MONTH') || '0');
  const newTotal = current + credits;
  props.setProperty('TAVILY_CREDITS_MONTH', String(newTotal));
  Logger.log(`  📊 Tavily credits this month: ${newTotal}/1000`);
  return newTotal;
}

function getTavilyMonthlyUsage() {
  return parseInt(
    PropertiesService.getScriptProperties().getProperty('TAVILY_CREDITS_MONTH') || '0'
  );
}


// ── Content quality check ─────────────────────────────────────────────────────

function looksLikeJobContent(text) {
  if (!text || text.length < 150) return false;
  const lower = text.toLowerCase();

  // Task/role signals — what the candidate will DO in the role
  const taskSignals = [
    'aufgaben', 'tätigkeiten', 'deine aufgaben', 'ihre aufgaben',
    'verantwortlichkeiten', 'was du tust', 'was du machst',
    'was du bei uns machst', 'deine rolle', 'in dieser rolle',
    'responsibilities', 'what you will do', 'you will be', 'your role',
    'wir suchen', 'we are looking for', 'your responsibilities',
    'du verantwortest', 'du übernimmst', 'zu deinen aufgaben'
  ];

  // Requirement/qualification signals — what the candidate must HAVE
  const reqSignals = [
    'anforderungen', 'qualifikation', 'voraussetzungen',
    'was du mitbringst', 'was wir erwarten', 'dein profil',
    'berufserfahrung', 'abgeschlossenes studium', 'must-have', 'must have',
    'requirements', 'qualifications', 'you bring', 'required experience',
    'years of experience', 'jahre erfahrung', 'nachweisbare erfahrung',
    'du hast erfahrung', 'du verfügst', 'du bringst mit'
  ];

  const hasTask = taskSignals.some(s => lower.includes(s));
  const hasReq  = reqSignals.some(s => lower.includes(s));

  return hasTask && hasReq;
}

/**
 * Returns false if the text is clearly a partial/truncated job description.
 * Aggregator sites often show excerpts ending with "read more" links.
 */
function isCompleteJobDescription(text) {
  if (!text) return false;
  const lower = text.toLowerCase();
  const partialSignals = [
    'zur kompletten stellenbeschreibung',
    'zur vollständigen stellenbeschreibung',
    'vollständige stellenbeschreibung',
    'zum vollständigen stellenangebot',
    'vollständige jobbeschreibung',
    'zur jobbeschreibung',
    'see full job description',
    'view full job',
    'read the full',
    'weiterlesen',
    'mehr anzeigen',
    'show more',
    'jobviewtrack.com',      // known redirect tracker used by aggregators
    'klicken sie hier für'
  ];
  const isPartial = partialSignals.some(s => lower.includes(s));
  if (isPartial) Logger.log(`  ✗ Truncated JD detected — discarding`);
  return !isPartial;
}

/**
 * Returns false if the URL is clearly a category/search-results page
 * rather than a single specific job posting.
 * Prevents Tavily from using aggregator listing pages as JD sources.
 */
function isJobDetailPage(url) {
  if (!url) return false;
  const lower = url.toLowerCase();

  // Existing job board search/listing patterns
  if (/stepstone\.[a-z]+\/jobs\/[^?#/]+\/in-[^?#/]+/.test(lower))  return false;
  if (/indeed\.com\/(jobs|jobsearch|\?q=)/.test(lower))             return false;
  if (/xing\.com\/jobs(?:\/search|\/?$)/.test(lower))               return false;
  if (/glassdoor\.[a-z]+\/(job-listings?|Jobs\/jobs)/.test(lower))  return false;
  if (/jobs\.google\.com\/search/.test(lower))                      return false;
  if (/linkedin\.com\/company\//.test(lower))                       return false;
  if (/linkedin\.com\/jobs\/search/.test(lower))                    return false;

  // NEW: reject root domains and known careers homepage paths
  // e.g. https://media-karriere.de  →  root domain, no job ID
  try {
    const parsed    = new URL(url);
    const pathParts = parsed.pathname.replace(/\/$/, '').split('/').filter(Boolean);
    // Root domain with no path
    if (pathParts.length === 0) return false;
    // Shallow careers/jobs landing pages with no job-specific ID segment
    if (pathParts.length === 1 &&
        /^(karriere|jobs?|stellen|stellenangebote?|vacancies|careers?|arbeiten|offene-stellen|jobs-de)$/.test(pathParts[0])) {
      return false;
    }
  } catch(e) {} // malformed URL — fall through to true

  return true;
}

/**
 * Returns true if the fetched text is actually about the right job.
 * Prevents JABA from using a competitor's careers page or unrelated content.
 */
function isJdRelevantToJob(text, company, jobTitle, trustedSource) {
  if (!text || text.length < 200) return false;
  const lower = text.toLowerCase();
 
  const titleWords = jobTitle
    .toLowerCase()
    .replace(/\([mwfdx*\/\s]{3,15}\)/gi, '')
    .replace(/\(all\s*genders?\)/gi, '')
    .replace(/\(gn\)/gi, '')
    .replace(/\(mensch\)/gi, '')
    .split(/\s+/)
    .filter(w => w.length > 4);
 
  // ── Trusted source path (Remotive, Jobicy, Arbeitnow, RemoteOK) ───────────
  // Company name is often stripped from raw API descriptions.
  // Title match alone is sufficient evidence.
  if (trustedSource) {
    const titleMatch = titleWords.some(w => lower.includes(w));
    if (!titleMatch) {
      Logger.log(`  ✗ Relevance (trusted): no title words from "${jobTitle}" in text`);
    }
    return titleMatch;
  }
 
  // ── External/fetched source path — require both company AND title ──────────
  const companyWords = company
    .toLowerCase()
    .replace(/\s+(gmbh\s*&\s*co\.?\s*kg|gmbh|ag|se|kg|ohg|ug|ltd|inc|corp|llc|sas|bv|nv|ab)\.?/gi, '')
    .replace(/[&.,\-]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length > 3);
 
  // Split text into individual tokens to prevent substring false positives.
  // "Stadt" (company name) must not match inside "Vollzeit", "Stadtbahn", etc.
  const textTokens = new Set(
    lower.split(/[\s,.\-\/()[\]{}|:!?;*+]+/).filter(w => w.length > 2)
  );
  const companyMatch = companyWords.length > 0 &&
    companyWords.some(w => textTokens.has(w));
 
  if (!companyMatch) {
    Logger.log(`  ✗ Relevance: company "${company}" not found in fetched text`);
    return false;
  }
 
  const titleMatch = titleWords.some(w => lower.includes(w));
  if (!titleMatch) {
    Logger.log(`  ✗ Relevance: company found but no title words from "${jobTitle}"`);
    return false;
  }
 
  return true;
}

/**
 * Returns true if a job location string indicates Germany.
 * Used to decide whether remote-only filter applies.
 */
function isJobRemoteEligible(job) {
  if (/^(remote|homeoffice|home\s*office|home-office)$/i.test((job.city || '').trim())) return true;
  if ((job.country || '').toLowerCase() === 'de') return true;
  if (isGermanLocation(job.city || '')) return true;
  const combined = ((job.title || '') + ' ' +
    (job.description || '').substring(0, 600)).toLowerCase();
  return /remote|homeoffice|home\s*office|home-office/.test(combined);
}

function isGermanLocation(location) {
  if (!location || location.trim() === '') return true; // no location = assume Germany
  const lower = location.toLowerCase();
  const germanSignals = [
    'deutschland', 'germany', 'berlin', 'münchen', 'munich', 'hamburg',
    'frankfurt', 'köln', 'cologne', 'stuttgart', 'düsseldorf', 'dortmund',
    'essen', 'bremen', 'hannover', 'nürnberg', 'nuremberg', 'leipzig',
    'dresden', 'bonn', 'mannheim', 'karlsruhe', 'augsburg', 'wiesbaden',
    'freiburg', 'mainz', 'rostock', 'kassel', 'potsdam', 'saarbrücken',
    'darmstadt', 'heidelberg', 'regensburg', 'würzburg', 'wolfsburg',
    'ulm', 'heilbronn', 'erfurt', 'magdeburg', 'kiel', 'lübeck',
    'osnabrück', 'oldenburg', 'braunschweig', 'aachen', 'sassnitz',
    ' de,', ', de', '(de)', ' de '
  ];
  return germanSignals.some(s => lower.includes(s));
}

// ── Tavily advanced extractor ─────────────────────────────────────────────────

function tavilyExtractAdvanced(url) {
  const key = getTavilyKey();
  if (!key) { Logger.log('  TAVILY_API_KEY not set'); return null; }

  try {
    const res = UrlFetchApp.fetch('https://api.tavily.com/extract', {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify({ api_key: key, urls: [url], extract_depth: 'advanced' }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log(`  Tavily advanced HTTP ${res.getResponseCode()}`);
      return null;
    }

    const data   = JSON.parse(res.getContentText());
    const result = data.results && data.results[0];
    if (!result || !result.raw_content) return null;

    incrementTavilyCounter(2);
    return result.raw_content.substring(0, 10000);

  } catch (e) {
    Logger.log(`  Tavily advanced exception: ${e.message}`);
    return null;
  }
}


// ── Direct HTML fetch ─────────────────────────────────────────────────────────

function fetchJobPageDirectly(url) {
  if (!url) return null;
  if (DIRECT_FETCH_BLOCKED.some(d => url.includes(d))) return null;

  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects:    true,
      headers: {
        'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept-Language': 'de-DE,de;q=0.9,en;q=0.8',
        'Accept':          'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
      }
    });

    if (res.getResponseCode() !== 200) return null;

    const text = stripHtmlToText(res.getContentText());
    return text && text.length >= 100 ? text.substring(0, 10000) : null;

  } catch (e) {
    Logger.log(`  Direct fetch exception: ${e.message}`);
    return null;
  }
}


// ── Adzuna API ────────────────────────────────────────────────────────────────

function getAdzunaAppId()  { return getScriptProperty('ADZUNA_APP_ID');  }
function getAdzunaAppKey() { return getScriptProperty('ADZUNA_APP_KEY'); }

/**
 * Extract the most specific city from an Adzuna location object.
 * location.area is an array like ["Germany", "Bavaria", "Munich"].
 */
function extractAdzunaCity(location) {
  if (!location) return '';
  const area = location.area || [];
  for (let i = area.length - 1; i >= 0; i--) {
    const part = area[i];
    if (part && part !== 'Germany' && part !== 'Deutschland') return part;
  }
  return (location.display_name || '').replace(/, Germany$/, '').trim();
}

/**
 * Search Adzuna for one keyword in Germany.
 * Returns array of { title, company, city, url, description, id }.
 * The description field is already clean text — used as Tier 1 JD source.
 */
function fetchAdzunaJobs(keyword) {
  const appId  = getAdzunaAppId();
  const appKey = getAdzunaAppKey();
  if (!appId || !appKey) {
    Logger.log('ADZUNA_APP_ID or ADZUNA_APP_KEY not set in Script Properties');
    return [];
  }

  const url = 'https://api.adzuna.com/v1/api/jobs/de/search/1?' +
    `app_id=${encodeURIComponent(appId)}&` +
    `app_key=${encodeURIComponent(appKey)}&` +
    `results_per_page=50&` +
    `what=${encodeURIComponent(keyword)}&` +
    `sort_by=date&` +
    `content-type=application/json`;

  try {
    const res = UrlFetchApp.fetch(url, {
      method:             'get',
      headers:            { 'Accept': 'application/json' },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log(`Adzuna HTTP ${res.getResponseCode()} for "${keyword}": ${res.getContentText().substring(0, 200)}`);
      return [];
    }

    const data    = JSON.parse(res.getContentText());
    const results = data.results || [];

    const jobs = results.map(job => ({
      title:       (job.title || '').trim(),
      company:     (job.company && job.company.display_name ? job.company.display_name : 'Unknown').trim(),
      city:        extractAdzunaCity(job.location),
      url:         job.redirect_url || '',
      description: stripHtmlToText(job.description || '').substring(0, 10000),
      id:          String(job.id || ''),
      pubDate:     job.created || ''
    })).filter(j => j.title && j.url);

    Logger.log(`Adzuna "${keyword}": ${jobs.length} jobs`);
    return jobs;

  } catch (e) {
    Logger.log(`Adzuna error for "${keyword}": ${e.message}`);
    return [];
  }
}

// ── REPLACE fetchAllJobSources() ─────────────────────────────────────────────
// Change: added sourceStats tracking + structured summary log at end.
// Everything else identical.
 
function fetchAllJobSources() {
  const seenIds    = new Set();
  const all        = [];
  const sourceStats = {}; // ← new: per-source diagnostics
 
  function addJobs(jobs, sourceName) {
    let added = 0;
    for (const job of jobs) {
      if (!seenIds.has(job.id)) {
        seenIds.add(job.id);
        all.push(job);
        added++;
      }
    }
    sourceStats[sourceName] = { fetched: jobs.length, added }; // ← new
    Logger.log(`  ${sourceName}: +${added} new (${jobs.length} fetched)`);
  }
 
  Logger.log('--- Fetching job sources ---');
 
  // Full-JD sources (no Tavily needed)
  addJobs(fetchRemotiveJobs(),      'Remotive');
  Utilities.sleep(500);
  addJobs(fetchJobicyJobs(),        'Jobicy');
  Utilities.sleep(500);
  addJobs(fetchArbeitnowJobs(),     'Arbeitnow');
  Utilities.sleep(500);
  addJobs(fetchRemoteOkJobs(),      'Remote OK');
  Utilities.sleep(500);
  addJobs(fetchWorkingNomadsJobs(), 'Working Nomads');
  Utilities.sleep(500);
 
  // Adzuna multi-country (snippet, Tavily used when needed)
  Logger.log('Fetching Adzuna multi-country...');
  const adzunaKeywords = [
    'marketing manager', 'digital marketing',
    'crm manager', 'marketing automation'
  ];
  for (const keyword of adzunaKeywords) {
    for (const country of ADZUNA_SEARCH_COUNTRIES) {
      addJobs(fetchAdzunaJobsForCountry(keyword, country), `Adzuna-${country}-${keyword}`);
      Utilities.sleep(300);
    }
  }
 
  // ── NEW: emit per-source summary ──────────────────────────────────────────
  let totalFetched = 0;
  let totalAdded   = 0;
  const sourceLines = Object.entries(sourceStats).map(([src, s]) => {
    totalFetched += s.fetched;
    totalAdded   += s.added;
    return `    ${src.padEnd(32)}: fetched=${String(s.fetched).padStart(3)}  added=${String(s.added).padStart(3)}`;
  });
  Logger.log([
    '',
    '━━━ SOURCE SUMMARY ━━━━━━━━━━━━━━━━━━━━━━━━━━',
    ...sourceLines,
    `    ${'TOTAL'.padEnd(32)}: fetched=${String(totalFetched).padStart(3)}  added(dedup)=${String(totalAdded).padStart(3)}`,
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    ''
  ].join('\n'));
  // ── END NEW ───────────────────────────────────────────────────────────────
 
  Logger.log(`Total jobs all sources (deduplicated): ${all.length}`);
  return all;
}

// ── ADD NEW HELPER: emitDiagnosticsSummary() ─────────────────────────────────
// Add this as a new standalone function anywhere in Code.js (e.g. after fetchAllJobSources).
 
function emitDiagnosticsSummary(diag) {
  const delta = (diag.tavily_end || 0) - (diag.tavily_start || 0);
  Logger.log([
    '',
    '━━━ RUN DIAGNOSTICS ━━━━━━━━━━━━━━━━━━━━━━━━━',
    '  ── pipeline gates ─────────────────────────',
    `  fetched_total     : ${diag.fetched_total}`,
    `  excluded_title    : ${diag.excluded_title}`,
    `  excluded_geo      : ${diag.excluded_geo}`,
    `  excluded_cache    : ${diag.excluded_cache}`,
    `  excluded_applied  : ${diag.excluded_applied}`,
    `  candidate_selected: ${diag.candidate_selected}`,
    '  ── processing ──────────────────────────────',
    `  processed_count   : ${diag.processed_count}`,
    `  jd_fetch_failed   : ${diag.jd_fetch_failed}`,
    `  jd_irrelevant     : ${diag.jd_irrelevant}`,
    `  smm_failed        : ${diag.smm_failed}`,
    '  ── scores ──────────────────────────────────',
    `  scored_m0         : ${diag.scored_m0}`,
    `  scored_m1         : ${diag.scored_m1}`,
    `  scored_m2_plus    : ${diag.scored_m2_plus}`,
    `  report_jobs_count : ${diag.report_jobs_count}`,
    '  ── tavily ──────────────────────────────────',
    `  tavily_start      : ${diag.tavily_start}/1000`,
    `  tavily_end        : ${diag.tavily_end}/1000`,
    `  tavily_delta      : +${delta} credits this run`,
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    ''
  ].join('\n'));
}

// ── ADD NEW HELPER: sendDebugDiagnosticsEmail() ──────────────────────────────
// Called when a run completes with zero M2+ jobs.
// Uses REPORT_EMAIL script property (same one the job report uses).
 
function sendDebugDiagnosticsEmail(diag) {
  const recipient = getScriptProperty('REPORT_EMAIL');
  if (!recipient) {
    Logger.log('⚠ REPORT_EMAIL not set — debug email skipped.');
    return;
  }
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy HH:mm');
  const delta   = (diag.tavily_end || 0) - (diag.tavily_start || 0);
 
  function row(label, value, highlight) {
    const style = highlight
      ? 'padding:6px 12px;font-weight:700;color:#ea4335;'
      : 'padding:6px 12px;color:#5f6368;';
    return `<tr>
      <td style="${style}">${label}</td>
      <td style="padding:6px 12px;font-weight:700;text-align:right;">${value}</td>
    </tr>`;
  }
 
  function section(title) {
    return `<tr style="background:#f8f9fa;">
      <td colspan="2" style="padding:8px 12px;font-weight:700;font-size:12px;
        color:#3c4043;text-transform:uppercase;letter-spacing:0.4px;">${title}</td>
    </tr>`;
  }
 
  const html = `<!DOCTYPE html>
<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;background:#f8f9fa;padding:20px;margin:0;">
<div style="max-width:520px;margin:0 auto;">
 
  <div style="background:#ea4335;border-radius:10px;padding:16px 20px;margin-bottom:20px;">
    <h2 style="color:white;margin:0 0 4px;font-size:17px;">🤖 JABA — No M2+ Jobs Found</h2>
    <div style="color:rgba(255,255,255,0.85);font-size:13px;">${dateStr}</div>
  </div>
 
  <div style="background:white;border-radius:10px;border:1px solid #e0e0e0;
    padding:4px 0;margin-bottom:16px;">
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
 
      ${section('Pipeline gates')}
      ${row('Total fetched',       diag.fetched_total)}
      ${row('Excluded — title',    diag.excluded_title,   diag.excluded_title > 20)}
      ${row('Excluded — geo',      diag.excluded_geo,     diag.excluded_geo > 10)}
      ${row('Excluded — cache',    diag.excluded_cache,   diag.excluded_cache > 20)}
      ${row('Excluded — applied',  diag.excluded_applied, diag.excluded_applied > 10)}
      ${row('Candidates selected', diag.candidate_selected)}
 
      ${section('Processing')}
      ${row('Processed',           diag.processed_count)}
      ${row('JD fetch failed',     diag.jd_fetch_failed,  diag.jd_fetch_failed > 2)}
      ${row('JD irrelevant',       diag.jd_irrelevant,    diag.jd_irrelevant > 2)}
      ${row('SMM failed',          diag.smm_failed,       diag.smm_failed > 1)}
 
      ${section('Score distribution')}
      ${row('M0 (0–10)',           diag.scored_m0)}
      ${row('M1 (11–20)',          diag.scored_m1)}
      ${row('M2+ (21+) ✓',        diag.scored_m2_plus)}
 
      ${section('Tavily')}
      ${row('Credits used this run', '+' + delta)}
      ${row('Total this month',      diag.tavily_end + '/1000',
            diag.tavily_end > 900)}
 
    </table>
  </div>
 
  <div style="font-size:11px;color:#9aa0a6;text-align:center;padding-bottom:20px;">
    JABA 🤖 · Phase 1 Diagnostics · Values in red exceeded expected thresholds
  </div>
 
</div>
</body>
</html>`;
 
  GmailApp.sendEmail(
    recipient,
    `JABA Debug — No M2+ jobs · ${dateStr}`,
    '',
    { htmlBody: html, name: 'JABA 🤖' }
  );
  Logger.log(`Debug diagnostics email sent to ${recipient}`);
}

/* ── Remotive API (free, full JD, remote-first) ── */
function fetchRemotiveJobs() {
  const categories = ['marketing', 'business'];
  const seenIds = new Set();
  const all = [];

  for (const cat of categories) {
    try {
      const res = UrlFetchApp.fetch(
        `https://remotive.com/api/remote-jobs?category=${cat}&limit=50`,
        { muteHttpExceptions: true }
      );
      if (res.getResponseCode() !== 200) continue;
      const data = JSON.parse(res.getContentText());
      for (const job of (data.jobs || [])) {
        const id = 'remotive_' + job.id;
        if (seenIds.has(id)) continue;
        seenIds.add(id);
        all.push({
          title:           (job.title || '').trim(),
          company:         (job.company_name || 'Unknown').trim(),
          city:            'Remote',
          url:             job.url || '',
          description:     stripHtmlToText(job.description || '').substring(0, 10000),
          id:              id,
          pubDate:         job.publication_date || '',
          descriptionFull: true   // ← full JD, skip Tavily
        });
      }
      Utilities.sleep(300);
    } catch(e) {
      Logger.log(`Remotive error (${cat}): ${e.message}`);
    }
  }
  Logger.log(`Remotive: ${all.length} jobs`);
  return all;
}

/* ── Jobicy API (free, full JD, remote) ── */
function fetchJobicyJobs() {
  try {
    const res = UrlFetchApp.fetch(
      'https://jobicy.com/api/v0/remote-jobs?count=50&industry=marketing',
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return [];
    const data = JSON.parse(res.getContentText());
    return (data.jobs || []).map(job => ({
      title:           (job.jobTitle || '').trim(),
      company:         (job.companyName || 'Unknown').trim(),
      city:            'Remote',
      url:             job.url || '',
      description:     stripHtmlToText(job.jobDescription || '').substring(0, 10000),
      id:              'jobicy_' + (job.id || Math.random()),
      pubDate:         job.pubDate || '',
      descriptionFull: true
    }));
  } catch(e) {
    Logger.log(`Jobicy error: ${e.message}`);
    return [];
  }
}

/* ── Arbeitnow API (free, EU jobs, full JD) ── */
function fetchArbeitnowJobs() {
  const allJobs = [];
  const seenIds = new Set();
  for (let page = 1; page <= 3; page++) {
    try {
      const res = UrlFetchApp.fetch(
        `https://www.arbeitnow.com/api/job-board-api?page=${page}`,
        { muteHttpExceptions: true }
      );
      if (res.getResponseCode() !== 200) {
        Logger.log(`Arbeitnow page ${page}: HTTP ${res.getResponseCode()}`);
        break;
      }
      const data = JSON.parse(res.getContentText());
      const jobs  = data.data || [];
      if (jobs.length === 0) break;
      for (const job of jobs) {
        const id    = 'arbeitnow_' + (job.slug || '');
        if (seenIds.has(id)) continue;
        seenIds.add(id);
        const title = (job.title || '').trim();
        if (!isRelevantJobTitle(title)) continue;  // title filter only; no geo gate
        allJobs.push({
          title,
          company:         (job.company_name || 'Unknown').trim(),
          city:            (job.location || 'Remote').trim(),
          url:             job.url || '',
          description:     stripHtmlToText(job.description || '').substring(0, 10000),
          id,
          pubDate:         job.created_at || '',
          descriptionFull: true
        });
      }
      Utilities.sleep(400);
    } catch(e) {
      Logger.log(`Arbeitnow page ${page} error: ${e.message}`);
      break;
    }
  }
  Logger.log(`Arbeitnow: ${allJobs.length} relevant jobs (no geo filter)`);
  return allJobs;
}

/* ── Remote OK API (free, international remote, full JD) ── */
function fetchRemoteOkJobs() {
  try {
    const res = UrlFetchApp.fetch('https://remoteok.com/api', {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; personal job search bot)'
      }
    });
    if (res.getResponseCode() !== 200) {
      Logger.log(`Remote OK HTTP ${res.getResponseCode()}`);
      return [];
    }

    const data = JSON.parse(res.getContentText());
    // First element is API metadata — skip it
    const jobs = Array.isArray(data) ? data.slice(1) : [];

    const results = [];
    for (const job of jobs) {
      const title = (job.position || '').trim();
      const tags  = Array.isArray(job.tags) ? job.tags.join(' ').toLowerCase() : '';

      // All Remote OK jobs are remote by nature — filter only by relevance
      if (!isRelevantJobTitle(title) &&
          !RELEVANT_KEYWORDS.some(k => tags.includes(k.toLowerCase()))) continue;

      results.push({
        title:           title,
        company:         (job.company || 'Unknown').trim(),
        city:            'Remote',
        url:             job.url || `https://remoteok.com/l/${job.slug || ''}`,
        description:     stripHtmlToText(job.description || '').substring(0, 10000),
        id:              'remoteok_' + (job.id || Math.random()),
        pubDate:         job.date || '',
        descriptionFull: true
      });
    }

    Logger.log(`Remote OK: ${results.length} relevant jobs`);
    return results;

  } catch(e) {
    Logger.log(`Remote OK error: ${e.message}`);
    return [];
  }
}

/* ── Working Nomads API (free, no key, remote jobs) ── */
function fetchWorkingNomadsJobs() {
  try {
    const res = UrlFetchApp.fetch(
      'https://www.workingnomads.com/api/exposed_jobs/?category=marketing',
      {
        method: 'get',
        headers: { 'Accept': 'application/json' },
        muteHttpExceptions: true
      }
    );
    if (res.getResponseCode() !== 200) {
      Logger.log(`Working Nomads HTTP ${res.getResponseCode()}`);
      return [];
    }

    const data = JSON.parse(res.getContentText());
    const jobs = Array.isArray(data) ? data : (data.jobs || []);

    const results = [];
    for (const job of jobs) {
      const title = (job.title || '').trim();
      if (!isRelevantJobTitle(title)) continue;

      // Working Nomads descriptions are HTML — strip tags
      const rawDesc = job.description || job.short_description || '';
      const desc    = stripHtmlToText(rawDesc).substring(0, 10000);

      results.push({
        title:           title,
        company:         (job.company_name || job.company || 'Unknown').trim(),
        city:            'Remote',
        url:             job.url || job.job_url || '',
        description:     desc,
        id:              'workingnomads_' + (job.id || Math.random()),
        pubDate:         job.pub_date || job.created_at || '',
        descriptionFull: desc.length > 300  // treat as full if description is substantial
      });
    }

    Logger.log(`Working Nomads: ${results.length} relevant jobs`);
    return results;

  } catch(e) {
    Logger.log(`Working Nomads error: ${e.message}`);
    return [];
  }
}

/* ── Adzuna multi-country (remote filter) ── */
const ADZUNA_SEARCH_COUNTRIES = [
  'gb', 'nl', 'se', 'no', 'be', 'at', 'ch', 'es', 'pl', 'fr', 'de'
];

function fetchAdzunaJobsForCountry(keyword, countryCode) {
  const appId  = getAdzunaAppId();
  const appKey = getAdzunaAppKey();
  if (!appId || !appKey) return [];
 
  const url = `https://api.adzuna.com/v1/api/jobs/${countryCode}/search/1?` +
    `app_id=${encodeURIComponent(appId)}&` +
    `app_key=${encodeURIComponent(appKey)}&` +
    `results_per_page=20&` +
    `what=${encodeURIComponent(countryCode === 'de' ? keyword : keyword + ' remote')}&` +
    `sort_by=date&` +
    `content-type=application/json`;
 
  try {
    const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return [];
    const data = JSON.parse(res.getContentText());
    return (data.results || []).map(job => ({
      title:           (job.title || '').trim(),
      company:         (job.company?.display_name || 'Unknown').trim(),
      city:            extractAdzunaCity(job.location) || countryCode.toUpperCase(),
      country:         countryCode,   // ← NEW: store country code directly
      url:             job.redirect_url || '',
      description:     stripHtmlToText(job.description || '').substring(0, 10000),
      id:              `adzuna_${countryCode}_${job.id}`,
      pubDate:         job.created || '',
      descriptionFull: false
    })).filter(j => j.title && j.url);
  } catch(e) {
    Logger.log(`Adzuna ${countryCode} error for "${keyword}": ${e.message}`);
    return [];
  }
}

// ── Smart JD extractor ────────────────────────────────────────────────────────
// Tier 1: Adzuna API description (free, already fetched — covers most cases)
// Tier 2: Direct UrlFetchApp HTML fetch (free, unlimited)
// Tier 3: Tavily advanced (2 credits — only fires when tiers 1 & 2 fail)

function smartExtractJD(url, apiDescription, descriptionFull) {
  // Remotive / Jobicy / Arbeitnow / RemoteOK — full JD already in API payload
  if (descriptionFull && apiDescription && looksLikeJobContent(apiDescription)) {
    Logger.log(`  ✓ Full API description (${apiDescription.length} chars)`);
    return { text: apiDescription, source: 'api_full' };
  }

  // Tier 1 — Direct UrlFetchApp
  // Requires 500+ chars AND no truncation markers to prevent score inflation
  const direct = fetchJobPageDirectly(url);
  if (direct && direct.length >= 500 && looksLikeJobContent(direct) && isCompleteJobDescription(direct)) {
    Logger.log(`  ✓ Direct fetch OK (${direct.length} chars)`);
    return { text: direct, source: 'direct' };
  }
  if (direct) Logger.log(`  ✗ Direct discarded — length: ${direct.length}, complete: ${isCompleteJobDescription(direct)}`);

  // Tier 2 — Tavily advanced
  const tavily = tavilyExtractAdvanced(url);
  if (tavily && tavily.length >= 500 && looksLikeJobContent(tavily) && isCompleteJobDescription(tavily)) {
    Logger.log(`  ✓ Tavily advanced OK (${tavily.length} chars)`);
    return { text: tavily, source: 'tavily_advanced' };
  }
  if (tavily) Logger.log(`  ✗ Tavily discarded — length: ${tavily.length}, complete: ${isCompleteJobDescription(tavily)}`);

  // api_snippet fallback intentionally removed.
  // Short Adzuna snippets (~200-400 chars) inflate scores because Mistral
  // sees only the attractive summary without the hard requirements.
  Logger.log(`  ✗ No complete JD found — skipping`);
  return null;
}


// ── Title filter ──────────────────────────────────────────────────────────────

function isValidJobTitleForSearch(title) {
  if (!title) return false;
  const lower = title.toLowerCase();
  return !JOB_SEARCH_EXCLUDE_TERMS.some(t => lower.includes(t));
}


// ── CV type detector (DE / EN only) ──────────────────────────────────────────

function detectCvTypeForSearch(jdText, jobTitle) {
  const combined = ((jobTitle || '') + ' ' + (jdText || '').substring(0, 800)).toLowerCase();

  // 'token' intentionally excluded — appears in too many non-Web3 contexts
  // (API tokens, auth tokens, loyalty tokens, promo codes).
  const web3Signals = ['web3', 'blockchain', 'crypto', 'defi', 'nft', 'dao',
                       'smart contract', 'solidity', 'decentralized'];
  if (web3Signals.some(s => combined.includes(s))) return 'Web3 Marketing Manager';

  const deSignals = [
    '(m/w/d)', '(w/m/d)', '(w/d/m)',
    '(m/w/x)',   // Allianz, Deutsche companies
    '(m/f/d)',   // EN-in-DE postings
    'vollzeit', 'bewerbung', 'berufserfahrung', 'aufgaben'
  ];
  if (deSignals.some(s => combined.includes(s))) return 'DE Web2 Marketing Manager';

  return 'EN Web2 Marketing Manager';
}


// ── Job_Search_Cache tab ──────────────────────────────────────────────────────

function getOrCreateJobCacheSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet   = ss.getSheetByName('Job_Search_Cache');
  const HEADERS = [
    'Timestamp','Company','Job_Title','URL','CV_Type',
    'Match_Level','Score','Fetch_Source','Fetched_URL','Source',
    'Review_Status','JD_Text'
  ];
 
  function applyDropdowns() {
    function dv(col, values) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(values, true).setAllowInvalid(true).build();
      sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
    }
    dv(5,  ['DE Web2 Marketing Manager','EN Web2 Marketing Manager','Web3 Marketing Manager','—']);
    dv(6,  ['M0','M1','M2','M3','M4','SKIP','QUEUED','already_applied']);
    dv(8,  ['api_full','direct','tavily_advanced','failed','irrelevant_jd',
            'queued','no_jd','already_applied','phase1_prefetch']);
    dv(10, ['JobSearch','Indeed','BA']);
    dv(11, ['in process','fit','discarded','registered','auto']);
  }
 
  if (!sheet) {
    sheet = ss.insertSheet('Job_Search_Cache');
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setValues([HEADERS])
         .setBackground('#34a853').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');
    for (let i = 1; i <= HEADERS.length; i++) sheet.autoResizeColumn(i);
    applyDropdowns();
    Logger.log('Job_Search_Cache sheet created (12 columns).');
  } else {
    const colCount = sheet.getLastColumn();
    if (colCount < 11) {
      // Migrate 10 → 12 columns
      sheet.getRange(1, 11, 1, 2)
           .setValues([['Review_Status','JD_Text']])
           .setBackground('#34a853').setFontColor('white').setFontWeight('bold');
      if (sheet.getLastRow() > 1) {
        const levels   = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues();
        const statuses = levels.map(r => {
          const n = parseInt((r[0] || '').toString().replace(/\D/g, '')) || 0;
          return [n >= 1 ? 'in process' : 'auto'];
        });
        sheet.getRange(2, 11, statuses.length, 1).setValues(statuses);
      }
      sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('@STRING@');
      applyDropdowns();
      Logger.log('Job_Search_Cache: migrated to 12 columns, Review_Status backfilled.');
    } else if (colCount < 12) {
      sheet.getRange(1, 12, 1, 1).setValue('JD_Text')
           .setBackground('#34a853').setFontColor('white').setFontWeight('bold');
      Logger.log('Job_Search_Cache: JD_Text column added.');
    }
  }
  return sheet;
}

function cleanJobCache() {
  const sheet = getOrCreateJobCacheSheet();
  if (sheet.getLastRow() < 2) return;
  const now            = new Date();
  const cutoffAnalyzed = new Date(now.getTime() - 14 * 24 * 60 * 60 * 1000);
  const cutoffSkip     = new Date(now.getTime() -  7 * 24 * 60 * 60 * 1000);
  const data           = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const toDelete       = [];
  data.forEach((row, i) => {
    const entryDate = parseCacheTimestamp(row[0]);
    if (!entryDate) return;
    const level     = (row[5] || '').toString().toUpperCase().trim();
    const shortLive = level === 'SKIP' || level === 'QUEUED';
    if (entryDate < (shortLive ? cutoffSkip : cutoffAnalyzed)) toDelete.push(i + 2);
  });
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
  if (toDelete.length > 0) {
    SpreadsheetApp.flush();
    Logger.log(`Job cache cleaned: ${toDelete.length} entries removed.`);
  }
}

function isJobInCache(company, title) {
  const sheet = getOrCreateJobCacheSheet();
  if (sheet.getLastRow() < 2) return false;
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
  const nc   = company.toLowerCase().trim();
  const nt   = normalizeJobTitle(title || '').toLowerCase().trim();
  return data.some(r =>
    r[0].toString().toLowerCase().trim() === nc &&
    normalizeJobTitle(r[1].toString()).toLowerCase().trim() === nt
  );
}
// Loads the entire Job_Search_Cache into a Set — call once per run
function buildCachedJobsSet() {
  const cached = new Set();
  const sheet  = getOrCreateJobCacheSheet();
  if (sheet.getLastRow() < 2) return cached;
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
  data.forEach(r => {
    cached.add(
      (r[0] || '').toString().toLowerCase().trim() + '||' +
      normalizeJobTitle((r[1] || '').toString()).toLowerCase().trim()
    );
  });
  return cached;
}

// Loads all applied jobs from all monthly sheets into a Set — call once per run
function buildAppliedJobsSet() {
  const applied    = new Set();
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const skipSheets = new Set(['Sankey_Data', 'Geo_Data', 'SMM_Raw_Data',
                               'Job_Search_Cache', 'Pending_SMM']);
  for (const sheet of ss.getSheets()) {
    const name = sheet.getName();
    if (skipSheets.has(name) || !/\d{4}/.test(name)) continue;
    if (sheet.getLastRow() < 2) continue;
    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
    data.forEach(function(r) {
      applied.add(
        (r[0] || '').toString().toLowerCase().trim() + '||' +
        normalizeJobTitle((r[1] || '').toString()).toLowerCase().trim()
      );
    });
  }
  return applied;
}

function isJobAlreadyApplied(company, title) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const nc  = company.toLowerCase().trim();
  const nt  = normalizeJobTitle(title || '').toLowerCase().trim();
  const skipSheets = new Set(['Sankey_Data', 'Geo_Data', 'SMM_Raw_Data', 'Job_Search_Cache']);
 
  for (const sheet of ss.getSheets()) {
    const name = sheet.getName();
    if (skipSheets.has(name) || !/\d{4}/.test(name)) continue;
    if (sheet.getLastRow() < 2) continue;
 
    // Read company (col B) + position (col C) together
    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
    if (data.some(r => {
      const rc = (r[0] || '').toString().toLowerCase().trim();
      const rt = normalizeJobTitle((r[1] || '').toString()).toLowerCase().trim();
      return rc === nc && rt === nt;
    })) return true;
  }
  return false;
}

function addJobToCache(job, smmResult, cvType, fetchSource, fetchedUrl, source, jdText) {
  const sheet        = getOrCreateJobCacheSheet();
  const timestamp    = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd-MM-yyyy HH:mm:ss');
  const levelNum     = parseInt((smmResult.match_level || '').replace(/\D/g, '')) || 0;
  const reviewStatus = levelNum >= 1 ? 'in process' : 'auto';
  const safeJdText   = (jdText || '').substring(0, 49000);
  sheet.appendRow([
    timestamp,
    job.company || '',
    job.title   || '',
    job.url     || '',
    cvType      || '',
    smmResult.match_level || 'M0',
    smmResult.total_score || 0,
    fetchSource  || '',
    fetchedUrl   || '',
    source       || 'JobSearch',
    reviewStatus,
    safeJdText
  ]);
}


// ── Email report builder ──────────────────────────────────────────────────────

function buildJobDots(smmResult) {
  const skills = smmResult.skills || [];
  function check(imp, dot) {
    const g = skills.filter(s => s.importance === imp);
    return g.length > 0 && g.every(s => (s.score || 0) >= 1) ? dot : '';
  }
  return check('Crucial', '🟢') + check('Necessary', '🟡') + check('Optional', '🔵');
}

function escapeHtmlEmail(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildJobReportHtml(jobs) {
  const dateStr     = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy');
  const tavilyMonth = getTavilyMonthlyUsage();
  const impDot      = { 'Crucial': '🟢', 'Necessary': '🟡', 'Optional': '🔵' };

  function levelColor(score) {
    if (score >= 36) return '#34a853';
    if (score >= 30) return '#4fc3f7';
    if (score >= 21) return '#f4b400';
    return '#ff6d00';
  }

  const cards = jobs.map(job => {
    const smm    = job.smmResult;
    const score  = smm.total_score || 0;
    const level  = smm.match_level || 'M0';
    const dots   = buildJobDots(smm);
    const color  = levelColor(score);
    const skills = smm.skills || [];

    const skillRows = skills.map(s => `
      <tr>
        <td style="padding:5px 10px;font-size:13px;color:#3c4043;border-bottom:1px solid #f1f3f4;">${escapeHtmlEmail(s.name)}</td>
        <td style="padding:5px 10px;font-size:13px;font-weight:700;color:#3c4043;text-align:center;border-bottom:1px solid #f1f3f4;white-space:nowrap;">${s.score || 0}/5</td>
        <td style="padding:5px 10px;font-size:12px;text-align:center;border-bottom:1px solid #f1f3f4;white-space:nowrap;">${impDot[s.importance] || '⚪'} ${escapeHtmlEmail(s.importance)}</td>
      </tr>`).join('');

    return `
    <div style="background:#ffffff;border-radius:10px;border:1px solid #e0e0e0;padding:18px 20px;margin-bottom:22px;box-shadow:0 1px 3px rgba(0,0,0,0.06);">
      <div style="font-size:16px;font-weight:700;color:#202124;margin-bottom:4px;">
        🏢 ${escapeHtmlEmail(job.title)} – ${escapeHtmlEmail(job.company)}
      </div>
      <div style="font-size:13px;color:#5f6368;margin-bottom:8px;">
        📍 ${escapeHtmlEmail(job.city || 'Germany')}
        &nbsp;|&nbsp;
        <span style="font-weight:700;color:${color};">Score: ${score}/40 &nbsp;|&nbsp; ${level}</span>
        &nbsp;${dots}
        &nbsp;|&nbsp;
        <span style="font-size:11px;background:#f1f3f4;padding:2px 6px;border-radius:4px;color:#5f6368;">JD: ${escapeHtmlEmail(job.fetchSource || 'unknown')}</span>
      </div>
      <div style="margin-bottom:14px;">
        <a href="${escapeHtmlEmail(job.url)}" style="display:inline-block;background:#4285f4;color:white;font-size:13px;font-weight:700;padding:6px 14px;border-radius:6px;text-decoration:none;">🔗 View &amp; Apply →</a>
      </div>
      <table style="width:100%;border-collapse:collapse;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">
        <thead>
          <tr style="background:#f8f9fa;">
            <th style="padding:7px 10px;font-size:11px;color:#5f6368;text-align:left;text-transform:uppercase;letter-spacing:0.4px;font-weight:600;">Skill</th>
            <th style="padding:7px 10px;font-size:11px;color:#5f6368;text-align:center;text-transform:uppercase;letter-spacing:0.4px;font-weight:600;">Score</th>
            <th style="padding:7px 10px;font-size:11px;color:#5f6368;text-align:center;text-transform:uppercase;letter-spacing:0.4px;font-weight:600;">Importance</th>
          </tr>
        </thead>
        <tbody>
          ${skillRows}
          <tr style="background:#f8f9fa;">
            <td colspan="2" style="padding:7px 10px;font-size:13px;font-weight:700;text-align:right;color:#3c4043;">Total</td>
            <td style="padding:7px 10px;font-size:15px;font-weight:800;text-align:center;color:${color};">${score}/40</td>
          </tr>
        </tbody>
      </table>
    </div>`;
  }).join('');

  return `<!DOCTYPE html>
<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;background:#f8f9fa;padding:20px;margin:0;">
<div style="max-width:660px;margin:0 auto;">
  <div style="background:#4285f4;border-radius:10px;padding:18px 22px;margin-bottom:24px;">
    <h2 style="color:white;margin:0 0 4px;font-size:18px;">🤖 JABA Job Report — ${escapeHtmlEmail(dateStr)}</h2>
    <div style="color:rgba(255,255,255,0.88);font-size:13px;">
      ${jobs.length} job(s) at M2 or above &nbsp;·&nbsp; Tavily this month: ${tavilyMonth}/1000 credits
    </div>
  </div>
  ${cards}
  <div style="font-size:11px;color:#9aa0a6;text-align:center;margin-top:12px;padding-bottom:20px;">
    Generated by JABA · Only M2 / M3 / M4 matches shown
  </div>
</div>
</body>
</html>`;
}

function sendJobReportEmail(htmlBody) {
  const recipient = getScriptProperty('REPORT_EMAIL');
  if (!recipient) {
    Logger.log('⚠ REPORT_EMAIL not set in Script Properties — email not sent.');
    return;
  }
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy');
  GmailApp.sendEmail(recipient, `JABA Job Report ${dateStr}`, '', { htmlBody, name: 'JABA 🤖' });
  Logger.log(`Job report sent to ${recipient}`);
}

function scoreJobTitleQuality(title) {
  if (!title) return 0;
  const lower = title.toLowerCase();
 
  const highValue = [
    'crm', 'lifecycle', 'retention', 'email marketing', 'marketing automation',
    'automation manager', 'growth marketing', 'performance marketing',
    'digital marketing', 'demand generation', 'demand gen', 'acquisition',
    'martech', 'onlinemarketing', 'online marketing', 'campaign manager',
    'campaign marketing', 'b2b marketing', 'wachstum', 'email-marketing',
    'marketing manager', 'digital manager'
  ];
 
  const lowValue = [
    'field marketing', 'design director', 'designer',
    'support specialist', 'shopper', 'coordinator'
  ];
 
  // Base scoring
  const highSpecific = ['crm', 'lifecycle', 'retention', 'email marketing', 'marketing automation'];
  let score = 0;
  highValue.forEach(t => {
    if (lower.includes(t)) score += highSpecific.includes(t) ? 3 : 2;
  });
  lowValue.forEach(t => {
    if (lower.includes(t)) score -= 1;
  });
 
  // ── Conditional penalties ────────────────────────────────────────────────
 
  // "Account manager" is a sales role UNLESS it contains marketing/growth context
  if (lower.includes('account manager')) {
    const hasMarketingContext = ['crm', 'lifecycle', 'retention', 'automation', 'marketing', 'growth']
      .some(t => lower.includes(t));
    if (!hasMarketingContext) score -= 4; // push below marketing roles
  }
 
  // "Freelance" without a specific marketing discipline = low priority
  if (lower.includes('freelance')) {
    const hasMarketingDiscipline = ['marketing', 'crm', 'email', 'digital', 'campaign', 'growth']
      .some(t => lower.includes(t));
    if (!hasMarketingDiscipline) score -= 3;
    else score -= 1; // small penalty even with discipline (vs. permanent role)
  }
 
  // "Content creator" — consistently M0 in scoring history, deprioritize hard
  if (lower.includes('content creator')) score -= 4;
 
  // "Community manager" alone — too narrow, unlikely M2+
  if (lower.includes('community manager') &&
      !['crm', 'growth', 'lifecycle', 'retention'].some(t => lower.includes(t))) {
    score -= 2;
  }
 
  return score;
}

// ── Main orchestrator ─────────────────────────────────────────────────────────

// ── REPLACE runDailyJobSearch() ──────────────────────────────────────────────
// Changes: added diag counter object; increment at each gate;
// call emitDiagnosticsSummary() at end; call sendDebugDiagnosticsEmail()
// when no M2+ jobs found. All filtering logic identical.
 
function runDailyJobSearch() {
  const MAX_JOBS       = 10;
  const CANDIDATE_CAP  = 40;
  const TIME_BUDGET_MS = 290000; // 4.8 min
 
  const runStart = Date.now();
 
  const diag = {
    fetched_total:      0,
    excluded_title:     0,
    excluded_geo:       0,
    excluded_cache:     0,
    excluded_applied:   0,
    candidate_selected: 0,
    candidates_deduped: 0, // ← new counter
    jd_fetch_failed:    0,
    jd_irrelevant:      0,
    smm_failed:         0,
    scored_m0:          0,
    scored_m1:          0,
    scored_m2_plus:     0,
    report_jobs_count:  0,
    processed_count:    0,
    tavily_start:       0,
    tavily_end:         0
  };
 
  Logger.log('=== JABA Daily Job Search started ===');
  diag.tavily_start = getTavilyMonthlyUsage();
  Logger.log(`Tavily credits at start of run: ${diag.tavily_start}/1000`);
 
  cleanJobCache();
 
  const allJobs = fetchAllJobSources();
  diag.fetched_total = allJobs.length;
 
  if (allJobs.length === 0) {
    Logger.log('No jobs returned today. Exiting.');
    diag.tavily_end = getTavilyMonthlyUsage();
    emitDiagnosticsSummary(diag);
    return;
  }
 
  // ── Step 3: filter candidates ─────────────────────────────────────────────
  Logger.log('Preloading cache and applied sets...');
  const cachedSet  = buildCachedJobsSet();
  const appliedSet = buildAppliedJobsSet();
  Logger.log(`Cache: ${cachedSet.size} | Applied: ${appliedSet.size}`);

  const candidates = [];
  for (const job of allJobs) {

    if (!isRelevantJobTitle(job.title)) {
      Logger.log(`  ✗ [title] "${job.title}"`);
      diag.excluded_title++;
      continue;
    }

    if (!isJobRemoteEligible(job)) {
      Logger.log(`  ✗ [geo] "${job.title}" @ ${job.city || job.country || '?'}`);
      diag.excluded_geo++;
      continue;
    }

    const cacheKey = job.company.toLowerCase().trim() + '||' +
                     normalizeJobTitle(job.title).toLowerCase().trim();

    if (cachedSet.has(cacheKey)) {
      Logger.log(`  ✗ [cache] "${job.title}" @ ${job.company}`);
      diag.excluded_cache++;
      continue;
    }

    if (appliedSet.has(cacheKey)) {
      Logger.log(`  ✗ [applied] "${job.title}" @ ${job.company}`);
      diag.excluded_applied++;
      addJobToCache(job, { match_level: 'already_applied', total_score: 0 },
                    '', 'already_applied', '', 'JobSearch', '');
      continue;
    }

    candidates.push(job);
    if (candidates.length >= CANDIDATE_CAP) break;
  }
 
  // ── NEW: deduplicate candidates by normalized company + title ─────────────
  // Prevents two Adzuna listings for the same role from consuming two slots.
  const dedupSeen = new Set();
  const dedupedCandidates = [];
  for (const job of candidates) {
    const key = `${job.company.toLowerCase().trim()}||${normalizeJobTitle(job.title).toLowerCase().trim()}`;
    if (!dedupSeen.has(key)) {
      dedupSeen.add(key);
      dedupedCandidates.push(job);
    }
  }
  diag.candidates_deduped = candidates.length - dedupedCandidates.length;
  if (diag.candidates_deduped > 0) {
    Logger.log(`  Removed ${diag.candidates_deduped} duplicate candidate(s) before scoring`);
  }
  // ── END dedup ─────────────────────────────────────────────────────────────
 
  // Sort best candidates first
  dedupedCandidates.sort((a, b) => scoreJobTitleQuality(b.title) - scoreJobTitleQuality(a.title));
  if (dedupedCandidates.length > 0) {
    Logger.log(`Top 5 after sort: ${dedupedCandidates.slice(0, 5).map(j => `"${j.title}"`).join(' | ')}`);
  }
 
  diag.candidate_selected = dedupedCandidates.length;
  Logger.log(`Candidates after filtering + dedup: ${dedupedCandidates.length}`);
 
  if (dedupedCandidates.length === 0) {
    Logger.log('No new candidates today — no email sent.');
    diag.tavily_end = getTavilyMonthlyUsage();
    emitDiagnosticsSummary(diag);
    return;
  }
 
  // ── Step 4: extract JD, score, collect M2+ ───────────────────────────────
  const reportJobs = [];
  let   processed  = 0;
 
  for (const job of dedupedCandidates) {
    if (processed >= MAX_JOBS) break;
 
    const elapsed = Date.now() - runStart;
    if (elapsed > TIME_BUDGET_MS) {
      Logger.log(`⏱ Time budget reached (${Math.round(elapsed/1000)}s) after ${processed} jobs — stopping gracefully`);
      break;
    }
 
    Logger.log(`\n[${processed + 1}/${MAX_JOBS}] "${job.title}" — ${job.company} [${job.country || 'remote'}]`);
 
    const extracted = smartExtractJD(job.url, job.description, job.descriptionFull);
    if (!extracted) {
      Logger.log(`  ✗ JD fetch failed — skipping`);
      diag.jd_fetch_failed++;
      addJobToCache(job, { match_level: 'SKIP', total_score: 0 }, '', 'failed', job.url, 'JobSearch', '');
      processed++;
      Utilities.sleep(500);
      continue;
    }
 
    const isTrustedSource = job.descriptionFull === true || extracted.source === 'api_full';
    if (!isJdRelevantToJob(extracted.text, job.company, job.title, isTrustedSource)) {
      Logger.log(`  ✗ JD failed relevance check for "${job.company}" — skipping`);
      diag.jd_irrelevant++;
      addJobToCache(job, { match_level: 'SKIP', total_score: 0 }, '', 'irrelevant_jd', job.url, 'JobSearch', '');
      processed++;
      Utilities.sleep(500);
      continue;
    }
 
    const cvType = detectCvTypeForSearch(extracted.text, job.title);
    Logger.log(`  CV type: ${cvType}`);

    // ── Pre-SMM time budget check ─────────────────────────────────────────
    const preSmm = Date.now() - runStart;
    if (preSmm > TIME_BUDGET_MS - 90000) {
      Logger.log(`⏱ Pre-SMM: only ${Math.round((TIME_BUDGET_MS - preSmm)/1000)}s left — stopping run`);
      break;
    }
    // ─────────────────────────────────────────────────────────────────────

    let smmResult;
    try {
      const raw = analyzeSkillsMatch(extracted.text, cvType, runStart, TIME_BUDGET_MS);
      smmResult = JSON.parse(raw);
      if (smmResult.error) throw new Error(smmResult.error);
    } catch (e) {
      Logger.log(`  ✗ SMM error: ${e.message}`);
      diag.smm_failed++;
      processed++;
      Utilities.sleep(3000);
      continue;
    }
 
    const score    = smmResult.total_score || 0;
    const level    = smmResult.match_level  || 'M0';
    const levelNum = parseInt(level.replace(/\D/g, '')) || 0;
 
    if      (levelNum === 0) diag.scored_m0++;
    else if (levelNum === 1) diag.scored_m1++;
    else                     diag.scored_m2_plus++;
 
    // ── Location enrichment ───────────────────────────────────────────────
    enrichSmmLocation(smmResult, job.company);

    Logger.log(`  Score: ${score}/40 | ${level} | Source: ${extracted.source} | Location: ${JSON.stringify(smmResult.job_location)} (${smmResult.location_source || 'unknown'})`);

    addJobToCache(job, smmResult, cvType, extracted.source, job.url, 'JobSearch', extracted.text);
 
    if (levelNum >= 2) {
      reportJobs.push({ ...job, smmResult, cvType, fetchSource: extracted.source });
      Logger.log(`  ✓ Added to report (${level})`);
    }
 
    processed++;
    Utilities.sleep(1500);
  }
 
  diag.processed_count   = processed;
  diag.report_jobs_count = reportJobs.length;
  diag.tavily_end        = getTavilyMonthlyUsage();
 
  emitDiagnosticsSummary(diag);
 
  if (reportJobs.length > 0) {
    Logger.log(`${reportJobs.length} M2+ job(s) stored in Job_Search_Cache with "in process" status.`);
    reportJobs.forEach(j => Logger.log(`  → ${j.company} — ${j.title} (${j.smmResult.match_level})`));
  } else {
    Logger.log('No M2+ jobs today. All scored jobs are in Job_Search_Cache for review.');
  }
 
  Logger.log('=== JABA Daily Job Search complete ===');
}
/* ============================================================
   CACHE UTILITIES
   ============================================================ */

/**
 * Wipes the entire Job_Search_Cache sheet.
 * Run once from the menu to clear cache saturation.
 * After running, the next daily search will evaluate fresh candidates.
 */
function clearAllJobCache() {
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();
  const sheet = getOrCreateJobCacheSheet();
  if (sheet.getLastRow() < 2) {
    if (ui) ui.alert('Job cache is already empty.');
    return;
  }
  const count = sheet.getLastRow() - 1;
  sheet.deleteRows(2, count);
  SpreadsheetApp.flush();
  Logger.log(`Job cache cleared: ${count} entries removed.`);
  if (ui) ui.alert(`✅ Job cache cleared — ${count} entries removed.\nThe next daily search run will evaluate fresh candidates.`);
}


/* ============================================================
   BA (BUNDESAGENTUR FÜR ARBEIT) JOB ALERT PROCESSOR
   Processes email alerts from jobsuche@arbeitsagentur.de
   Separate label tree: JABA BA Alert/
   Separate timestamp cursor: LAST_BA_ALERT_SCAN
   ============================================================ */

/**
 * Extracts job listings from a BA alert email HTML body.
 * Matches any arbeitsagentur.de href links that look like job titles.
 */
function extractBAJobListingsFromHtml(htmlBody) {
  if (!htmlBody) return [];
  const jobs = [];
  const seen = new Set();

  const html = htmlBody
    .replace(/&amp;/g, '&').replace(/&#38;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&#\d+;/g, ' ').replace(/&nbsp;/g, ' ')
    .replace(/&shy;/g, '');  // BA emails use soft hyphens in titles

  // BA alert email structure:
  //   <a href="...jobdetail/ID...">
  //     <span class="preventVisitedLinkStyle">N.</span>
  //     <span class="hover-underline">JOB TITLE HERE</span>
  //   </a>
  // The title link and the "Stelle ansehen" button share the same jobdetail URL.
  // We target the hover-underline span — this is always the actual job title.

  const jobPattern = /href="(https:\/\/www\.arbeitsagentur\.de\/jobsuche\/jobdetail\/([^"]+))"[^>]*>[\s\S]{0,800}?<span[^>]*class="hover-underline"[^>]*>([\s\S]*?)<\/span>/gi;

  let match;
  while ((match = jobPattern.exec(html)) !== null) {
    const url      = match[1];
    const jobId    = match[2];
    const rawTitle = match[3]
      .replace(/<[^>]+>/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    if (!rawTitle || rawTitle.length < 4 || rawTitle.length > 120) continue;
    if (/^(stelle ansehen|jetzt bewerben|zur stellenanzeige|abbestellen|verwalten|stellensuche)/i
        .test(rawTitle)) continue;
    if (seen.has(jobId)) continue;
    seen.add(jobId);

    // Company: extract from HTML in the ~1200 chars after this title block
    const afterTitle = html.substring(
      match.index + match[0].length,
      match.index + match[0].length + 1200
    );
    const company = extractCompanyFromHtmlContext(afterTitle) || 'Unknown';

    jobs.push({
      title:   rawTitle,
      company: company,
      url:     url,
      id:      'ba_' + jobId.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 40)
    });
  }

  Logger.log('  BA HTML extraction: ' + jobs.length + ' job(s) found');
  return jobs;
}
/* ============================================================
BA (BUNDESAGENTUR) ALERT PHASE 1
   Same pattern as Indeed Phase 1 but for BA job emails.
   Replaces processArbeitsagenturAlertEmails (kept as dead code).
   ============================================================ */
function processArbeitsagenturAlertEmails_Phase1() {
  const props    = PropertiesService.getScriptProperties();
  const timezone = CONFIG.TIMEZONE;
  const runStart = Date.now();
  const TIME_BUDGET_MS = 280000;
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();

  cleanAlertResults();

  const lastScanStr    = props.getProperty('LAST_BA_ALERT_SCAN');
  const sinceDate      = lastScanStr
    ? new Date(lastScanStr)
    : new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000);
  const queryDate      = new Date(sinceDate.getTime() - 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(queryDate, timezone, 'yyyy/MM/dd');
  Logger.log(`BA Phase 1 — since: ${sinceFormatted}`);

  const labelInternship = getOrCreateLabel('JABA BA Alert/🎓 Internship');
  const labelQueued     = getOrCreateLabel('JABA BA Alert/⏳ Queued');

  const query = [
    'from:jobsuche@arbeitsagentur.de',
    `after:${sinceFormatted}`,
    '-label:JABA-BA-Alert/Processed'
  ].join(' ');

  const threads = GmailApp.search(query, 0, 30);
  Logger.log(`Found ${threads.length} BA alert thread(s)`);

  if (threads.length === 0) {
    props.setProperty('LAST_BA_ALERT_SCAN', new Date().toISOString());
    if (ui) ui.alert('No new BA alert emails found since last scan.');
    return;
  }

  const pendingSheet   = getOrCreatePendingSmmSheet();
  const dateStr        = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');
  const MAX_JD_FETCHES = 12;
  let jdFetches    = 0;
  let totalQueued  = 0;
  let totalSkipped = 0;
  let totalCached  = 0;
  let totalApplied = 0;

  // Load applied jobs once — same pattern as Indeed Phase 1 and job search.
  const appliedSetBa = buildAppliedJobsSet();

  for (const thread of threads) {
    if (Date.now() - runStart > TIME_BUDGET_MS - 30000) {
      Logger.log('⏱ BA Phase 1 time budget — stopping');
      break;
    }

    const threadId = thread.getId();
    const message  = thread.getMessages()[thread.getMessages().length - 1];
    const htmlBody = message.getBody();
    const subject  = message.getSubject();
    Logger.log(`\nBA thread: "${subject}"`);

    const allJobs      = extractBAJobListingsFromHtml(htmlBody);
    const relevantJobs = allJobs.filter(j => isRelevantJobTitle(j.title));
    Logger.log(`  Jobs: ${allJobs.length} total | ${relevantJobs.length} relevant`);

    if (relevantJobs.length === 0) {
      const hasInternship = allJobs.some(j =>
        /werkstudent|praktikum|pflichtpraktikum|internship/i.test(j.title)
      );
      if (hasInternship) thread.addLabel(labelInternship);
      thread.addLabel(getOrCreateLabel('JABA BA Alert/Processed'));
      continue;
    }

    let threadQueued         = 0;
    let threadFullyProcessed = true;

    for (const job of relevantJobs) {
      if (isJobInCache(job.company, job.title)) {
        Logger.log(`    [cache] "${job.title}" @ ${job.company} — skip`);
        totalCached++;
        continue;
      }

      if (jdFetches >= MAX_JD_FETCHES) {
        Logger.log(`  MAX_JD_FETCHES (${MAX_JD_FETCHES}) reached — thread not fully processed`);
        threadFullyProcessed = false;
        break;
      }

      if (Date.now() - runStart > TIME_BUDGET_MS - 30000) {
        threadFullyProcessed = false;
        break;
      }

      Logger.log(`  → "${job.title}" at "${job.company}"`);
      jdFetches++;

      let jdText    = null;
      let fetchedUrl = job.url;

      // Tier 1: direct UrlFetchApp — BA pages are public
      const direct = fetchJobPageDirectly(job.url);
      if (direct && looksLikeJobContent(direct) && isCompleteJobDescription(direct) &&
          isJdRelevantToJob(direct, job.company, job.title, false)) {
        jdText = direct;
        Logger.log(`    Tier 1 (direct) OK (${direct.length} chars)`);
      }

      // Tier 2: Tavily advanced extract
      if (!jdText) {
        Logger.log(`    Tier 2: Tavily extract`);
        const tavily = tavilyExtractAdvanced(job.url);
        if (tavily && looksLikeJobContent(tavily) && isCompleteJobDescription(tavily) &&
            isJdRelevantToJob(tavily, job.company, job.title, false)) {
          jdText = tavily;
          Logger.log(`    Tier 2 (Tavily) OK (${tavily.length} chars)`);
        }
      }

      if (!jdText) {
        Logger.log(`    ✗ No valid JD — writing skipped to Alert_Results`);
        const skipDate = Utilities.formatDate(new Date(), timezone, 'dd.MM.yyyy HH:mm');
        writeToAlertResults(skipDate, 'BA', job.company, job.title,
                            job.url, '', 0, 'Skip', 'No valid JD found', '', '', threadId);
        addJobToCache(job, { match_level: 'SKIP', total_score: 0 }, '', 'no_jd', job.url, 'BA', '');
        totalSkipped++;
        Utilities.sleep(500);
        continue;
      }

      // ── Already applied check (post-JD-fetch, consistent with job search) ──
      const baApplyKey = job.company.toLowerCase().trim() + '||' +
                         normalizeJobTitle(job.title).toLowerCase().trim();
      if (appliedSetBa.has(baApplyKey)) {
        Logger.log(`    [already_applied] "${job.title}" @ ${job.company} — skipping`);
        const skipDate = Utilities.formatDate(new Date(), timezone, 'dd.MM.yyyy HH:mm');
        writeToAlertResults(skipDate, 'BA', job.company, job.title,
                            job.url, fetchedUrl, 0, 'already_applied',
                            'Already applied', '', '', threadId);
        addJobToCache(job, { match_level: 'already_applied', total_score: 0 },
                      '', 'already_applied', fetchedUrl, 'BA', '');
        totalApplied++;
        Utilities.sleep(300);
        continue;
      }
      // ─────────────────────────────────────────────────────────────────────

      pendingSheet.appendRow([
        dateStr, job.company, job.title, job.url,
        '', '',
        jdText.substring(0, 10000),
        'TRUE',
        scoreJobTitleQuality(job.title),
        job.id || '',
        '', '', '', '',
        'BA',
        threadId,
        'TRUE',
        fetchedUrl                        // Source_URL — actual JD source (may differ from job.url)
      ]);
      addJobToCache(job, { match_level: 'QUEUED', total_score: 0 }, '', 'queued', fetchedUrl, 'BA', '');
      totalQueued++;
      threadQueued++;
      Logger.log(`    ✓ Queued for SMM`);

      Utilities.sleep(1000);
    }

    if (threadFullyProcessed) {
      thread.addLabel(getOrCreateLabel('JABA BA Alert/Processed'));
      if (threadQueued > 0) thread.addLabel(labelQueued);
    } else {
      Logger.log(`  Thread not fully processed — will retry on next trigger run`);
    }

    Utilities.sleep(300);
  }

  props.setProperty('LAST_BA_ALERT_SCAN', new Date().toISOString());
  SpreadsheetApp.flush();
  Logger.log(`BA Phase 1: ${totalQueued} queued | ${totalSkipped} skipped | ${totalCached} cached | ${totalApplied} already_applied`);

  if (totalQueued > 0) {
    deletePhase2Triggers();
    ScriptApp.newTrigger('runPhase2').timeBased().after(PHASE2_DELAY_MS).create();
    Logger.log(`Phase 2 trigger created — runs in ${PHASE2_DELAY_MS / 60000} min`);
  }

  const summary = [
    '📧 BA Alert Phase 1 Complete',
    '─────────────────────────────',
    `✅ Queued for SMM: ${totalQueued}`,
    `⚠  Skipped (no valid JD): ${totalSkipped}`,
    `⏭  Already cached: ${totalCached}`,
    `✅ Already applied (skipped): ${totalApplied}`,
    totalQueued > 0 ? `\nSMM analysis runs automatically in ${PHASE2_DELAY_MS / 60000} minutes.` : ''
  ].filter(Boolean).join('\n');

  if (ui) ui.alert(summary);
}

/**
 * DEPRECATED: Main BA alert processor.
 * Replaced by processArbeitsagenturAlertEmails_Phase1 + runPhase2.
 * Kept as dead code reference only — do not run or trigger.
 */
function _DEPRECATED_processArbeitsagenturAlertEmails() {
  const props    = PropertiesService.getScriptProperties();
  const timezone = CONFIG.TIMEZONE;
  const runStart = Date.now();
  const TIME_BUDGET_MS = 280000; // 4.67 minutes
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();

  // ── Timestamp cursor ──────────────────────────────────────────────────────
  const lastScanStr = props.getProperty('LAST_BA_ALERT_SCAN');
  const sinceDate   = lastScanStr
    ? new Date(lastScanStr)
    : new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000);
  // Gmail's after: is exclusive — subtract 1 day so emails on the last-scan date are not missed.
  // The -label:Processed filter prevents re-processing already-handled threads.
  const queryDate      = new Date(sinceDate.getTime() - 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(queryDate, timezone, 'yyyy/MM/dd');
  Logger.log(`BA alert scan — searching since: ${sinceFormatted}`);

  const labelInternship = getOrCreateLabel('JABA BA Alert/🎓 Internship');

  const query = [
    'from:jobsuche@arbeitsagentur.de',
    `after:${sinceFormatted}`,
    '-label:JABA-BA-Alert/Processed'
  ].join(' ');

  const threads = GmailApp.search(query, 0, 30);
  Logger.log(`Found ${threads.length} BA alert thread(s) to process.`);

  if (threads.length === 0) {
    props.setProperty('LAST_BA_ALERT_SCAN', new Date().toISOString());
    if (ui) ui.alert('No new BA alert emails found since last scan.');
    return;
  }

  let totalReviewed = 0;
  let totalLow      = 0;
  let totalSkipped  = 0;
  const summaryLines  = [];
  const MAX_JOBS_PER_RUN = 10;
  let jobsProcessed   = 0;

  for (const thread of threads) {
    if (jobsProcessed >= MAX_JOBS_PER_RUN) {
      Logger.log(`MAX_JOBS_PER_RUN (${MAX_JOBS_PER_RUN}) reached — run again for remaining emails.`);
      break;
    }

    const message  = thread.getMessages()[thread.getMessages().length - 1];
    const htmlBody = message.getBody();
    const subject  = message.getSubject();
    Logger.log(`\nBA Email: "${subject}"`);

    const allJobs      = extractBAJobListingsFromHtml(htmlBody);
    const relevantJobs = allJobs.filter(j => isRelevantJobTitle(j.title));
    Logger.log(`BA jobs: ${allJobs.length} total | ${relevantJobs.length} passed keyword filter`);

    if (relevantJobs.length === 0) {
      const hasInternship = allJobs.some(j =>
        /werkstudent|praktikum|pflichtpraktikum|internship/i.test(j.title));
      if (hasInternship) thread.addLabel(labelInternship);
      thread.addLabel(getOrCreateLabel('JABA BA Alert/Processed'));
      continue;
    }

    const threadResults = [];

    for (const job of relevantJobs) {
      if (jobsProcessed >= MAX_JOBS_PER_RUN) break;

      // ── Time budget check ───────────────────────────────────────────────
      if (Date.now() - runStart > TIME_BUDGET_MS - 60000) {
        Logger.log(`⏱ BA scan: time budget nearly exhausted — stopping gracefully`);
        break;
      }

      Logger.log(`→ "${job.title}" at "${job.company}"`);

      // ── JD fetch: direct → Tavily ───────────────────────────────────────
      let jdText   = null;
      let jdSource = 'unknown';

      // Tier 1: direct UrlFetchApp (BA pages are public)
      const direct = fetchJobPageDirectly(job.url);
      if (direct && looksLikeJobContent(direct) && isCompleteJobDescription(direct)) {
        jdText   = direct;
        jdSource = 'direct';
        Logger.log(`  Tier 1 (direct) OK: ${direct.length} chars`);
      }

      // Tier 2: Tavily advanced
      if (!jdText) {
        Logger.log(`  Tier 2: Tavily extract for BA job`);
        const tavily = tavilyExtractAdvanced(job.url);
        if (tavily && looksLikeJobContent(tavily) && isCompleteJobDescription(tavily)) {
          jdText   = tavily;
          jdSource = 'tavily';
          Logger.log(`  Tier 2 (Tavily) OK: ${tavily.length} chars`);
        }
      }

      if (!jdText) {
        Logger.log(`  ⚠ No valid JD found — skipping`);
        totalSkipped++;
        threadResults.push({ title: job.title, company: job.company, skipped: true, reason: 'no_jd' });
        jobsProcessed++;
        continue;
      }

      // ── CV type detection ───────────────────────────────────────────────
      const cvType = detectCvTypeFromText(job.title + ' ' + jdText.substring(0, 500));

      // ── SMM analysis ────────────────────────────────────────────────────
      let smmResult;
      try {
        const smmRaw = analyzeSkillsMatch(jdText, cvType, runStart, TIME_BUDGET_MS);
        smmResult    = JSON.parse(smmRaw);
        if (smmResult.error) throw new Error(smmResult.error);
      } catch(e) {
        Logger.log(`  ✗ SMM error: ${e.message}`);
        totalSkipped++;
        threadResults.push({ title: job.title, company: job.company, skipped: true });
        jobsProcessed++;
        Utilities.sleep(3000);
        continue;
      }

      const score      = smmResult.total_score || 0;
      const matchLevel = smmResult.match_level  || 'M0';
      const skills     = smmResult.skills        || [];
      const levelNum   = parseInt(matchLevel.replace(/\D/g, '')) || 0;

      const crucialSkills  = skills.filter(s => s.importance === 'Crucial');
      const zeroCrucial    = crucialSkills.length === 0;
      const allCrucialPass = !zeroCrucial && crucialSkills.every(s => (s.score || 0) >= 1);
      const qualifies      = levelNum >= 1 && allCrucialPass;

      Logger.log(`  Score: ${score}/40 | ${matchLevel} | Crucial: ${crucialSkills.length} | Pass: ${allCrucialPass} | Qualifies: ${qualifies}`);

      if (zeroCrucial) {
        totalLow++;
        threadResults.push({ title: job.title, company: job.company, smmResult, score, matchLevel, qualifies: false, reason: 'no-crucial' });
        summaryLines.push(`⚠ MANUAL REVIEW: ${job.company} — ${job.title} — ${score}/40`);
      } else if (qualifies) {
        totalReviewed++;
        threadResults.push({ title: job.title, company: job.company, smmResult, score, matchLevel, qualifies: true });
        summaryLines.push(`✅ ${job.company} — ${job.title} — ${score}/40 (${matchLevel})`);
      } else {
        totalLow++;
        threadResults.push({ title: job.title, company: job.company, smmResult, score, matchLevel, qualifies: false });
        summaryLines.push(`⬇ ${job.company} — ${job.title} — ${score}/40 (${matchLevel})`);
      }

      jobsProcessed++;
      Utilities.sleep(2500);
    }

    // ── Apply Gmail labels ────────────────────────────────────────────────
    const analyzed  = threadResults.filter(r => r.smmResult);
    const noJd      = threadResults.filter(r => r.skipped && r.reason === 'no_jd');
    const otherSkip = threadResults.filter(r => r.skipped && r.reason !== 'no_jd');

    if (analyzed.length > 0) {
      const best       = analyzed.reduce((a, b) =>
        (a.score || 0) >= (b.score || 0) ? a : b);
      // Reuse buildAlertLabel but replace the JABA Alert prefix with JABA BA Alert
      const scoreLabel = buildAlertLabel(best.smmResult)
        .replace('JABA Alert/', 'JABA BA Alert/');
      thread.addLabel(getOrCreateLabel(scoreLabel));
    } else if (noJd.length > 0 && analyzed.length === 0) {
      thread.addLabel(getOrCreateLabel('JABA BA Alert/⏭ No JD Found'));
    } else if (otherSkip.length > 0 && analyzed.length === 0) {
      thread.addLabel(getOrCreateLabel('JABA BA Alert/⏭ Skipped'));
    }

    thread.addLabel(getOrCreateLabel('JABA BA Alert/Processed'));
    Utilities.sleep(500);
  }

  props.setProperty('LAST_BA_ALERT_SCAN', new Date().toISOString());

  const remaining = Math.max(0, threads.length - jobsProcessed);
  const summary = [
    '📧 BA Alert Scan Complete',
    '─────────────────────────────',
    `✅ Qualifying jobs: ${totalReviewed}`,
    `⬇  Below threshold: ${totalLow}`,
    `⚠  Skipped (fetch failed): ${totalSkipped}`,
    '',
    ...summaryLines,
    '',
    `Jobs processed this run: ${jobsProcessed}/${MAX_JOBS_PER_RUN}`,
    remaining > 0 ? `${remaining} email(s) still pending — run again to continue.` : ''
  ].filter(l => l !== undefined).join('\n');

  Logger.log(summary);
  if (ui) ui.alert(summary);
}

/* ============================================================
   ALERT_RESULTS & M2_NOTIFICATIONS — Schema helpers
   ============================================================ */

function getOrCreateAlertResultsSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet   = ss.getSheetByName('Alert_Results');
  const HEADERS = [
    'Date_Added', 'Source', 'Company', 'Title',
    'Source_URL', 'Fetched_JD_URL', 'Score', 'Level',
    'Status', 'Dots', 'CV_Type', 'Read', 'Thread_ID', 'Label_Applied'
  ];
  if (!sheet) {
    sheet = ss.insertSheet('Alert_Results');
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setValues([HEADERS])
         .setBackground('#ea4335').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
    for (let i = 1; i <= HEADERS.length; i++) sheet.autoResizeColumn(i);
    Logger.log('Alert_Results sheet created.');
  }
  return sheet;
}

function getOrCreateM2NotificationsSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('M2_Notifications');
  if (!sheet) {
    sheet = ss.insertSheet('M2_Notifications');
    sheet.getRange('A1').setValue('');
    sheet.hideSheet();
    Logger.log('M2_Notifications sheet created (hidden).');
  }
  return sheet;
}

/* ── Insert newest row at top, below frozen header ── */
function writeToAlertResults(dateStr, source, company, title,
                              sourceUrl, fetchedUrl, score, level,
                              status, dots, cvType, threadId) {
  const sheet = getOrCreateAlertResultsSheet();
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, 14).setValues([[
    dateStr, source, company, title,
    sourceUrl  || '', fetchedUrl || '',
    score, level, status, dots,
    cvType || '', false,         // Read = false
    threadId || '', false        // Label_Applied = false
  ]]);
}

/* ── Delete rows older than 7 days — called at Phase 1 start ── */
/**
 * Parses an Alert_Results date string stored as "dd.MM.yyyy HH:mm".
 * Returns a Date object, or null if the string cannot be parsed.
 * Uses explicit DD/MM/YYYY parsing because new Date("05.06.2026 14:30")
 * produces Invalid Date in V8/GAS — the dot-separated format is not ISO.
 */
function parseAlertResultsDate(str) {
  if (!str) return null;
  const m = str.toString().match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (!m) return null;
  // months are 0-indexed in JS Date
  return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
}

function cleanAlertResults() {
  const sheet = getOrCreateAlertResultsSheet();
  if (sheet.getLastRow() < 2) return;
  const cutoff   = new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000);
  const dates    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const toDelete = [];
  dates.forEach(function(row, i) {
    const entryDate = parseAlertResultsDate(row[0]);
    if (entryDate && entryDate < cutoff) toDelete.push(i + 2);
  });
  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);
  if (toDelete.length > 0) {
    SpreadsheetApp.flush();
    Logger.log('Alert_Results: removed ' + toDelete.length + ' expired row(s).');
  }
}

/* ── Write to hidden M2_Notifications cell (triggers mobile push) ── */
function updateM2NotificationCell(count, jobLines) {
  const sheet   = getOrCreateM2NotificationsSheet();
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy HH:mm');
  sheet.getRange('A1').setValue(
    `🚨 ${count} new M2+ job${count > 1 ? 's' : ''} · ${dateStr}\n${jobLines.join('\n')}`
  );
  Logger.log(`M2_Notifications updated: ${count} M2+ alert job(s).`);
}

/* ── Apply M-level summary label to each alert thread ── */
function applyAlertThreadLabels() {
  const sheet = getOrCreateAlertResultsSheet();
  if (sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();

  const threadMap = {};
  data.forEach((row, i) => {
    const threadId     = (row[12] || '').toString().trim();
    const labelApplied =  row[13];
    const level        = (row[7]  || '').toString().trim();
    const source       = (row[1]  || 'Indeed').toString().trim();
    if (!threadId || labelApplied) return;
    const levelNum = parseInt(level.replace(/\D/g, '')) || 0;
    if (!threadMap[threadId]) {
      threadMap[threadId] = { source, counts: {}, rowIndices: [] };
    }
    if (levelNum >= 1) {
      const key = 'M' + Math.min(levelNum, 4);
      threadMap[threadId].counts[key] = (threadMap[threadId].counts[key] || 0) + 1;
    }
    threadMap[threadId].rowIndices.push(i + 2);
  });

  let labeled = 0;
  Object.entries(threadMap).forEach(([threadId, info]) => {
    try {
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) { Logger.log(`Thread not found: ${threadId}`); return; }

      const parts = ['M1','M2','M3','M4']
        .filter(k => (info.counts[k] || 0) > 0)
        .map(k => `${k}=${info.counts[k]}`);

      if (parts.length > 0) {
        const prefix    = info.source === 'BA' ? 'JABA BA Alert' : 'JABA Alert';
        const labelName = `${prefix}/${parts.join(' ')}`;
        thread.addLabel(getOrCreateLabel(labelName));
        Logger.log(`Thread ${threadId}: label applied — "${labelName}"`);
        labeled++;
      }
      info.rowIndices.forEach(rowNum => sheet.getRange(rowNum, 14).setValue(true));
    } catch(e) {
      Logger.log(`applyAlertThreadLabels error (${threadId}): ${e.message}`);
    }
  });
  if (labeled > 0) SpreadsheetApp.flush();
  Logger.log(`Alert thread labels applied: ${labeled} thread(s).`);
}

/* ── Desktop dialog on sheet open — shows unread M2+ alert jobs ── */
function checkUnreadM2Alerts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Alert_Results');
    if (!sheet || sheet.getLastRow() < 2) return;

    const data   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    const unread = [];
    data.forEach((row, i) => {
      const levelNum = parseInt((row[7] || '').toString().replace(/\D/g, '')) || 0;
      if (levelNum >= 2 && !row[11]) {
        unread.push({ rowIndex: i + 2, company: row[2], title: row[3],
                      level: row[7], score: row[6], source: row[1] });
      }
    });
    if (unread.length === 0) return;

    const lines = unread.map(j =>
      `${j.source === 'BA' ? '🏛' : '🔔'} ${j.company} — ${j.title} (${j.level}: ${j.score}/40)`
    ).join('\n');

    SpreadsheetApp.getUi().alert(
      `🚨 ${unread.length} new M2+ Alert Job${unread.length > 1 ? 's' : ''}`,
      `${lines}\n\nOpen the Alert_Results tab for links and details.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    unread.forEach(j => sheet.getRange(j.rowIndex, 12).setValue(true));
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log(`checkUnreadM2Alerts error: ${e.message}`);
  }
}

/* ── Called from runPhase2 when all pending rows processed ── */
function finalizeAlertResults(m2PlusRows) {
  applyAlertThreadLabels();
  if (!m2PlusRows || m2PlusRows.length === 0) {
    Logger.log('finalizeAlertResults: no M2+ alert rows.');
    return;
  }
  const jobLines = m2PlusRows.map(row =>
    `• ${(row[14] || 'Indeed')} | ${row[1]} — ${row[2]}`
  );
  updateM2NotificationCell(m2PlusRows.length, jobLines);
  Logger.log(`finalizeAlertResults: ${m2PlusRows.length} M2+ alert job(s) notified.`);
}

/* ============================================================
   PHASE 1 / PHASE 2 — Batched daily job search
   Phase 1: fetch all sources, filter, write top 24 to Pending_SMM,
            create a one-time trigger for Phase 2.
   Phase 2: pick 8 rows from Pending_SMM, fetch JD, run SMM,
            delete processed rows, self-schedule until queue is empty,
            then send the report.
   Only one manual trigger needed: runJobSearchPhase1 daily at 07:00.
   ============================================================ */

const PHASE1_CANDIDATE_CAP = 24;  // top N candidates saved per day
const PHASE2_BATCH_SIZE    = 8;   // SMM calls per Phase 2 run (~5 min)
const PHASE2_DELAY_MS      = 15 * 60 * 1000; // 15 minutes between batches

/* ── Pending_SMM sheet ─────────────────────────────────────── */
function getOrCreatePendingSmmSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet   = ss.getSheetByName('Pending_SMM');
  const HEADERS = [
    'Date_Added', 'Company', 'Title', 'URL', 'City', 'Country',
    'Description', 'DescriptionFull', 'QualityScore', 'ID',
    'Status', 'CvType', 'FetchSource', 'SmmData',
    'Source', 'Thread_ID', 'JD_Fetched', 'Source_URL'
  ];
  if (!sheet) {
    sheet = ss.insertSheet('Pending_SMM');
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setValues([HEADERS])
         .setBackground('#34a853').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
    for (let i = 1; i <= HEADERS.length; i++) sheet.autoResizeColumn(i);
    Logger.log('Pending_SMM sheet created (18 columns).');
  } else if (sheet.getLastColumn() < 18) {
    // One-time migration: add the 4 new columns to existing sheet
    sheet.getRange(1, 15, 1, 4)
         .setValues([['Source', 'Thread_ID', 'JD_Fetched', 'Source_URL']])
         .setBackground('#34a853').setFontColor('white').setFontWeight('bold');
    Logger.log('Pending_SMM: migrated to 18 columns.');
  }
  return sheet;
}

/* ── Trigger helpers ───────────────────────────────────────── */
function deletePhase2Triggers() {
  ScriptApp.getProjectTriggers()
    .filter(t =>
      t.getHandlerFunction() === 'runJobSearchPhase2' ||
      t.getHandlerFunction() === 'runPhase2'
    )
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/* ── PHASE 1 ───────────────────────────────────────────────── */
function runJobSearchPhase1() {

  try {
    _runJobSearchPhase1Impl();
  } catch (e) {
    if (e.message && (e.message.includes('INTERNAL') || e.message.includes('server error occurred'))) {
      Logger.log('⚠️ GAS INTERNAL storage error — scheduling retry in 10 min: ' + e.message);
      ScriptApp.newTrigger('runJobSearchPhase1')
        .timeBased()
        .after(10 * 60 * 1000)
        .create();
      return;
    }
    throw e;
  }
}


function _runJobSearchPhase1Impl() {

  Logger.log('=== JABA Phase 1: Fetch & Filter ===');

  cleanJobCache();

  const allJobs = fetchAllJobSources();
  Logger.log('Phase 1: ' + allJobs.length + ' total jobs fetched');
  if (allJobs.length === 0) { Logger.log('Phase 1: no jobs returned — exiting.'); return; }

  // ── Preload cache + applied data ONCE (fixes timeout) ────────────────────
  Logger.log('Phase 1: preloading cache and applied sets...');
  const cachedSet  = buildCachedJobsSet();
  const appliedSet = buildAppliedJobsSet();
  Logger.log('Phase 1: cache=' + cachedSet.size + ' applied=' + appliedSet.size);
  // ─────────────────────────────────────────────────────────────────────────

  let excludedTitle = 0, excludedGeo = 0, excludedCache = 0, excludedApplied = 0;
  const candidates = [];

  for (const job of allJobs) {
    if (!isRelevantJobTitle(job.title)) { excludedTitle++; continue; }

    if (!isJobRemoteEligible(job)) { excludedGeo++; continue; }

    const cacheKey = job.company.toLowerCase().trim() + '||' +
                     normalizeJobTitle(job.title).toLowerCase().trim();
 
    if (cachedSet.has(cacheKey))  { excludedCache++; continue; }
    if (appliedSet.has(cacheKey)) {
      excludedApplied++;
      addJobToCache(job, { match_level: 'already_applied', total_score: 0 },
                    '', 'already_applied', '', 'JobSearch', '');
      continue;
    }

    candidates.push(job);
  }

  // ── Deduplicate ───────────────────────────────────────────────────────────
  const seen    = new Set();
  const deduped = [];
  for (const job of candidates) {
    const key = job.company.toLowerCase().trim() + '||' +
                normalizeJobTitle(job.title).toLowerCase().trim();
    if (!seen.has(key)) { seen.add(key); deduped.push(job); }
  }

  deduped.sort((a, b) => scoreJobTitleQuality(b.title) - scoreJobTitleQuality(a.title));
  const selected = deduped.slice(0, PHASE1_CANDIDATE_CAP);

  Logger.log(
    'Phase 1 gates — title: ' + excludedTitle + ', geo: ' + excludedGeo +
    ', cache: ' + excludedCache + ', applied: ' + excludedApplied +
    ' | candidates: ' + deduped.length + ' | selected (cap ' + PHASE1_CANDIDATE_CAP + '): ' + selected.length
  );

  if (selected.length === 0) {
    Logger.log('Phase 1: no candidates after filtering — no Phase 2 needed.');
    return;
  }

  // ── Write to Pending_SMM ──────────────────────────────────────────────────
  const sheet   = getOrCreatePendingSmmSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const oldStatuses = sheet.getRange(2, 11, lastRow - 1, 1).getValues();
    const allOldRows  = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    const oldM2Plus   = allOldRows.filter((row, i) =>
      (oldStatuses[i][0] || '').toString().trim() === 'M2+'
    );
    if (oldM2Plus.length > 0) {
      Logger.log('Phase 1: found ' + oldM2Plus.length + ' unreported M2+ rows — routing by source');
      const oldJobSearch = oldM2Plus.filter(r => (r[14] || 'JobSearch').toString().trim() === 'JobSearch');
      const oldAlerts    = oldM2Plus.filter(r => (r[14] || '').toString().trim() !== 'JobSearch');
      if (oldJobSearch.length > 0) sendPhase2Report(oldJobSearch);
      if (oldAlerts.length > 0)    finalizeAlertResults(oldAlerts);
    } else {
      sheet.deleteRows(2, lastRow - 1);
    }
  }

  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const rows    = selected.map(function(job) {
    return [
      dateStr,
      job.company  || '',
      job.title    || '',
      job.url      || '',
      job.city     || '',
      job.country  || '',
      (job.description || '').substring(0, 10000),
      job.descriptionFull ? 'TRUE' : 'FALSE',
      scoreJobTitleQuality(job.title),
      job.id       || '',
      '', '', '', '',           // Status, CvType, FetchSource, SmmData
      'JobSearch', '', 'FALSE', job.url || ''  // Source, Thread_ID, JD_Fetched, Source_URL
    ];
  });

  sheet.getRange(2, 1, rows.length, 18).setValues(rows);
  SpreadsheetApp.flush();
  Logger.log('Phase 1 complete: ' + rows.length + ' candidates written to Pending_SMM');

  deletePhase2Triggers();
  ScriptApp.newTrigger('runPhase2')
    .timeBased()
    .after(PHASE2_DELAY_MS)
    .create();
  Logger.log('Phase 1: Phase 2 trigger created — runs in ' + (PHASE2_DELAY_MS / 60000) + ' min');
}

/* ============================================================
   UNIVERSAL PHASE 2 — handles JobSearch, Indeed, BA
   Replaces runJobSearchPhase2 (keep old function until Deploy 3)
   ============================================================ */
function runPhase2() {
  const TIME_BUDGET_MS = 280000;
  const runStart       = Date.now();

  Logger.log('=== JABA Universal Phase 2: SMM Batch ===');
  deletePhase2Triggers();

  const sheet = getOrCreatePendingSmmSheet();
  if (sheet.getLastRow() < 2) {
    Logger.log('Phase 2: Pending_SMM empty — nothing to do.');
    return;
  }

  const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 18).getValues();

  const pendingEntries  = [];
  const m2PlusJobSearch = [];
  const m2PlusAlerts    = [];

  allData.forEach((row, i) => {
    const status = (row[10] || '').toString().trim();
    const source = (row[14] || 'JobSearch').toString().trim();
    if (status === 'M2+') {
      if (source === 'JobSearch') m2PlusJobSearch.push(row);
      else                        m2PlusAlerts.push(row);
    } else {
      pendingEntries.push({ sheetRow: i + 2, data: row });
    }
  });

  Logger.log(`Phase 2: ${pendingEntries.length} pending | ${m2PlusJobSearch.length} JS-M2+ | ${m2PlusAlerts.length} Alert-M2+ accumulated`);
  PropertiesService.getScriptProperties().setProperty('PHASE2_STATUS', JSON.stringify({
    running: true, pending: pendingEntries.length, startedAt: Date.now()
  }));

  if (pendingEntries.length === 0) {
    Logger.log('Phase 2: no pending rows — finalizing.');
    sendPhase2Report(m2PlusJobSearch);
    finalizeAlertResults(m2PlusAlerts);
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().deleteProperty('PHASE2_STATUS');
    return;
  }

  const batch        = pendingEntries.slice(0, PHASE2_BATCH_SIZE);
  const rowsToDelete = [];

  for (const { sheetRow, data } of batch) {
    if (Date.now() - runStart > TIME_BUDGET_MS - 90000) {
      Logger.log('⏱ Phase 2 time budget — stopping early');
      break;
    }

    const company         = data[1];
    const title           = data[2];
    const url             = data[3];
    const city            = data[4];
    const country         = data[5];
    const description     = data[6];
    const descriptionFull = data[7] === 'TRUE' || data[7] === true;
    const jobId           = data[9];
    const source          = (data[14] || 'JobSearch').toString().trim();
    const threadId        = (data[15] || '').toString().trim();
    const jdFetched       = data[16] === 'TRUE' || data[16] === true;
    const sourceUrl       = (data[17] || url).toString().trim();
    const job             = { title, company, city, country, url, description, descriptionFull, id: jobId };
    const dateStr         = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy HH:mm');

    Logger.log(`\n[${source}] "${title}" — ${company}`);

    let extractedText = null;
    let fetchSource   = 'phase1_prefetch';
    let fetchedUrl    = url;

    if (jdFetched && description && description.length > 300 && looksLikeJobContent(description)) {
      // Alert jobs: JD already validated in Phase 1
      // Use sourceUrl (col 18) which now stores the actual JD source URL from Phase 1
      extractedText = description;
      fetchedUrl    = sourceUrl;
      Logger.log(`  ✓ Pre-fetched JD (${description.length} chars) from ${sourceUrl}`);
    } else {
      // Job Search jobs: fetch JD now (existing logic unchanged)
      const extracted = smartExtractJD(url, description, descriptionFull);
      if (!extracted) {
        Logger.log('  ✗ JD fetch failed — skipping');
        addJobToCache(job, { match_level: 'SKIP', total_score: 0 }, '', 'failed', url, source, '');
        if (source !== 'JobSearch') {
          writeToAlertResults(dateStr, source, company, title, sourceUrl, url, 0, 'Skip', 'No JD found', '', '', threadId);
        }
        rowsToDelete.push(sheetRow);
        Utilities.sleep(500);
        continue;
      }
      const isTrusted = descriptionFull || extracted.source === 'api_full';
      if (!isJdRelevantToJob(extracted.text, company, title, isTrusted)) {
        Logger.log('  ✗ JD irrelevant — skipping');
        addJobToCache(job, { match_level: 'SKIP', total_score: 0 }, '', 'irrelevant_jd', url, source, '');
        if (source !== 'JobSearch') {
          writeToAlertResults(dateStr, source, company, title, sourceUrl, url, 0, 'Skip', 'JD irrelevant', '', '', threadId);
        }
        rowsToDelete.push(sheetRow);
        Utilities.sleep(500);
        continue;
      }
      extractedText = extracted.text;
      fetchSource   = extracted.source;
    }

    const cvType = source === 'JobSearch'
      ? detectCvTypeForSearch(extractedText, title)
      : detectCvTypeFromText(title + ' ' + extractedText.substring(0, 500));

    let smmResult;
    try {
      const raw = analyzeSkillsMatch(extractedText, cvType, runStart, TIME_BUDGET_MS);
      smmResult = JSON.parse(raw);
      if (smmResult.error) throw new Error(smmResult.error);
    } catch(e) {
      Logger.log(`  ✗ SMM error: ${e.message}`);
      if (source !== 'JobSearch') {
        writeToAlertResults(dateStr, source, company, title, sourceUrl, fetchedUrl, 0, 'Err', 'SMM error', '', '', threadId);
      }
      rowsToDelete.push(sheetRow);
      Utilities.sleep(3000);
      continue;
    }

    const score    = smmResult.total_score || 0;
    const level    = smmResult.match_level  || 'M0';
    const levelNum = parseInt(level.replace(/\D/g, '')) || 0;
    const dots     = buildJobDots(smmResult);

    // ── Location enrichment — fills job_location when JD had none ────────
    enrichSmmLocation(smmResult, company);

    Logger.log(`  Score: ${score}/40 | ${level} | Source: ${source} | Location: ${JSON.stringify(smmResult.job_location)} (${smmResult.location_source || 'unknown'})`);
    addJobToCache(job, smmResult, cvType, fetchSource, fetchedUrl, source, extractedText || '');

    // All alert jobs (including M0) go to Alert_Results for full visibility
    if (source !== 'JobSearch') {
      writeToAlertResults(dateStr, source, company, title, sourceUrl, fetchedUrl, score, level, 'Analyzed', dots, cvType, threadId);
    }

    if (levelNum >= 2) {
      sheet.getRange(sheetRow, 11).setValue('M2+');
      sheet.getRange(sheetRow, 12).setValue(cvType);
      sheet.getRange(sheetRow, 13).setValue(fetchSource);
      sheet.getRange(sheetRow, 14).setValue(JSON.stringify(smmResult));
      Logger.log(`  ✓ M2+ — kept for final report`);
    } else {
      rowsToDelete.push(sheetRow);
    }

    Utilities.sleep(1500);
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    try { sheet.deleteRow(rowsToDelete[i]); } catch(e) {}
  }
  SpreadsheetApp.flush();

  let remainingPending = 0;
  if (sheet.getLastRow() > 1) {
    remainingPending = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1)
      .getValues()
      .filter(r => (r[0] || '').toString().trim() !== 'M2+')
      .length;
  }

  Logger.log(`Phase 2 batch done. Remaining pending: ${remainingPending}`);

  PropertiesService.getScriptProperties().deleteProperty('PHASE2_STATUS');
  if (remainingPending > 0) {
    ScriptApp.newTrigger('runPhase2').timeBased().after(PHASE2_DELAY_MS).create();
    Logger.log(`Phase 2 rescheduled — ${PHASE2_DELAY_MS / 60000} min`);
  } else {
    const finalLastRow = sheet.getLastRow();
    const finalData    = finalLastRow > 1
      ? sheet.getRange(2, 1, finalLastRow - 1, 18).getValues()
      : [];

    const finalJS = finalData.filter(r =>
      (r[10] || '').toString().trim() === 'M2+' &&
      (r[14] || 'JobSearch').toString().trim() === 'JobSearch'
    );
    const finalAlerts = finalData.filter(r =>
      (r[10] || '').toString().trim() === 'M2+' &&
      (r[14] || '').toString().trim() !== 'JobSearch'
    );

    sendPhase2Report(finalJS);
    finalizeAlertResults(finalAlerts);

    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    SpreadsheetApp.flush();
    Logger.log('Phase 2 complete — Pending_SMM cleared.');
  }
}


// runJobSearchPhase2 removed — superseded by runPhase2 (universal handler).
// deletePhase2Triggers() still cleans up this name as a safety net.

/* ── Send report and clear Pending_SMM ─────────────────────── */
function sendPhase2Report(m2PlusRows) {
  if (!m2PlusRows || m2PlusRows.length === 0) {
    Logger.log('sendPhase2Report: no M2+ rows — no email sent.');
    return;
  }
  const reportJobs = m2PlusRows.map(row => {
    let smmResult;
    try { smmResult = JSON.parse(row[13]); } catch(e) { return null; }
    if (!smmResult) return null;
    return {
      company: (row[1] || '').toString(),
      title:   (row[2] || '').toString(),
      url:     (row[3] || '').toString(),
      city:    (row[4] || '').toString(),
      cvType:      (row[11] || '').toString(),
      fetchSource: (row[12] || '').toString(),
      smmResult
    };
  }).filter(Boolean);
  if (reportJobs.length === 0) return;
  Logger.log(`sendPhase2Report: sending email for ${reportJobs.length} M2+ job(s)`);
  reportJobs.forEach(j => Logger.log(`  → ${j.company} — ${j.title} (${j.smmResult.match_level})`));
  const html = buildJobReportHtml(reportJobs);
  sendJobReportEmail(html);
}

/** Debug utility — resets the BA scan window to 7 days ago. */
function resetBAAlertScanTimestamp() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_BA_ALERT_SCAN');
  Logger.log('BA alert scan timestamp cleared. Next run will scan last 7 days.');
}

function debugBASearch() {
  const threads = GmailApp.search('from:jobsuche@arbeitsagentur.de', 0, 10);
  Logger.log(`Total threads found: ${threads.length}`);
  threads.forEach(t => {
    const msg = t.getMessages()[0];
    Logger.log(`Subject: ${msg.getSubject()} | Date: ${msg.getDate()} | Labels: ${t.getLabels().map(l => l.getName()).join(', ')}`);
  });
}

function diagnoseBaEmails() {
  // Broad: any sender from the arbeitsagentur.de domain
  const broad = GmailApp.search('from:arbeitsagentur.de', 0, 20);
  Logger.log('Broad search (from:arbeitsagentur.de): ' + broad.length + ' thread(s)');

  broad.forEach(function(t) {
    const msg = t.getMessages()[0];
    Logger.log(
      'FROM: "' + msg.getFrom() + '" | ' +
      'SUBJ: "' + msg.getSubject() + '" | ' +
      'DATE: ' + msg.getDate()
    );
  });

  // Narrow: the old hardcoded address
  const narrow = GmailApp.search('from:jobsuche@arbeitsagentur.de', 0, 10);
  Logger.log('Narrow search (jobsuche@): ' + narrow.length + ' thread(s)');

  // Current timestamp cursor
  const ts = PropertiesService.getScriptProperties().getProperty('LAST_BA_ALERT_SCAN');
  Logger.log('LAST_BA_ALERT_SCAN = "' + (ts || 'not set') + '"');
}

function checkBaEmailLabels() {
  const threads = GmailApp.search('from:jobsuche@arbeitsagentur.de', 0, 20);
  threads.forEach(function(t) {
    const msg    = t.getMessages()[0];
    const labels = t.getLabels().map(function(l) { return l.getName(); }).join(', ');
    Logger.log(
      'SUBJ: "' + msg.getSubject() + '" | ' +
      'DATE: ' + msg.getDate() + ' | ' +
      'LABELS: [' + (labels || 'none') + ']'
    );
  });
}
function diagnoseBaHtmlExtraction() {
  const threads = GmailApp.search('from:jobsuche@arbeitsagentur.de', 0, 2);
  if (threads.length === 0) { Logger.log('No BA threads found.'); return; }

  const msg      = threads[0].getMessages()[0];
  const htmlBody = msg.getBody();
  const subject  = msg.getSubject();
  Logger.log('Email: "' + subject + '"');
  Logger.log('HTML length: ' + htmlBody.length + ' chars');

  // Show what the current extraction actually finds
  const jobs = extractBAJobListingsFromHtml(htmlBody);
  Logger.log('Extracted ' + jobs.length + ' item(s):');
  jobs.forEach(function(j, i) {
    Logger.log('  [' + i + '] title: "' + j.title + '" | company: "' + j.company + '" | url: "' + j.url + '"');
  });

  // Also dump the first 3000 chars of raw HTML so we can see the structure
  Logger.log('--- RAW HTML PREVIEW (first 3000 chars) ---');
  Logger.log(htmlBody.substring(0, 3000));
}

function diagnoseBaHtmlStructure() {
  const threads = GmailApp.search('from:jobsuche@arbeitsagentur.de', 0, 1);
  if (threads.length === 0) { Logger.log('No threads found.'); return; }

  const msg  = threads[0].getMessages()[0];
  const html = msg.getBody()
    .replace(/&amp;/g, '&').replace(/&#38;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ');

  // Find the first jobdetail URL and show 1200 chars of HTML before it
  const pattern = /href="(https:\/\/www\.arbeitsagentur\.de\/jobsuche\/jobdetail\/[^"]+)"/i;
  const match   = pattern.exec(html);

  if (!match) {
    Logger.log('No jobdetail URL found in HTML.');
    return;
  }

  Logger.log('Job URL: ' + match[1]);
  Logger.log('--- HTML BEFORE LINK (1200 chars) ---');
  Logger.log(html.substring(Math.max(0, match.index - 1200), match.index));
  Logger.log('--- HTML AFTER LINK (300 chars) ---');
  Logger.log(html.substring(match.index, match.index + 300));
}

/**
 * Parses a cache timestamp string in either:
 *   new format: dd-MM-yyyy HH:mm:ss
 *   old format: yyyy-MM-dd
 * Returns a Date object or null.
 */
function parseCacheTimestamp(val) {
  if (!val) return null;
  const str = val.toString().trim();
  const m   = str.match(/^(\d{2})-(\d{2})-(\d{4})(?:\s+(\d{2}):(\d{2}):(\d{2}))?$/);
  if (m) {
    return new Date(
      parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]),
      parseInt(m[4] || '0'), parseInt(m[5] || '0'), parseInt(m[6] || '0')
    );
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return new Date(str); // legacy
  return null;
}
 
/**
 * Returns the count of Job_Search_Cache rows with Review_Status = 'fit'.
 * Called from the sidebar to keep the fit-jobs badge current.
 */
function getFitJobCount() {
  try {
    const sheet = getOrCreateJobCacheSheet();
    if (sheet.getLastRow() < 2) return 0;
    const data = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).getValues();
    return data.filter(r => (r[0] || '').toString().trim() === 'fit').length;
  } catch(e) {
    Logger.log(`getFitJobCount error: ${e.message}`);
    return 0;
  }
}
 
/**
 * Returns the highest M-level 'fit' job from Job_Search_Cache as a JSON string.
 * skipIndices: array of rowIndex values already loaded in this sidebar session.
 */
function getNextFitJob(skipIndices) {
  try {
    skipIndices = skipIndices || [];
    const sheet = getOrCreateJobCacheSheet();
    if (sheet.getLastRow() < 2) return JSON.stringify({ empty: true, fitCount: 0 });
 
    const data    = sheet.getDataRange().getValues();
    const fitJobs = [];
 
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if ((row[10] || '').toString().trim() !== 'fit') continue;
      const rowIndex = i + 1;
      if (skipIndices.includes(rowIndex)) continue;
 
      const matchLevel = (row[5] || '').toString().trim();
      fitJobs.push({
        rowIndex,
        company:          (row[1]  || '').toString(),
        title:            (row[2]  || '').toString(),
        url:              (row[3]  || '').toString(),
        cvType:           (row[4]  || '').toString(),
        matchLevel,
        levelNum:         parseInt(matchLevel.replace(/\D/g, '')) || 0,
        score:            row[6]   || 0,
        fetchedUrl:       (row[8]  || '').toString(),
        source:           (row[9]  || '').toString(),
        jdText:           (row[11] || '').toString()
      });
    }
 
    // Total fit count (before skip filter) for badge display
    const totalFit = data.slice(1)
      .filter(r => (r[10] || '').toString().trim() === 'fit').length;
 
    if (fitJobs.length === 0) return JSON.stringify({ empty: true, fitCount: totalFit });
 
    fitJobs.sort((a, b) =>
      b.levelNum !== a.levelNum ? b.levelNum - a.levelNum : b.rowIndex - a.rowIndex
    );
 
    const best = fitJobs[0];
    return JSON.stringify({
      rowIndex:          best.rowIndex,
      company:           best.company,
      title:             best.title,
      url:               best.url,
      cvType:            best.cvType,
      matchLevel:        best.matchLevel,
      score:             best.score,
      fetchedUrl:        best.fetchedUrl,
      source:            best.source,
      jdText:            best.jdText,
      urlMatchesFetched: best.url === best.fetchedUrl,
      fitCount:          totalFit - skipIndices.length    // remaining after loaded ones
    });
  } catch(e) {
    Logger.log(`getNextFitJob error: ${e.message}`);
    return JSON.stringify({ error: e.message });
  }
}
 
/**
 * Re-fetches fetchedUrl, computes Jaccard word-overlap similarity against
 * the cached JD text, and returns { accessible, score } as a JSON string.
 * Tries direct fetch first, then Tavily as fallback.
 * Returns accessible: false for blocked domains (LinkedIn, Indeed, etc.).
 */
function checkJdSimilarity(fetchedUrl, cachedJdText) {
  try {
    if (!fetchedUrl || !cachedJdText || cachedJdText.length < 100) {
      return JSON.stringify({ accessible: false, score: 0 });
    }
    if (DIRECT_FETCH_BLOCKED.some(d => fetchedUrl.includes(d))) {
      return JSON.stringify({ accessible: false, score: 0, reason: 'blocked_domain' });
    }
 
    let liveText = fetchJobPageDirectly(fetchedUrl);
 
    if (!liveText || liveText.length < 100) {
      liveText = tavilyExtractAdvanced(fetchedUrl);
    }
 
    if (!liveText || liveText.length < 100) {
      return JSON.stringify({ accessible: false, score: 0, reason: 'fetch_failed' });
    }
 
    const score = computeTextSimilarity(cachedJdText, liveText);
    return JSON.stringify({ accessible: true, score: Math.round(score * 100) });
  } catch(e) {
    Logger.log(`checkJdSimilarity error: ${e.message}`);
    return JSON.stringify({ accessible: false, score: 0, reason: 'error' });
  }
}
 
/**
 * Jaccard similarity between two texts.
 * Tokenises into unique lowercase words, removes stop words and short tokens.
 * Returns a float 0.0–1.0.
 */
function computeTextSimilarity(text1, text2) {
  const stopWords = new Set([
    'die','der','das','und','in','zu','den','ist','für','von','mit','wir','sie',
    'ihr','ein','eine','einen','dem','des','einer','werden','kann','auch','an',
    'bei','nach','auf','hat','wird','durch','haben','oder','aber','als','sind',
    'the','and','for','are','you','our','your','will','have','with','this',
    'that','from','not','but','all','can','we','be','to','of','in','at','on',
    'or','an','as','by','is','it','if','do','no','has','been','who','they','their'
  ]);
  function tokenize(text) {
    return new Set(
      text.toLowerCase()
          .replace(/[^a-zäöüßàáâèéêëìíîïòóôùúû0-9\s]/g, ' ')
          .split(/\s+/)
          .filter(w => w.length > 3 && !stopWords.has(w))
    );
  }
  const set1 = tokenize(text1);
  const set2 = tokenize(text2);
  if (set1.size === 0 || set2.size === 0) return 0;
  let intersection = 0;
  set1.forEach(w => { if (set2.has(w)) intersection++; });
  const union = set1.size + set2.size - intersection;
  return union === 0 ? 0 : intersection / union;
}
 
/**
 * Updates Review_Status in Job_Search_Cache to newStatus for ALL rows
 * matching company + normalised job title (handles minor name variations).
 */
function markCacheRowStatus(company, title, newStatus) {
  try {
    const sheet = getOrCreateJobCacheSheet();
    if (sheet.getLastRow() < 2) return;
    const nc   = company.toLowerCase().trim();
    const nt   = normalizeJobTitle(title).toLowerCase().trim();
    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
    let updated = 0;
    data.forEach((row, i) => {
      if (row[0].toString().toLowerCase().trim() === nc &&
          normalizeJobTitle(row[1].toString()).toLowerCase().trim() === nt) {
        sheet.getRange(i + 2, 11).setValue(newStatus);
        updated++;
      }
    });
    if (updated > 0) {
      SpreadsheetApp.flush();
      Logger.log(`markCacheRowStatus: "${company}" / "${title}" → "${newStatus}" (${updated} row(s))`);
    }
  } catch(e) {
    Logger.log(`markCacheRowStatus error: ${e.message}`);
  }
}
 
/**
 * Fires at 15:00 Berlin time (time-based trigger).
 * Counts 'in process' M1–M4 rows and:
 *  1. Writes summary to M2_Notifications A1  →  triggers Sheets mobile push
 *     if the user enabled "Content changes" notifications in the Sheets app.
 *  2. Stores summary in PENDING_NOTIFICATION_TEXT script property
 *     so onOpen() can display a desktop popup.
 */
function sendDailyNotificationSummary() {
  try {
    const sheet = getOrCreateJobCacheSheet();
    if (sheet.getLastRow() < 2) return;
    const data   = sheet.getDataRange().getValues();
    const counts = { M1: 0, M2: 0, M3: 0, M4: 0 };
    for (let i = 1; i < data.length; i++) {
      if ((data[i][10] || '').toString().trim() !== 'in process') continue;
      const lvl = parseInt((data[i][5] || '').toString().replace(/\D/g, '')) || 0;
      if (lvl >= 1 && lvl <= 4) counts['M' + lvl]++;
    }
    const total = counts.M1 + counts.M2 + counts.M3 + counts.M4;
    if (total === 0) { Logger.log('sendDailyNotificationSummary: no in-process jobs.'); return; }
 
    const parts = [];
    if (counts.M4 > 0) parts.push(`${counts.M4}× M4`);
    if (counts.M3 > 0) parts.push(`${counts.M3}× M3`);
    if (counts.M2 > 0) parts.push(`${counts.M2}× M2`);
    if (counts.M1 > 0) parts.push(`${counts.M1}× M1`);
 
    const ts       = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd-MM-yyyy HH:mm');
    const message  = `📋 ${total} job${total > 1 ? 's' : ''} in review: ${parts.join(', ')}`;
    const fullText = `${message}\nUpdated: ${ts}`;
 
    getOrCreateM2NotificationsSheet().getRange('A1').setValue(fullText);
    PropertiesService.getScriptProperties().setProperty('PENDING_NOTIFICATION_TEXT', fullText);
    SpreadsheetApp.flush();
    Logger.log(`Daily notification written: ${message}`);
  } catch(e) {
    Logger.log(`sendDailyNotificationSummary error: ${e.message}`);
  }
}
 
/**
 * Called from onOpen(). If a pending notification exists (written by
 * sendDailyNotificationSummary), shows a desktop popup and clears the property.
 */
function checkPendingNotification() {
  try {
    const props = PropertiesService.getScriptProperties();
    const text  = props.getProperty('PENDING_NOTIFICATION_TEXT');
    if (!text) return;
    SpreadsheetApp.getUi().alert(
      '📋 JABA Job Review',
      text + '\n\nOpen Job_Search_Cache → change promising rows to "fit" → use the sidebar to register.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    props.deleteProperty('PENDING_NOTIFICATION_TEXT');
  } catch(e) {
    Logger.log(`checkPendingNotification error: ${e.message}`);
  }
}
 
/**
 * Creates the daily 15:00 Berlin trigger for sendDailyNotificationSummary.
 * Run once from the menu: 🤖 AI Recruitment → 🔔 Setup Daily Notification (15:00).
 *
 * MOBILE SETUP (one-time, manual):
 *   1. Open Google Sheets app on your phone.
 *   2. Tap the JABA spreadsheet → ⋮ → Notifications → enable "When content changes".
 *   JABA writes to the hidden M2_Notifications sheet at 15:00 → phone push notification.
 */
function createDailyNotificationTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyNotificationSummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyNotificationSummary')
    .timeBased().atHour(15).everyDays(1).inTimezone('Europe/Berlin').create();
  SpreadsheetApp.getUi().alert(
    '✅ Daily notification trigger set for 15:00 Berlin time.\n\n' +
    'MOBILE PUSH (one-time setup):\n' +
    '1. Open Google Sheets app on your phone\n' +
    '2. Long-press or open ⋮ menu on the JABA spreadsheet\n' +
    '3. Tap Notifications → enable "When content changes"\n' +
    'JABA writes to the hidden M2_Notifications sheet at 15:00 → ' +
    'your phone receives a push notification.'
  );
}
/**
 * Scans Job_Search_Cache for 'in process' and 'fit' entries whose stored
 * JD_Text no longer passes the looksLikeJobContent quality check.
 * These entries were likely scored on a wrong/incomplete page (e.g. a
 * company homepage instead of the actual job posting).
 * Flagged entries are marked "suspicious" in the Review_Status column.
 * Run from: 🤖 AI Recruitment → 🔍 Audit Cache Quality
 */
function auditCacheQuality() {
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();
  const sheet = getOrCreateJobCacheSheet();

  if (sheet.getLastRow() < 2) {
    if (ui) ui.alert('Job_Search_Cache is empty — nothing to audit.');
    return;
  }

  const data      = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  const flagged   = [];
  const toUpdate  = []; // { rowIndex, newStatus }

  data.forEach(function(row, i) {
    const reviewStatus = (row[10] || '').toString().trim().toLowerCase();
    if (reviewStatus !== 'in process' && reviewStatus !== 'fit') return;

    const jdText = (row[11] || '').toString();

    // Flag if JD is missing or fails the new stricter quality check
    if (!jdText || !looksLikeJobContent(jdText)) {
      toUpdate.push({ rowIndex: i + 2, newStatus: 'suspicious' });
      flagged.push({
        company: (row[1] || '').toString(),
        title:   (row[2] || '').toString(),
        level:   (row[5] || '').toString(),
        score:   row[6]  || 0
      });
    }
  });

  // Write status updates in a batch
  toUpdate.forEach(function(u) {
    sheet.getRange(u.rowIndex, 11).setValue(u.newStatus);
  });

  if (toUpdate.length > 0) SpreadsheetApp.flush();

  if (flagged.length === 0) {
    if (ui) ui.alert(
      '✅ Cache Quality Audit',
      'All "in process" and "fit" entries passed the JD quality check.\nNo suspicious entries found.',
      ui.ButtonSet.OK
    );
    Logger.log('Cache quality audit: all entries passed.');
    return;
  }

  const lines = flagged.map(function(j) {
    return '• ' + j.company + ' — ' + j.title + ' (' + j.level + ': ' + j.score + '/40)';
  }).join('\n');

  const message =
    flagged.length + ' entry(ies) were scored on a JD that no longer passes quality checks\n' +
    '(likely a company homepage or truncated page, not the actual job posting).\n\n' +
    lines + '\n\n' +
    'These have been marked "suspicious" in Job_Search_Cache.\n\n' +
    'Options:\n' +
    '  • Open the sidebar → paste the real JD → run SMM to re-score\n' +
    '  • Change status to "discarded" if you\'ve already reviewed manually';

  if (ui) ui.alert('⚠️ Cache Quality Audit — ' + flagged.length + ' Suspicious Entry(ies)', message, ui.ButtonSet.OK);
  Logger.log('Cache quality audit: ' + flagged.length + ' suspicious entries flagged.');
}

/* ============================================================
   WEB APP — 3 NEW FUNCTIONS FOR Code.js
   
   WHERE TO PASTE: Open Code.js in Cloud Shell Editor.
   Scroll to the very bottom of the file.
   Paste everything below this comment block at the end.
   ============================================================ */


/**
 * Serves the JABA Web App when the URL is opened in a browser.
 * This is the "front door" for the Web App.
 * The sidebar (showSidebar) continues to work independently.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('JABA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Returns ALL Job_Search_Cache rows where Review_Status = 'fit',
 * sorted M4 → M3 → M2, then by score descending within each level.
 * Called by the Queue tab in Index.html on page load and after refresh.
 */
function getAllFitJobs() {
  try {
    const sheet = getOrCreateJobCacheSheet();
    if (sheet.getLastRow() < 2) return JSON.stringify([]);

    const data    = sheet.getDataRange().getValues();
    const fitJobs = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Column 11 (index 10) = Review_Status — filter for 'fit' only
      if ((row[10] || '').toString().trim() !== 'fit') continue;

      const matchLevel = (row[5] || '').toString().trim();
      const levelNum   = parseInt(matchLevel.replace(/\D/g, '')) || 0;

      fitJobs.push({
        rowIndex:    i + 1,                             // 1-based sheet row number
        company:     (row[1]  || '').toString().trim(),  // col 2: Company
        title:       (row[2]  || '').toString().trim(),  // col 3: Job_Title
        url:         (row[3]  || '').toString().trim(),  // col 4: URL
        cvType:      (row[4]  || '').toString().trim(),  // col 5: CV_Type
        matchLevel,                                      // col 6: Match_Level
        levelNum,
        score:       Number(row[6])  || 0,              // col 7: Score
        fetchSource: (row[7]  || '').toString().trim(),  // col 8: Fetch_Source
        fetchedUrl:  (row[8]  || '').toString().trim(),  // col 9: Fetched_URL
        source:      (row[9]  || '').toString().trim(),  // col 10: Source
        jdText:      (row[11] || '').toString()          // col 12: JD_Text
      });
    }

    // Sort: highest level first, then highest score within same level
    fitJobs.sort((a, b) =>
      b.levelNum !== a.levelNum
        ? b.levelNum - a.levelNum
        : b.score - a.score
    );

    Logger.log(`getAllFitJobs: ${fitJobs.length} fit job(s) returned`);
    return JSON.stringify(fitJobs);

  } catch (e) {
    Logger.log(`getAllFitJobs error: ${e.message}`);
    return JSON.stringify({ error: e.message });
  }
}
/* ============================================================
   FEATURE: MONTHLY SHEET DATA READER
   Returns all data from a named monthly tab as JSON for the
   web app Scan table. Includes headers + all rows.
   Also returns the list of all available monthly tab names
   so the web app can build the month navigator.
   ============================================================ */
function getMonthlySheetNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const SKIP_SHEETS = new Set([
      'Sankey_Data', 'Geo_Data', 'SMM_Raw_Data', 'Interview_Geo_Data',
      'Job_Search_Cache', 'Pending_SMM', 'Alert_Results', 'M2_Notifications',
      'Weekly_Data'
    ]);
 
    // Month parsing helper — returns a sortable date from "Jun 2026" etc.
    const MONTH_NAMES = {
      jan:0,feb:1,mar:2,apr:3,may:4,jun:5,
      jul:6,aug:7,sep:8,oct:9,nov:10,dec:11
    };
 
    const months = [];
    ss.getSheets().forEach(function(s) {
      const name = s.getName();
      if (SKIP_SHEETS.has(name)) return;
      const parts = name.trim().split(/\s+/);
      if (parts.length !== 2) return;
      const mon = MONTH_NAMES[parts[0].toLowerCase().slice(0, 3)];
      const yr  = parseInt(parts[1], 10);
      if (mon === undefined || isNaN(yr)) return;
      months.push({ name, sortKey: yr * 12 + mon });
    });
 
    // Sort newest first
    months.sort((a, b) => b.sortKey - a.sortKey);
    return JSON.stringify(months.map(m => m.name));
  } catch (e) {
    Logger.log(`getMonthlySheetNames error: ${e.message}`);
    return JSON.stringify([]);
  }
}
 
function getMonthlySheetData(tabName) {
  try {
    if (!tabName) return JSON.stringify({ error: 'No tab name provided' });
 
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return JSON.stringify({ error: `Tab "${tabName}" not found` });
 
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return JSON.stringify({ rows: [], tabName });
 
    // Read all 20 columns (A–T).  Cols J–R (10–18) are binary stage columns.
    const NUM_COLS = 20;
    const allData  = sheet.getRange(1, 1, lastRow, NUM_COLS).getValues();
 
    // Row 0 = headers, rows 1+ = data
    const headers = allData[0].map(h => (h || '').toString());
    const rows    = [];
 
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      // Skip completely empty rows (no company in col B)
      if (!row[1]) continue;
      rows.push(row.map(cell => {
        // Dates come back as Date objects from Sheets — format them
        if (cell instanceof Date && !isNaN(cell)) {
          return Utilities.formatDate(cell, CONFIG.TIMEZONE, 'dd.MM.yyyy');
        }
        return cell === null || cell === undefined ? '' : cell;
      }));
    }
 
    return JSON.stringify({ headers, rows, tabName });
  } catch (e) {
    Logger.log(`getMonthlySheetData error: ${e.message}`);
    return JSON.stringify({ error: e.message });
  }
}

/**
 * Marks a single Job_Search_Cache row as 'discarded'.
 * Called when the user clicks Skip in the Queue tab and selects a reason.
 * rowIndex : 1-based sheet row number (comes from getAllFitJobs).
 * reason   : string — logged only, not stored in the sheet.
 */
function skipFitJob(rowIndex, reason) {
  try {
    const sheet = getOrCreateJobCacheSheet();
    // Column 11 = Review_Status
    sheet.getRange(rowIndex, 11).setValue('discarded');
    SpreadsheetApp.flush();
    Logger.log(`skipFitJob: row ${rowIndex} → discarded. Reason: ${reason || 'none given'}`);
    return JSON.stringify({ success: true });
  } catch (e) {
    Logger.log(`skipFitJob error: ${e.message}`);
    return JSON.stringify({ error: e.message });
  }
}
/* ============================================================
   FEATURE: MONTHLY SHEET CELL UPDATER
   Called from the web app Scan table on every cell edit.
   Writes one value to a specific row + column in a monthly tab.
   For Status column (F = col 6): also calls updateRowStatusLogic
   and handles Rejected → Email Rejection timestamp.
   tabName  : e.g. "Jun 2026"
   rowIndex : 1-based sheet row (2 = first data row)
   colIndex : 1-based column (1 = A, 6 = F, etc.)
   value    : new cell value as string
   ============================================================ */
function updateMonthlyCell(tabName, rowIndex, colIndex, value) {
  try {
    if (!tabName || !rowIndex || !colIndex) {
      return JSON.stringify({ error: 'Missing parameters' });
    }
 
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return JSON.stringify({ error: `Tab "${tabName}" not found` });
 
    const cell = sheet.getRange(rowIndex, colIndex);
 
    // Column 7 (G) = Application Date — store as-is (string dd.MM.yyyy)
    // Column 10–18 (J–R) = binary stage columns — store as integer
    if (colIndex >= 10 && colIndex <= 18) {
      cell.setValue(value === '1' || value === true || value === 1 ? 1 : 0);
    } else {
      cell.setValue(value);
    }
 
    // Column 6 (F) = Status — trigger binary column update + Rejected timestamp
    if (colIndex === 6) {
      updateRowStatusLogic(sheet, rowIndex, value);
 
      if (value === 'Rejected') {
        const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy');
        sheet.getRange(rowIndex, 20).setValue(`${dateStr} 🙅🏽‍♂️`);
      }
 
      // Flag Interview_Reached in SMM_Raw_Data for HR/1st Interview
      const interviewStatuses = ['HR Interview', '1st Interview'];
      if (interviewStatuses.includes(value)) {
        const company = sheet.getRange(rowIndex, 2).getValue().toString().trim();
        const appDate = sheet.getRange(rowIndex, 7).getValue();
        flagSmmInterviewReached(company, appDate);
      }
 
      // Clear binary columns if status is cleared
      if (!value) {
        sheet.getRange(rowIndex, 9, 1, 10).clearContent();
      }
    }
 
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  } catch (e) {
    Logger.log(`updateMonthlyCell error: ${e.message}`);
    return JSON.stringify({ error: e.message });
  }
}
function focusJdContent(text) {
  if (!text || text.length < 300) return text;
 
  const lower = text.toLowerCase();
 
  // Patterns that mark the TRUE start of job content
  const jobStartPatterns = [
    // German
    'aufgaben', 'deine aufgaben', 'ihre aufgaben', 'zu deinen aufgaben',
    'was du machst', 'was du tust', 'was du bei uns',
    'in dieser rolle', 'wir suchen', 'du verantwortest',
    'stellenbeschreibung', 'über die stelle',
    'dein profil', 'anforderungen', 'voraussetzungen',
    'was wir uns wünschen', 'was du mitbringst',
    // English
    'responsibilities', 'what you will do', "what you'll do",
    'about the role', 'about this role', 'the role',
    'your role', 'in this role', 'job description',
    'requirements', "what we're looking for", 'who you are',
    'you will be', 'we are looking for', "we're looking for",
    'about the job', 'about the position'
  ];
 
  let bestStart = -1;
 
  for (const pattern of jobStartPatterns) {
    const idx = lower.indexOf(pattern);
    // Only trim noise that is between 150 and 3000 chars from the start.
    // Below 150: probably already at the job start — don't trim.
    // Above 3000: the "pattern" found is likely inside the job itself — don't trim.
    if (idx > 150 && idx < 6000) {
      if (bestStart === -1 || idx < bestStart) {
        bestStart = idx;
      }
    }
  }
 
  if (bestStart > 150) {
    // Include 150 chars of context before the pattern (catches section headings)
    const trimmed = text.substring(Math.max(0, bestStart - 150));
    Logger.log(`focusJdContent: trimmed ${bestStart - 150} chars of leading noise`);
    return trimmed;
  }
 
  return text;
}
 
/**
 * Resolves a company's headquarters city when the JD contains no location.
 * Uses Tavily Search (1 credit) + Groq extraction.
 * Results cached in Script Properties — same company never looked up twice.
 * Returns { city, country, fromCompany: true } or { city: null, country: null }.
 */
function resolveCompanyLocation(companyName) {
  if (!companyName || companyName === 'Unknown') {
    return { city: null, country: null };
  }

  // ── Cache check ───────────────────────────────────────────────────────────
  const props      = PropertiesService.getScriptProperties();
  const cacheKey   = 'COMPLOC_' + companyName
    .toLowerCase()
    .replace(/ä/g,'ae').replace(/ö/g,'oe').replace(/ü/g,'ue').replace(/ß/g,'ss')
    .replace(/[^a-z0-9]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'')
    .substring(0, 40);

  try {
    const cached = props.getProperty(cacheKey);
    if (cached) {
      if (cached === 'NOTFOUND') {
        Logger.log(`resolveCompanyLocation: cached miss → "${companyName}"`);
        return { city: null, country: null };
      }
      const parsed = JSON.parse(cached);
      Logger.log(`resolveCompanyLocation: cache hit → "${companyName}" → ${parsed.city}, ${parsed.country}`);
      return { ...parsed, fromCompany: true };
    }
  } catch(e) {
    Logger.log(`resolveCompanyLocation: cache read error: ${e.message}`);
  }

  // ── Tavily Search (1 credit) ──────────────────────────────────────────────
  const tavilyKey = getTavilyKey();
  if (!tavilyKey) {
    Logger.log('resolveCompanyLocation: TAVILY_API_KEY not set');
    return { city: null, country: null };
  }

  let snippet = null;
  try {
    const searchRes = UrlFetchApp.fetch('https://api.tavily.com/search', {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify({
        api_key:      tavilyKey,
        query:        `"${companyName}" Unternehmen Standort Hauptsitz`,
        search_depth: 'basic',
        max_results:  3,
        include_raw_content: false,
        exclude_domains: ['linkedin.com', 'xing.com', 'glassdoor.com', 'kununu.com']
      }),
      muteHttpExceptions: true
    });

    if (searchRes.getResponseCode() === 200) {
      const data    = JSON.parse(searchRes.getContentText());
      const results = data.results || [];
      // Concatenate the top 3 snippets for better extraction accuracy
      snippet = results
        .map(r => (r.title || '') + ' ' + (r.content || ''))
        .join(' ')
        .substring(0, 1500);
      Logger.log(`resolveCompanyLocation: Tavily got ${results.length} results for "${companyName}"`);
    } else {
      Logger.log(`resolveCompanyLocation: Tavily HTTP ${searchRes.getResponseCode()}`);
    }
  } catch(e) {
    Logger.log(`resolveCompanyLocation: Tavily exception: ${e.message}`);
  }

  if (!snippet || snippet.trim().length < 30) {
    try { props.setProperty(cacheKey, 'NOTFOUND'); } catch(e) {}
    return { city: null, country: null };
  }

  // ── Groq extraction ───────────────────────────────────────────────────────
  const systemPrompt = 'You extract company headquarters location from text snippets. JSON only, no markdown.';
  const userPrompt   = `Extract the city and country where "${companyName}" is headquartered.

TEXT:
${snippet}

Rules:
- city: the specific city name only (e.g. "Berlin"), or null if not clearly stated
- country: the country name in English (e.g. "Germany"), or null if not clearly stated
- If multiple cities are mentioned, pick the one most likely to be the HQ
- If the company is clearly remote-only with no physical HQ, set both to null

Respond ONLY with JSON: {"city": "Berlin", "country": "Germany"}`;

  const raw = callGroqApi(systemPrompt, userPrompt, 'llama-3.1-8b-instant', 60);

  if (!raw) {
    Logger.log(`resolveCompanyLocation: Groq returned nothing for "${companyName}"`);
    try { props.setProperty(cacheKey, 'NOTFOUND'); } catch(e) {}
    return { city: null, country: null };
  }

  try {
    const extracted = JSON.parse(raw.replace(/```json|```/g, '').trim());
    const city      = extracted.city    || null;
    const country   = extracted.country || null;

    if (!city && !country) {
      try { props.setProperty(cacheKey, 'NOTFOUND'); } catch(e) {}
      return { city: null, country: null };
    }

    // Cache the successful result
    try {
      props.setProperty(cacheKey, JSON.stringify({ city, country }));
    } catch(e) {}

    Logger.log(`resolveCompanyLocation: resolved "${companyName}" → ${city}, ${country}`);
    return { city, country, fromCompany: true };

  } catch(e) {
    Logger.log(`resolveCompanyLocation: Groq parse error: ${e.message} | raw: ${raw}`);
    try { props.setProperty(cacheKey, 'NOTFOUND'); } catch(e2) {}
    return { city: null, country: null };
  }
}

/**
 * Enriches an SMM result object with company-lookup location when the JD
 * provided no location. Modifies parsed in place and returns it.
 *
 * companyName : string — the company name from the job listing
 * parsed      : object — already-parsed SMM JSON result
 *
 * Sets parsed.location_source to:
 *   "jd"             — location came from JD text (already set, untouched)
 *   "company_lookup" — location found via company name search
 *   "unknown"        — no location found anywhere
 */
function enrichSmmLocation(parsed, companyName) {
  // Already has a JD location — nothing to do
  if (parsed.location_source === 'jd' &&
      parsed.job_location &&
      (parsed.job_location.city || parsed.job_location.country)) {
    return parsed;
  }

  const resolved = resolveCompanyLocation(companyName);

  if (resolved.city || resolved.country) {
    parsed.job_location    = { city: resolved.city, country: resolved.country };
    parsed.location_source = 'company_lookup';
    Logger.log(`enrichSmmLocation: "${companyName}" → ${resolved.city}, ${resolved.country}`);
  } else {
    parsed.location_source = 'unknown';
  }

  return parsed;
}

/**
 * Geocodes a city + country pair via Nominatim (OpenStreetMap).
 * Called from the sidebar/web app when the user switches to the Map chart view.
 * Returns JSON string: { lat, lng, display } or { lat: null, lng: null }.
 * Note: Nominatim requires a proper User-Agent — never call it client-side.
 */
/**
 * Geocodes a city + country pair.
 * Tier 1: hardcoded lookup table (instant, zero API calls)
 * Tier 2: Open-Meteo Geocoding API (free, no key, works from GCP IPs)
 * Results cached in Script Properties for instant repeat lookups.
 * Same return format: JSON string { lat, lng, display } or { lat: null, lng: null }
 */
function geocodeLocation(city, country) {
  if (!city && !country) return JSON.stringify({ lat: null, lng: null });
  const cityClean = (city || '').trim();
  if (/^(remote|homeoffice|home\s*office|home-office)$/i.test(cityClean)) {
    return JSON.stringify({ lat: null, lng: null, isRemote: true });
  }

  // ── Build a stable cache key ──────────────────────────────────────────────
  const cacheKey = 'GEO2_' + cityClean
    .toLowerCase()
    .replace(/ä/g,'ae').replace(/ö/g,'oe').replace(/ü/g,'ue').replace(/ß/g,'ss')
    .replace(/[^a-z0-9]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'');

  const props = PropertiesService.getScriptProperties();

  // ── Check cache first (both success and known-failures) ──────────────────
  try {
    const cached = props.getProperty(cacheKey);
    if (cached) {
      if (cached === 'NOTFOUND') {
        Logger.log(`geocodeLocation: cached miss → "${cityClean}"`);
        return JSON.stringify({ lat: null, lng: null });
      }
      const parts = cached.split(',');
      const lat = parseFloat(parts[0]);
      const lng = parseFloat(parts[1]);
      if (!isNaN(lat) && !isNaN(lng)) {
        Logger.log(`geocodeLocation: cache hit → "${cityClean}" (${lat}, ${lng})`);
        return JSON.stringify({
          lat, lng,
          display: cityClean + (country ? ', ' + country : ''),
          source: 'cache'
        });
      }
    }
  } catch(e) {
    Logger.log(`geocodeLocation: cache read error: ${e.message}`);
  }

  // ── Tier 1: hardcoded lookup table ────────────────────────────────────────
  // Covers the most common German cities + major European job hubs.
  const COORDS = {
    // German cities — alphabetical
    'aachen':        [50.7753, 6.0839],
    'augsburg':      [48.3705, 10.8978],
    'berlin':        [52.5200, 13.4050],
    'bielefeld':     [52.0302, 8.5325],
    'bochum':        [51.4818, 7.2162],
    'bonn':          [50.7374, 7.0982],
    'braunschweig':  [52.2689, 10.5268],
    'bremen':        [53.0793, 8.8017],
    'chemnitz':      [50.8279, 12.9214],
    'cologne':       [50.9333, 6.9500],
    'darmstadt':     [49.8728, 8.6512],
    'dortmund':      [51.5136, 7.4653],
    'dresden':       [51.0504, 13.7373],
    'duesseldorf':   [51.2217, 6.7762],
    'duisburg':      [51.4344, 6.7623],
    'erfurt':        [50.9848, 11.0299],
    'essen':         [51.4556, 7.0116],
    'frankfurt':     [50.1109, 8.6821],
    'freiburg':      [47.9990, 7.8421],
    'gelsenkirchen': [51.5177, 7.0857],
    'halle':         [51.4828, 11.9697],
    'hamburg':       [53.5753, 10.0153],
    'hannover':      [52.3759, 9.7320],
    'heidelberg':    [49.3988, 8.6724],
    'heilbronn':     [49.1427, 9.2109],
    'ingolstadt':    [48.7665, 11.4257],
    'karlsruhe':     [49.0069, 8.4037],
    'kassel':        [51.3127, 9.4797],
    'kiel':          [54.3233, 10.1228],
    'koblenz':       [50.3569, 7.5890],
    'koeln':         [50.9333, 6.9500],
    'langenfeld':    [51.1088, 6.9481],
    'leipzig':       [51.3397, 12.3731],
    'luebeck':       [53.8655, 10.6866],
    'ludwigshafen':  [49.4774, 8.4453],
    'magdeburg':     [52.1205, 11.6276],
    'mainz':         [49.9929, 8.2473],
    'mannheim':      [49.4875, 8.4660],
    'muenchen':      [48.1351, 11.5820],
    'munich':        [48.1351, 11.5820],
    'muenster':      [51.9607, 7.6261],
    'nuernberg':     [49.4521, 11.0767],
    'nuremberg':     [49.4521, 11.0767],
    'oberhausen':    [51.4697, 6.8509],
    'oldenburg':     [53.1434, 8.2146],
    'osnabrueck':    [52.2799, 8.0472],
    'potsdam':       [52.3906, 13.0645],
    'regensburg':    [49.0134, 12.1016],
    'rostock':       [54.0887, 12.1404],
    'saarbruecken':  [49.2354, 6.9965],
    'sassnitz':      [54.5167, 13.6333],
    'stuttgart':     [48.7758, 9.1829],
    'ulm':           [48.3984, 9.9922],
    'wiesbaden':     [50.0782, 8.2398],
    'wolfsburg':     [52.4227, 10.7865],
    'wuerzburg':     [49.7913, 9.9534],
    // Major European job hubs
    'amsterdam':     [52.3676, 4.9041],
    'antwerp':       [51.2194, 4.4025],
    'athens':        [37.9838, 23.7275],
    'barcelona':     [41.3851, 2.1734],
    'brussels':      [50.8503, 4.3517],
    'budapest':      [47.4979, 19.0402],
    'copenhagen':    [55.6761, 12.5683],
    'dublin':        [53.3498, -6.2603],
    'dubai':         [25.2048, 55.2708],
    'helsinki':      [60.1699, 24.9384],
    'istanbul':      [41.0082, 28.9784],
    'lisbon':        [38.7223, -9.1393],
    'london':        [51.5074, -0.1278],
    'luxembourg':    [49.6116, 6.1319],
    'madrid':        [40.4168, -3.7038],
    'milan':         [45.4642, 9.1900],
    'oslo':          [59.9139, 10.7522],
    'paris':         [48.8566, 2.3522],
    'prague':        [50.0755, 14.4378],
    'riga':          [56.9496, 24.1052],
    'rome':          [41.9028, 12.4964],
    'stockholm':     [59.3293, 18.0686],
    'tallinn':       [59.4370, 24.7536],
    'vienna':        [48.2082, 16.3738],
    'vilnius':       [54.6872, 25.2797],
    'warsaw':        [52.2297, 21.0122],
    'zurich':        [47.3769, 8.5417],
    // Aliases / common misspellings
    'wien':          [48.2082, 16.3738],
    'bruessel':      [50.8503, 4.3517],
    'zuerich':       [47.3769, 8.5417],
    'kopenhagen':    [55.6761, 12.5683],
    'prag':          [50.0755, 14.4378],
    'lissabon':      [38.7223, -9.1393],
    'mailand':       [45.4642, 9.1900],
    'rom':           [41.9028, 12.4964],
    'warschau':      [52.2297, 21.0122],
    'moskau':        [55.7558, 37.6173],
  };

  // Normalise city for lookup: lowercase + umlaut substitution
  const cityNorm = cityClean
    .toLowerCase()
    .replace(/ä/g,'ae').replace(/ö/g,'oe').replace(/ü/g,'ue').replace(/ß/g,'ss')
    .replace(/[^a-z]/g,'');

  if (COORDS[cityNorm]) {
    const [lat, lng] = COORDS[cityNorm];
    Logger.log(`geocodeLocation: hardcoded hit → "${cityClean}" (${lat}, ${lng})`);
    try { props.setProperty(cacheKey, lat + ',' + lng); } catch(e) {}
    return JSON.stringify({
      lat, lng,
      display: cityClean + (country ? ', ' + country : ''),
      source: 'hardcoded'
    });
  }

  // ── Tier 2: Open-Meteo Geocoding API (free, no key, GCP-friendly) ─────────
  try {
    const query   = encodeURIComponent(cityClean + (country ? ', ' + country : ''));
    const apiUrl  = 'https://geocoding-api.open-meteo.com/v1/search?name=' +
                    encodeURIComponent(cityClean) +
                    '&count=5&language=en&format=json';

    const res = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      headers: { 'Accept': 'application/json' },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log(`geocodeLocation: Open-Meteo HTTP ${res.getResponseCode()} for "${cityClean}"`);
      try { props.setProperty(cacheKey, 'NOTFOUND'); } catch(e) {}
      return JSON.stringify({ lat: null, lng: null });
    }

    const data    = JSON.parse(res.getContentText());
    const results = data.results || [];

    if (results.length === 0) {
      Logger.log(`geocodeLocation: Open-Meteo no results for "${cityClean}"`);
      try { props.setProperty(cacheKey, 'NOTFOUND'); } catch(e) {}
      return JSON.stringify({ lat: null, lng: null });
    }

    // Pick the best result: prefer country match if available
    let best = results[0];
    if (country) {
      const countryLower = country.toLowerCase();
      const countryMatch = results.find(r =>
        (r.country || '').toLowerCase().includes(countryLower) ||
        (r.country_code || '').toLowerCase() === countryLower.substring(0, 2)
      );
      if (countryMatch) best = countryMatch;
    }

    const lat     = best.latitude;
    const lng     = best.longitude;
    const display = best.name + (best.admin1 ? ', ' + best.admin1 : '') +
                    (best.country ? ', ' + best.country : '');

    try { props.setProperty(cacheKey, lat + ',' + lng); } catch(e) {}
    Logger.log(`geocodeLocation: Open-Meteo OK → "${cityClean}" → ${display} (${lat}, ${lng})`);

    return JSON.stringify({ lat, lng, display, source: 'open-meteo' });

  } catch(e) {
    Logger.log(`geocodeLocation: Open-Meteo exception: ${e.message}`);
    try { props.setProperty(cacheKey, 'NOTFOUND'); } catch(ee) {}
    return JSON.stringify({ lat: null, lng: null });
  }
}
 
 
/**
 * Returns live stats for the sidebar/web app stats strip:
 * - applied : total application rows across all monthly sheets
 * - inReview: 'in process' + 'fit' rows in Job_Search_Cache
 * - fit     : 'fit' rows in Job_Search_Cache
 */
function getQueueStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const skipSheets = new Set([
      'Sankey_Data', 'Geo_Data', 'SMM_Raw_Data', 'Interview_Geo_Data',
      'Job_Search_Cache', 'Pending_SMM', 'Alert_Results', 'M2_Notifications',
      'Weekly_Data'
    ]);
 
    // Count total applied rows across all monthly sheets
    let applied = 0;
    ss.getSheets().forEach(function(sheet) {
      const name = sheet.getName();
      if (skipSheets.has(name) || !/[A-Za-z]+ \d{4}/.test(name)) return;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][1]) applied++; // col B = Company
      }
    });
 
    // Count cache statuses
    let inReview = 0;
    let fit      = 0;
    const cacheSheet = ss.getSheetByName('Job_Search_Cache');
    if (cacheSheet && cacheSheet.getLastRow() > 1) {
      const statuses = cacheSheet
        .getRange(2, 11, cacheSheet.getLastRow() - 1, 1)
        .getValues();
      statuses.forEach(function(row) {
        const s = (row[0] || '').toString().trim();
        if (s === 'fit')        { fit++; inReview++; }
        if (s === 'in process') { inReview++; }
      });
    }
 
    return JSON.stringify({ applied, inReview, fit });
 
  } catch (e) {
    Logger.log(`getQueueStats error: ${e.message}`);
    return JSON.stringify({ error: e.message });
  }
}
/* ============================================================
   FEATURE: DUPLICATE DETECTION
   Called from sidebar / web app BEFORE triggering SMM.
   Checks ALL monthly tabs for an exact match on
   Company (col B) + normalised Position (col C).
   Returns JSON string:
     { found: false }
     { found: true, tabName: "Jun 2026", company: "ACME", position: "CRM Manager" }
   ============================================================ */
function checkIfAlreadyApplied(company, position) {
  try {
    if (!company || !position) return JSON.stringify({ found: false });
 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nc = company.toLowerCase().trim();
    const np = normalizeJobTitle(position).toLowerCase().trim();
 
    const SKIP_SHEETS = new Set([
      'Sankey_Data', 'Geo_Data', 'SMM_Raw_Data', 'Interview_Geo_Data',
      'Job_Search_Cache', 'Pending_SMM', 'Alert_Results', 'M2_Notifications',
      'Weekly_Data'
    ]);
 
    for (const sheet of ss.getSheets()) {
      const name = sheet.getName();
      if (SKIP_SHEETS.has(name) || !/[A-Za-z]+ \d{4}/.test(name)) continue;
      if (sheet.getLastRow() < 2) continue;
 
      const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
      for (const row of data) {
        const rc = (row[0] || '').toString().toLowerCase().trim();
        const rp = normalizeJobTitle((row[1] || '').toString()).toLowerCase().trim();
        if (rc === nc && rp === np) {
          return JSON.stringify({
            found:    true,
            tabName:  name,
            company:  (row[0] || '').toString().trim(),
            position: (row[1] || '').toString().trim()
          });
        }
      }
    }
    return JSON.stringify({ found: false });
  } catch (e) {
    Logger.log(`checkIfAlreadyApplied error: ${e.message}`);
    return JSON.stringify({ found: false });
  }
} 
 
/**
 * Applies conditional formatting to Job_Search_Cache based on Review_Status
 * (column K = col 11). Run once from the menu after deployment.
 * Colors use JABA's dark theme palette.
 */
function applyJobCacheFormatting() {
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();
  const sheet = getOrCreateJobCacheSheet();
 
  sheet.clearConditionalFormatRules();
 
  const maxRow  = sheet.getMaxRows();
  const dataRng = sheet.getRange(2, 1, maxRow - 1, 12);
 
  function rule(value, bg, fg) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$K2="' + value + '"')
      .setBackground(bg)
      .setFontColor(fg)
      .setRanges([dataRng])
      .build();
  }
 
  sheet.setConditionalFormatRules([
    rule('fit',         '#0d2e1e', '#34d399'),  // green — ready to act on
    rule('in process',  '#0f1f3a', '#4f8ef7'),  // blue  — scored, under review
    rule('registered',  '#0a1f12', '#22c55e'),  // bright green — applied
    rule('discarded',   '#1a1a1a', '#6b7280'),  // gray  — skipped
    rule('suspicious',  '#2d1f0a', '#f59e0b'),  // amber — needs re-check
    rule('auto',        '#141414', '#4b5563'),  // very muted — M0, auto-dismissed
  ]);
 
  SpreadsheetApp.flush();
  Logger.log('applyJobCacheFormatting: conditional formatting applied.');
 
  if (ui) ui.alert(
    '🎨 Job Cache Color Coding Applied',
    '🟢 fit         — green\n' +
    '🔵 in process  — blue\n' +
    '✅ registered  — bright green\n' +
    '⚫ discarded   — dark gray\n' +
    '🟡 suspicious  — amber\n' +
    '⬛ auto        — very muted\n\n' +
    'Run again anytime to refresh.',
    ui.ButtonSet.OK
  );
}
 
 
/**
 * Builds a Weekly_Data sheet with application counts per ISO week.
 * Called by refreshAllData() and can be triggered manually from the menu.
 * Feeds the weekly trend chart in Looker Studio.
 */
function updateWeeklyData() {
  const ss         = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const sheets     = ss.getSheets();
  const weekCounts = {};
 
  const skipSheets = new Set([
    'Sankey_Data', 'Geo_Data', 'SMM_Raw_Data', 'Interview_Geo_Data',
    'Job_Search_Cache', 'Pending_SMM', 'Alert_Results', 'M2_Notifications',
    'Weekly_Data'
  ]);
 
  sheets.forEach(function(sheet) {
    const sheetName = sheet.getName().replace(/\./g, '');
    if (skipSheets.has(sheetName) || !/[A-Za-z]+ \d{4}/.test(sheetName)) return;
 
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() < 2) return;
    const data = dataRange.getValues();
 
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1]) continue;  // empty company = skip
 
      // Column G (index 6) = Application Date (dd.MM.yyyy or Date object)
      const rawDate = row[6];
      let dateObj;
 
      if (rawDate instanceof Date && !isNaN(rawDate)) {
        dateObj = rawDate;
      } else if (rawDate) {
        const parts = rawDate.toString().split('.');
        if (parts.length === 3) {
          dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        } else {
          continue;
        }
      } else {
        continue;
      }
 
      if (isNaN(dateObj)) continue;
 
      // ISO week number
      const year        = dateObj.getFullYear();
      const startOfYear = new Date(year, 0, 1);
      const weekNum     = Math.ceil(
        ((dateObj - startOfYear) / 86400000 + startOfYear.getDay() + 1) / 7
      );
      const weekLabel   = `W${String(weekNum).padStart(2, '0')} ${year}`;
 
      if (!weekCounts[weekLabel]) {
        weekCounts[weekLabel] = { year, weekNum, count: 0 };
      }
      weekCounts[weekLabel].count++;
    }
  });
 
  // Sort chronologically
  const sorted = Object.keys(weekCounts).sort(function(a, b) {
    const wa = weekCounts[a], wb = weekCounts[b];
    return wa.year !== wb.year ? wa.year - wb.year : wa.weekNum - wb.weekNum;
  });
 
  let weeklySheet = ss.getSheetByName('Weekly_Data');
  if (!weeklySheet) {
    weeklySheet = ss.insertSheet('Weekly_Data');
  } else {
    weeklySheet.clearContents();
    weeklySheet.clearFormats();
  }
 
  const rows    = sorted.map(function(key) {
    const d = weekCounts[key];
    return [String(d.year), String(d.weekNum), key, d.count];
  });
  const allRows = [['Year', 'Week_Num', 'Week_Label', 'Applications'], ...rows];
  const range   = weeklySheet.getRange(1, 1, allRows.length, 4);
  range.setNumberFormat('@STRING@');  // Prevent Looker Studio date conversion
  range.setValues(allRows);
 
  Logger.log(`updateWeeklyData: ${rows.length} week(s) written.`);
}