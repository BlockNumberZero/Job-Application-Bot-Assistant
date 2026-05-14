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
  SMM:           'v1.4',
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

function getSmmCacheKey(jdText, cvType) {
  const input   = cvType + '|' + jdText.replace(/\s+/g, ' ').substring(0, 3000);
  const digest  = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
  return 'smm_' + digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
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
    .addItem('Open AI Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Scan Gmail for Applications', 'processGmailApplications')
    .addItem('Scan Gmail for Rejections', 'processRejectionEmails')
    .addSeparator()
    .addItem('🔄 Refresh Dashboard Data', 'refreshAllData')
    .addItem('🧠 Refresh SMM Categories (min. 20 apps)', 'batchRefreshMasterCategories')
    .addSeparator()
    .addItem('📧 Process Indeed Job Alerts', 'processIndeedAlertEmails')
    .addSeparator()
    .addItem('🔍 Run Daily Job Search', 'runDailyJobSearch')
    .addToUi();
}

function refreshAllData() {
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();
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
    const cleanedJD = jdInput.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '').substring(0, 6000);

// Guard: JD too short for meaningful SMM analysis
if (cleanedJD.length < 300) {
  Logger.log(`SMM skipped — JD too short: ${cleanedJD.length} chars`);
  return JSON.stringify({ error: 'JD too short for SMM analysis (< 300 chars). Try fetching the full job page.' });
}
        // Check cache before calling Mistral
    const cacheKey    = getSmmCacheKey(cleanedJD, cvType);
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
   Language requirements — apply this logic exactly:
- German C2, "Muttersprache", "native speaker", or "verhandlungssicher": 
  include as a skill, classify as Crucial, score it 0/5 — the candidate holds 
  C1 which does not satisfy these levels.
- German C1, "fließend", "fluent", "gute Kenntnisse", or any English language 
  requirement: EXCLUDE entirely — not a differentiating factor for this candidate.
- No language level specified: EXCLUDE entirely.
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
3. Classify JD importance: "Crucial" = role cannot be done without it. "Necessary" = strongly preferred. "Optional" = mentioned once or as a plus.
4. Evidence: copy max 10 words verbatim from the CV, or write exactly "Not found in CV".
5. Gap tip: max 10 words, start with a verb (e.g. "Add", "Quantify", "Include").

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
  const parsed = repairAndParseSmm(content); // uses the new helper below
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
    const resultStr = JSON.stringify(parsed);
scriptCache.put(cacheKey, resultStr, 43200); // cache for 12 hours
return resultStr;

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
  'marketing','crm','growth','content','social media','web3','blockchain',
  'community','automation','digital','campaign','brand','branding','seo','sea',
  'performance','email marketing','influencer','analytics','communications',
  'copywriter','storytelling','acquisition','retention','engagement','e-commerce',
  // German
  'wachstum','inhalt','gemeinschaft','automatisierung','kampagne','marke',
  'leistung','kommunikation','öffentlichkeitsarbeit','digitalmarketing',
  'onlinemarketing','online-marketing','markenführung','reichweite'
];

const SKIP_KEYWORDS = [
  // English
  'sales','engineer','software developer','software engineer','lawyer',
  'accountant','nurse','driver','warehouse','key account','key-account',
  'recruiter','finance controller','electrician','plumber','mechanic',
  'chef','cook','cleaner','internship','intern ','fellowship','security guard',
  // German
  'vertrieb','verkauf','ingenieur','softwareentwickler','entwickler',
  'rechtsanwalt','buchhalter','steuerberater','krankenschwester','pfleger',
  'fahrer','lagerarbeiter','lagermitarbeiter','schlüsselkunde','key account',
  'personalvermittler','werkstudent','werkstudentin','praktikum','praktikant','praktikantin','pflichtpraktikum','elektriker','klempner','mechaniker','koch','reinigungskraft'
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
      const text = tavilyExtract(r.url);
      if (!text || !looksLikeJobContent(text)) { Utilities.sleep(400); continue; }
      if (!isJdRelevantToJob(text, company, jobTitle)) {
        Logger.log(`  Tavily result rejected (irrelevant): ${r.url}`);
        Utilities.sleep(400);
        continue;
      }
      Logger.log(`  Tavily search validated OK: ${r.url}`);
      return text;
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
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ')
    .replace(/&#\d+;/g, ' ')
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
   MAIN: processIndeedAlertEmails
   Run from Apps Script menu.
   ============================================================ */
function processIndeedAlertEmails() {
  const props    = PropertiesService.getScriptProperties();
  const timezone = CONFIG.TIMEZONE;
  const ui = (() => { try { return SpreadsheetApp.getUi(); } catch(e) { return null; } })();

  // Determine scan window
  const lastScanStr    = props.getProperty('LAST_ALERT_SCAN');
  const sinceDate      = lastScanStr
    ? new Date(lastScanStr)
    : new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000);
  const sinceFormatted = Utilities.formatDate(sinceDate, timezone, 'yyyy/MM/dd');
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
        const smmRaw = analyzeSkillsMatch(jdText, cvType);
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
      sheet.getRange(startRow, 1, rows.length, 13).setValues(rows);
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
    'nach reiflicher überlegung',
    'nach sorgfältiger prüfung',
    'nach eingehender prüfung',
    'anderen kandidaten den vorzug',
    'anderen bewerbern den vorzug'
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
  'senior', 'sr.', 'lead ', 'head of', 'director',
  'vp ', 'vice president', 'sales', 'vertrieb', 'ausbildung',
  'verkauf', 'key account', 'leiter', 'leitung', 'cmo', 'praktikum', 
  'praktikant', 'werkstudent', 'werkstudentin',
  'internship', 'intern ' 
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
  const signals = [
    'aufgaben', 'anforderungen', 'erfahrung', 'kenntnisse', 'qualifikation',
    'requirements', 'responsibilities', 'experience', 'skills', 'qualifications',
    'bewerb', 'stelle', 'vollzeit', 'teilzeit', 'wir suchen', 'we are looking',
    'marketing', 'automation', 'crm', 'manager', 'kampagne', 'campaign'
  ];
  const lower = text.toLowerCase();
  return signals.filter(s => lower.includes(s)).length >= 2;
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
 * Returns true if the fetched text is actually about the right job.
 * Prevents JABA from using a competitor's careers page or unrelated content.
 */
function isJdRelevantToJob(text, company, jobTitle) {
  if (!text || text.length < 200) return false;
  const lower = text.toLowerCase();

  const companyWords = company
    .toLowerCase()
    .replace(/\s+(gmbh|ag|se|kg|ltd|inc|llc|bv|sas|co\.)\b/gi, '')
    .replace(/[&.,\-]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length > 3);

  const titleWords = jobTitle
    .toLowerCase()
    .replace(/\(m\/w\/d\)|\(w\/m\/d\)|\(f\/m\/d\)/gi, '')
    .split(/\s+/)
    .filter(w => w.length > 4);

  // Company name MUST appear — no company match = wrong page, full stop
  const companyMatch = companyWords.length > 0 &&
    companyWords.some(w => lower.includes(w));

  if (!companyMatch) {
    Logger.log(`  ✗ Relevance: company "${company}" not found in fetched text`);
    return false;
  }

  // With company confirmed, require at least one title word too
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
  addJobs(fetchRemotiveJobs(),   'Remotive');
  Utilities.sleep(500);
  addJobs(fetchJobicyJobs(),     'Jobicy');
  Utilities.sleep(500);
  addJobs(fetchArbeitnowJobs(),  'Arbeitnow');
  Utilities.sleep(500);
  addJobs(fetchRemoteOkJobs(),   'Remote OK');
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

        const title    = (job.title || '').trim();
        const location = (job.location || '').trim();
        const isRemote = job.remote === true;
        const inGermany = isGermanLocation(location);

        // Geographic rule: Germany = all jobs; abroad = remote only
        if (!inGermany && !isRemote) continue;

        // Keyword relevance filter
        if (!isRelevantJobTitle(title)) continue;

        allJobs.push({
          title:           title,
          company:         (job.company_name || 'Unknown').trim(),
          city:            location || (inGermany ? 'Germany' : 'Remote'),
          url:             job.url || '',
          description:     stripHtmlToText(job.description || '').substring(0, 10000),
          id:              id,
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

  Logger.log(`Arbeitnow: ${allJobs.length} relevant jobs`);
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
  // Remotive and Jobicy return complete JDs — use directly, no fetching needed
  if (descriptionFull && apiDescription && looksLikeJobContent(apiDescription)) {
    Logger.log(`  ✓ Using full API description (${apiDescription.length} chars)`);
    return { text: apiDescription, source: 'api_full' };
  }

  // Tier 1 — Direct UrlFetchApp
  const direct = fetchJobPageDirectly(url);
  if (direct && looksLikeJobContent(direct)) {
    Logger.log(`  ✓ Direct fetch OK (${direct.length} chars)`);
    return { text: direct, source: 'direct' };
  }

  // Tier 2 — Tavily advanced
  const tavily = tavilyExtractAdvanced(url);
  if (tavily && looksLikeJobContent(tavily)) {
    Logger.log(`  ✓ Tavily advanced OK`);
    return { text: tavily, source: 'tavily_advanced' };
  }

  // Tier 3 — Adzuna snippet fallback
  if (apiDescription && apiDescription.length > 200 && looksLikeJobContent(apiDescription)) {
    Logger.log(`  ✓ Using API snippet fallback (${apiDescription.length} chars)`);
    return { text: apiDescription, source: 'api_snippet' };
  }

  Logger.log(`  ✗ All tiers failed for: ${url}`);
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
  const web3Signals = ['web3','blockchain','crypto','defi','nft','token','dao'];
  if (web3Signals.some(s => combined.includes(s))) return 'Web3 Marketing Manager';
  const deSignals = ['(m/w/d)','vollzeit','bewerbung','berufserfahrung','aufgaben'];
  if (deSignals.some(s => combined.includes(s))) return 'DE Web2 Marketing Manager';
  return 'EN Web2 Marketing Manager'; // default for international
}


// ── Job_Search_Cache tab ──────────────────────────────────────────────────────

function getOrCreateJobCacheSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Job_Search_Cache');
  if (!sheet) {
    sheet = ss.insertSheet('Job_Search_Cache');
    const headers = ['Date', 'Company', 'Job_Title', 'URL', 'CV_Type', 'Match_Level', 'Score', 'Fetch_Source'];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setBackground('#34a853')
         .setFontColor('white')
         .setFontWeight('bold');
    sheet.setFrozenRows(1);
    for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
    Logger.log('Job_Search_Cache sheet created.');
  }
  return sheet;
}

function cleanJobCache() {
  const sheet = getOrCreateJobCacheSheet();
  if (sheet.getLastRow() < 2) return;

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);

  const dates    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const toDelete = [];
  dates.forEach((row, i) => {
    if (row[0] && new Date(row[0]) < cutoff) toDelete.push(i + 2);
  });

  for (let i = toDelete.length - 1; i >= 0; i--) sheet.deleteRow(toDelete[i]);

  if (toDelete.length > 0) {
    SpreadsheetApp.flush();
    Logger.log(`Job cache: removed ${toDelete.length} entries older than 30 days.`);
  }
}

function isJobInCache(company, title) {
  const sheet = getOrCreateJobCacheSheet();
  if (sheet.getLastRow() < 2) return false;
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
  const nc   = company.toLowerCase().trim();
  const nt   = title.toLowerCase().trim();
  return data.some(r =>
    r[0].toString().toLowerCase().trim() === nc &&
    r[1].toString().toLowerCase().trim() === nt
  );
}

function isJobAlreadyApplied(company) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nc = company.toLowerCase().trim();
  const skipSheets = new Set(['Sankey_Data', 'Geo_Data', 'SMM_Raw_Data', 'Job_Search_Cache']);

  for (const sheet of ss.getSheets()) {
    const name = sheet.getName();
    if (skipSheets.has(name) || !/\d{4}/.test(name)) continue;
    if (sheet.getLastRow() < 2) continue;
    const companies = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    if (companies.some(r => r[0].toString().toLowerCase().trim() === nc)) return true;
  }
  return false;
}

function addJobToCache(job, smmResult, cvType, fetchSource) {
  const sheet   = getOrCreateJobCacheSheet();
  const dateStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  sheet.appendRow([
    dateStr,
    job.company  || '',
    job.title    || '',
    job.url      || '',
    cvType       || '',
    smmResult.match_level  || 'M0',
    smmResult.total_score  || 0,
    fetchSource  || ''
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


// ── Main orchestrator ─────────────────────────────────────────────────────────

// ── REPLACE runDailyJobSearch() ──────────────────────────────────────────────
// Changes: added diag counter object; increment at each gate;
// call emitDiagnosticsSummary() at end; call sendDebugDiagnosticsEmail()
// when no M2+ jobs found. All filtering logic identical.
 
function runDailyJobSearch() {
  const MAX_JOBS = 6; // unchanged
 
  // ── diagnostics counter object ────────────────────────────────────────────
  const diag = {
    fetched_total:      0,
    excluded_title:     0,
    excluded_geo:       0,
    excluded_cache:     0,
    excluded_applied:   0,
    candidate_selected: 0,
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
  // ── end diagnostics setup ─────────────────────────────────────────────────
 
  Logger.log('=== JABA Daily Job Search started ===');
  diag.tavily_start = getTavilyMonthlyUsage(); // ← record start credits
  Logger.log(`Tavily credits at start of run: ${diag.tavily_start}/1000`);
 
  // Step 1 — clean cache entries older than 30 days
  cleanJobCache();
 
  // Step 2 — fetch all jobs (deduplicated by job ID)
  const allJobs = fetchAllJobSources();
  diag.fetched_total = allJobs.length; // ← total from all sources
 
  if (allJobs.length === 0) {
    Logger.log('No jobs returned today. Exiting.');
    diag.tavily_end = getTavilyMonthlyUsage();
    emitDiagnosticsSummary(diag);
    return;
  }
 
  // Step 3 — apply filters
  const candidates = [];
  for (const job of allJobs) {
    if (!isRelevantJobTitle(job.title)) {
      Logger.log(`  ✗ [title] "${job.title}"`);
      diag.excluded_title++; // ← title gate
      continue;
    }
    if (!job.descriptionFull) {
      const inGermany   = isGermanLocation(job.city || '');
      const titleLower  = job.title.toLowerCase();
      const descLower   = (job.description || '').toLowerCase();
      const mentionsRemote = titleLower.includes('remote') ||
                             descLower.includes('remote')  ||
                             descLower.includes('homeoffice') ||
                             descLower.includes('home office');
      if (!inGermany && !mentionsRemote) {
        Logger.log(`  ✗ [geo] "${job.title}" @ ${job.city}`);
        diag.excluded_geo++; // ← geo gate
        continue;
      }
    }
    if (isJobInCache(job.company, job.title)) {
      Logger.log(`  ✗ [cache] "${job.title}" @ ${job.company}`);
      diag.excluded_cache++; // ← cache gate
      continue;
    }
    if (isJobAlreadyApplied(job.company)) {
      Logger.log(`  ✗ [applied] ${job.company}`);
      diag.excluded_applied++; // ← applied gate
      continue;
    }
    candidates.push(job);
    if (candidates.length >= MAX_JOBS * 2) break;
  }
 
  diag.candidate_selected = candidates.length; // ← candidates that passed all gates
  Logger.log(`Candidates after filtering: ${candidates.length}`);
 
  if (candidates.length === 0) {
    Logger.log('No new candidates today — no email sent.');
    diag.tavily_end = getTavilyMonthlyUsage();
    emitDiagnosticsSummary(diag);
    sendDebugDiagnosticsEmail(diag); // ← debug email even when candidates = 0
    return;
  }
 
  // Step 4 — extract JD, run SMM, collect M2+ results
  const reportJobs = [];
  let   processed  = 0;
 
  for (const job of candidates) {
    if (processed >= MAX_JOBS) break;
 
    Logger.log(`\n[${processed + 1}/${MAX_JOBS}] "${job.title}" — ${job.company}`);
 
    const extracted = smartExtractJD(job.url, job.description, job.descriptionFull);
    if (!extracted) {
      Logger.log(`  ✗ JD fetch failed — skipping`);
      diag.jd_fetch_failed++; // ← JD fetch gate
      addJobToCache(job, { match_level: 'SKIP', total_score: 0 }, 'unknown', 'failed');
      processed++;
      Utilities.sleep(500);
      continue;
    }
 
    if (!isJdRelevantToJob(extracted.text, job.company, job.title)) {
      Logger.log(`  ✗ JD failed relevance check for "${job.company}" — skipping`);
      diag.jd_irrelevant++; // ← relevance gate
      addJobToCache(job, { match_level: 'SKIP', total_score: 0 }, 'unknown', 'irrelevant_jd');
      processed++;
      Utilities.sleep(500);
      continue;
    }
 
    const cvType = detectCvTypeForSearch(extracted.text, job.title);
    Logger.log(`  CV type: ${cvType}`);
 
    let smmResult;
    try {
      const raw = analyzeSkillsMatch(extracted.text, cvType);
      smmResult = JSON.parse(raw);
      if (smmResult.error) throw new Error(smmResult.error);
    } catch (e) {
      Logger.log(`  ✗ SMM error: ${e.message}`);
      diag.smm_failed++; // ← SMM gate
      processed++;
      Utilities.sleep(3000);
      continue;
    }
 
    const score    = smmResult.total_score || 0;
    const level    = smmResult.match_level  || 'M0';
    const levelNum = parseInt(level.replace(/\D/g, '')) || 0;
 
    // ── score bucket ────────────────────────────────────────────────────────
    if      (levelNum === 0) diag.scored_m0++;
    else if (levelNum === 1) diag.scored_m1++;
    else                     diag.scored_m2_plus++;
    // ── end score bucket ────────────────────────────────────────────────────
 
    Logger.log(`  Score: ${score}/40 | ${level} | Source: ${extracted.source}`);
 
    addJobToCache(job, smmResult, cvType, extracted.source);
 
    if (levelNum >= 2) {
      reportJobs.push({ ...job, smmResult, cvType });
      Logger.log(`  ✓ Added to report (${level})`);
    }
 
    processed++;
    Utilities.sleep(1500);
  }
 
  diag.processed_count   = processed;
  diag.report_jobs_count = reportJobs.length;
  diag.tavily_end        = getTavilyMonthlyUsage(); // ← record end credits
 
  // ── emit structured diagnostics to log ───────────────────────────────────
  emitDiagnosticsSummary(diag);
 
  // Step 5 — send email
  if (reportJobs.length > 0) {
    const html = buildJobReportHtml(reportJobs);
    sendJobReportEmail(html);
    Logger.log(`Report email sent with ${reportJobs.length} job(s).`);
  } else {
    Logger.log('No M2+ jobs found today — no email sent.');
    sendDebugDiagnosticsEmail(diag); // ← debug email with full breakdown
  }
 
  Logger.log('=== JABA Daily Job Search complete ===');
}