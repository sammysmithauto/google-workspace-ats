/**
 * =========================================================================
 * 1. USER INTERFACE & SETUP
 * =========================================================================
 */

// Creates the custom menu when the spreadsheet opens
// Creates the custom menu when the spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Job Tools')
    .addItem('🔄 Sync Gmail Now', 'syncJobsFromGmail')
    .addSeparator()
    .addItem('1. Setup Tracker & Settings', 'setupTracker')
    .addItem('2. Setup Cover Letter Tab', 'setupCoverLetterTab')
    .addSeparator()
    .addItem('⚡ Generate Cover Letter', 'generateCoverLetter')
    .addToUi();
}

// Builds the "Applications" and "Settings" tabs
function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Setup Applications Sheet
  let appSheet = ss.getSheetByName("Applications");
  if (!appSheet) {
    appSheet = ss.insertSheet("Applications");
  }
  const appHeaders = ["Date", "Role / Job Title", "Company", "Entry Date", "Employment Type", "Work Mode", "Source", "Status", "GmailID", "GmailLink", "Notes"];
  if (appSheet.getLastRow() === 0) {
    appSheet.appendRow(appHeaders);
    appSheet.getRange("A1:K1").setFontWeight("bold");
  }
  
  // Setup Settings Sheet
  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    const defaultSettings = [
      ["Setting", "Value"],
      ["Gmail Search Query", '(label:jobs OR subject:(application OR applied OR "thank you for applying" OR ansökan OR ansökan mottagen OR "we received your application") OR from:(jobs@ OR careers@ OR no-reply@ OR teamtailor OR greenhouse OR lever OR workable OR smartrecruiters OR icims)) newer_than:365d'],
      ["Default Status", "Applied"],
      ["Target Sheet Name", "Applications"],
      ["Max Emails Per Run", "50"],
      ["Cover Letter Template URL", "PASTE_YOUR_BASE_COVER_LETTER_DOC_LINK_HERE"],
      ["Cover Letters Folder ID", "PASTE_YOUR_DRIVE_FOLDER_ID_HERE"]
    ];
    settingsSheet.getRange(1, 1, defaultSettings.length, 2).setValues(defaultSettings);
    settingsSheet.getRange("A1:B1").setFontWeight("bold");
    settingsSheet.setColumnWidth(1, 200);
    settingsSheet.setColumnWidth(2, 600);
  }
  SpreadsheetApp.getUi().alert("Setup Complete! Check the 'Settings' tab to add your URLs.");
}

// Builds the input interface for the Cover Letter generator
function setupCoverLetterTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let inputSheet = ss.getSheetByName("Cover Letter Input");
  
  if (!inputSheet) {
    inputSheet = ss.insertSheet("Cover Letter Input", 0); // Puts it at the front
    const labels = [
      ["Company Name", ""],
      ["Job Title", ""],
      ["Location / Work Mode", ""],
      ["Start Date (If known)", ""],
      ["Form Questions / Limits", ""],
      ["Job Description", "Paste full description below this cell\n⬇️⬇️⬇️"]
    ];
    inputSheet.getRange("A1:B6").setValues(labels);
    inputSheet.getRange("A1:A6").setFontWeight("bold");
    inputSheet.setColumnWidth(1, 200);
    inputSheet.setColumnWidth(2, 600);
    inputSheet.getRange("B6").setWrap(true);
    
    // Make a big merged space for the job description
    inputSheet.getRange("A7:B30").merge();
    inputSheet.getRange("A7").setVerticalAlignment("top").setWrap(true);
  } else {
    SpreadsheetApp.getUi().alert("Cover Letter Input tab already exists!");
  }
}

/**
 * =========================================================================
 * 2. BACKGROUND GMAIL TRACKER
 * =========================================================================
 */

/**
 * =========================================================================
 * 2. BACKGROUND GMAIL TRACKER (OPTIMIZED)
 * =========================================================================
 */

function syncJobsFromGmail() {
  const startTime = Date.now(); // Start the stopwatch!
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) return;

  let rawQuery = settingsSheet.getRange("B2").getValue();
  const query = rawQuery.replace(/newer_than:\w+/g, '') + ' newer_than:1m';
  
  const defaultStatus = settingsSheet.getRange("B3").getValue();
  const targetSheetName = settingsSheet.getRange("B4").getValue();
  const maxThreads = parseInt(settingsSheet.getRange("B5").getValue()) || 50;

  const appSheet = ss.getSheetByName(targetSheetName);
  if (!appSheet) return;

  const headers = appSheet.getRange(1, 1, 1, appSheet.getLastColumn()).getValues()[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);

  let existingIDs = new Set();
  if (colMap["GmailID"] !== undefined && appSheet.getLastRow() > 1) {
    let ids = appSheet.getRange(2, colMap["GmailID"] + 1, appSheet.getLastRow() - 1, 1).getValues().flat();
    existingIDs = new Set(ids);
  }

  const threads = GmailApp.search(query, 0, maxThreads);
  let rowsToAdd = [];
  
  for (let i = threads.length - 1; i >= 0; i--) { 
    // THE FAILSAFE: Check if we've been running for more than 5.5 minutes (330,000 milliseconds)
    if (Date.now() - startTime > 330000) {
      console.warn("Approaching 6-minute limit. Stopping early and saving current progress.");
      break; 
    }

    const thread = threads[i];
    const firstMessage = thread.getMessages()[0];
    const msgId = firstMessage.getId();
    
    if (existingIDs.has(msgId)) continue; 
    
    const subject = firstMessage.getSubject();
    const body = firstMessage.getPlainBody().toLowerCase();
    const sender = firstMessage.getFrom();
    
    let newRow = new Array(headers.length).fill("");
    if (colMap["Date"] !== undefined) newRow[colMap["Date"]] = ""; 
    if (colMap["Entry Date"] !== undefined) newRow[colMap["Entry Date"]] = firstMessage.getDate();
    
    if (colMap["Role / Job Title"] !== undefined) newRow[colMap["Role / Job Title"]] = extractRole(subject, body) || "Unknown";
    if (colMap["Company"] !== undefined) newRow[colMap["Company"]] = extractCompany(subject, sender, body) || "Unknown";
    if (colMap["Employment Type"] !== undefined) newRow[colMap["Employment Type"]] = determineEmploymentType(subject.toLowerCase() + " " + body);
    if (colMap["Work Mode"] !== undefined) newRow[colMap["Work Mode"]] = determineWorkMode(subject.toLowerCase() + " " + body);
    if (colMap["Source"] !== undefined) newRow[colMap["Source"]] = extractSource(sender);
    if (colMap["Status"] !== undefined) newRow[colMap["Status"]] = defaultStatus;
    if (colMap["GmailID"] !== undefined) newRow[colMap["GmailID"]] = msgId;
    if (colMap["GmailLink"] !== undefined) newRow[colMap["GmailLink"]] = thread.getPermalink();

    rowsToAdd.push(newRow);
  }

  if (rowsToAdd.length > 0) {
    appSheet.getRange(appSheet.getLastRow() + 1, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
    ss.toast(`Successfully synced ${rowsToAdd.length} new jobs.`, "Sync Complete", 5);
  } else {
    ss.toast("Tracker is up to date! No new jobs found.", "Sync Complete", 3);
  }
}

// Extraction Helpers

function cleanText(text) {
  if (!text) return "";
  return text.replace(/["“”*!]/g, '').trim(); // Strips quotes, asterisks, and exclamation marks
}

function extractRole(subject, body) {
  let sub = subject.replace(/["“”*!]/g, '').trim();
  
  // Split the email body into clean, individual lines
  let bodyLines = body.split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);

  // 1. Catch LinkedIn: "Your application was sent to..."
  if (/application was sent to/i.test(sub)) {
     for (let i = 0; i < bodyLines.length; i++) {
        if (/application was sent to/i.test(bodyLines[i])) {
           for (let j = i + 1; j < i + 5 && j < bodyLines.length; j++) {
              let candidate = bodyLines[j];
              // Increased safety limit to 120 to catch long corporate titles
              if (!/^http/i.test(candidate) && !/application was sent/i.test(candidate) && candidate.length < 120) {
                 return cleanText(candidate);
              }
           }
        }
     }
  }

  // 2. Catch Body text (Greenhouse/Lever/Amazon/Workday)
  if (/thank you for applying|application received|applying to amazon/i.test(sub)) {
     for (let line of bodyLines) {
        // Expanded phrasing to catch Amazon's exact wording
        let match = line.match(/(?:applying for the|applying to the|application for the|application for|position of|interest in the) (.*?)(?: at | role| position|\.|$)/i);
        // Increased limit to 120 characters
        if (match && match[1].length < 120) {
           let foundTitle = cleanText(match[1]);
           // Bonus: Strip out the ugly "(ID: 123456)" from the end of Amazon titles
           foundTitle = foundTitle.replace(/\s*\(ID:.*?\)/i, '');
           return foundTitle;
        }
     }
  }

  // 3. Standard Subject Match
  let match1 = sub.match(/(?:application for|applied for|ansökan till) (.*?)(?: at | på | - |$)/i);
  if (match1) return cleanText(match1[1]);

  // 4. Pattern: "Role": Company
  let match2 = sub.match(/^"?([^"]+)"?:\s*([^-]+)/);
  if (match2) return cleanText(match2[1]);

  // 5. Broad Fallback
  for (let line of bodyLines) {
     let fallbackMatch = line.match(/(?:role of|position of|applying for the|application for the) ([a-zA-Z0-9\s&,\-\.\/\(\)]+?)(?: position| role| at |\.|!|$)/i);
     if (fallbackMatch && fallbackMatch[1].length < 120) {
         let foundTitle = cleanText(fallbackMatch[1]);
         return foundTitle.replace(/\s*\(ID:.*?\)/i, '');
     }
  }

  // 6. Short subject fallback
  if (sub.length < 60 && !/(application|applied|thank you)/i.test(sub)) {
      return cleanText(sub);
  }

  return "Unknown";
}

function extractCompany(subject, sender, body) {
  let sub = subject.replace(/["“”*!]/g, '').trim();

  // 1. LinkedIn Pattern: "Sammy, your application was sent to [Company]"
  let m1 = sub.match(/application was sent to (.*?)$/i);
  if (m1) return cleanText(m1[1]);

  // 2. Greenhouse/Lever Pattern: "Thank you for applying to [Company]"
  let m2 = sub.match(/applying to (.*?)$/i);
  if (m2) return cleanText(m2[1]);

  // 3. Application to ...
  let m3 = sub.match(/(?:application|applied) to (.*?)(?: -|$)/i);
  if (m3) return cleanText(m3[1]);

  // 4. Standard "... at [Company]"
  let m4 = sub.match(/(?: at | på )([^-\(]+)/i);
  if (m4) return cleanText(m4[1]);

  // 5. Pattern: "Role": [Company]
  let m5 = sub.match(/^"?[^"]+"?:\s*([^-]+)/);
  if (m5) return cleanText(m5[1]);

  // Fallback: Try to get the sender's name if all else fails
  let senderNameMatch = sender.match(/^"?(.*?)"?\s*</);
  if (senderNameMatch) {
     let sName = senderNameMatch[1].trim();
     // Ensure we don't accidentally log "LinkedIn" or "Greenhouse" as the company
     if (!/teamtailor|greenhouse|lever|workable|smartrecruiters|icims|linkedin|glassdoor|alerts/i.test(sName) && sName.length > 2) {
        return cleanText(sName);
     }
  }

  return "Unknown";
}

function extractSource(sender) {
  const s = sender.toLowerCase();
  if (s.includes("linkedin")) return "LinkedIn";
  if (s.includes("teamtailor")) return "Teamtailor";
  if (s.includes("greenhouse")) return "Greenhouse";
  if (s.includes("lever")) return "Lever";
  if (s.includes("workable")) return "Workable";
  if (s.includes("smartrecruiters")) return "SmartRecruiters";
  if (s.includes("icims")) return "iCIMS";
  return "Other";
}

function determineEmploymentType(text) {
  if (/(full-time|permanent|heltid)/.test(text)) return "Full-time";
  if (/(part-time|deltid|30%|50%)/.test(text)) return "Part-time";
  if (/(contract|consultant|konsult|freelance|interim)/.test(text)) return "Contract/Freelance";
  return "Unknown";
}

function determineWorkMode(text) {
  if (/(remote|work from home|fjärrjobb)/.test(text)) return "Remote";
  if (/(hybrid|hybridarbete)/.test(text)) return "Hybrid";
  if (/(on-site|onsite|på plats)/.test(text)) return "On-site";
  return "Unknown";
}

// Manual Historical Cleanup Function
function fillDownDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Applications");
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateColIndex = headers.indexOf("Date") + 1;
  if (dateColIndex === 0) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, dateColIndex, lastRow - 1, 1);
  const values = range.getValues();
  let lastValidDate = "";
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) {
      lastValidDate = values[i][0];
    } else if (lastValidDate !== "") {
      values[i][0] = lastValidDate;
    }
  }
  range.setValues(values);
}

/**
 * =========================================================================
 * 3. AI COVER LETTER GENERATOR
 * =========================================================================
 */

function generateCoverLetter() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Check API Key
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    ui.alert("Missing API Key! Please add GEMINI_API_KEY to your Script Properties.");
    return;
  }

  // 2. Get Settings
  const settingsSheet = ss.getSheetByName("Settings");
  const templateUrl = settingsSheet.getRange("B6").getValue(); 
  const folderId = settingsSheet.getRange("B7").getValue(); 
  
  const templateIdMatch = templateUrl.match(/[-\w]{25,}/);
  if (!templateIdMatch) {
    ui.alert("Invalid Cover Letter Template URL in Settings B6.");
    return;
  }

  // 3. Get Input Data
  const inputSheet = ss.getSheetByName("Cover Letter Input");
  if (!inputSheet) {
    ui.alert("Missing 'Cover Letter Input' sheet. Run 'Setup Cover Letter Tab' from the menu.");
    return;
  }
  
  const companyName = inputSheet.getRange("B1").getValue();
  const jobTitle = inputSheet.getRange("B2").getValue();
  const location = inputSheet.getRange("B3").getValue();
  const startDate = inputSheet.getRange("B4").getValue();
  const formQuestions = inputSheet.getRange("B5").getValue();
  const jobDescription = inputSheet.getRange("A7").getValue();

  if (!companyName || !jobDescription) {
    ui.alert("Please provide at least a Company Name and Job Description.");
    return;
  }

  ss.toast("Reading template and calling Gemini...", "Processing", 5);

  // 4. Read Template
  let templateText = "";
  try {
    const templateDoc = DocumentApp.openById(templateIdMatch[0]);
    templateText = templateDoc.getBody().getText();
  } catch (e) {
    ui.alert("Error reading template. Ensure the script has access to the Doc.");
    return;
  }

  // 5. Construct Strict Prompt
  const prompt = `
    You are an expert career coach writing a tailored cover letter for me.
    Use my cover letter template as the ONLY base. Tailor the cover letter to the job provided by adjusting:
    - The role title and company name
    - 3-5 subheadings to match the job requirements
    - The closing line (start date / location / on-site or remote)

    RULES:
    - UK English
    - Clear, concise, professional
    - No jargon, no fluff, no invented details
    - Use "I am" style wording
    - Keep the same structure as the template: short intro + Title Case subheadings + Tools + closing
    - Do NOT reference my CV
    - Do NOT add claims, metrics, clients, or tools that are not already in the template
    - Keep it around 180-250 words unless the job asks for more

    REQUIRED OPENING (Keep exactly as template):
    Keep my name and contact block as-is.
    Start with:
    Hi,
    My name is Sammy and I’m a Stockholm-based Video Producer and Creative Developer, originally from London.

    TEMPLATE:
    ${templateText}

    INPUT DETAILS:
    Role: ${jobTitle}
    Company: ${companyName}
    Location/Mode: ${location}
    Start Date: ${startDate}
    Form Questions: ${formQuestions}

    JOB DESCRIPTION:
    ${jobDescription}

    OUTPUT INSTRUCTIONS:
    Return ONLY the final, formatted text of the tailored cover letter.
    If there are form questions listed in the input details, provide short, professional answers to them two lines below the end of the cover letter.
  `;

  // 6. Call Gemini API
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const payload = { "contents": [{ "parts": [{"text": prompt}] }] };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  let tailoredText = "";
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const data = JSON.parse(response.getContentText());
    if (data.error) throw new Error(data.error.message);
    tailoredText = data.candidates[0].content.parts[0].text;
  } catch (e) {
    ui.alert("Error with AI: " + e.toString());
    return;
  }

  // 7. Create the Document
  const cleanCompanyName = companyName.toString().replace(/[^a-zA-Z0-9]/g, "_");
  const fileName = `Sammy_Smith_Cover_Letter_${cleanCompanyName}`;
  
  try {
    const outputFolder = DriveApp.getFolderById(folderId);
    const newDoc = DocumentApp.create(fileName);
    
    // Safety check just in case the AI returns empty text
    if (!tailoredText || tailoredText.trim() === "") {
      tailoredText = "Error: The AI returned an empty response. Please check your prompt or job description.";
    }
    
    // Write the text
    newDoc.getBody().setText(tailoredText);
    
    // CRITICAL FIX: Force the document to save the text before moving it!
    newDoc.saveAndClose();
    
    // Move to folder
    DriveApp.getFileById(newDoc.getId()).moveTo(outputFolder);

    ui.alert("Success!", `Cover letter created: ${fileName}\nSaved to your specified folder.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Document creation failed.\nError: " + e.toString());
  }
}
