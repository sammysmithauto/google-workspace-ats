/**
 * =========================================================================
 * 1. USER INTERFACE & SETUP
 * =========================================================================
 */

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

function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let appSheet = ss.getSheetByName("Applications");
  if (!appSheet) {
    appSheet = ss.insertSheet("Applications");
  }
  const appHeaders = ["Date", "Role / Job Title", "Company", "Entry Date", "Employment Type", "Work Mode", "Source", "Status", "GmailID", "GmailLink", "Notes"];
  if (appSheet.getLastRow() === 0) {
    appSheet.appendRow(appHeaders);
    appSheet.getRange("A1:K1").setFontWeight("bold");
  }
  
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

function setupCoverLetterTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let inputSheet = ss.getSheetByName("Cover Letter Input");
  
  if (!inputSheet) {
    inputSheet = ss.insertSheet("Cover Letter Input", 0); 
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
    
    inputSheet.getRange("A7:B30").merge();
    inputSheet.getRange("A7").setVerticalAlignment("top").setWrap(true);
  } else {
    SpreadsheetApp.getUi().alert("Cover Letter Input tab already exists!");
  }
}

/**
 * =========================================================================
 * 2. BACKGROUND GMAIL TRACKER (OPTIMIZED)
 * =========================================================================
 */

function syncJobsFromGmail() {
  const startTime = Date.now(); 
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

function cleanText(text) {
  if (!text) return "";
  return text.replace(/["“”*!]/g, '').trim(); 
}

function extractRole(subject, body) {
  let sub = subject.replace(/["“”*!]/g, '').trim();
  let bodyLines = body.split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);

  if (/application was sent to/i.test(sub)) {
     for (let i = 0; i < bodyLines.length; i++) {
        if (/application was sent to/i.test(bodyLines[i])) {
           for (let j = i + 1; j < i + 5 && j < bodyLines.length; j++) {
              let candidate = bodyLines[j];
              if (!/^http/i.test(candidate) && !/application was sent/i.test(candidate) && candidate.length < 120) {
                 return cleanText(candidate);
              }
           }
        }
     }
  }

  if (/thank you for applying|application received|applying to amazon/i.test(sub)) {
     for (let line of bodyLines) {
        let match = line.match(/(?:applying for the|applying to the|application for the|application for|position of|interest in the) (.*?)(?: at | role| position|\.|$)/i);
        if (match && match[1].length < 120) {
           let foundTitle = cleanText(match[1]);
           foundTitle = foundTitle.replace(/\s*\(ID:.*?\)/i, '');
           return foundTitle;
        }
     }
  }

  let match1 = sub.match(/(?:application for|applied for|ansökan till) (.*?)(?: at | på | - |$)/i);
  if (match1) return cleanText(match1[1]);

  let match2 = sub.match(/^"?([^"]+)"?:\s*([^-]+)/);
  if (match2) return cleanText(match2[1]);

  for (let line of bodyLines) {
     let fallbackMatch = line.match(/(?:role of|position of|applying for the|application for the) ([a-zA-Z0-9\s&,\-\.\/\(\)]+?)(?: position| role| at |\.|!|$)/i);
     if (fallbackMatch && fallbackMatch[1].length < 120) {
         let foundTitle = cleanText(fallbackMatch[1]);
         return foundTitle.replace(/\s*\(ID:.*?\)/i, '');
     }
  }

  if (sub.length < 60 && !/(application|applied|thank you)/i.test(sub)) {
      return cleanText(sub);
  }

  return "Unknown";
}

function extractCompany(subject, sender, body) {
  let sub = subject.replace(/["“”*!]/g, '').trim();

  let m1 = sub.match(/application was sent to (.*?)$/i);
  if (m1) return cleanText(m1[1]);

  let m2 = sub.match(/applying to (.*?)$/i);
  if (m2) return cleanText(m2[1]);

  let m3 = sub.match(/(?:application|applied) to (.*?)(?: -|$)/i);
  if (m3) return cleanText(m3[1]);

  let m4 = sub.match(/(?: at | på )([^-\(]+)/i);
  if (m4) return cleanText(m4[1]);

  let m5 = sub.match(/^"?[^"]+"?:\s*([^-]+)/);
  if (m5) return cleanText(m5[1]);

  let senderNameMatch = sender.match(/^"?(.*?)"?\s*</);
  if (senderNameMatch) {
     let sName = senderNameMatch[1].trim();
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

/**
 * =========================================================================
 * 3. AI COVER LETTER GENERATOR (STRICT TEMPLATE METHOD)
 * =========================================================================
 */

function generateCoverLetter() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    ui.alert("Missing API Key! Please add GEMINI_API_KEY to your Script Properties.");
    return;
  }

  const settingsSheet = ss.getSheetByName("Settings");
  const templateUrl = settingsSheet.getRange("B6").getValue(); 
  const folderId = settingsSheet.getRange("B7").getValue(); 
  
  const templateIdMatch = templateUrl.match(/[-\w]{25,}/);
  if (!templateIdMatch) {
    ui.alert("Invalid Cover Letter Template URL in Settings B6.");
    return;
  }

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

  ss.toast("Reading strict template and calling Gemini...", "Processing", 5);

  let templateText = "";
  try {
    const templateDoc = DocumentApp.openById(templateIdMatch[0]);
    templateText = templateDoc.getBody().getText();
  } catch (e) {
    ui.alert("Error reading template. Ensure the script has access to the Doc.");
    return;
  }

  // THE STRICT PROMPT
  const prompt = `
    You are an expert career coach writing the body of a tailored cover letter for me.
    Use my Cover Letter Template as the ONLY base. Tailor the middle paragraphs to fit the new role.

    CRITICAL RULES:
    - UK English
    - Clear, concise, professional
    - No jargon, no fluff, no invented details
    - Use "I am" style wording
    - Keep the same structure as the template: Title Case subheadings followed by short paragraphs.
    - Do NOT reference a CV.
    - Do NOT add claims, metrics, clients, titles, or tools that are not already explicitly in the template.
    - If I lack direct experience for a specific requirement in the job description, DO NOT fabricate it. Instead, map my existing transferable skills (like process optimization or workflow reliability) to the requirement, and express a strong, enthusiastic intent to grow my career in that specific direction (e.g., learning new platforms).
    - Keep it around 180-250 words.

    TEMPLATE TO ADAPT:
    ${templateText}

    INPUT DETAILS:
    Role: ${jobTitle}
    Company: ${companyName}
    Location/Mode: ${location}
    Form Questions: ${formQuestions}

    JOB DESCRIPTION:
    ${jobDescription}

    OUTPUT INSTRUCTIONS:
    Return ONLY the body paragraphs and subheadings in plain text. Do not include the "Hi" intro or the "Tools" footer, as that is already handled in the document.
    If there are form questions, answer them two lines below the end of the text.
  `;

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
    tailoredText = data.candidates[0].content.parts[0].text.trim();
  } catch (e) {
    ui.alert("Error with AI: " + e.toString());
    return;
  }

  const cleanCompanyName = companyName.toString().replace(/[^a-zA-Z0-9]/g, "_");
  const fileName = `Sammy_Smith_Cover_Letter_${cleanCompanyName}`;
  
  try {
    const templateFile = DriveApp.getFileById(templateIdMatch[0]);
    const outputFolder = DriveApp.getFolderById(folderId);
    
    // Create Document
    const newFile = templateFile.makeCopy(fileName, outputFolder);
    const newDoc = DocumentApp.openById(newFile.getId());
    const body = newDoc.getBody();
    
    if (!tailoredText || tailoredText.trim() === "") {
      tailoredText = "Error: The AI returned an empty response. Please check your prompt or job description.";
    }
    
    // Inject the AI text perfectly into the tag
    body.replaceText("{{AI_COVER_LETTER}}", tailoredText);

    // SMART LOCATION SWAP: Automatically update the closing line if the job is remote
    if (location.toLowerCase().includes("remote")) {
      body.replaceText("I can work from the office and I am available to start immediately.", "I can work remotely and I am available to start immediately.");
    }
    
    newDoc.saveAndClose();
    ui.alert("Success!", `Cover letter created: ${fileName}\nSaved to your specified folder.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Document creation failed.\nError: " + e.toString());
  }
}
