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
    .addItem('3. Setup Interview Q&A Tab', 'setupQATab')
    .addSeparator()
    .addItem('⚡ Generate Cover Letter', 'generateCoverLetter')
    .addItem('📝 Answer App Questions', 'generateQuestionAnswers')
    .addSeparator()
    .addItem('🧹 Clear Current Inputs', 'clearInputTab')
    .addToUi();
}

function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let appSheet = ss.getSheetByName("Applications");
  if (!appSheet) {
    appSheet = ss.insertSheet("Applications");
    const appHeaders = ["Date", "Role / Job Title", "Company", "Entry Date", "Employment Type", "Work Mode", "Source", "Status", "GmailID", "GmailLink", "Notes"];
    appSheet.appendRow(appHeaders);
    appSheet.getRange("A1:K1").setFontWeight("bold");
  }
  
  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    const defaultSettings = [
      ["Setting", "Value"],
      ["Gmail Search Query", 'subject:("your application was sent" OR "thank you for applying" OR "application received" OR "ansökan mottagen" OR "we received your application" OR "application confirmed" OR "application for") -subject:("job alert" OR "jobs similar to" OR "hiring for" OR "new jobs" OR "discover the" OR "spotlight on") -from:(alerts@) -label:Tracker-Logged'],
      ["Default Status", "Applied"],
      ["Target Sheet Name", "Applications"],
      ["Max Emails Per Run", "50"],
      ["Cover Letter Template URL", "PASTE_YOUR_BASE_COVER_LETTER_DOC_LINK_HERE"],
      ["Cover Letters Folder ID", "PASTE_YOUR_DRIVE_FOLDER_ID_HERE"],
      ["CV Doc URL", "PASTE_YOUR_CV_DOC_LINK_HERE"] 
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
      ["Job Description", "Paste full job description text below this cell\n⬇️⬇️⬇️"]
    ];
    inputSheet.getRange("A1:B5").setValues(labels);
    inputSheet.getRange("A1:A5").setFontWeight("bold");
    inputSheet.setColumnWidth(1, 200);
    inputSheet.setColumnWidth(2, 600);
    inputSheet.getRange("B5").setWrap(true);
    
    // Clean Job Description Box (No more Q&A boxes below this)
    inputSheet.getRange("A6:B30").merge();
    inputSheet.getRange("A6").setVerticalAlignment("top").setWrap(true);
    
    SpreadsheetApp.getUi().alert("Cover Letter Input tab created cleanly.");
  } else {
    SpreadsheetApp.getUi().alert("Cover Letter Input tab already exists!");
  }
}

function setupQATab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let qaSheet = ss.getSheetByName("Interview Q&A");
  
  if (!qaSheet) {
    qaSheet = ss.insertSheet("Interview Q&A", 1);
    const headers = ["Application Question", "AI Answer", "Feedback / Tweak (Optional)"];
    qaSheet.appendRow(headers);
    qaSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
    qaSheet.setColumnWidth(1, 350);
    qaSheet.setColumnWidth(2, 500);
    qaSheet.setColumnWidth(3, 250);
    qaSheet.setFrozenRows(1);
    qaSheet.getRange("A:C").setWrap(true).setVerticalAlignment("top");
    
    SpreadsheetApp.getUi().alert("Interview Q&A tab created! Paste one question per row in Column A.");
  } else {
    SpreadsheetApp.getUi().alert("Interview Q&A tab already exists!");
  }
}

function clearInputTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear Cover Letter Data
  const inputSheet = ss.getSheetByName("Cover Letter Input");
  if (inputSheet) {
    inputSheet.getRange("B1:B4").clearContent();
    inputSheet.getRange("A6").clearContent();
  }
  
  // Clear Q&A Data
  const qaSheet = ss.getSheetByName("Interview Q&A");
  if (qaSheet && qaSheet.getLastRow() > 1) {
    qaSheet.getRange(2, 1, qaSheet.getLastRow() - 1, 3).clearContent();
  }
  
  ss.toast("All input fields cleared!", "Ready", 3);
}

/**
 * =========================================================================
 * 2. BACKGROUND GMAIL TRACKER & AI EXTRACTOR
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

  let trackingLabel = GmailApp.getUserLabelByName("Tracker-Logged");
  if (!trackingLabel) {
    trackingLabel = GmailApp.createLabel("Tracker-Logged");
  }

  const threads = GmailApp.search(query, 0, maxThreads);
  let rowsToAdd = [];
  
  for (let i = threads.length - 1; i >= 0; i--) { 
    if (Date.now() - startTime > 300000) {
      logDebug("Sync", "Timeout", "Approaching 5-minute limit. Stopping early.");
      break; 
    }

    const thread = threads[i];
    const firstMessage = thread.getMessages()[0];
    const msgId = firstMessage.getId();
    
    const subject = firstMessage.getSubject();
    const body = firstMessage.getPlainBody();
    
    const aiData = extractJobDataWithGemini(subject, body) || {};
    
    let newRow = new Array(headers.length).fill("");
    if (colMap["Date"] !== undefined) newRow[colMap["Date"]] = ""; 
    if (colMap["Entry Date"] !== undefined) newRow[colMap["Entry Date"]] = firstMessage.getDate();
    
    if (colMap["Role / Job Title"] !== undefined) newRow[colMap["Role / Job Title"]] = aiData.jobTitle || "Review Manually";
    if (colMap["Company"] !== undefined) newRow[colMap["Company"]] = aiData.company || "Review Manually";
    if (colMap["Employment Type"] !== undefined) newRow[colMap["Employment Type"]] = aiData.employmentType || "Unknown";
    if (colMap["Work Mode"] !== undefined) newRow[colMap["Work Mode"]] = aiData.workMode || "Unknown";
    if (colMap["Source"] !== undefined) newRow[colMap["Source"]] = aiData.source || "Unknown";
    if (colMap["Status"] !== undefined) newRow[colMap["Status"]] = aiData.status || defaultStatus;
    if (colMap["GmailID"] !== undefined) newRow[colMap["GmailID"]] = msgId;
    if (colMap["GmailLink"] !== undefined) newRow[colMap["GmailLink"]] = thread.getPermalink();

    rowsToAdd.push(newRow);
    thread.addLabel(trackingLabel);
    Utilities.sleep(500); 
  }

  if (rowsToAdd.length > 0) {
    appSheet.getRange(appSheet.getLastRow() + 1, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
    logDebug("Sync", "Success", `Synced ${rowsToAdd.length} jobs.`);
    ss.toast(`Successfully synced ${rowsToAdd.length} new jobs.`, "Sync Complete", 5);
  } else {
    logDebug("Sync", "Success", "No new jobs found.");
    ss.toast("Tracker is up to date! No new jobs found.", "Sync Complete", 3);
  }
}

function extractJobDataWithGemini(emailSubject, emailBody) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  
  const prompt = `
    You are an expert data extraction bot for a job application tracker. 
    Read the following job application email subject and body, and extract the details.
    Return ONLY a raw, valid JSON object with the exact keys below. Do not include markdown formatting or backticks.
    If a value is not explicitly stated, use "Unknown".

    {
      "jobTitle": "String (Extract the specific role)",
      "company": "String (Extract the hiring company)",
      "employmentType": "Full-time, Part-time, Contract/Freelance, or Unknown",
      "workMode": "Remote, Hybrid, On-site, or Unknown",
      "source": "LinkedIn, Teamtailor, Greenhouse, SmartRecruiters, Direct, or Other",
      "status": "Applied, Assessment Received, Interview Requested, Rejected, or Offer"
    }

    Subject: ${emailSubject}
    Body: ${emailBody}
  `;

  const payload = { "contents": [{ "parts": [{"text": prompt}] }] };
  const options = { 
    "method": "post", 
    "contentType": "application/json", 
    "payload": JSON.stringify(payload), 
    "muteHttpExceptions": true 
  };

  try {
    const response = fetchWithBackoff(apiUrl, options);
    const data = JSON.parse(response.getContentText());
    if (data.error) throw new Error(data.error.message);
    
    const rawJsonString = data.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
    return JSON.parse(rawJsonString);
  } catch (e) {
    logDebug("AI Parsing Error", "Failed", e.toString());
    return null; 
  }
}

/**
 * =========================================================================
 * 3. AI COVER LETTER GENERATOR 
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
  const cvDocUrl = settingsSheet.getRange("B8").getValue(); 
  
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
  const jobDescription = inputSheet.getRange("A6").getValue(); // Reverted to text input

  if (!companyName || !jobDescription) {
    ss.toast("Please provide a Company Name and Job Description.", "Missing Info", 4);
    return;
  }

  ss.toast("Reading template & CV, then calling Gemini...", "Processing", 5);

  let templateText = "";
  try {
    const templateDoc = DocumentApp.openById(templateIdMatch[0]);
    templateText = templateDoc.getBody().getText();
  } catch (e) {
    logDebug("Template Read", "Error", e.toString());
    ss.toast("Could not read template document. Check settings.", "Error", 5);
    return;
  }

  let cvText = "";
  try {
    if (cvDocUrl) {
      const cvDocIdMatch = cvDocUrl.match(/[-\w]{25,}/);
      if (cvDocIdMatch) {
        cvText = DocumentApp.openById(cvDocIdMatch[0]).getBody().getText();
      }
    }
  } catch (e) {
    logDebug("CV Read", "Warning", "Could not load CV Doc.");
  }

  const prompt = `
    You are an expert copywriter and career coach. Write the body paragraphs of a highly tailored cover letter for me.

    CRITICAL RULES:
    - UK English.
    - Write in a natural, engaging, and professional tone. DO NOT sound like a robotic AI.
    - DO NOT use subheadings or bullet points. Write 2 to 3 cohesive, flowing paragraphs.
    - You MUST explicitly mention the company (${companyName}) and the role (${jobTitle}) in the text.
    - STRICT FOCUS: Do NOT assume or make up any experiences, tools, or metrics not explicitly stated in the CV context. Rely ONLY on the provided background details.
    - Connect the dots: Explain exactly how my specific background solves the specific problems listed in the Job Description.
    - Keep it around 180-250 words total.

    HERE IS MY CV CONTEXT (Use this to pull specific facts and skills):
    ${cvText ? cvText : "(No CV context provided)"}

    TEMPLATE TO ADAPT (Maintain this structure):
    ${templateText}

    INPUT DETAILS:
    Role: ${jobTitle}
    Company: ${companyName}
    Location/Mode: ${location}

    JOB DESCRIPTION:
    ${jobDescription}

    OUTPUT INSTRUCTIONS:
    Return ONLY the tailored body paragraphs in plain text. Do not include the "Hi" intro or the footer.
  `;

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const payload = { "contents": [{ "parts": [{"text": prompt}] }] };
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };

  let tailoredText = "";
  try {
    const response = fetchWithBackoff(apiUrl, options);
    const data = JSON.parse(response.getContentText());
    if (data.error) throw new Error(data.error.message);
    tailoredText = data.candidates[0].content.parts[0].text.trim();
  } catch (e) {
    logDebug("Generate Cover Letter", "Error", e.toString());
    ss.toast("Failed to connect to AI. Check Debug Logs.", "Error", 5);
    return;
  }

  const cleanCompanyName = companyName.toString().replace(/[^a-zA-Z0-9]/g, "_");
  const fileName = `My_Cover_Letter_${cleanCompanyName}`;
  
  try {
    const templateFile = DriveApp.getFileById(templateIdMatch[0]);
    const outputFolder = DriveApp.getFolderById(folderId);
    
    const newFile = templateFile.makeCopy(fileName, outputFolder);
    const newDoc = DocumentApp.openById(newFile.getId());
    const body = newDoc.getBody();
    
    if (!tailoredText || tailoredText.trim() === "") {
      tailoredText = "Error: The AI returned an empty response. Please check your prompt or job description.";
    }
    
    body.replaceText("{{AI_COVER_LETTER}}", tailoredText);

    if (location.toLowerCase().includes("remote")) {
      body.replaceText("I can work from the office and I am available to start immediately.", "I can work remotely and I am available to start immediately.");
    }
    
    newDoc.saveAndClose();
    ss.toast(`Cover letter created: ${fileName}`, "Success!", 5);
    logDebug("Generate Cover Letter", "Success", `Generated for ${companyName}`);
  } catch (e) {
    logDebug("Document Creation", "Error", e.toString());
    ss.toast("Failed to save Document. Check Debug Logs.", "Error", 5);
  }
}

/**
 * =========================================================================
 * 4. DYNAMIC TWO-COLUMN Q&A ANSWERER
 * =========================================================================
 */

function generateQuestionAnswers() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const qaSheet = ss.getSheetByName("Interview Q&A");
  const inputSheet = ss.getSheetByName("Cover Letter Input");
  const settingsSheet = ss.getSheetByName("Settings");
  
  if (!qaSheet) {
    ss.toast("Missing Interview Q&A tab. Please run 'Setup Interview Q&A Tab' from the menu.", "Error", 5);
    return;
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    ui.alert("Missing API Key! Please add GEMINI_API_KEY to your Script Properties.");
    return;
  }

  // Gather Job Context
  const companyName = inputSheet ? inputSheet.getRange("B1").getValue() : "Unknown Company";
  const jobTitle = inputSheet ? inputSheet.getRange("B2").getValue() : "Unknown Role";
  const jobDescription = inputSheet ? inputSheet.getRange("A6").getValue() : "";
  const cvDocUrl = settingsSheet ? settingsSheet.getRange("B8").getValue() : ""; 

  // Gather Questions from Q&A Tab
  const lastRow = qaSheet.getLastRow();
  if (lastRow < 2) {
    ss.toast("Please paste your questions into Column A of the Interview Q&A tab.", "No Questions", 4);
    return;
  }

  const dataRange = qaSheet.getRange(2, 1, lastRow - 1, 3);
  const rows = dataRange.getValues();
  let questionsToProcess = [];

  for (let i = 0; i < rows.length; i++) {
    const question = rows[i][0];
    const feedback = rows[i][2]; // User can leave notes here to guide the AI
    
    if (question && question.trim() !== "") {
      questionsToProcess.push({
        rowIndex: i + 2, // Keep track of the exact row number
        questionText: question,
        userFeedback: feedback
      });
    }
  }

  if (questionsToProcess.length === 0) {
    ss.toast("No valid questions found in Column A.", "No Questions", 4);
    return;
  }

  ss.toast(`Analyzing ${questionsToProcess.length} questions...`, "AI Working", 5);

  // Fetch CV Context 
  let cvText = "";
  try {
    if (cvDocUrl) {
      const cvDocIdMatch = cvDocUrl.match(/[-\w]{25,}/);
      if (cvDocIdMatch) {
        cvText = DocumentApp.openById(cvDocIdMatch[0]).getBody().getText();
      }
    }
  } catch (e) {
    logDebug("CV Read", "Warning", "Could not load CV Doc for Q&A.");
  }

  const prompt = `
    You are an expert career coach. I am applying for the role of ${jobTitle} at ${companyName}.
    I have provided a JSON array of specific application questions.
    Please draft professional, concise, and high-impact answers for each question based on my background.

    CRITICAL RULES:
    - UK English.
    - STRICT FOCUS: Do NOT assume or make up any experiences, tools, or metrics not explicitly stated in the CV context. Rely ONLY on the provided background details.
    - If "userFeedback" is provided for a question, you MUST alter your answer to follow that feedback exactly (e.g., if feedback says "make it shorter", keep it under 50 words).
    - Return ONLY a raw, valid JSON array. Do not include markdown formatting or backticks.
    
    EXPECTED OUTPUT FORMAT:
    [
      { "rowIndex": [Keep the original rowIndex], "answer": "The tailored answer..." }
    ]

    MY CV BACKGROUND:
    ${cvText ? cvText : "Video Producer, expert in DaVinci Resolve, Python automation, GenAI workflows."}

    JOB DESCRIPTION FOR CONTEXT:
    ${jobDescription}

    QUESTIONS TO ANSWER:
    ${JSON.stringify(questionsToProcess)}
  `;

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const payload = { "contents": [{ "parts": [{"text": prompt}] }] };
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };

  try {
    const response = fetchWithBackoff(apiUrl, options);
    const data = JSON.parse(response.getContentText());
    if (data.error) throw new Error(data.error.message);
    
    // Clean markdown if Gemini accidentally includes it
    const rawJsonString = data.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
    const aiResponses = JSON.parse(rawJsonString);

    // Map answers back to the exact correct rows in Column B
    aiResponses.forEach(item => {
      qaSheet.getRange(item.rowIndex, 2).setValue(item.answer);
    });

    ss.toast("All answers successfully generated!", "Success", 3);
    logDebug("Q&A Generation", "Success", `Answered ${aiResponses.length} questions.`);
  } catch (e) {
    logDebug("Q&A Generation", "Error", e.toString());
    ss.toast("Failed to parse AI responses. Check Debug Logs.", "Error", 5);
  }
}

/**
 * =========================================================================
 * 5. UTILITIES
 * =========================================================================
 */

function logDebug(action, status, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Debug_Logs");
  
  if (!logSheet) {
    logSheet = ss.insertSheet("Debug_Logs");
    logSheet.appendRow(["Timestamp", "Action", "Status", "Details"]);
    logSheet.getRange("A1:D1").setFontWeight("bold");
    logSheet.hideSheet();
  }
  
  if (logSheet.getLastRow() > 500) {
    logSheet.deleteRows(2, 100);
  }
  
  logSheet.appendRow([new Date(), action, status, details]);
}

function fetchWithBackoff(url, options, maxRetries = 3) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 429 || response.getResponseCode() >= 500) {
         throw new Error("API Server Busy: HTTP " + response.getResponseCode());
      }
      return response;
    } catch (e) {
      if (i === maxRetries - 1) throw e; 
      const delay = Math.pow(2, i) * 1000 + Math.round(Math.random() * 1000); 
      Utilities.sleep(delay);
      logDebug("API Retry", "Warning", `Waiting ${delay}ms before retry. Error: ${e.toString()}`);
    }
  }
}
