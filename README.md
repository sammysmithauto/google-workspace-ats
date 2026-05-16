# 🚀 Automated ATS & AI Cover Letter Generator (Google Workspace)

A fully automated, private Applicant Tracking System (ATS) and AI Application Command Center that lives entirely inside your personal Google Workspace.

Tired of manually logging jobs, rewriting cover letters, and struggling over custom application questions? This tool runs in the background to automatically scan your Gmail for job application confirmations, logs them cleanly into a Google Sheet, and uses the Google Gemini 2.5 API to generate tailored cover letters and application answers based strictly on your main template and personal CV.

### ✨ Features

* **Automated Gmail Tracking:** Scans your inbox every 6 hours and uses a Semantic JSON AI Pipeline to cleanly extract the Role, Company, Work Mode, and Status from ATS emails (Greenhouse, Workday, LinkedIn, etc.).
* **Smart Rejection Filter:** The AI actively analyzes email context and automatically ignores rejection emails (even if the subject line is misleading), ensuring your tracker only logs valid applications.
* **Auto-Clean & Sort:** Every time the tracker syncs, it silently scrubs duplicate entries and perfectly organizes your sheet chronologically (newest first), pushing manual, undated notes safely to the bottom.
* **Contextual AI Cover Letter Tailor:** Generates a highly tailored cover letter in Google Docs. The AI reads your actual CV/Resume document and maps your existing transferable skills to the new role rather than fabricating fake job experience.
* **Dynamic Q&A Database:** Replaces messy text boxes with a dedicated, two-column interactive database. Paste custom questions from a job board, and the AI drafts professional answers. Add a note in the "Feedback" column (e.g., "Make this shorter") to have the AI instantly refine its response.
* **Formatting Preservation:** Duplicates your exact Google Doc template, keeping your layout and hyperlinks intact. It seamlessly injects the new AI paragraphs in your exact font and size using a simple text tag.
* **Remote-Aware:** Automatically updates your closing sentence from "office" to "remotely" if the job location is listed as remote in your input tab.
* **Silent Debugging:** Features a hidden background logger that tracks API limits and errors without interrupting your workflow with pop-up alerts.
* **100% Private & Free:** Runs locally on your Google Account. No third-party extensions, no subscriptions, and your data never leaves your Google Drive.

---

## 📋 Prerequisites & Requirements

Before you install the script, ensure you have the following ready:

1. **A Standard Google Account:** You need access to Gmail, Google Sheets, Google Docs, and Google Drive.
2. **A Free Google Gemini API Key:** Get one instantly from [Google AI Studio](https://aistudio.google.com/app/apikey).
3. **A Cover Letter Template (Google Doc):** Create a base cover letter in Google Docs. Keep your standard header (contact info, links) and footer (tools, sign-off). Delete the middle paragraphs and replace them with this exact tag: `{{AI_COVER_LETTER}}`. Ensure this tag is formatted in the exact font and size you want the final generated text to be. Have the document URL ready. *(Note: Ensure your closing sentence says "I can work from the office..." so the script can automatically swap it for remote roles).*
4. **A CV / Background Document (Google Doc):** A standard Google Doc containing your resume, portfolio links, and key projects for the AI to reference when tailoring responses.
5. **An Output Folder (Google Drive):** Create an empty folder in your Google Drive where the script will save your tailored cover letters. Have the Folder ID ready (the long string of letters/numbers at the end of the folder's URL).

---

## 🛠️ Installation & Setup Guide

### Step 1: Create the Spreadsheet & Add the Code

1. Open a brand new, completely blank Google Sheet (type `sheets.new` into your browser).
2. Name the spreadsheet (e.g., "Automated Job Tracker").
3. Click **Extensions > Apps Script** from the top menu.
4. Delete any existing code in the editor, and paste in the full code from `Code.gs` (provided in this repository).
5. Click the **Save** icon 💾 (or press Cmd/Ctrl + S).

### Step 2: Securely Add Your Gemini API Key

*We use Google's secure Script Properties so your API key is never exposed in the spreadsheet cells.*

1. In the Apps Script editor, click the **Project Settings** icon (gear ⚙️) on the left sidebar.
2. Scroll down to **Script Properties** and click **Add script property**.
3. Under **Property**, type exactly: `GEMINI_API_KEY`
4. Under **Value**, paste your secret key from Google AI Studio.
5. Click **Save script properties**.

### Step 3: Initialize the System

1. Close the Apps Script tab and return to your Google Sheet. Refresh the page.
2. You will now see a custom menu at the top called **🤖 Job Tools**.
3. Click **🤖 Job Tools > 1. Setup Tracker & Settings**.

* *Security Note: Because this is a custom script running on your own account, Google will ask for permission and show an "Unverified App" warning. Click **Advanced > Go to [Project Name] (unsafe)** to allow it to build your tracker.*

4. Click **🤖 Job Tools > 2. Setup Cover Letter Tab** to generate your AI input interface.
5. Click **🤖 Job Tools > 3. Setup Interview Q&A Tab** to generate your Q&A database.

### Step 4: Configure Your Settings

Go to the newly created **Settings** tab in your spreadsheet and fill in the bottom three rows:

* **Cover Letter Template URL:** Paste the full link to your base Google Doc template (Ensure your document contains the `{{AI_COVER_LETTER}}` tag).
* **Cover Letters Folder ID:** Paste the ID of your designated Google Drive output folder.
* **CV Doc URL:** Paste the full link to your CV Google Doc.

### Step 5: Automate the Background Tracker

To make the script scan your email automatically:

1. Open the Apps Script editor (**Extensions > Apps Script**).
2. Click the **Triggers** icon (alarm clock ⏰) on the left sidebar.
3. Click **+ Add Trigger** in the bottom right corner.
4. Set it up exactly like this:

* **Function to run:** `syncJobsFromGmail`
* **Event source:** `Time-driven`
* **Type of time based trigger:** `Hour timer`
* **Hour interval:** `Every 6 hours` (or your preference)

5. Click **Save**.

---

## 💻 How to Use It

### The Automated Tracker

* Your spreadsheet will automatically update in the background as you apply for jobs. It will securely extract data, filter out rejections, and sort your list chronologically.
* To force a manual scan immediately, click **🤖 Job Tools > 🔄 Sync Gmail Now**.
* To manually tidy up the sheet after typing in your own notes, click **🤖 Job Tools > 🧹 Clean & Sort Tracker**.

### The Cover Letter Generator

1. Navigate to the **Cover Letter Input** tab.
2. Fill in the specific details for the job you are applying for (Company, Role, Location).
3. Paste the full Job Description into the large input box.
4. Click **🤖 Job Tools > ⚡ Generate Cover Letter**.
5. Wait a few seconds. Your perfectly tailored, formatted cover letter is now saved as a new Google Doc in your specified Drive folder!

### Answer App Questions (Interactive Database)

1. Ensure the Company, Role, and Job Description are filled out on the `Cover Letter Input` tab so the AI has context.
2. Navigate to the **Interview Q&A** tab.
3. Paste your custom application questions into **Column A** (one question per row).
4. Click **🤖 Job Tools > 📝 Answer App Questions**.
5. The AI will output professional, tailored responses directly into **Column B**.
6. *Want to tweak an answer?* Type instructions (e.g., "Make this punchier") into **Column C**, run the tool again, and the AI will update that specific row!

### Reset the Interface

* Click **🤖 Job Tools > 🧹 Clear Current Inputs** to clear all fields and text boxes across your input tabs when you are ready to start your next application!
