# google-workspace-ats
An automated Applicant Tracking System and AI Cover Letter Generator built entirely in Google Sheets.

# 🚀 Automated ATS & AI Cover Letter Generator (Google Workspace)

A fully automated, private Applicant Tracking System (ATS) and AI Cover Letter Generator that lives entirely inside your personal Google Workspace.

Tired of manually logging jobs and rewriting cover letters? This tool runs in the background to automatically scan your Gmail for job application confirmations, logs them cleanly into a Google Sheet, and uses the Google Gemini API to generate perfectly tailored cover letters based on your master template.

### ✨ Features

* **Automated Gmail Tracking:** Scans your inbox every 6 hours and cleanly extracts the Role, Company, and Status from ATS emails (Greenhouse, Workday, LinkedIn, etc.).
* **Smart Duplication Prevention:** Never logs the same job twice.
* **AI Cover Letter Tailor:** Generates a highly-tailored cover letter in Google Docs based strictly on your personal template and the specific job description.
* **100% Private & Free:** Runs locally on your Google Account. No third-party extensions, no subscriptions, and your data never leaves your Google Drive.

---

## 📋 Prerequisites & Requirements

Before you install the script, ensure you have the following ready:

1. **A Standard Google Account:** You need access to Gmail, Google Sheets, Google Docs, and Google Drive.
2. **A Free Google Gemini API Key:** Get one instantly from [Google AI Studio](https://aistudio.google.com/app/apikey).
3. **A Cover Letter Template (Google Doc):** Create a base cover letter in Google Docs that the AI will use as its strict structural template. Have the URL ready.
4. **An Output Folder (Google Drive):** Create an empty folder in your Google Drive where the AI will save your new, tailored cover letters. Have the Folder ID ready (the long string of letters/numbers at the end of the folder's URL).

---

## 🛠️ Installation & Setup Guide

### Step 1: Create the Spreadsheet & Add the Code

1. Open a brand new, completely blank Google Sheet (type `sheets.new` into your browser).
2. Name the spreadsheet (e.g., "Automated Job Tracker").
3. Click **Extensions > Apps Script** from the top menu.


<img width="985" height="989" alt="Screenshot 2026-03-03 at 04 10 37" src="https://github.com/user-attachments/assets/39d6055a-ca7a-4e93-b5ad-7992339122a3" />


4. Delete any existing code in the editor, and paste in the full code from `Code.gs` (provided in this repository).
5. Click the **Save** icon 💾 (or press Cmd/Ctrl + S).

### Step 2: Securely Add Your Gemini API Key

*We use Google's secure Script Properties so your API key is never exposed in the spreadsheet cells.*

1. In the Apps Script editor, click the **Project Settings** icon (gear ⚙️) on the left sidebar.
2. Scroll down to **Script Properties** and click **Add script property**.
3. Under **Property**, type exactly: `GEMINI_API_KEY`
4. Under **Value**, paste your secret key from Google AI Studio.
5. Click **Save script properties**.

<img width="1553" height="978" alt="Screenshot 2026-03-03 at 04 21 38" src="https://github.com/user-attachments/assets/cd20fe57-790f-4487-b9db-cdebc7af1f01" />


### Step 3: Initialize the System

1. Close the Apps Script tab and return to your Google Sheet. Refresh the page.
2. You will now see a custom menu at the top called **🤖 Job Tools**.
3. Click **🤖 Job Tools > 1. Setup Tracker & Settings**.
* *Security Note: Because this is a custom script running on your own account, Google will ask for permission and show an "Unverified App" warning. Click **Advanced > Go to [Project Name] (unsafe)** to allow it to build your tracker.*

<img width="1076" height="992" alt="Screenshot 2026-03-03 at 04 10 22" src="https://github.com/user-attachments/assets/7f66a8c5-ce46-4342-8729-2c35282bddb4" />


4. Click **🤖 Job Tools > 2. Setup Cover Letter Tab** to generate your AI input interface.

### Step 4: Configure Your Settings

Go to the newly created **Settings** tab in your spreadsheet and fill in the bottom two rows:

* **Cover Letter Template URL:** Paste the full link to your base Google Doc template.
* **Cover Letters Folder ID:** Paste the ID of your designated Google Drive output folder.

<img width="1593" height="999" alt="Screenshot 2026-03-03 at 04 28 27" src="https://github.com/user-attachments/assets/11450d74-de2a-4c5e-9b11-74b3a8fc334e" />


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

<img width="1720" height="987" alt="Screenshot 2026-03-03 at 04 21 04" src="https://github.com/user-attachments/assets/12cd13f9-e69c-4826-81cb-ced3d5b841ee" />

---

## 💻 How to Use It

### The Automated Tracker

* Your spreadsheet will automatically update in the background as you apply for jobs.
* To force a manual scan immediately, click **🤖 Job Tools > 🔄 Sync Gmail Now**.

### The AI Cover Letter Generator

1. Navigate to the **Cover Letter Input** tab.
2. Fill in the specific details for the job you are applying for (Company, Role, Location).
3. Paste the full Job Description into the large input box.
4. Click **🤖 Job Tools > ⚡ Generate Cover Letter**.
5. Wait a few seconds for the "Success" pop-up. Your perfectly tailored, formatted cover letter is now saved as a new Google Doc in your specified Drive folder!

<img width="1092" height="999" alt="Screenshot 2026-03-03 at 04 30 32" src="https://github.com/user-attachments/assets/58f04778-69da-4ae9-80ab-f1d5dbe66026" />



