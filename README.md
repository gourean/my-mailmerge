# My Mail Merge
My Mail Merge is a professional-grade Google Sheets add-on that automates personalized email campaigns. It allows you to send custom messages with physical attachments or Drive links directly from your spreadsheet data.

# Getting Started (For Users)
**1. One-Click Installation**
To create your own private copy of the spreadsheet and the automation script, click the link below:

[Click here to copy the My Mail Merge Template](https://docs.google.com/spreadsheets/d/1t65bne3jl7QsNoAdNC-Va5hKbb5RSy3mD3wsLcgiM7o/copy)

**2. Authorization**
Because the app interacts with your Gmail and Google Drive, you must grant permissions on the first run:

Open your new spreadsheet copy.

Click the **My Mail Merge** menu > Open My Mail Merge.

Follow the Google prompts: Continue > Select Account > Advanced > Go to My Mail Merge (unsafe) > Allow.

# Manual Setup (For Developers)
If you prefer to build this into an existing sheet, follow these steps to use the source code provided in this repository:

1. Open your Google Sheet.

2. Go to Extensions > Apps Script.

3. Create a new script file named **Code.gs** and paste the provided Javascript code.

4. Create a new HTML file named **Page.html** and paste the provided HTML/CSS code.

5. Click Services (+) and add the **Drive API**.

6. Save and refresh your spreadsheet to see the new menu.

# How to Use
**1. Prepare Your Spreadsheet**
Ensure your headers are in Row 1. Your sheet should include columns for:

**- Email Address**
**- File Name** (Must include extension, e.g., invoice.pdf).
- Any personalization fields like Name etc.

**2. Configure the Sidebar**
Open the sidebar from the **My Mail Merge** menu:

**-Mapping**: Select your target columns from the dropdowns.
**-Drive Folder ID**: Open your Google Drive folder and copy the ID from the URL.
**-ID Location**: drive.google.com/drive/folders/**YOUR_FOLDER_ID_HERE**
**-Reply-To**: Enter a manual email address where you want recipients to reply.

**3. Personalize Your Message**
Use {{Column Name}} syntax in the subject or message body.

**-Example**: Hi {{Name}}, find your {{Month}} report attached.

**-Formatting**: Press Enter to create new paragraphs; the script preserves your layout.

# Features
**Save Settings**: Stores your subject, message, and folder ID for future sessions.

**Send Test Email**: Sends a preview of the first data row to your own inbox.

**Error Handling**: Flags missing files as "Error: File Not Found" in the Log Sent column.

# Important Limits
**Daily Quotas**: Personal accounts are limited to **100 emails/day**; Workspace accounts are limited to **2,000/day**.

**Browser Issue**s: If the sidebar doesn't load, ensure you are **NOT** logged into **multiple Google accounts** or use an Incognito Window.

# Security & Privacy
This is a **Container-Bound Script**. All data stays within your Google account. **No third-party servers** ever see your contacts or files.
