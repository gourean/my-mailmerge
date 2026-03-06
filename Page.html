function onOpen() {
  SpreadsheetApp.getUi().createMenu('My Mail Merge')
      .addItem('Open My Mail Merge', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('My Mail Merge')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetHeaders() {
  return SpreadsheetApp.getActiveSheet().getRange(1, 1, 1, SpreadsheetApp.getActiveSheet().getLastColumn()).getValues()[0];
}

function saveSettings(config) {
  PropertiesService.getUserProperties().setProperty('saved_config', JSON.stringify(config));
  return "Settings saved!";
}

function loadSettings() {
  const saved = PropertiesService.getUserProperties().getProperty('saved_config');
  return saved ? JSON.parse(saved) : null;
}

function clearLogs() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const logIdx = headers.indexOf("Log Sent");
  const linkIdx = headers.indexOf("Drive Link");
  
  if (logIdx > -1) sheet.getRange(2, logIdx + 1, sheet.getLastRow()).clearContent();
  if (linkIdx > -1) sheet.getRange(2, linkIdx + 1, sheet.getLastRow()).clearContent();
  return "Logs cleared!";
}

function checkPermissions(folderId) {
  if (!folderId) return "❌ Error: Please enter a Folder ID first.";
  try {
    const folder = DriveApp.getFolderById(folderId);
    return "✅ Success! Access confirmed for: " + folder.getName();
  } catch (e) {
    return "❌ Access Denied: Check ID and ensure folder is shared with this account.";
  }
}

function generateDriveLinks(folderId, fileNameCol) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const fileColIdx = headers.indexOf(fileNameCol);
  
  if (fileColIdx === -1) return "❌ Error: Could not find column: " + fileNameCol;

  let linkColIdx = headers.indexOf("Drive Link");
  if (linkColIdx === -1) {
    linkColIdx = headers.length;
    sheet.getRange(1, linkColIdx + 1).setValue("Drive Link");
  }

  const folder = DriveApp.getFolderById(folderId);
  data.forEach((row, i) => {
    const files = folder.getFilesByName(row[fileColIdx]);
    if (files.hasNext()) {
      const file = files.next();
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheet.getRange(i + 2, linkColIdx + 1).setValue(file.getUrl());
    } else {
      sheet.getRange(i + 2, linkColIdx + 1).setValue("File Not Found");
    }
  });
  return "🔗 Links generated in 'Drive Link' column!";
}

function runMerge(config) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const userEmail = Session.getActiveUser().getEmail();
  
  let logIdx = headers.indexOf("Log Sent");
  if (logIdx === -1 && !config.isTest) {
    logIdx = headers.length;
    sheet.getRange(1, logIdx + 1).setValue("Log Sent");
  }

  const rowsToProcess = config.isTest ? [data[0]] : data;
  const total = rowsToProcess.length;

  rowsToProcess.forEach((row, i) => {
    const recipient = config.isTest ? userEmail : row[headers.indexOf(config.emailCol)];
    
    // Rectify paragraphs: Convert newlines to HTML breaks
    let body = config.messageBody.replace(/\n/g, '<br>');
    
    headers.forEach((h, idx) => {
      const regex = new RegExp(`{{${h}}}`, "g");
      body = body.replace(regex, row[idx]);
    });

    const options = {
      cc: config.cc, 
      bcc: config.bcc, 
      replyTo: config.replyTo,
      htmlBody: body,
      attachments: []
    };

    if (config.folderId && row[headers.indexOf(config.fileCol)]) {
      try {
        const folder = DriveApp.getFolderById(config.folderId);
        const files = folder.getFilesByName(row[headers.indexOf(config.fileCol)]);
        if (files.hasNext()) {
          const file = files.next();
          if (config.mode === 'attach') options.attachments.push(file.getBlob());
        } else if (!config.isTest) {
          sheet.getRange(i + 2, logIdx + 1).setValue("Error: File Not Found");
          return; 
        }
      } catch(e) {
        if (!config.isTest) sheet.getRange(i + 2, logIdx + 1).setValue("Error: Folder Access");
        return;
      }
    }

    MailApp.sendEmail(recipient, config.subjectLine, body, options);
    if (!config.isTest) sheet.getRange(i + 2, logIdx + 1).setValue(new Date());
    PropertiesService.getUserProperties().setProperty('mergeProgress', `${i + 1}/${total}`);
  });
  return config.isTest ? "Test email sent!" : "Merge Complete!";
}

function getProgress() {
  return PropertiesService.getUserProperties().getProperty('mergeProgress') || "0/0";
}
