/**
 * IGC Market at a Glance Web Scraper for Google Apps Script
 * Website: https://www.igc.int/en/default.aspx
 * 
 * Features:
 * - Robust Error Handling & Retries (Exponential Backoff)
 * - Automatic Failure Notifications (MS Teams & Email alerts on error)
 * - Scrapes daily prices for Wheat, Maize, Barley, Soyabeans, Rice
 * - Updates Google Sheets (Consolidated & Group Tabs)
 * - Detects NEW daily price dates
 * - Sends MS Teams Channel Notification with Google Drive CSV link attached
 * - Sends Email to 3 recipients with CSV file attached
 * 
 * Author: Antigravity AI
 */

// =========================================================================
// 🔑 SCRIPT PROPERTY NAMES (Configure in Project Settings > Script Properties)
// =========================================================================
// 1. "SPREADSHEET_ID"    : Google Spreadsheet ID (Leave blank to auto-create or use active sheet)
// 2. "SHEET_NAME"        : Sheet tab name for consolidated data (Default: "All_Data")
// 3. "TEAMS_WEBHOOK_URL" : MS Teams Channel Incoming Webhook URL
// 4. "EMAIL_RECIPIENTS"  : Comma-separated email addresses (e.g. "a@email.com, b@email.com, c@email.com")
// 5. "CSV_FOLDER_ID"     : (Optional) Google Drive Folder ID to store CSV exports
// 6. "LAST_PROCESSED_DATE": (Auto-updated) Tracks the latest notified price date

/**
 * Main execution function with Global Error Handling.
 */
function runScraper() {
  const props = PropertiesService.getScriptProperties();
  const teamsWebhookUrl = props.getProperty('TEAMS_WEBHOOK_URL');
  const emailRecipients = props.getProperty('EMAIL_RECIPIENTS');

  try {
    Logger.log('Starting IGC Market Data Scraper...');
    executeScraper(props);
    Logger.log('Scraper execution completed successfully!');
  } catch (error) {
    const errorMsg = error.stack || error.toString();
    Logger.log('❌ CRITICAL ERROR OCCURRED: ' + errorMsg);

    // Send Error Alerts via MS Teams & Email
    if (teamsWebhookUrl && teamsWebhookUrl.trim() !== '') {
      sendTeamsErrorAlert(teamsWebhookUrl, errorMsg);
    }
    if (emailRecipients && emailRecipients.trim() !== '') {
      sendEmailErrorAlert(emailRecipients, errorMsg);
    }

    // Re-throw to record error in Apps Script Executions log
    throw error;
  }
}

/**
 * Core Scraper Execution Engine
 */
function executeScraper(props) {
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  const sheetName = props.getProperty('SHEET_NAME') || 'All_Data';
  const teamsWebhookUrl = props.getProperty('TEAMS_WEBHOOK_URL');
  const emailRecipients = props.getProperty('EMAIL_RECIPIENTS');
  const lastProcessedDate = props.getProperty('LAST_PROCESSED_DATE') || '';

  // 1. Get or Open Spreadsheet
  let ss;
  if (spreadsheetId && spreadsheetId.trim() !== '') {
    ss = SpreadsheetApp.openById(spreadsheetId.trim());
  } else {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {
      ss = null;
    }
    if (!ss) {
      ss = SpreadsheetApp.create('IGC_Market_Data');
      props.setProperty('SPREADSHEET_ID', ss.getId());
      Logger.log('Created new Spreadsheet with ID: ' + ss.getId());
    }
  }

  // 2. Fetch main page with Retry logic
  const baseUrl = 'https://www.igc.int/en/default.aspx';
  const initialHeaders = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
  };

  const initialResponse = fetchWithRetry(baseUrl, {
    'method': 'get',
    'headers': initialHeaders,
    'muteHttpExceptions': true
  });

  const html = initialResponse.getContentText();
  const cookies = extractCookies(initialResponse.getAllHeaders());

  // Extract ASP.NET Hidden Fields
  const viewState = extractHiddenInput('__VIEWSTATE', html);
  const generator = extractHiddenInput('__VIEWSTATEGENERATOR', html);
  const validation = extractHiddenInput('__EVENTVALIDATION', html);

  if (!viewState || !generator || !validation) {
    throw new Error('Failed to extract ASP.NET WebForms tokens (__VIEWSTATE, __VIEWSTATEGENERATOR, __EVENTVALIDATION) from IGC website.');
  }

  const tabs = [
    { name: 'Wheat', button: 'WheatPriceButton' },
    { name: 'Maize', button: 'MaizePriceButton' },
    { name: 'Barley', button: 'BarleyPriceButton' },
    { name: 'Soyabeans', button: 'SoyabeansPriceButton' },
    { name: 'Rice', button: 'RicePriceButton' }
  ];

  const allFlatData = [];

  // 3. Loop over each commodity group tab with Retries
  tabs.forEach(tab => {
    Logger.log('Scraping group: ' + tab.name + '...');

    const payload = {
      '__EVENTTARGET': tab.button,
      '__EVENTARGUMENT': '',
      '__VIEWSTATE': viewState,
      '__VIEWSTATEGENERATOR': generator,
      '__EVENTVALIDATION': validation
    };

    const postHeaders = {
      'User-Agent': initialHeaders['User-Agent'],
      'Content-Type': 'application/x-www-form-urlencoded'
    };
    if (cookies) {
      postHeaders['Cookie'] = cookies;
    }

    const postResponse = fetchWithRetry(baseUrl, {
      'method': 'post',
      'headers': postHeaders,
      'payload': payload,
      'muteHttpExceptions': true
    });

    const postHtml = postResponse.getContentText();
    const groupData = parseTableData(postHtml, tab.name);

    if (groupData.length === 0) {
      Logger.log('⚠️ Warning: No price data parsed for commodity group: ' + tab.name);
    }

    allFlatData.push(...groupData);
  });

  if (allFlatData.length === 0) {
    throw new Error('Scraper completed but 0 total records were parsed across all commodity groups.');
  }

  Logger.log('Total records scraped successfully: ' + allFlatData.length);

  // 4. Update Google Sheets
  updateConsolidatedSheet(ss, sheetName, allFlatData);
  updateGroupSheets(ss, allFlatData);

  // 5. Detect NEW DATE
  const latestDateIso = getLatestDateIso(allFlatData);
  const latestDateDisplay = getLatestDateDisplay(allFlatData);
  Logger.log('Latest Date in Scraped Data: ' + latestDateDisplay + ' (ISO: ' + latestDateIso + ')');
  Logger.log('Last Notified Processed Date: ' + lastProcessedDate);

  const isNewDateDetected = (latestDateIso > lastProcessedDate);

  if (isNewDateDetected) {
    Logger.log('🎉 NEW DATE DETECTED! Processing notifications...');

    // A. Export CSV to Google Drive
    const csvFile = createCsvFileInDrive(allFlatData, latestDateIso);
    const csvFileUrl = csvFile.getUrl();
    Logger.log('CSV file created in Drive: ' + csvFileUrl);

    // B. Send MS Teams Notification
    if (teamsWebhookUrl && teamsWebhookUrl.trim() !== '') {
      sendMSTeamsNotification(teamsWebhookUrl, latestDateDisplay, allFlatData, csvFileUrl, ss.getUrl());
    }

    // C. Send Email to Recipients
    if (emailRecipients && emailRecipients.trim() !== '') {
      sendEmailWithAttachment(emailRecipients, latestDateDisplay, csvFile, ss.getUrl());
    }

    // D. Update Last Processed Date
    props.setProperty('LAST_PROCESSED_DATE', latestDateIso);
    Logger.log('Updated LAST_PROCESSED_DATE property to: ' + latestDateIso);

  } else {
    Logger.log('No new date detected. Current latest date (' + latestDateDisplay + ') already processed. No notifications sent.');
  }
}

/**
 * Robust HTTP Fetcher with Exponential Backoff Retry Logic
 */
function fetchWithRetry(url, options, maxRetries = 3) {
  let attempt = 0;
  while (attempt < maxRetries) {
    attempt++;
    try {
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();

      if (statusCode === 200) {
        return response;
      }

      Logger.log(`[Attempt ${attempt}/${maxRetries}] HTTP Status ${statusCode} received from ${url}`);
      if (attempt >= maxRetries) {
        throw new Error(`HTTP Request failed with status code ${statusCode} after ${maxRetries} attempts.`);
      }
    } catch (e) {
      Logger.log(`[Attempt ${attempt}/${maxRetries}] Fetch exception: ${e.message}`);
      if (attempt >= maxRetries) {
        throw new Error(`Network/Fetch error after ${maxRetries} attempts: ${e.message}`);
      }
    }

    // Exponential Backoff Wait (2s, 4s, 8s...)
    const sleepMs = Math.pow(2, attempt) * 1000;
    Logger.log(`Waiting ${sleepMs / 1000}s before retrying...`);
    Utilities.sleep(sleepMs);
  }
}

/**
 * Sends Error Alert to MS Teams Channel
 */
function sendTeamsErrorAlert(webhookUrl, errorMsg) {
  try {
    const cardPayload = {
      "@type": "MessageCard",
      "@context": "http://schema.org/extensions",
      "themeColor": "D93025", // Red error color
      "summary": "❌ IGC Scraper System Error Alert",
      "sections": [{
        "activityTitle": "❌ IGC Market Scraper Execution Failed",
        "activitySubtitle": `Time: ${new Date().toLocaleString()}`,
        "text": `**An error occurred while running the automated IGC scraper:**\n\`\`\`\n${errorMsg.substring(0, 1000)}\n\`\`\``,
        "markdown": true
      }]
    };

    UrlFetchApp.fetch(webhookUrl.trim(), {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(cardPayload),
      'muteHttpExceptions': true
    });
  } catch (e) {
    Logger.log('Failed to send Teams Error Alert: ' + e.message);
  }
}

/**
 * Sends Error Alert Email to Recipients
 */
function sendEmailErrorAlert(recipientsStr, errorMsg) {
  try {
    const subject = `[Alert] IGC Market Scraper Execution Error`;
    const body = `Dear Admin/Team,\n\nAn error occurred while executing the IGC Market Scraper System:\n\nError Details:\n${errorMsg}\n\nPlease check the Google Apps Script Executions log for details.\n\nSystem Timestamp: ${new Date().toString()}`;

    MailApp.sendEmail({
      to: recipientsStr.trim(),
      subject: subject,
      body: body
    });
  } catch (e) {
    Logger.log('Failed to send Email Error Alert: ' + e.message);
  }
}

/**
 * Converts date strings like "29/07/2026" to ISO "2026-07-29"
 */
function parseDateToIso(dateStr) {
  if (!dateStr) return '';
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    const day = parts[0].padStart(2, '0');
    const month = parts[1].padStart(2, '0');
    const year = parts[2];
    return `${year}-${month}-${day}`;
  }
  return dateStr;
}

function getLatestDateIso(allData) {
  let maxIso = '';
  allData.forEach(item => {
    const iso = parseDateToIso(item.date);
    if (iso > maxIso) maxIso = iso;
  });
  return maxIso;
}

function getLatestDateDisplay(allData) {
  const maxIso = getLatestDateIso(allData);
  for (let item of allData) {
    if (parseDateToIso(item.date) === maxIso) {
      return item.date;
    }
  }
  return maxIso;
}

function createCsvFileInDrive(allFlatData, dateIso) {
  let csvContent = 'Group,SubCommodity,Date,Price_USD,Updated_At\n';
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  allFlatData.forEach(item => {
    const cleanGroup = `"${item.group.replace(/"/g, '""')}"`;
    const cleanSub = `"${item.subCommodity.replace(/"/g, '""')}"`;
    const cleanDate = `"${item.date}"`;
    const priceVal = item.price !== null ? item.price : '';
    csvContent += `${cleanGroup},${cleanSub},${cleanDate},${priceVal},"${nowStr}"\n`;
  });

  const fileName = `IGC_Market_Data_${dateIso}.csv`;
  const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);

  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('CSV_FOLDER_ID');
  
  let file;
  if (folderId && folderId.trim() !== '') {
    const folder = DriveApp.getFolderById(folderId.trim());
    file = folder.createFile(blob);
  } else {
    file = DriveApp.createFile(blob);
  }

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log('Note on setSharing: ' + e.message);
  }

  return file;
}

function sendMSTeamsNotification(webhookUrl, dateDisplay, allData, csvUrl, sheetUrl) {
  Logger.log('Sending MS Teams Notification...');

  const latestIso = parseDateToIso(dateDisplay);
  const latestRecords = allData.filter(i => parseDateToIso(i.date) === latestIso);

  let summaryText = `**🌾 IGC Commodity Prices Updated for ${dateDisplay}**\n\n`;
  latestRecords.forEach(rec => {
    const priceStr = rec.price !== null ? `$${rec.price}` : 'N/A';
    summaryText += `• **[${rec.group}]** ${rec.subCommodity}: **${priceStr}**\n`;
  });

  const cardPayload = {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    "themeColor": "0076D7",
    "summary": `IGC Market Price Update (${dateDisplay})`,
    "sections": [{
      "activityTitle": "📢 IGC Market Daily Price Alert (New Date Detected!)",
      "activitySubtitle": `Date: ${dateDisplay}`,
      "text": summaryText,
      "markdown": true
    }],
    "potentialAction": [
      {
        "@type": "OpenUri",
        "name": "📥 Download / View CSV File",
        "targets": [{ "os": "default", "uri": csvUrl }]
      },
      {
        "@type": "OpenUri",
        "name": "📊 Open Google Sheet",
        "targets": [{ "os": "default", "uri": sheetUrl }]
      }
    ]
  };

  const response = UrlFetchApp.fetch(webhookUrl.trim(), {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(cardPayload),
    'muteHttpExceptions': true
  });

  Logger.log('MS Teams Webhook Response: ' + response.getResponseCode());
}

function sendEmailWithAttachment(recipientsStr, dateDisplay, csvFile, sheetUrl) {
  Logger.log('Sending Email with CSV attachment to: ' + recipientsStr);

  const subject = `[IGC Market Update] New Commodity Prices Alert - ${dateDisplay}`;
  
  const bodyText = `Dear Team,\n\n` +
    `New daily commodity market prices for ${dateDisplay} have been detected and updated from IGC.\n\n` +
    `📎 Attached File: ${csvFile.getName()}\n` +
    `🌐 Direct Google Drive CSV Link: ${csvFile.getUrl()}\n` +
    `📊 Google Sheet Link: ${sheetUrl}\n\n` +
    `Best regards,\n` +
    `Automated IGC Market Scraper System`;

  const htmlBody = `<div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">` +
    `<h2 style="color: #1b5e20;">🌾 IGC Market Daily Price Update</h2>` +
    `<p>New daily commodity prices for <strong>${dateDisplay}</strong> have been published.</p>` +
    `<div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin: 15px 0;">` +
    `<p style="margin: 5px 0;">📁 <strong>Attached File:</strong> ${csvFile.getName()}</p>` +
    `<p style="margin: 5px 0;">🔗 <a href="${csvFile.getUrl()}" target="_blank" style="color: #1a73e8; font-weight: bold;">Download / Open CSV from Google Drive</a></p>` +
    `<p style="margin: 5px 0;">📊 <a href="${sheetUrl}" target="_blank" style="color: #1a73e8; font-weight: bold;">Open Live Google Sheet</a></p>` +
    `</div>` +
    `<p style="font-size: 12px; color: #777;">This is an automated notification from IGC Market Scraper System.</p>` +
    `</div>`;

  MailApp.sendEmail({
    to: recipientsStr.trim(),
    subject: subject,
    body: bodyText,
    htmlBody: htmlBody,
    attachments: [csvFile.getAs(MimeType.CSV)]
  });

  Logger.log('Email sent successfully!');
}

function parseTableData(html, groupName) {
  const tableRegex = /<table[^>]*id="GridViewHiddenPrices"[^>]*>([\s\S]*?)<\/table>/i;
  const tableMatch = html.match(tableRegex);
  if (!tableMatch) return [];

  const tableHtml = tableMatch[1];
  
  const thRegex = /<th[^>]*>([\s\S]*?)<\/th>/gi;
  const rawHeaders = [];
  let thMatch;
  while ((thMatch = thRegex.exec(tableHtml)) !== null) {
    rawHeaders.push(stripHtmlTags(thMatch[1]));
  }
  const subCommodities = rawHeaders.slice(1);

  const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  const rows = [];
  let trMatch;
  let isFirstRow = true;
  while ((trMatch = trRegex.exec(tableHtml)) !== null) {
    if (isFirstRow) {
      isFirstRow = false;
      continue;
    }

    const rowHtml = trMatch[1];
    const tdRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
    const cols = [];
    let tdMatch;
    while ((tdMatch = tdRegex.exec(rowHtml)) !== null) {
      cols.push(stripHtmlTags(tdMatch[1]));
    }

    if (cols.length >= 2) {
      const dateVal = cols[0];
      subCommodities.forEach((subComm, idx) => {
        const rawPrice = cols[idx + 1];
        let numPrice = null;
        if (rawPrice && rawPrice !== '-') {
          numPrice = parseFloat(rawPrice);
          if (isNaN(numPrice)) numPrice = rawPrice;
        }
        rows.push({
          group: groupName,
          subCommodity: subComm,
          date: dateVal,
          price: numPrice
        });
      });
    }
  }

  return rows;
}

function updateConsolidatedSheet(ss, sheetName, dataList) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Group', 'SubCommodity', 'Date', 'Price_USD', 'Updated_At']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#EFEFEF');
  }

  const existingData = sheet.getDataRange().getValues();
  const existingMap = new Map();
  for (let i = 1; i < existingData.length; i++) {
    const key = `${existingData[i][0]}|${existingData[i][1]}|${existingData[i][2]}`;
    existingMap.set(key, i + 1);
  }

  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const rowsToAppend = [];

  dataList.forEach(item => {
    const key = `${item.group}|${item.subCommodity}|${item.date}`;
    if (existingMap.has(key)) {
      const rowIndex = existingMap.get(key);
      sheet.getRange(rowIndex, 4, 1, 2).setValues([[item.price, nowStr]]);
    } else {
      rowsToAppend.push([item.group, item.subCommodity, item.date, item.price, nowStr]);
    }
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
  }

  sheet.autoResizeColumns(1, 5);
}

function updateGroupSheets(ss, dataList) {
  const grouped = {};
  dataList.forEach(item => {
    if (!grouped[item.group]) grouped[item.group] = [];
    grouped[item.group].push(item);
  });

  Object.keys(grouped).forEach(groupName => {
    let sheet = ss.getSheetByName(groupName);
    if (!sheet) {
      sheet = ss.insertSheet(groupName);
    }

    const items = grouped[groupName];
    const subComms = [...new Set(items.map(i => i.subCommodity))];
    const dates = [...new Set(items.map(i => i.date))];

    const headerRow = ['Date', ...subComms, 'Updated_At'];
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const priceMap = {};
    items.forEach(i => {
      priceMap[`${i.date}|${i.subCommodity}`] = i.price;
    });

    const matrixRows = dates.map(d => {
      const row = [d];
      subComms.forEach(sc => {
        const val = priceMap[`${d}|${sc}`];
        row.push(val !== undefined ? val : '');
      });
      row.push(nowStr);
      return row;
    });

    sheet.clear();
    sheet.appendRow(headerRow);
    sheet.getRange(1, 1, 1, headerRow.length).setFontWeight('bold').setBackground('#D9EAD3');
    if (matrixRows.length > 0) {
      sheet.getRange(2, 1, matrixRows.length, headerRow.length).setValues(matrixRows);
    }
    sheet.autoResizeColumns(1, headerRow.length);
  });
}

function extractHiddenInput(name, html) {
  const regex = new RegExp(`id="${name}"[^>]*value="([^"]*)"`, 'i');
  const match = html.match(regex);
  if (match) return match[1];

  const nameRegex = new RegExp(`name="${name}"[^>]*value="([^"]*)"`, 'i');
  const nameMatch = html.match(nameRegex);
  return nameMatch ? nameMatch[1] : '';
}

function extractCookies(headers) {
  if (!headers) return '';
  const setCookie = headers['Set-Cookie'] || headers['set-cookie'];
  if (!setCookie) return '';
  
  if (Array.isArray(setCookie)) {
    return setCookie.map(c => c.split(';')[0]).join('; ');
  } else {
    return setCookie.split(';')[0];
  }
}

function stripHtmlTags(str) {
  if (!str) return '';
  return str.replace(/<[^>]+>/g, '').trim();
}

function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runScraper') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('runScraper')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Daily trigger created for runScraper at 8:00 AM!');
}
