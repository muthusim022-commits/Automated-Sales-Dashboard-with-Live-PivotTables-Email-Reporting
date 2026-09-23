/**
 * Dashboard PDF & Email Dispatcher
 * ---------------------------------
 * Exports the "Dashboard" tab as a PDF, saves it into a Drive folder,
 * and emails it to a list of stakeholders — either on a weekly schedule
 * or on demand via a checkbox "button" placed on the Dashboard tab.
 *
 * SETUP
 * 1. Open the spreadsheet → Extensions → Apps Script.
 * 2. Paste this whole file in, replacing the default Code.gs contents.
 * 3. Edit the CONFIG block below if you want different sheet/cell names.
 * 4. Run `setup()` once from the editor (it will ask you to authorize
 *    Gmail/Drive/Sheets access). This will:
 *      - create a "Recipients" tab with a sample email table
 *      - add a checkbox "button" + label on the Dashboard tab
 *      - install the weekly trigger
 *      - install the trigger that watches the button
 * 5. Open the "Recipients" tab and replace the sample emails with your
 *    real stakeholder list — one email per row.
 * 6. On the Dashboard tab, tick the checkbox any time to send a report
 *    immediately. It unticks itself once the send completes.
 */

const CONFIG = {
  // Tab that gets exported as the PDF.
  DASHBOARD_SHEET_NAME: 'Dashboard',

  // Tab holding the recipient list (created automatically if missing).
  RECIPIENTS_SHEET_NAME: 'Recipients',

  // Drive folder PDFs are saved into (created if missing).
  DRIVE_FOLDER_NAME: 'Portfolio',

  // Cell on the Dashboard tab that holds the "Send Now" checkbox.
  BUTTON_CELL: 'K2',
  BUTTON_LABEL_CELL: 'L2',
  BUTTON_LABEL_TEXT: '◀ Tick to send the report now',

  // Email subject/body. {{date}} is replaced automatically.
  EMAIL_SUBJECT: 'Sales Dashboard Summary — {{date}}',
  EMAIL_BODY:
    'Hi team,\n\n' +
    'Attached is the latest Sales Dashboard summary as of {{date}}.\n\n' +
    'This is an automated message.',

  // Day/time for the recurring trigger. DAY_OF_WEEK options:
  // MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY, SUNDAY
  TRIGGER_DAY: ScriptApp.WeekDay.MONDAY,
  TRIGGER_HOUR: 8, // 24-hour clock, in the script's timezone
};

/**
 * One-time setup. Run this manually once from the Apps Script editor.
 * Safe to re-run — it won't duplicate the Recipients sheet, the button,
 * or the triggers.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureRecipientsSheet(ss);
  ensureSendButton(ss);
  installTriggers();

  Logger.log('Setup complete. Weekly send is scheduled; the Dashboard checkbox is live.');
}

/**
 * Creates the Recipients tab with a header + sample rows, if it
 * doesn't already exist.
 */
function ensureRecipientsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.RECIPIENTS_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(CONFIG.RECIPIENTS_SHEET_NAME);
  sheet.getRange('A1:B1')
    .setValues([['Email', 'Name (optional)']])
    .setFontWeight('bold')
    .setBackground('#1F3864')
    .setFontColor('#FFFFFF');
  sheet.getRange('A2:B3').setValues([
    ['stakeholder1@example.com', 'Stakeholder One'],
    ['stakeholder2@example.com', 'Stakeholder Two'],
  ]);
  sheet.setColumnWidths(1, 2, 220);
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Adds a checkbox + label to the Dashboard tab, if not already present.
 */
function ensureSendButton(ss) {
  const dash = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
  if (!dash) {
    throw new Error('Sheet "' + CONFIG.DASHBOARD_SHEET_NAME + '" not found.');
  }

  const buttonCell = dash.getRange(CONFIG.BUTTON_CELL);
  if (!buttonCell.getDataValidation()) {
    buttonCell.insertCheckboxes();
  }
  buttonCell.setValue(false);

  dash.getRange(CONFIG.BUTTON_LABEL_CELL)
    .setValue(CONFIG.BUTTON_LABEL_TEXT)
    .setFontWeight('bold')
    .setFontColor('#1F3864');
}

/**
 * Installs both triggers, clearing any duplicates from a previous run.
 */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    const fn = trigger.getHandlerFunction();
    if (fn === 'sendDashboardPDF' || fn === 'onEditInstallable') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Weekly scheduled send.
  ScriptApp.newTrigger('sendDashboardPDF')
    .timeBased()
    .onWeekDay(CONFIG.TRIGGER_DAY)
    .atHour(CONFIG.TRIGGER_HOUR)
    .create();

  // Watches the checkbox button. An *installable* onEdit trigger is
  // required (not a simple onEdit) because sending mail and hitting
  // Drive/UrlFetch needs full authorization.
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

/**
 * Installable onEdit handler. Fires on every edit to the spreadsheet;
 * only acts when the Dashboard button checkbox was just ticked.
 */
function onEditInstallable(e) {
  const range = e.range;
  const sheet = range.getSheet();

  const isButton =
    sheet.getName() === CONFIG.DASHBOARD_SHEET_NAME &&
    range.getA1Notation() === CONFIG.BUTTON_CELL &&
    e.value === 'TRUE';

  if (!isButton) return;

  try {
    sendDashboardPDF();
  } finally {
    // Reset the checkbox so it behaves like a momentary button.
    range.setValue(false);
  }
}

/**
 * Main routine: export the dashboard tab to PDF, save to Drive, email it
 * to everyone listed on the Recipients tab.
 */
function sendDashboardPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + CONFIG.DASHBOARD_SHEET_NAME + '" not found.');
  }

  const recipients = getRecipients(ss);
  if (recipients.length === 0) {
    Logger.log('No recipients found on the "' + CONFIG.RECIPIENTS_SHEET_NAME + '" tab — nothing sent.');
    return;
  }

  const pdfBlob = exportSheetAsPdf(ss, sheet);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  pdfBlob.setName(CONFIG.DASHBOARD_SHEET_NAME + '_' + today + '.pdf');

  const folder = getOrCreateFolder(CONFIG.DRIVE_FOLDER_NAME);
  const file = folder.createFile(pdfBlob);

  emailStakeholders(recipients, pdfBlob, today, file);

  Logger.log('PDF saved to Drive: ' + file.getUrl());
  Logger.log('Email sent to: ' + recipients.join(', '));
}

/**
 * Reads every non-empty email in column A of the Recipients tab
 * (skipping the header row).
 */
function getRecipients(ss) {
  const sheet = ss.getSheetByName(CONFIG.RECIPIENTS_SHEET_NAME);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .map(function (row) { return String(row[0]).trim(); })
    .filter(function (email) { return email.length > 0 && email.indexOf('@') > -1; });
}

/**
 * Exports a single sheet (by gid) as a clean, print-formatted PDF blob.
 */
function exportSheetAsPdf(ss, sheet) {
  const url = ss.getUrl().replace(/edit.*$/, '');
  const exportUrl = url +
    'export?format=pdf' +
    '&gid=' + sheet.getSheetId() +
    '&size=A4' +
    '&portrait=false' +          // landscape — better for wide dashboards
    '&fitw=true' +               // fit to page width
    '&sheetnames=false&printtitle=false&pagenumbers=false' +
    '&gridlines=false' +
    '&fzr=false' +                // don't repeat frozen rows
    '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  });

  return response.getBlob().setContentType('application/pdf');
}

/**
 * Finds (or creates) the Drive folder PDFs are saved into.
 */
function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(name);
}

/**
 * Sends the PDF to every recipient.
 */
function emailStakeholders(recipients, pdfBlob, dateStr, driveFile) {
  const subject = CONFIG.EMAIL_SUBJECT.replace('{{date}}', dateStr);
  const body = CONFIG.EMAIL_BODY.replace('{{date}}', dateStr) +
    '\n\nDrive link: ' + driveFile.getUrl();

  recipients.forEach(function (recipient) {
    GmailApp.sendEmail(recipient, subject, body, {
      attachments: [pdfBlob],
      name: 'Muthus Cooporation',
    });
  });
}
