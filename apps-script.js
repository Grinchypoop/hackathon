// ============================================
// FRICTION HACKATHON — Google Apps Script
// ============================================
// 1. Open your Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this entire file into the editor (replace any existing code)
// 4. Click Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the Web App URL and paste it into index.html (replace APPS_SCRIPT_URL)
// ============================================

const SHEET_ID = '1WpbFdN8lA9y0PjnSYXUjxYmZanPL-LkkfugUkz-IQfM';
const FRICTION_EMAIL = 'frictionteam@noverseinc.com';

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

// Set up headers on first run
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let reg = ss.getSheetByName('Registrations');
  if (!reg) {
    reg = ss.insertSheet('Registrations');
  }
  reg.getRange(1, 1, 1, 22).setValues([[
    'Timestamp', 'Team Code', 'Team Name', 'Team Description',
    'Member 1 Name', 'Member 1 Email', 'Member 1 University', 'Member 1 Age',
    'Member 2 Name', 'Member 2 Email', 'Member 2 University', 'Member 2 Age',
    'Member 3 Name', 'Member 3 Email', 'Member 3 University', 'Member 3 Age',
    'Member 4 Name', 'Member 4 Email', 'Member 4 University', 'Member 4 Age',
    'Payment Status', 'Battle Pass Code'
  ]]);

  let sub = ss.getSheetByName('Submissions');
  if (!sub) {
    sub = ss.insertSheet('Submissions');
  }
  sub.getRange(1, 1, 1, 7).setValues([[
    'Timestamp', 'Team Name', 'Battle Pass Code',
    'Project URL', 'Source Code URL', 'Demo Video URL', 'Notes'
  ]]);
}

// Handle incoming POST requests
function doPost(e) {
  try {
    let data;
    // Handle both form submissions and raw JSON
    if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      return jsonResponse({ success: false, message: 'No data received' });
    }

    if (data.type === 'register') {
      return handleRegistration(data);
    } else if (data.type === 'submit') {
      return handleSubmission(data);
    }

    return jsonResponse({ success: false, message: 'Unknown request type' });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function doGet(e) {
  return jsonResponse({ status: 'ok' });
}

function handleRegistration(data) {
  const sheet = getSheet('Registrations');
  const members = data.members || [];

  // Check for duplicate team name
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const teamNames = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    for (let i = 0; i < teamNames.length; i++) {
      if (teamNames[i][0].toString().toLowerCase() === data.teamName.toLowerCase()) {
        return jsonResponse({ success: false, message: 'A team with this name already exists. Please choose a different name.' });
      }
    }
  }

  // Team code = row number (row 2 = team 1, row 3 = team 2, etc.)
  const teamCode = lastRow; // header is row 1, so first team is row 2 = code 1, etc.

  // Build row
  const row = [
    new Date().toISOString(),
    teamCode,
    data.teamName,
    data.teamDescription
  ];

  // Add member data (up to 4 members, 4 fields each)
  for (let i = 0; i < 4; i++) {
    if (members[i]) {
      row.push(members[i].name, members[i].email, members[i].university, members[i].age);
    } else {
      row.push('', '', '', '');
    }
  }

  // Payment status, Battle Pass Code
  row.push('pending_payment', '');

  sheet.appendRow(row);

  // Send confirmation email to team lead
  const leadEmail = members[0] ? members[0].email : '';
  if (leadEmail) {
    sendRegistrationEmail(data.teamName, teamCode, leadEmail);
  }

  return jsonResponse({
    success: true,
    message: 'Registration received! Check your email for payment instructions.',
    teamCode: teamCode
  });
}

function sendRegistrationEmail(teamName, teamCode, leadEmail) {
  const subject = 'Welcome to Friction! Your Submission is Confirmed';

  const htmlBody = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #000000; color: #e8edf2; padding: 2.5rem; border-radius: 12px;">
      <h2 style="color: #f5c542; margin-bottom: 1.5rem;">Welcome to Friction!</h2>

      <p>Hi <strong style="color: #2dd4bf;">${teamName}</strong>,</p>

      <p>Congratulations — your submission for Friction has been successfully received! We're thrilled to have you on board.</p>

      <p>To complete your registration, please make your participation payment using one of the methods below.</p>

      <div style="background: #111; border: 1px solid #1a1a1a; padding: 1.5rem; border-radius: 8px; margin: 1.5rem 0; text-align: center;">
        <p style="color: #6b8299; font-size: 0.8rem; margin-bottom: 0.5rem;">YOUR TEAM CODE</p>
        <h1 style="color: #f5c542; letter-spacing: 0.2em; font-size: 2.5rem; margin: 0;">${teamCode}</h1>
      </div>

      <p>You'll need this code when making your payment — please keep it handy.</p>

      <hr style="border: none; border-top: 1px solid #1a1a1a; margin: 2rem 0;">

      <h3 style="color: #2dd4bf;">Option 1: bKash</h3>
      <p><strong>Step 1</strong> — Open your bKash app and ensure your account is active.</p>
      <p><strong>Step 2</strong> — Send payment of <strong style="color: #f5c542;">৳500 BDT</strong> to:<br>
      <span style="color: #f5c542;">+8801537252941</span></p>
      <p><strong>Step 3</strong> — Enter your Team Code <strong style="color: #f5c542;">${teamCode}</strong> in the Reference field.</p>
      <p><strong>Step 4</strong> — Take a screenshot of the confirmation and reply to this email.</p>

      <hr style="border: none; border-top: 1px solid #1a1a1a; margin: 2rem 0;">

      <h3 style="color: #2dd4bf;">Option 2: Bank Transfer</h3>
      <table style="color: #e8edf2; font-size: 0.9rem; margin: 0.5rem 0;">
        <tr><td style="padding: 4px 12px 4px 0; color: #6b8299;">Account Name</td><td>TASFIA HASANAT KHAN</td></tr>
        <tr><td style="padding: 4px 12px 4px 0; color: #6b8299;">Account Number</td><td>1071999750001</td></tr>
        <tr><td style="padding: 4px 12px 4px 0; color: #6b8299;">Bank</td><td>BRAC Bank PLC</td></tr>
        <tr><td style="padding: 4px 12px 4px 0; color: #6b8299;">Branch</td><td>DONIA BRANCH</td></tr>
        <tr><td style="padding: 4px 12px 4px 0; color: #6b8299;">Routing Number</td><td>060271424</td></tr>
        <tr><td style="padding: 4px 12px 4px 0; color: #6b8299;">Reference</td><td style="color: #f5c542;">Team Code: ${teamCode}</td></tr>
      </table>
      <p>After transfer, take a screenshot and reply to this email.</p>

      <hr style="border: none; border-top: 1px solid #1a1a1a; margin: 2rem 0;">

      <p style="background: #1a1a1a; padding: 1rem; border-radius: 6px; border-left: 3px solid #e8395b;">
        <strong style="color: #e8395b;">Important:</strong> Payments without a Team Code in the reference field cannot be verified. Please double-check before submitting.
      </p>

      <p>If you run into any issues, don't hesitate to reach out — we're happy to help.</p>

      <p>See you at Friction!</p>

      <p style="color: #6b8299; margin-top: 2rem;">
        Warm regards,<br>
        <strong style="color: #e8edf2;">Friction Hackathon Team</strong>
      </p>
    </div>
  `;

  MailApp.sendEmail({
    to: leadEmail,
    replyTo: FRICTION_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}

function handleSubmission(data) {
  const regSheet = getSheet('Registrations');
  const subSheet = getSheet('Submissions');

  // Validate battle pass
  const regData = regSheet.getDataRange().getValues();
  let valid = false;

  for (let i = 1; i < regData.length; i++) {
    const row = regData[i];
    // Columns: 0=timestamp, 1=teamCode, 2=teamName, 3=desc, 4-7=m1, 8-11=m2, 12-15=m3, 16-19=m4, 20=paymentStatus, 21=battlePass
    const teamName = row[2].toString().toLowerCase();
    const paymentStatus = row[20];
    const battlePass = row[21];

    if (teamName === data.teamName.toLowerCase() &&
        battlePass.toString().toUpperCase() === data.battlePassCode.toUpperCase() &&
        paymentStatus === 'approved') {
      valid = true;
      break;
    }
  }

  if (!valid) {
    return jsonResponse({ success: false, message: 'Invalid Battle Pass code or team name. Make sure your payment has been approved and you\'re using the correct team name.' });
  }

  // Check for duplicate submission
  const subData = subSheet.getDataRange().getValues();
  for (let i = 1; i < subData.length; i++) {
    if (subData[i][2].toString().toUpperCase() === data.battlePassCode.toUpperCase()) {
      return jsonResponse({ success: false, message: 'A submission with this Battle Pass has already been received.' });
    }
  }

  subSheet.appendRow([
    new Date().toISOString(),
    data.teamName,
    data.battlePassCode,
    data.projectUrl,
    data.sourceCodeUrl || '',
    data.demoVideoUrl,
    data.notes || ''
  ]);

  return jsonResponse({ success: true, message: 'Project submitted! Good luck!' });
}

// ============================================
// BATTLE PASS — Approve & Send
// ============================================
// Select a row in the Registrations sheet, then run
// Friction > Approve & Send Battle Pass

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Friction')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Approve & Send Battle Pass', 'sendBattlePassForSelectedRow')
    .addToUi();
}

function sendBattlePassForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a registration row (not the header).');
    return;
  }

  const teamName = sheet.getRange(row, 3).getValue();       // Column C = Team Name
  const leadEmail = sheet.getRange(row, 6).getValue();       // Column F = Member 1 Email
  const existingCode = sheet.getRange(row, 22).getValue();   // Column V = Battle Pass Code
  const currentStatus = sheet.getRange(row, 21).getValue();  // Column U = Payment Status

  if (!leadEmail) {
    SpreadsheetApp.getUi().alert('No lead email found for this row.');
    return;
  }

  if (existingCode) {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert('Battle Pass already exists: ' + existingCode + '\n\nResend email?', ui.ButtonSet.YES_NO);
    if (result === ui.Button.NO) return;
  }

  // Generate battle pass code
  const code = existingCode || generateBattlePass();

  // Update sheet
  sheet.getRange(row, 21).setValue('approved');  // Payment Status
  sheet.getRange(row, 22).setValue(code);        // Battle Pass Code

  // Email the battle pass
  MailApp.sendEmail({
    to: leadEmail,
    replyTo: FRICTION_EMAIL,
    subject: 'FRICTION Hackathon — Your Battle Pass is Here!',
    htmlBody: `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #000000; color: #e8edf2; padding: 2.5rem; border-radius: 12px;">
        <h2 style="color: #f5c542;">Your Payment is Approved!</h2>

        <p>Team <strong style="color: #2dd4bf;">${teamName}</strong>, you're officially in.</p>

        <div style="background: #111; border: 1px solid #1a1a1a; padding: 1.5rem; border-radius: 8px; margin: 1.5rem 0; text-align: center;">
          <p style="color: #6b8299; font-size: 0.8rem; margin-bottom: 0.5rem;">YOUR BATTLE PASS</p>
          <h1 style="color: #f5c542; letter-spacing: 0.2em; font-size: 2.5rem; margin: 0;">${code}</h1>
        </div>

        <p>Use this code along with your <strong>exact team name</strong> to submit your final project when the time comes.</p>

        <p style="background: #1a1a1a; padding: 1rem; border-radius: 6px; border-left: 3px solid #f5c542;">
          Keep this code safe. <strong>One code per team.</strong> You won't be able to submit without it.
        </p>

        <p>Stay tuned for the theme reveal and get ready to build!</p>

        <p style="color: #6b8299; margin-top: 2rem;">
          Warm regards,<br>
          <strong style="color: #e8edf2;">Friction Hackathon Team</strong>
        </p>
      </div>
    `
  });

  SpreadsheetApp.getUi().alert('Battle Pass ' + code + ' sent to ' + leadEmail);
}

function generateBattlePass() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 4; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  code += '-';
  for (let i = 0; i < 4; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
