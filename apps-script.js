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
const FRICTION_EMAIL = 'friction@noverseinc.com';

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
  reg.getRange(1, 1, 1, 23).setValues([[
    'Timestamp', 'Team Code', 'Team Name', 'Team Description', 'Lead Phone',
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
  if (e.parameter && e.parameter.payload) {
    try {
      const data = JSON.parse(e.parameter.payload);

      if (data.type === 'register') {
        return handleRegistration(data);
      } else if (data.type === 'submit') {
        return handleSubmission(data);
      }

      return jsonResponse({ success: false, message: 'Unknown type' });
    } catch (err) {
      return jsonResponse({ success: false, message: err.toString() });
    }
  }
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
    data.teamDescription,
    data.leadPhone || ''
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
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #000000; color: #e8edf2; border-radius: 12px; overflow: hidden;">

      <!-- Header -->
      <div style="background: linear-gradient(135deg, #111 0%, #000 100%); padding: 2.5rem 2.5rem 1.5rem; border-bottom: 2px solid #f5c542;">
        <h1 style="color: #f5c542; margin: 0; font-size: 1.8rem; letter-spacing: 0.05em;">FRICTION</h1>
        <p style="color: #6b8299; margin: 0.3rem 0 0; font-size: 0.85rem;">Hackathon 2026</p>
      </div>

      <div style="padding: 2rem 2.5rem 2.5rem;">

        <p style="font-size: 1.1rem;">Hi <strong style="color: #2dd4bf;">${teamName}</strong>,</p>

        <p>Congratulations — your submission for Friction has been successfully received! We're thrilled to have you on board.</p>

        <p>To complete your registration, please make your participation payment of <strong style="color: #f5c542; font-size: 1.1rem;">&#2547;500 BDT</strong> using one of the options below.</p>

        <!-- Team Code Box -->
        <div style="background: #111; border: 2px solid #f5c542; padding: 1.5rem; border-radius: 10px; margin: 2rem 0; text-align: center;">
          <p style="color: #6b8299; font-size: 0.75rem; margin: 0 0 0.5rem; text-transform: uppercase; letter-spacing: 0.15em;">Your Team Code</p>
          <h1 style="color: #f5c542; letter-spacing: 0.3em; font-size: 3rem; margin: 0; font-family: 'Courier New', monospace;">${teamCode}</h1>
          <p style="color: #6b8299; font-size: 0.8rem; margin: 0.5rem 0 0;">You'll need this code when making your payment</p>
        </div>

        <!-- IMPORTANT WARNING -->
        <div style="background: #1a0a0a; padding: 1.2rem 1.5rem; border-radius: 8px; border-left: 4px solid #e8395b; margin: 1.5rem 0 2rem;">
          <p style="margin: 0; font-size: 0.95rem;">
            <strong style="color: #e8395b;">IMPORTANT:</strong> You <strong>must</strong> include your Team Code <strong style="color: #f5c542;">${teamCode}</strong> in the reference/note field when making payment. Payments without a Team Code cannot be verified.
          </p>
        </div>

        <!-- Option 1: bKash -->
        <div style="background: #0a0a0a; border: 1px solid #1a1a1a; border-radius: 10px; padding: 1.5rem; margin-bottom: 1.5rem;">
          <h3 style="color: #2dd4bf; margin: 0 0 1rem; font-size: 1rem;">Option 1 — bKash (Personal)</h3>
          <table style="width: 100%; color: #e8edf2; font-size: 0.9rem;">
            <tr><td style="padding: 6px 0; color: #6b8299; width: 100px;">Send to</td><td style="padding: 6px 0;"><strong style="color: #f5c542;">+8801537252941</strong> (Personal)</td></tr>
            <tr><td style="padding: 6px 0; color: #6b8299;">Amount</td><td style="padding: 6px 0;"><strong>&#2547;500 BDT</strong></td></tr>
            <tr><td style="padding: 6px 0; color: #6b8299;">Reference</td><td style="padding: 6px 0;"><strong style="color: #f5c542;">${teamCode}</strong></td></tr>
          </table>
        </div>

        <!-- Option 2: Bank -->
        <div style="background: #0a0a0a; border: 1px solid #1a1a1a; border-radius: 10px; padding: 1.5rem; margin-bottom: 2rem;">
          <h3 style="color: #2dd4bf; margin: 0 0 1rem; font-size: 1rem;">Option 2 — Bank Transfer</h3>
          <table style="width: 100%; color: #e8edf2; font-size: 0.9rem;">
            <tr><td style="padding: 6px 0; color: #6b8299; width: 130px;">Account Name</td><td style="padding: 6px 0;">TASFIA HASANAT KHAN</td></tr>
            <tr><td style="padding: 6px 0; color: #6b8299;">Account Number</td><td style="padding: 6px 0;">1071999750001</td></tr>
            <tr><td style="padding: 6px 0; color: #6b8299;">Bank</td><td style="padding: 6px 0;">BRAC Bank PLC</td></tr>
            <tr><td style="padding: 6px 0; color: #6b8299;">Branch</td><td style="padding: 6px 0;">DONIA BRANCH</td></tr>
            <tr><td style="padding: 6px 0; color: #6b8299;">Routing Number</td><td style="padding: 6px 0;">060271424</td></tr>
            <tr><td style="padding: 6px 0; color: #6b8299;">Reference</td><td style="padding: 6px 0;"><strong style="color: #f5c542;">Team Code: ${teamCode}</strong></td></tr>
          </table>
        </div>

        <!-- Steps -->
        <div style="background: #0d1117; border: 1px solid #1a1a1a; border-radius: 10px; padding: 1.5rem;">
          <h3 style="color: #e8edf2; margin: 0 0 1rem; font-size: 1rem;">After Payment</h3>
          <p style="margin: 0.5rem 0;"><span style="color: #f5c542; font-weight: bold;">Step 1</span> — Take a screenshot of the payment confirmation</p>
          <p style="margin: 0.5rem 0;"><span style="color: #f5c542; font-weight: bold;">Step 2</span> — Email the screenshot to <strong style="color: #2dd4bf;">frictionteam@gmail.com</strong></p>
          <p style="margin: 0.5rem 0;"><span style="color: #f5c542; font-weight: bold;">Step 3</span> — We'll verify and send your Battle Pass within <strong style="color: #2dd4bf;">24 hours</strong></p>
        </div>

        <p style="margin-top: 2rem;">If you run into any issues, email us at <strong style="color: #2dd4bf;">frictionteam@gmail.com</strong> — we're happy to help.</p>

        <p>See you at Friction!</p>

        <div style="border-top: 1px solid #1a1a1a; margin-top: 2rem; padding-top: 1.5rem;">
          <p style="color: #6b8299; margin: 0;">
            Warm regards,<br>
            <strong style="color: #e8edf2;">Friction Hackathon Team</strong><br>
            <span style="font-size: 0.8rem;">Noverse Inc</span>
          </p>
        </div>

      </div>
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
    // Columns: 0=timestamp, 1=teamCode, 2=teamName, 3=desc, 4=leadPhone, 5-8=m1, 9-12=m2, 13-16=m3, 17-20=m4, 21=paymentStatus, 22=battlePass
    const teamName = row[2].toString().toLowerCase();
    const paymentStatus = row[21];
    const battlePass = row[22];

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
    .addSeparator()
    .addItem('Send Launch Email to All Approved Teams', 'sendLaunchEmail')
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
  const leadEmail = sheet.getRange(row, 7).getValue();       // Column G = Member 1 Email
  const existingCode = sheet.getRange(row, 23).getValue();   // Column W = Battle Pass Code
  const currentStatus = sheet.getRange(row, 22).getValue();  // Column V = Payment Status

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
  sheet.getRange(row, 22).setValue('approved');  // Payment Status
  sheet.getRange(row, 23).setValue(code);        // Battle Pass Code

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

function sendLaunchEmail() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Send Launch Email',
    'This will email ALL approved teams that the hackathon has started and reveal the theme. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const sheet = getSheet('Registrations');
  const data = sheet.getDataRange().getValues();
  let sent = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const teamName = row[2];
    const leadEmail = row[6]; // Member 1 Email
    const paymentStatus = row[21];
    const battlePass = row[22];

    if (paymentStatus !== 'approved' || !leadEmail) continue;

    MailApp.sendEmail({
      to: leadEmail,
      replyTo: FRICTION_EMAIL,
      subject: 'FRICTION HAS STARTED — The Theme is Live!',
      htmlBody: `
        <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #000000; color: #e8edf2; border-radius: 12px; overflow: hidden;">

          <div style="background: linear-gradient(135deg, #111 0%, #000 100%); padding: 2.5rem 2.5rem 1.5rem; border-bottom: 2px solid #f5c542;">
            <h1 style="color: #f5c542; margin: 0; font-size: 1.8rem; letter-spacing: 0.05em;">FRICTION</h1>
            <p style="color: #e8395b; margin: 0.3rem 0 0; font-size: 0.85rem; font-weight: bold;">IT'S GO TIME</p>
          </div>

          <div style="padding: 2rem 2.5rem 2.5rem;">

            <p style="font-size: 1.1rem;">Team <strong style="color: #2dd4bf;">${teamName}</strong>,</p>

            <p>The wait is over. The theme has been revealed and the clock is ticking.</p>

            <!-- Theme -->
            <div style="background: #111; border: 2px solid #f5c542; padding: 2rem; border-radius: 10px; margin: 2rem 0; text-align: center;">
              <p style="color: #6b8299; font-size: 0.75rem; margin: 0 0 0.5rem; text-transform: uppercase; letter-spacing: 0.15em;">The Theme</p>
              <h1 style="color: #f5c542; font-size: 3rem; margin: 0; letter-spacing: 0.1em;">FRICTION</h1>
              <p style="color: #e8edf2; margin: 1rem 0 0; font-size: 0.95rem;">One word. Infinite interpretations.</p>
            </div>

            <div style="background: #0a0a0a; border: 1px solid #1a1a1a; border-radius: 10px; padding: 1.5rem; margin: 1.5rem 0;">
              <p style="margin: 0; line-height: 1.7;">
                Something slows everything down. We've spent decades trying to remove it — faster payments, smoother interfaces, instant everything.<br><br>
                But maybe we got that wrong.<br><br>
                Your job is not to remove friction. Your job is to <strong style="color: #f5c542;">understand</strong> it. Build something that takes a position.
              </p>
            </div>

            <p>We don't care what you build. We care <em>why</em>. Interpret the theme however you want — social friction, mechanical friction, cognitive friction, creative friction. Take a stance. Ship it.</p>

            <!-- Reminder -->
            <div style="background: #0d1117; border: 1px solid #1a1a1a; border-radius: 10px; padding: 1.5rem; margin: 1.5rem 0;">
              <h3 style="color: #2dd4bf; margin: 0 0 1rem; font-size: 1rem;">Quick Reminder</h3>
              <p style="margin: 0.3rem 0;"><strong style="color: #f5c542;">Time:</strong> 72 hours starting now</p>
              <p style="margin: 0.3rem 0;"><strong style="color: #f5c542;">Battle Pass:</strong> ${battlePass}</p>
              <p style="margin: 0.3rem 0;"><strong style="color: #f5c542;">Submit at:</strong> friction hackathon website > Final Submission tab</p>
            </div>

            <p>You've got 72 hours. Make them count.</p>

            <p>Good luck.</p>

            <div style="border-top: 1px solid #1a1a1a; margin-top: 2rem; padding-top: 1.5rem;">
              <p style="color: #6b8299; margin: 0;">
                <strong style="color: #e8edf2;">Friction Hackathon Team</strong><br>
                <span style="font-size: 0.8rem;">Noverse Inc</span>
              </p>
            </div>

          </div>
        </div>
      `
    });
    sent++;
  }

  ui.alert('Launch email sent to ' + sent + ' approved team(s)!');
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
