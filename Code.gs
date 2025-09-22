// --- MAIN FUNCTION: run once daily ---
function scheduleEmailsAllSheets() {
  const now = new Date();

  // STOP if after 11 AM
  if (now.getHours() >= 11) {
    Logger.log("After 11 AM — no emails scheduled today.");
    return;
  }

  const day = now.getDay();
  if (day < 2 || day > 5) { // Tue-Fri only
    Logger.log("Not a weekday for sending emails.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Remove old triggers for sendSingleEmail
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendSingleEmail") {
      ScriptApp.deleteTrigger(t);
    }
  });

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const nameIndex = headers.indexOf("Name");
    const emailIndex = headers.indexOf("Email");
    const schoolIndex = headers.indexOf("School");
    const companyIndex = headers.indexOf("Company");
    const statusIndex = headers.indexOf("Status");

    const pendingContacts = [];
    for (let i = 1; i < data.length && pendingContacts.length < 3; i++) {
      if (data[i][statusIndex] === "Pending") {
        pendingContacts.push({
          row: i + 1,
          name: data[i][nameIndex],
          email: data[i][emailIndex],
          school: data[i][schoolIndex],
          company: data[i][companyIndex],
          sheetName: sheet.getName()
        });
      }
    }

    if (pendingContacts.length === 0) {
      Logger.log(`No pending contacts on sheet: ${sheet.getName()}`);
      return;
    }

    // Generate random times for these contacts ≥45 min apart
    const randomTimes = generateSpacedRandomTimes(pendingContacts.length, 8, 11);

    // Schedule each email individually
    pendingContacts.forEach((contact, i) => {
      const sendTime = randomTimes[i];

      // Save contact info to ScriptProperties uniquely
      const propKey = `contact_${contact.sheetName}_${contact.row}`;
      const props = PropertiesService.getScriptProperties();
      props.setProperty(propKey, JSON.stringify(contact));

      ScriptApp.newTrigger("sendSingleEmail")
        .timeBased()
        .at(sendTime)
        .create();

      Logger.log(`Scheduled email for ${contact.name} (${contact.sheetName}) at ${sendTime}`);
    });
  });
}

// --- HELPER: generate times ≥45 min apart ---
function generateSpacedRandomTimes(count, startHour, endHour) {
  const times = [];
  while (times.length < count) {
    const hour = startHour + Math.floor(Math.random() * (endHour - startHour));
    let minute = Math.floor(Math.random() * 60);
    if (minute % 5 === 0) continue; // skip multiples of 5

    const candidate = new Date();
    candidate.setHours(hour, minute, 0, 0);

    const tooClose = times.some(t => Math.abs((candidate - t) / (1000 * 60)) < 45);
    if (!tooClose) {
      times.push(candidate);
    }
  }
  times.sort((a, b) => a - b);
  return times;
}

// --- FUNCTION CALLED BY TRIGGER ---
function sendSingleEmail(e) {
  const props = PropertiesService.getScriptProperties();
  const keys = Object.keys(props.getProperties());

  // Each trigger should only send the email for its contact
  for (let key of keys) {
    if (!key.startsWith("contact_")) continue;

    const contact = JSON.parse(props.getProperty(key));
    if (!contact) continue;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(contact.sheetName);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusIndex = headers.indexOf("Status");

    // Send the email
    const dayGreeting = getDayGreeting(new Date().getDay());
    const subject = `Texas A&M Student interested in IB at ${contact.company}`;
    const body = buildEmailBody(contact.name, contact.school, contact.company, dayGreeting);
    GmailApp.sendEmail(contact.email, subject, body);

    // Update Status
    sheet.getRange(contact.row, statusIndex + 1).setValue("Sent");

    Logger.log(`Sent email to ${contact.name} (${contact.sheetName}) at ${new Date()}`);

    // Remove property so trigger won’t resend
    props.deleteProperty(key);
  }
}

// --- DAILY STATUS REPORT ---
function sendDailyStatusReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sentCount = 0;

  sheets.forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusIndex = headers.indexOf("Status");

    for (let i = 1; i < data.length; i++) {
      if (data[i][statusIndex] === "Sent") sentCount++;
    }
  });

  const subject = "Daily Email Status Report";
  const body = sentCount === 0
    ? "No emails were sent today."
    : `A total of ${sentCount} emails were sent today.`;

  GmailApp.sendEmail("nikhilprabh32@gmail.com", subject, body);
}

// --- HELPERS ---
function getDayGreeting(day) {
  switch(day) {
    case 1: return "a great start to your week";
    case 2: return "a great start to your week";
    case 3: return "a great week so far!";
    case 4: return "a great week so far!";
    case 5: return "a great week";
    default: return "a great day";
  }
}

function buildEmailBody(name, school, company, dayGreeting) {
  return `Hi ${name},

I hope you're having ${dayGreeting}. My name is Nik, and I am a sophomore finance major at Texas A&M University interested in Investment Banking opportunities at ${company}. 

If possible, would you be open to a brief 10 - 15-minute call in the coming weeks to speak about your experience at the firm as well as your journey from ${school}? I am more than happy to work around your schedule.

In case it's helpful to provide more context on my background, I have attached my resume below for your reference. I look forward to hearing back from you soon!

Best,

Nik`;
}
