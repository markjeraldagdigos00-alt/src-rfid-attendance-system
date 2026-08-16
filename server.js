const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const nodemailer = require('nodemailer');
const https = require('https');
const querystring = require('querystring');

const app = express();
const PORT = 3000;

// Setup Storage para sa Uploaded Logos
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, 'school_logo' + ext);
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) {
      cb(null, true);
    } else {
      cb(new Error('Only image files are allowed!'), false);
    }
  }
});

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/uploads', express.static(uploadsDir));

const STUDENTS_FILE = path.join(__dirname, 'students.json');
const ATTENDANCE_FILE = path.join(__dirname, 'attendance.json');
const CONFIG_FILE = path.join(__dirname, 'config.json');

const loadData = (file, fallback) => fs.existsSync(file) ? JSON.parse(fs.readFileSync(file)) : fallback;
const saveData = (file, data) => fs.writeFileSync(file, JSON.stringify(data, null, 2));

let students = loadData(STUDENTS_FILE, []);
let attendance = loadData(ATTENDANCE_FILE, []);
let config = loadData(CONFIG_FILE, { 
  systemName: 'RFID Attendance System',
  logoPath: '',
  events: ['General Event', 'Faculty Meeting', 'Orientation', 'Seminar'], 
  currentEvent: 'General Event', 
  cutoffTime: '08:00', 
  latestUid: '',
  enableEmail: false,
  gmailUser: '',
  gmailPass: '',
  enableSms: false,
  semaphoreApiKey: ''
});

if (!Array.isArray(config.events)) config.events = ['General Event'];

// Function para magpadala ng Email via Nodemailer
function sendEmailNotification(recipientEmail, studentName, scanType, status, eventName, timestamp, duration) {
  if (!config.enableEmail || !config.gmailUser || !config.gmailPass || !recipientEmail) return;

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: config.gmailUser,
      pass: config.gmailPass
    }
  });

  const durationText = duration ? `<li><strong>Duration:</strong> ${duration}</li>` : '';

  const mailOptions = {
    from: `"${config.systemName || 'RFID Attendance System'}" <${config.gmailUser}>`,
    to: recipientEmail,
    subject: `[${scanType}] Attendance Alert: ${studentName}`,
    html: `
      <div style="font-family: Arial, sans-serif; padding: 15px; border: 1px solid #ddd; border-radius: 6px;">
        <h2 style="color: #2c3e50;">Attendance Notification (${scanType})</h2>
        <p>Hello,</p>
        <p>This is an automated notification to confirm that <strong>${studentName}</strong> has logged <strong>${scanType}</strong>.</p>
        <ul>
          <li><strong>Event:</strong> ${eventName}</li>
          <li><strong>Scan Type:</strong> <span style="color:#2980b9; font-weight:bold;">${scanType}</span></li>
          <li><strong>Status:</strong> <span style="color:${status === 'LATE' ? '#e74c3c' : '#2ecc71'}; font-weight:bold;">${status}</span></li>
          <li><strong>Time:</strong> ${timestamp}</li>
          ${durationText}
        </ul>
        <p>Thank you!</p>
      </div>
    `
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log('[EMAIL ERROR]', error.message);
    } else {
      console.log('[EMAIL SENT]', info.response);
    }
  });
}

// Function para magpadala ng SMS via Semaphore API
function sendSMSNotification(phoneNumber, studentName, scanType, status, eventName, timestamp, duration) {
  if (!config.enableSms || !config.semaphoreApiKey || !phoneNumber) return;

  let message = `[Attendance] ${studentName} logged ${scanType} for ${eventName} at ${timestamp}. Status: ${status}.`;
  if (duration) {
    message += ` Duration: ${duration}.`;
  }

  const postData = querystring.stringify({
    apikey: config.semaphoreApiKey,
    number: phoneNumber,
    message: message
  });

  const options = {
    hostname: 'api.semaphore.co',
    port: 443,
    path: '/api/v4/messages',
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': postData.length
    }
  };

  const req = https.request(options, (res) => {
    res.on('data', (d) => {
      console.log('[SMS RESPONSE]', d.toString());
    });
  });

  req.on('error', (e) => {
    console.error('[SMS ERROR]', e.message);
  });

  req.write(postData);
  req.end();
}

// Helper: Calculate Time Difference
function calculateDuration(timeInStr, timeOutStr) {
  const tIn = new Date(timeInStr);
  const tOut = new Date(timeOutStr);
  const diffMs = tOut - tIn;

  if (isNaN(diffMs) || diffMs < 0) return 'N/A';

  const totalMinutes = Math.floor(diffMs / (1000 * 60));
  const hours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;

  if (hours > 0) {
    return `${hours}h ${mins}m`;
  }
  return `${mins}m`;
}

// API: ESP8266 RFID Scan (With Auto Time-In / Time-Out Logic)
app.post('/api/scan', (req, res) => {
  const { uid } = req.body;
  if (!uid) return res.status(400).json({ status: 'error', message: 'No UID' });

  const cleanUid = uid.trim().toUpperCase();
  config.latestUid = cleanUid;
  saveData(CONFIG_FILE, config);

  const student = students.find(s => s.uid.trim().toUpperCase() === cleanUid);
  const now = new Date();
  const todayStr = now.toLocaleDateString();

  if (student) {
    const eventName = student.assignedEvent || config.currentEvent || 'General Event';

    // Hanapin ang pinakabagong scan ngayong araw ng kaparehong estudyante para sa parehong event
    const todayLogs = attendance.filter(a => 
      a.uid === cleanUid && 
      a.event === eventName && 
      new Date(a.rawTimestamp).toLocaleDateString() === todayStr
    );

    const lastLog = todayLogs.length > 0 ? todayLogs[todayLogs.length - 1] : null;

    let scanType = 'TIME-IN';
    let duration = '';
    let statusLabel = 'ON TIME';

    // Kung nakapag-TIME IN na ngayong araw at wala pang kasunod na TIME OUT -> magiging TIME OUT na ito
    if (lastLog && lastLog.scanType === 'TIME-IN') {
      scanType = 'TIME-OUT';
      statusLabel = 'COMPLETED';
      duration = calculateDuration(lastLog.rawTimestamp, now.toISOString());
    } else {
      // Kung bagong TIME-IN: Check Late Cutoff
      const currentTimeStr = now.toTimeString().slice(0, 5);
      const isLate = currentTimeStr > (config.cutoffTime || '08:00');
      statusLabel = isLate ? 'LATE' : 'ON TIME';
    }

    const record = {
      uid: cleanUid,
      name: student.name,
      studentId: student.studentId,
      yearLevel: student.yearLevel || 'N/A',
      section: student.section || 'N/A',
      position: student.position || 'N/A',
      email: student.email || '',
      phone: student.phone || '',
      event: eventName,
      scanType: scanType,
      status: statusLabel,
      duration: duration || 'N/A',
      timestamp: now.toLocaleString(),
      rawTimestamp: now.toISOString()
    };
    
    attendance.push(record);
    saveData(ATTENDANCE_FILE, attendance);

    // Trigger Notifications
    if (student.email) {
      sendEmailNotification(student.email, student.name, scanType, statusLabel, eventName, record.timestamp, duration);
    }
    if (student.phone) {
      sendSMSNotification(student.phone, student.name, scanType, statusLabel, eventName, record.timestamp, duration);
    }

    console.log(`[SCAN] ${student.name} logged [${scanType}] for ${eventName} [${statusLabel}]`);
    return res.json({ 
      status: 'success', 
      scanType, 
      isLate: statusLabel === 'LATE', 
      message: `${scanType} recorded for ${student.name}` 
    });
  } else {
    return res.json({ status: 'unknown', message: 'Card not registered' });
  }
});

// API: Realtime Data
app.get('/api/live-data', (req, res) => {
  res.json({
    latestUid: config.latestUid || '',
    attendance: attendance.slice().reverse(),
    students: students
  });
});

// API: Update System Name
app.post('/api/update-system-name', (req, res) => {
  const { systemName } = req.body;
  if (systemName) {
    config.systemName = systemName;
    saveData(CONFIG_FILE, config);
  }
  res.redirect('/');
});

// API: Upload Logo
app.post('/api/upload-logo', upload.single('logoFile'), (req, res) => {
  if (req.file) {
    config.logoPath = `/uploads/${req.file.filename}?v=${Date.now()}`;
    saveData(CONFIG_FILE, config);
  }
  res.redirect('/');
});

// API: Remove Logo
app.post('/api/remove-logo', (req, res) => {
  config.logoPath = '';
  saveData(CONFIG_FILE, config);
  res.redirect('/');
});

// API: Save Notification Config
app.post('/api/notification-settings', (req, res) => {
  const { enableEmail, gmailUser, gmailPass, enableSms, semaphoreApiKey } = req.body;
  
  config.enableEmail = enableEmail === 'on';
  config.gmailUser = gmailUser || '';
  if (gmailPass && gmailPass !== '******') config.gmailPass = gmailPass;
  
  config.enableSms = enableSms === 'on';
  config.semaphoreApiKey = semaphoreApiKey || '';

  saveData(CONFIG_FILE, config);
  res.redirect('/');
});

// API: Update Event & Cutoff Settings
app.post('/api/event-settings', (req, res) => {
  const { newEvent, activeEvent, cutoffTime } = req.body;
  
  if (newEvent && !config.events.includes(newEvent)) {
    config.events.push(newEvent);
    config.currentEvent = newEvent;
  } else if (activeEvent) {
    config.currentEvent = activeEvent;
  }
  if (cutoffTime) config.cutoffTime = cutoffTime;

  saveData(CONFIG_FILE, config);
  res.redirect('/');
});

// API: Delete Event
app.post('/api/delete-event', (req, res) => {
  const { eventToDelete } = req.body;
  if (eventToDelete) {
    config.events = config.events.filter(e => e !== eventToDelete);
    if (config.currentEvent === eventToDelete) {
      config.currentEvent = config.events[0] || 'General Event';
    }
    saveData(CONFIG_FILE, config);
  }
  res.redirect('/');
});

// API: Register Student / Person
app.post('/api/register', (req, res) => {
  const { uid, name, studentId, yearLevel, section, assignedEvent, position, customPosition, email, phone } = req.body;
  if (uid && name && studentId) {
    const cleanUid = uid.trim().toUpperCase();
    const existingIndex = students.findIndex(s => s.uid === cleanUid);
    
    let finalPosition = 'N/A';
    if (position === 'Other' && customPosition) {
      finalPosition = customPosition.trim();
    } else if (position) {
      finalPosition = position;
    }

    const studentData = { 
      uid: cleanUid, 
      name, 
      studentId, 
      yearLevel: yearLevel || 'Grade 7', 
      section: section || 'A', 
      position: finalPosition,
      email: email || '',
      phone: phone || '',
      assignedEvent 
    };

    if (existingIndex > -1) {
      students[existingIndex] = studentData;
    } else {
      students.push(studentData);
    }
    saveData(STUDENTS_FILE, students);
  }
  res.redirect('/');
});

// API: Delete Student
app.post('/api/delete-student', (req, res) => {
  const { uid } = req.body;
  if (uid) {
    students = students.filter(s => s.uid !== uid);
    saveData(STUDENTS_FILE, students);
  }
  res.redirect('/');
});

// API: Export CSV
app.get('/api/export-csv', (req, res) => {
  const selectedEvent = req.query.event;
  let filtered = attendance;
  
  if (selectedEvent && selectedEvent !== 'ALL') {
    filtered = attendance.filter(a => a.event === selectedEvent);
  }

  let csv = 'UID,ID Number,Name,Grade Level,Section,Position,Email,Phone,Event,Type,Status,Duration,Timestamp\n';
  filtered.forEach(row => {
    csv += `"${row.uid}","${row.studentId}","${row.name}","${row.yearLevel}","${row.section}","${row.position || 'N/A'}","${row.email || ''}","${row.phone || ''}","${row.event}","${row.scanType || 'TIME-IN'}","${row.status}","${row.duration || 'N/A'}","${row.timestamp}"\n`;
  });

  const filename = selectedEvent && selectedEvent !== 'ALL' 
    ? `attendance_${selectedEvent.replace(/\s+/g, '_')}.csv` 
    : 'attendance_all_events.csv';

  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.status(200).send(csv);
});

// API: Clear Logs
app.post('/api/clear-logs', (req, res) => {
  attendance = [];
  saveData(ATTENDANCE_FILE, attendance);
  res.redirect('/');
});

// Web Dashboard
app.get('/', (req, res) => {
  const eventList = Array.isArray(config.events) ? config.events : ['General Event'];
  const eventOptions = eventList.map(e => `<option value="${e}" ${e === config.currentEvent ? 'selected' : ''}>${e}</option>`).join('');
  
  let gradeOptions = '';
  for (let i = 7; i <= 12; i++) {
    gradeOptions += `<option value="Grade ${i}">Grade ${i}</option>`;
  }

  const logoHtml = config.logoPath 
    ? `<img src="${config.logoPath}" alt="School Logo" class="header-logo">` 
    : '';

  const html = `
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${config.systemName || 'RFID Attendance System'}</title>
    <style>
      body { font-family: 'Segoe UI', Tahoma, sans-serif; margin: 20px; background: #f0f2f5; color: #333; }
      .header-container { display: flex; align-items: center; gap: 15px; margin-bottom: 20px; }
      .header-logo { height: 65px; width: auto; object-fit: contain; border-radius: 6px; }
      h1, h2, h3 { color: #1a252f; margin: 0; }
      .container { display: flex; flex-direction: column; gap: 20px; }
      .card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); flex: 1; min-width: 320px; }
      input, select { width: 100%; padding: 10px; margin: 6px 0 12px 0; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
      button, input[type="submit"] { background: #27ae60; color: white; padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
      .btn-danger { background: #e74c3c; padding: 6px 12px; font-size: 0.85em; }
      .btn-secondary { background: #2980b9; }
      table { width: 100%; border-collapse: collapse; margin-top: 10px; background: white; }
      th, td { border: 1px solid #e1e8ed; padding: 10px; text-align: left; }
      th { background: #34495e; color: white; }
      .badge-ontime { background: #2ecc71; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 0.85em; }
      .badge-late { background: #e74c3c; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 0.85em; }
      .badge-type-in { background: #3498db; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 0.85em; }
      .badge-type-out { background: #9b59b6; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 0.85em; }
      .delete-form { display: inline; }
      .section-divider { border: 0; height: 1px; background: #e1e8ed; margin: 15px 0; }
      
      details.settings-card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); padding: 15px 20px; }
      details.settings-card summary { font-size: 1.2em; font-weight: bold; color: #1a252f; cursor: pointer; user-select: none; }
      details.settings-card[open] summary { margin-bottom: 15px; }
      .hidden-field { display: none; }
    </style>
  </head>
  <body>
    <div class="header-container">
      ${logoHtml}
      <h1>${config.systemName || 'RFID Attendance System'}</h1>
    </div>

    <div class="container">
      
      <!-- Collapsible System Settings -->
      <details class="settings-card">
        <summary> System, Event & Notification Settings </summary>
        
        <div style="display: flex; flex-wrap: wrap; gap: 20px; margin-top: 10px;">
          <!-- System Name & Logo -->
          <div style="flex: 1; min-width: 280px;">
            <h3>System Name & Logo</h3>
            <form action="/api/update-system-name" method="POST">
              <label><strong>System Name:</strong></label>
              <input type="text" name="systemName" value="${config.systemName || 'RFID Attendance System'}" required>
              <input type="submit" value="Save System Name" style="width: 100%;">
            </form>

            <hr class="section-divider">

            <form action="/api/upload-logo" method="POST" enctype="multipart/form-data">
              <label><strong>School Logo Image:</strong></label>
              <input type="file" name="logoFile" accept="image/*" required>
              <input type="submit" value="Upload Logo" class="btn-secondary" style="width: 100%; margin-bottom: 10px;">
            </form>

            ${config.logoPath ? `
            <form action="/api/remove-logo" method="POST">
              <button type="submit" class="btn-danger" style="width: 100%;">Remove Logo</button>
            </form>
            ` : ''}
          </div>

          <!-- Event Management -->
          <div style="flex: 1; min-width: 280px;">
            <h3>Event Management</h3>
            <form action="/api/event-settings" method="POST">
              <label><strong>Active Event:</strong></label>
              <select name="activeEvent">${eventOptions}</select>

              <label><strong>Add New Event Name:</strong></label>
              <input type="text" name="newEvent" placeholder="e.g. Faculty Meeting / Sports Day">

              <label><strong>Late Cut-off Time (HH:MM):</strong></label>
              <input type="time" name="cutoffTime" value="${config.cutoffTime || '08:00'}">

              <input type="submit" value="Save Event Settings" style="width: 100%;">
            </form>

            <hr class="section-divider">

            <h3>Delete Event</h3>
            <form action="/api/delete-event" method="POST" onsubmit="return confirm('Delete this event?');">
              <select name="eventToDelete">${eventOptions}</select>
              <button type="submit" class="btn-danger" style="width: 100%;">Delete Selected Event</button>
            </form>
          </div>

          <!-- Notification API Settings -->
          <div style="flex: 1; min-width: 280px;">
            <h3>SMS & Email API Settings</h3>
            <form action="/api/notification-settings" method="POST">
              <label>
                <input type="checkbox" name="enableEmail" ${config.enableEmail ? 'checked' : ''} style="width: auto;">
                <strong>Enable Email Notifications (Nodemailer)</strong>
              </label>
              <input type="email" name="gmailUser" placeholder="Gmail Address" value="${config.gmailUser || ''}">
              <input type="password" name="gmailPass" placeholder="Gmail App Password" value="${config.gmailPass ? '******' : ''}">

              <hr class="section-divider">

              <label>
                <input type="checkbox" name="enableSms" ${config.enableSms ? 'checked' : ''} style="width: auto;">
                <strong>Enable SMS Notifications (Semaphore API)</strong>
              </label>
              <input type="text" name="semaphoreApiKey" placeholder="Semaphore API Key" value="${config.semaphoreApiKey || ''}">

              <input type="submit" value="Save Notification API Settings" style="width: 100%; margin-top: 10px;">
            </form>
          </div>
        </div>
      </details>

      <!-- Registration Form -->
      <div class="card">
        <h2>Register Participant / Student</h2>
        <p>Last Scanned UID: <strong id="scannedUid" style="color: #e67e22;">${config.latestUid || 'None'}</strong></p>
        <form action="/api/register" method="POST">
          <input type="text" id="uidInput" name="uid" placeholder="RFID Card UID" required>
          <input type="text" name="studentId" placeholder="ID Number" required>
          <input type="text" name="name" placeholder="Full Name" required>

          <!-- Contact Fields for SMS and Email -->
          <div style="display: flex; gap: 10px;">
            <div style="flex: 1;">
              <label><strong>Email Address (Optional):</strong></label>
              <input type="email" name="email" placeholder="e.g. parent@gmail.com">
            </div>
            <div style="flex: 1;">
              <label><strong>Phone Number (Optional):</strong></label>
              <input type="tel" name="phone" placeholder="e.g. 09171234567">
            </div>
          </div>

          <label><strong>Assign to Event:</strong></label>
          <select id="eventSelect" name="assignedEvent" onchange="checkMeetingEvent()">${eventOptions}</select>

          <!-- Student Grade & Section Fields -->
          <div id="studentFields" style="display: flex; gap: 10px;">
            <div style="flex: 1;">
              <label><strong>Grade Level:</strong></label>
              <select name="yearLevel">${gradeOptions}</select>
            </div>
            <div style="flex: 1;">
              <label><strong>Section:</strong></label>
              <input type="text" name="section" placeholder="e.g. Diamond / A">
            </div>
          </div>

          <!-- Meeting Position Selector Field -->
          <div id="meetingFields" class="hidden-field">
            <label><strong>Position / Role (For Meeting):</strong></label>
            <select id="positionSelect" name="position" onchange="checkCustomPosition()">
              <option value="Member">Member</option>
              <option value="Officer">Officer</option>
              <option value="President">President</option>
              <option value="Vice President">Vice President</option>
              <option value="Secretary">Secretary</option>
              <option value="Treasurer">Treasurer</option>
              <option value="Teacher / Faculty">Teacher / Faculty</option>
              <option value="Guest">Guest</option>
              <option value="Other">Custom Position...</option>
            </select>

            <div id="customPositionBox" class="hidden-field">
              <label><strong>Specify Custom Position:</strong></label>
              <input type="text" id="customPositionInput" name="customPosition" placeholder="Enter custom position/role">
            </div>
          </div>

          <button type="button" class="btn-secondary" onclick="useLatestUid()" style="width: 100%; margin-top: 10px; margin-bottom: 10px;">Use Last Scanned UID</button>
          <input type="submit" value="Register Card" style="width: 100%;">
        </form>
      </div>

    </div>

    <!-- Attendance Logs -->
    <div class="card" style="margin-top: 20px;">
      <h2>Live Attendance Log (Time-In / Time-Out)</h2>
      
      <div style="display: flex; gap: 10px; align-items: center; margin-bottom: 15px;">
        <label><strong>Export CSV for Event:</strong></label>
        <select id="exportEventSelect" style="width: auto; margin: 0;">
          <option value="ALL">All Events</option>
          ${eventOptions}
        </select>
        <button type="button" onclick="downloadCSV()">Download CSV</button>
        <form action="/api/clear-logs" method="POST" style="margin-left: auto;" onsubmit="return confirm('Clear logs?');">
          <button type="submit" class="btn-danger">Clear All Logs</button>
        </form>
      </div>

      <table>
        <thead>
          <tr>
            <th>Name</th>
            <th>ID Number</th>
            <th>Grade / Position</th>
            <th>Event</th>
            <th>Type</th>
            <th>Status</th>
            <th>Duration</th>
            <th>Timestamp</th>
            <th>UID</th>
          </tr>
        </thead>
        <tbody id="attendanceTableBody"></tbody>
      </table>
    </div>

    <!-- Registered Database Table -->
    <div class="card" style="margin-top: 20px;">
      <h2>Registered Participants Database</h2>
      <table>
        <thead>
          <tr>
            <th>ID Number</th>
            <th>Name</th>
            <th>Grade / Position</th>
            <th>Email / Phone</th>
            <th>Assigned Event</th>
            <th>Card UID</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody id="studentsTableBody"></tbody>
      </table>
    </div>

    <script>
      function checkMeetingEvent() {
        const eventVal = document.getElementById('eventSelect').value.toLowerCase();
        const meetingFields = document.getElementById('meetingFields');
        const studentFields = document.getElementById('studentFields');

        if (eventVal.includes('meeting')) {
          meetingFields.classList.remove('hidden-field');
          studentFields.classList.add('hidden-field');
        } else {
          meetingFields.classList.add('hidden-field');
          studentFields.classList.remove('hidden-field');
        }
      }

      function checkCustomPosition() {
        const posVal = document.getElementById('positionSelect').value;
        const customBox = document.getElementById('customPositionBox');
        if (posVal === 'Other') {
          customBox.classList.remove('hidden-field');
        } else {
          customBox.classList.add('hidden-field');
        }
      }

      async function updateDashboard() {
        try {
          const res = await fetch('/api/live-data');
          const data = await res.json();
          if (data.latestUid) document.getElementById('scannedUid').innerText = data.latestUid;

          // Attendance Table
          const tbody = document.getElementById('attendanceTableBody');
          if (!data.attendance || data.attendance.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;">No attendance records found.</td></tr>';
          } else {
            tbody.innerHTML = data.attendance.map(row => {
              const detailInfo = (row.position && row.position !== 'N/A') 
                ? \`<strong style="color: #2980b9;">[\${row.position}]</strong>\` 
                : \`\${row.yearLevel} - \${row.section}\`;

              const typeBadge = row.scanType === 'TIME-OUT' 
                ? '<span class="badge-type-out">TIME-OUT</span>' 
                : '<span class="badge-type-in">TIME-IN</span>';

              const statusBadge = row.status === 'LATE' 
                ? '<span class="badge-late">LATE</span>' 
                : '<span class="badge-ontime">' + row.status + '</span>';

              return \`
                <tr>
                  <td><strong>\${row.name}</strong></td>
                  <td>\${row.studentId}</td>
                  <td>\${detailInfo}</td>
                  <td>\${row.event}</td>
                  <td>\${typeBadge}</td>
                  <td>\${statusBadge}</td>
                  <td><strong>\${row.duration || 'N/A'}</strong></td>
                  <td>\${row.timestamp}</td>
                  <td><code>\${row.uid}</code></td>
                </tr>
              \`;
            }).join('');
          }

          // Database Table
          const stBody = document.getElementById('studentsTableBody');
          if (!data.students || data.students.length === 0) {
            stBody.innerHTML = '<tr><td colspan="7" style="text-align:center;">No participants registered yet.</td></tr>';
          } else {
            stBody.innerHTML = data.students.map(st => {
              const detailInfo = (st.position && st.position !== 'N/A') 
                ? \`<strong style="color: #2980b9;">[\${st.position}]</strong>\` 
                : \`\${st.yearLevel} - \${st.section}\`;

              const contactInfo = [st.email, st.phone].filter(Boolean).join('<br>') || '<span style="color:#aaa;">None</span>';

              return \`
                <tr>
                  <td>\${st.studentId}</td>
                  <td><strong>\${st.name}</strong></td>
                  <td>\${detailInfo}</td>
                  <td><small>\${contactInfo}</small></td>
                  <td>\${st.assignedEvent || 'General Event'}</td>
                  <td><code>\${st.uid}</code></td>
                  <td>
                    <form action="/api/delete-student" method="POST" class="delete-form" onsubmit="return confirm('Remove participant?');">
                      <input type="hidden" name="uid" value="\${st.uid}">
                      <button type="submit" class="btn-danger">Delete</button>
                    </form>
                  </td>
                </tr>
              \`;
            }).join('');
          }

        } catch (err) {}
      }

      function useLatestUid() {
        const uid = document.getElementById('scannedUid').innerText;
        if (uid && uid !== 'None') document.getElementById('uidInput').value = uid;
      }

      function downloadCSV() {
        const selected = document.getElementById('exportEventSelect').value;
        window.location.href = '/api/export-csv?event=' + encodeURIComponent(selected);
      }

      checkMeetingEvent();
      setInterval(updateDashboard, 1000);
      updateDashboard();
    </script>
  </body>
  </html>
  `;
  res.send(html);
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});