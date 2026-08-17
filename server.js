const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const nodemailer = require('nodemailer');
const https = require('https');
const querystring = require('querystring');
const mongoose = require('mongoose');

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
const PORT = process.env.PORT || 3000;

// MONGOOSE DATABASE CONNECTION
const MONGO_URI = process.env.MONGO_URI || 'mongodb+srv://srcadmin:30005BNHS@cluster0.he7jspr.mongodb.net/scholarhub_db?appName=Cluster0';

mongoose.connect(MONGO_URI)
  .then(() => console.log('[DATABASE] Connected to MongoDB Atlas successfully!'))
  .catch(err => console.error('[DATABASE ERROR] Could not connect to MongoDB:', err.message));

// SCHEMAS & MODELS
const studentSchema = new mongoose.Schema({
  uid: { type: String, default: '' },
  studentId: { type: String, required: true },
  name: { type: String, required: true },
  yearLevel: { type: String, default: 'Grade 7' },
  section: { type: String, default: 'A' },
  position: { type: String, default: 'Officer' },
  email: { type: String, default: '' },
  phone: { type: String, default: '' },
  assignedEvent: { type: String, default: 'General Event' }
});

const attendanceSchema = new mongoose.Schema({
  uid: String,
  name: String,
  studentId: String,
  yearLevel: String,
  section: String,
  position: String,
  email: String,
  phone: String,
  event: String,
  scanType: String,
  status: String,
  duration: String,
  timestamp: String,
  rawTimestamp: { type: Date, default: Date.now }
});

const configSchema = new mongoose.Schema({
  systemName: { type: String, default: 'RFID Attendance System' },
  logoPath: { type: String, default: '' },
  events: { type: [String], default: ['General Event', 'Faculty Meeting', 'Orientation', 'Seminar'] },
  currentEvent: { type: String, default: 'General Event' },
  cutoffTime: { type: String, default: '08:00' },
  latestUid: { type: String, default: '' },
  enableEmail: { type: Boolean, default: true },
  gmailUser: { type: String, default: process.env.EMAIL_USER || 'markjeraldagdigos00@gmail.com' },
  gmailPass: { type: String, default: process.env.EMAIL_PASS || 'iidgggfvklwjezsm' },
  enableSms: { type: Boolean, default: false },
  semaphoreApiKey: { type: String, default: '' }
});

const Student = mongoose.model('Student', studentSchema);
const Attendance = mongoose.model('Attendance', attendanceSchema);
const Config = mongoose.model('Config', configSchema);

async function getConfig() {
  let config = await Config.findOne();
  if (!config) {
    config = await Config.create({
      enableEmail: true,
      gmailUser: process.env.EMAIL_USER || 'markjeraldagdigos00@gmail.com',
      gmailPass: process.env.EMAIL_PASS || 'iidgggfvklwjezsm'
    });
  }
  return config;
}

// UPLOADS SETUP FOR SCHOOL LOGO
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => cb(null, 'school_logo' + path.extname(file.originalname))
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) cb(null, true);
    else cb(new Error('Only image files are allowed!'), false);
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

// NOTIFICATIONS (EMAIL & SMS)
const { Resend } = require('resend');
const resend = new Resend(process.env.RESEND_API_KEY || 'YOUR_RESEND_API_KEY');

async function sendEmailNotification(config, recipientEmail, studentName, scanType, status, eventName, timestamp, duration) {
  if (!recipientEmail) return;
  const durationText = duration ? `<li><strong>Duration:</strong> ${duration}</li>` : '';
  try {
    await resend.emails.send({
      from: 'RFID System <onboarding@resend.dev>',
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
        </div>
      `
    });
  } catch (error) {
    console.error('[EMAIL ERROR]', error.message);
  }
}

function sendSMSNotification(config, phoneNumber, studentName, scanType, status, eventName, timestamp, duration) {
  if (!config.enableSms || !config.semaphoreApiKey || !phoneNumber) return;
  let message = `[Attendance] ${studentName} logged ${scanType} for ${eventName} at ${timestamp}. Status: ${status}.`;
  if (duration) message += ` Duration: ${duration}.`;

  const postData = querystring.stringify({ apikey: config.semaphoreApiKey, number: phoneNumber, message: message });
  const options = {
    hostname: 'api.semaphore.co', port: 443, path: '/api/v4/messages', method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': postData.length }
  };
  const req = https.request(options, (res) => { res.on('data', (d) => console.log('[SMS RESPONSE]', d.toString())); });
  req.on('error', (e) => console.error('[SMS ERROR]', e.message));
  req.write(postData); req.end();
}

function calculateDuration(timeInDate, timeOutDate) {
  const diffMs = timeOutDate - timeInDate;
  if (isNaN(diffMs) || diffMs < 0) return 'N/A';
  const totalMinutes = Math.floor(diffMs / (1000 * 60));
  const hours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;
  return hours > 0 ? `${hours}h ${mins}m` : `${mins}m`;
}

// API: ESP8266 SCANNER ENDPOINT
app.post('/api/scan', async (req, res) => {
  try {
    const { uid } = req.body;
    if (!uid) return res.status(400).json({ status: 'error', message: 'No UID' });

    const cleanUid = uid.trim().toUpperCase();
    const config = await getConfig();
    
    config.latestUid = cleanUid;
    await config.save();

    const student = await Student.findOne({ uid: cleanUid });
    const now = new Date();

    if (student) {
      const eventName = student.assignedEvent || config.currentEvent || 'General Event';
      const startOfDay = new Date(now.setHours(0, 0, 0, 0));
      const endOfDay = new Date(now.setHours(23, 59, 59, 999));
      const scanTime = new Date();

      const lastLog = await Attendance.findOne({
        uid: cleanUid,
        event: eventName,
        rawTimestamp: { $gte: startOfDay, $lte: endOfDay }
      }).sort({ rawTimestamp: -1 });

      let scanType = 'TIME-IN';
      let duration = '';
      let statusLabel = 'ON TIME';

      if (lastLog && lastLog.scanType === 'TIME-IN') {
        scanType = 'TIME-OUT';
        statusLabel = 'COMPLETED';
        duration = calculateDuration(new Date(lastLog.rawTimestamp), scanTime);
      } else {
        const currentTimeStr = scanTime.toTimeString().slice(0, 5);
        statusLabel = currentTimeStr > (config.cutoffTime || '08:00') ? 'LATE' : 'ON TIME';
      }

      const record = new Attendance({
        uid: cleanUid,
        name: student.name,
        studentId: student.studentId,
        yearLevel: student.yearLevel || 'N/A',
        section: student.section || 'N/A',
        position: student.position || 'Officer',
        email: student.email || '',
        phone: student.phone || '',
        event: eventName,
        scanType,
        status: statusLabel,
        duration: duration || 'N/A',
        timestamp: scanTime.toLocaleString(),
        rawTimestamp: scanTime
      });

      await record.save();

      if (config.enableEmail && student.email) {
        sendEmailNotification(config, student.email, student.name, scanType, statusLabel, eventName, record.timestamp, duration);
      }
      if (config.enableSms && student.phone) {
        sendSMSNotification(config, student.phone, student.name, scanType, statusLabel, eventName, record.timestamp, duration);
      }

      return res.json({ status: 'success', scanType, isLate: statusLabel === 'LATE', message: `${scanType} recorded for ${student.name}` });
    } else {
      return res.json({ status: 'unknown', message: 'Card not registered' });
    }
  } catch (err) {
    res.status(500).json({ status: 'error', message: 'Server Error' });
  }
});

// REALTIME DATA ENDPOINT
app.get('/api/live-data', async (req, res) => {
  try {
    const config = await getConfig();
    const students = await Student.find();
    const attendance = await Attendance.find().sort({ rawTimestamp: -1 });

    res.json({
      latestUid: config.latestUid || '',
      attendance: attendance,
      students: students
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// SETTINGS ENDPOINTS
app.post('/api/update-system-name', async (req, res) => {
  const { systemName } = req.body;
  if (systemName) {
    const config = await getConfig();
    config.systemName = systemName;
    await config.save();
  }
  res.redirect('/');
});

app.post('/api/upload-logo', upload.single('logoFile'), async (req, res) => {
  if (req.file) {
    const config = await getConfig();
    config.logoPath = `/uploads/${req.file.filename}?v=${Date.now()}`;
    await config.save();
  }
  res.redirect('/');
});

app.post('/api/remove-logo', async (req, res) => {
  const config = await getConfig();
  config.logoPath = '';
  await config.save();
  res.redirect('/');
});

app.post('/api/notification-settings', async (req, res) => {
  const { enableEmail, gmailUser, gmailPass, enableSms, semaphoreApiKey } = req.body;
  const config = await getConfig();

  config.enableEmail = enableEmail === 'on';
  config.gmailUser = gmailUser || 'markjeraldagdigos00@gmail.com';
  if (gmailPass && gmailPass !== '******') config.gmailPass = gmailPass;
  config.enableSms = enableSms === 'on';
  config.semaphoreApiKey = semaphoreApiKey || '';

  await config.save();
  res.redirect('/');
});

app.post('/api/event-settings', async (req, res) => {
  const { newEvent, activeEvent, cutoffTime } = req.body;
  const config = await getConfig();

  if (newEvent && !config.events.includes(newEvent)) {
    config.events.push(newEvent);
    config.currentEvent = newEvent;
  } else if (activeEvent) {
    config.currentEvent = activeEvent;
  }
  if (cutoffTime) config.cutoffTime = cutoffTime;

  await config.save();
  res.redirect('/');
});

app.post('/api/delete-event', async (req, res) => {
  const { eventToDelete } = req.body;
  if (eventToDelete) {
    const config = await getConfig();
    config.events = config.events.filter(e => e !== eventToDelete);
    if (config.currentEvent === eventToDelete) {
      config.currentEvent = config.events[0] || 'General Event';
    }
    await config.save();
  }
  res.redirect('/');
});

// ADMIN REGISTER / EDIT PARTICIPANT ENDPOINT
app.post('/api/register', async (req, res) => {
  try {
    const { mongoId, uid, name, studentId, yearLevel, section, assignedEvent, position, customPosition, email, phone } = req.body;
    let finalPosition = (position === 'Other' && customPosition) ? customPosition.trim() : position || 'Officer';
    const cleanUid = uid ? uid.trim().toUpperCase() : '';

    if (mongoId) {
      await Student.findByIdAndUpdate(mongoId, {
        uid: cleanUid,
        name,
        studentId,
        yearLevel: yearLevel || 'Grade 7',
        section: section || 'A',
        position: finalPosition,
        email: email || '',
        phone: phone || '',
        assignedEvent: assignedEvent || 'General Event'
      });
    } else {
      const newStudent = new Student({
        uid: cleanUid,
        name,
        studentId,
        yearLevel: yearLevel || 'Grade 7',
        section: section || 'A',
        position: finalPosition,
        email: email || '',
        phone: phone || '',
        assignedEvent: assignedEvent || 'General Event'
      });
      await newStudent.save();
    }
  } catch (err) {
    console.error('[SAVE ERROR]', err.message);
  }
  res.redirect('/');
});

app.post('/api/delete-student', async (req, res) => {
  const { id } = req.body;
  if (id) await Student.findByIdAndDelete(id);
  res.redirect('/');
});

app.get('/api/export-csv', async (req, res) => {
  const selectedEvent = req.query.event;
  let filter = {};
  if (selectedEvent && selectedEvent !== 'ALL') filter.event = selectedEvent;

  const attendance = await Attendance.find(filter).sort({ rawTimestamp: -1 });

  let csv = 'UID,ID Number,Name,Grade Level,Section,Position,Email,Phone,Event,Type,Status,Duration,Timestamp\n';
  attendance.forEach(row => {
    csv += `"${row.uid}","${row.studentId}","${row.name}","${row.yearLevel}","${row.section}","${row.position || 'Officer'}","${row.email || ''}","${row.phone || ''}","${row.event}","${row.scanType || 'TIME-IN'}","${row.status}","${row.duration || 'N/A'}","${row.timestamp}"\n`;
  });

  const filename = selectedEvent && selectedEvent !== 'ALL' 
    ? `attendance_${selectedEvent.replace(/\s+/g, '_')}.csv` 
    : 'attendance_all_events.csv';

  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.status(200).send(csv);
});

app.post('/api/clear-logs', async (req, res) => {
  await Attendance.deleteMany({});
  res.redirect('/');
});

// -------------------------------------------------------------
// 1. PUBLIC STUDENT REGISTRATION FORM (With Grade 7-12, Section & Position)
// -------------------------------------------------------------
app.get('/student-register', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Student Self-Registration</title>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f6f9; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px 0; }
        .card { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); width: 100%; max-width: 420px; box-sizing: border-box; }
        h2 { text-align: center; color: #2c3e50; margin-bottom: 20px; }
        label { font-weight: bold; font-size: 14px; color: #34495e; }
        input, select { width: 100%; padding: 10px; margin: 6px 0 16px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        .row { display: flex; gap: 10px; }
        .row > div { flex: 1; }
        button { width: 100%; background: #27ae60; color: white; padding: 12px; border: none; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 16px; margin-top: 10px; }
        button:hover { background: #219150; }
        .hidden { display: none; }
      </style>
    </head>
    <body>
      <div class="card">
        <h2>Student Registration Form</h2>
        <form action="/api/register-student" method="POST">
          <label>ID Number:</label>
          <input type="text" name="studentId" placeholder="e.g. 2026-1001" required>

          <label>Full Name:</label>
          <input type="text" name="name" placeholder="Juan Dela Cruz" required>

          <div class="row">
            <div>
              <label>Grade Level:</label>
              <select name="yearLevel" required>
                <option value="Grade 7">Grade 7</option>
                <option value="Grade 8">Grade 8</option>
                <option value="Grade 9">Grade 9</option>
                <option value="Grade 10">Grade 10</option>
                <option value="Grade 11">Grade 11</option>
                <option value="Grade 12">Grade 12</option>
              </select>
            </div>
            <div>
              <label>Section:</label>
              <input type="text" name="section" placeholder="e.g. Diamond" required>
            </div>
          </div>

          <label>Position / Role:</label>
          <select name="position" id="studentPosSelect" onchange="checkCustomPos()" required>
            <option value="Officer">Officer</option>
            <option value="Member">Member</option>
            <option value="President">President</option>
            <option value="Vice President">Vice President</option>
            <option value="Secretary">Secretary</option>
            <option value="Treasurer">Treasurer</option>
            <option value="Teacher / Faculty">Teacher / Faculty</option>
            <option value="Guest">Guest</option>
            <option value="Other">Custom Position...</option>
          </select>

          <div id="customPosBox" class="hidden">
            <label>Specify Custom Position:</label>
            <input type="text" name="customPosition" placeholder="Specify position">
          </div>

          <label>Email Address:</label>
          <input type="email" name="email" placeholder="juan@gmail.com" required>

          <label>Phone Number (Optional):</label>
          <input type="tel" name="phone" placeholder="09171234567">

          <button type="submit">Submit Registration</button>
        </form>
      </div>

      <script>
        function checkCustomPos() {
          const val = document.getElementById('studentPosSelect').value;
          const box = document.getElementById('customPosBox');
          if (val === 'Other') {
            box.classList.remove('hidden');
          } else {
            box.classList.add('hidden');
          }
        }
      </script>
    </body>
    </html>
  `);
});

app.post('/api/register-student', async (req, res) => {
  try {
    const { name, email, studentId, phone, yearLevel, section, position, customPosition } = req.body;
    let finalPosition = (position === 'Other' && customPosition) ? customPosition.trim() : position || 'Officer';

    const existing = await Student.findOne({ email });
    if (existing) {
      return res.send(`<div style="text-align:center; padding:50px; font-family:Arial;"><h2 style="color:#e74c3c;">May nakarehistro nang ganyang email!</h2><a href="/student-register">Bumalik sa Form</a></div>`);
    }

    const newStudent = new Student({
      name,
      email,
      studentId,
      phone: phone || '',
      yearLevel: yearLevel || 'Grade 7',
      section: section || 'A',
      position: finalPosition,
      uid: ''
    });

    await newStudent.save();
    res.send(`
      <div style="text-align:center; padding:50px; font-family:Arial;">
        <h2 style="color:#2ecc71;">Registration Successful!</h2>
        <p>Salamat <strong>${name}</strong>! Naisumite na ang iyong impormasyon (${yearLevel} - ${section} | ${finalPosition}). I-a-assign ng Admin ang iyong RFID Card.</p>
        <a href="/student-register" style="display:inline-block; margin-top:15px; text-decoration:none; color:#2980b9; font-weight:bold;">Mag-register ng iba pa</a>
      </div>
    `);
  } catch (err) {
    res.status(500).send('Error: ' + err.message);
  }
});

// -------------------------------------------------------------
// 2. DIRECT ADMIN DASHBOARD ( / )
// -------------------------------------------------------------
app.get('/', async (req, res) => {
  const config = await getConfig();
  const eventList = Array.isArray(config.events) ? config.events : ['General Event'];
  const eventOptions = eventList.map(e => `<option value="${e}" ${e === config.currentEvent ? 'selected' : ''}>${e}</option>`).join('');

  let gradeOptions = '';
  for (let i = 7; i <= 12; i++) {
    gradeOptions += `<option value="Grade ${i}">Grade ${i}</option>`;
  }

  const logoHtml = config.logoPath ? `<img src="${config.logoPath}" alt="School Logo" class="header-logo">` : '';

  res.send(`
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
      .btn-warning { background: #f39c12; color: white; padding: 6px 12px; font-size: 0.85em; border-radius: 4px; cursor: pointer; border: none; font-weight: bold; }
      .btn-secondary { background: #2980b9; }
      table { width: 100%; border-collapse: collapse; margin-top: 10px; background: white; }
      th, td { border: 1px solid #e1e8ed; padding: 10px; text-align: left; }
      th { background: #34495e; color: white; }
      .badge-ontime { background: #2ecc71; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 0.85em; }
      .badge-late { background: #e74c3c; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 0.85em; }
      .badge-type-in { background: #3498db; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 0.85em; }
      .badge-type-out { background: #9b59b6; color: white; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 0.85em; }
      .section-divider { border: 0; height: 1px; background: #e1e8ed; margin: 15px 0; }
      details.settings-card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); padding: 15px 20px; }
      details.settings-card summary { font-size: 1.2em; font-weight: bold; color: #1a252f; cursor: pointer; }
      .hidden-field { display: none; }
      .nav-links { margin-bottom: 15px; background: #e8f4f8; padding: 10px 15px; border-radius: 6px; }
      .nav-links a { font-weight: bold; color: #2980b9; text-decoration: none; }
    </style>
  </head>
  <body>

    <div id="adminContent">
      <div class="header-container">
        ${logoHtml}
        <h1>${config.systemName || 'RFID Attendance System'}</h1>
      </div>

      <div class="nav-links">
        <strong>Student Self-Registration Link:</strong>
        <a href="/student-register" target="_blank">/student-register (Ibigay sa mga Magre-rehistro)</a>
      </div>

      <div class="container">
        <details class="settings-card">
          <summary> System, Event & Notification Settings </summary>
          
          <div style="display: flex; flex-wrap: wrap; gap: 20px; margin-top: 10px;">
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

            <div style="flex: 1; min-width: 280px;">
              <h3>SMS & Email API Settings</h3>
              <form action="/api/notification-settings" method="POST">
                <label>
                  <input type="checkbox" name="enableEmail" ${config.enableEmail ? 'checked' : ''} style="width: auto;">
                  <strong>Enable Email Notifications</strong>
                </label>
                <input type="email" name="gmailUser" placeholder="Gmail Address" value="${config.gmailUser || 'markjeraldagdigos00@gmail.com'}">
                <input type="password" name="gmailPass" placeholder="Gmail App Password" value="${config.gmailPass ? '******' : ''}">

                <hr class="section-divider">

                <label>
                  <input type="checkbox" name="enableSms" ${config.enableSms ? 'checked' : ''} style="width: auto;">
                  <strong>Enable SMS Notifications (Semaphore)</strong>
                </label>
                <input type="text" name="semaphoreApiKey" placeholder="Semaphore API Key" value="${config.semaphoreApiKey || ''}">

                <input type="submit" value="Save Notification Settings" style="width: 100%; margin-top: 10px;">
              </form>
            </div>
          </div>
        </details>

        <div class="card" id="registrationCard">
          <h2 id="formTitle">Register / Edit Participant</h2>
          <p>Last Scanned RFID Card UID: <strong id="scannedUid" style="color: #e67e22;">${config.latestUid || 'None'}</strong></p>
          
          <form action="/api/register" method="POST" id="registerForm">
            <input type="hidden" id="mongoIdInput" name="mongoId">

            <label><strong>RFID Card UID (Pwedeng i-link sa huli):</strong></label>
            <input type="text" id="uidInput" name="uid" placeholder="RFID Card UID">
            
            <label><strong>ID Number:</strong></label>
            <input type="text" id="studentIdInput" name="studentId" placeholder="ID Number" required>
            
            <label><strong>Full Name:</strong></label>
            <input type="text" id="nameInput" name="name" placeholder="Full Name" required>

            <div style="display: flex; gap: 10px;">
              <div style="flex: 1;">
                <label><strong>Email Address:</strong></label>
                <input type="email" id="emailInput" name="email" placeholder="e.g. parent@gmail.com">
              </div>
              <div style="flex: 1;">
                <label><strong>Phone Number:</strong></label>
                <input type="tel" id="phoneInput" name="phone" placeholder="e.g. 09171234567">
              </div>
            </div>

            <label><strong>Assign to Event:</strong></label>
            <select id="eventSelect" name="assignedEvent" onchange="checkMeetingEvent()">${eventOptions}</select>

            <div id="studentFields" style="display: flex; gap: 10px;">
              <div style="flex: 1;">
                <label><strong>Grade Level:</strong></label>
                <select id="yearLevelSelect" name="yearLevel">${gradeOptions}</select>
              </div>
              <div style="flex: 1;">
                <label><strong>Section:</strong></label>
                <input type="text" id="sectionInput" name="section" placeholder="e.g. Diamond / A">
              </div>
            </div>

            <div id="meetingFields">
              <label><strong>Position / Role:</strong></label>
              <select id="positionSelect" name="position" onchange="checkCustomPosition()">
                <option value="Officer" selected>Officer</option>
                <option value="Member">Member</option>
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
                <input type="text" id="customPositionInput" name="customPosition" placeholder="Enter position/role">
              </div>
            </div>

            <button type="button" class="btn-secondary" onclick="useLatestUid()" style="width: 100%; margin-top: 10px; margin-bottom: 10px;">Link Last Scanned Card UID</button>
            <input type="submit" id="submitBtn" value="Save / Update Participant" style="width: 100%;">
            <button type="button" id="cancelEditBtn" onclick="resetForm()" class="btn-danger hidden-field" style="width: 100%; margin-top: 5px;">Cancel Edit</button>
          </form>
        </div>
      </div>

      <div class="card" style="margin-top: 20px;">
        <h2>Live Attendance Log (Time-In / Time-Out)</h2>
        
        <div style="display: flex; gap: 10px; align-items: center; margin-bottom: 15px;">
          <label><strong>Export CSV for Event:</strong></label>
          <select id="exportEventSelect" style="width: auto; margin: 0;">
            <option value="ALL">All Events</option>
            ${eventOptions}
          </select>
          <button type="button" onclick="downloadCSV()">Download CSV</button>
          <form action="/api/clear-logs" method="POST" style="margin-left: auto;" onsubmit="return confirm('Clear all logs?');">
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

      <div class="card" style="margin-top: 20px;">
        <h2>Registered Participants Database</h2>
        <table>
          <thead>
            <tr>
              <th>ID Number</th>
              <th>Name</th>
              <th>Grade / Section</th>
              <th>Position</th>
              <th>Email / Phone</th>
              <th>Assigned Event</th>
              <th>Card UID</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody id="studentsTableBody"></tbody>
        </table>
      </div>
    </div>

    <script>
      let registeredStudents = [];

      function checkCustomPosition() {
        const posVal = document.getElementById('positionSelect').value;
        const customBox = document.getElementById('customPositionBox');
        if (posVal === 'Other') {
          customBox.classList.remove('hidden-field');
        } else {
          customBox.classList.add('hidden-field');
        }
      }

      function editStudent(id) {
        const student = registeredStudents.find(s => s._id === id);
        if (!student) return;

        document.getElementById('mongoIdInput').value = student._id;
        document.getElementById('uidInput').value = student.uid || '';
        document.getElementById('studentIdInput').value = student.studentId;
        document.getElementById('nameInput').value = student.name;
        document.getElementById('emailInput').value = student.email || '';
        document.getElementById('phoneInput').value = student.phone || '';

        if (student.assignedEvent) {
          document.getElementById('eventSelect').value = student.assignedEvent;
        }

        if (student.yearLevel) document.getElementById('yearLevelSelect').value = student.yearLevel;
        if (student.section) document.getElementById('sectionInput').value = student.section;

        if (student.position) {
          const posSelect = document.getElementById('positionSelect');
          const options = Array.from(posSelect.options).map(o => o.value);
          if (options.includes(student.position)) {
            posSelect.value = student.position;
          } else {
            posSelect.value = 'Other';
            document.getElementById('customPositionInput').value = student.position;
          }
          checkCustomPosition();
        }

        document.getElementById('formTitle').innerText = 'Edit Details (' + student.name + ')';
        document.getElementById('submitBtn').value = 'Update Participant Record';
        document.getElementById('cancelEditBtn').classList.remove('hidden-field');
        document.getElementById('registrationCard').scrollIntoView({ behavior: 'smooth' });
      }

      function resetForm() {
        document.getElementById('registerForm').reset();
        document.getElementById('mongoIdInput').value = '';
        document.getElementById('formTitle').innerText = 'Register / Edit Participant';
        document.getElementById('submitBtn').value = 'Save / Update Participant';
        document.getElementById('cancelEditBtn').classList.add('hidden-field');
      }

      async function updateDashboard() {
        try {
          const res = await fetch('/api/live-data');
          const data = await res.json();
          if (data.latestUid) document.getElementById('scannedUid').innerText = data.latestUid;

          registeredStudents = data.students || [];

          const tbody = document.getElementById('attendanceTableBody');
          if (!data.attendance || data.attendance.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;">No attendance records found.</td></tr>';
          } else {
            tbody.innerHTML = data.attendance.map(row => {
              const detailInfo = \`\${row.yearLevel} - \${row.section} <br><small style="color:#2980b9;">[\${row.position || 'Officer'}]</small>\`;

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

          const stBody = document.getElementById('studentsTableBody');
          if (!registeredStudents || registeredStudents.length === 0) {
            stBody.innerHTML = '<tr><td colspan="8" style="text-align:center;">No participants registered yet.</td></tr>';
          } else {
            stBody.innerHTML = registeredStudents.map(st => {
              const contactInfo = [st.email, st.phone].filter(Boolean).join('<br>') || '<span style="color:#aaa;">None</span>';

              return \`
                <tr>
                  <td>\${st.studentId}</td>
                  <td><strong>\${st.name}</strong></td>
                  <td>\${st.yearLevel || 'N/A'} - \${st.section || 'N/A'}</td>
                  <td><strong style="color: #2980b9;">\${st.position || 'Officer'}</strong></td>
                  <td><small>\${contactInfo}</small></td>
                  <td>\${st.assignedEvent || 'General Event'}</td>
                  <td>\${st.uid ? '<code>' + st.uid + '</code>' : '<span style="color:#e67e22; font-weight:bold;">No Card Linked</span>'}</td>
                  <td>
                    <button type="button" class="btn-warning" onclick="editStudent('\${st._id}')">Edit / Link Card</button>
                    <form action="/api/delete-student" method="POST" style="display:inline;" onsubmit="return confirm('Remove participant?');">
                      <input type="hidden" name="id" value="\${st._id}">
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

      updateDashboard();
      setInterval(updateDashboard, 2000);
    </script>
  </body>
  </html>
  `);
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
