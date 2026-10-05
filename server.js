const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const nodemailer = require('nodemailer');
const https = require('https');
const querystring = require('querystring');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');
const { Resend } = require('resend');

const app = express();
const PORT = process.env.PORT || 3000;

// MONGOOSE DATABASE CONNECTION
const MONGO_URI = process.env.MONGO_URI || 'mongodb+srv://srcadmin:30005BNHS@cluster0.he7jspr.mongodb.net/school_attendance_db?appName=Cluster0';

mongoose.connect(MONGO_URI)
  .then(() => console.log('[DATABASE] Connected to MongoDB Atlas successfully!'))
  .catch(err => console.error('[DATABASE ERROR] Could not connect to MongoDB:', err.message));

// AUTOMATICALLY GENERATE TEMPLATE.XLSX IF MISSING
async function ensureExcelTemplateExists() {
  const templatePath = path.join(__dirname, 'template.xlsx');
  
  if (fs.existsSync(templatePath)) {
    console.log('[EXCEL TEMPLATE] "template.xlsx" already exists.');
    return;
  }

  console.log('[EXCEL TEMPLATE] Generating school attendance "template.xlsx"...');
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('School Attendance Log');

  worksheet.columns = [
    { key: 'studentId', width: 18 },
    { key: 'name', width: 28 },
    { key: 'gradeSection', width: 20 },
    { key: 'event', width: 24 },
    { key: 'scanType', width: 15 },
    { key: 'status', width: 15 },
    { key: 'duration', width: 15 },
    { key: 'timestamp', width: 25 }
  ];

  // Merge Header Title Area
  worksheet.mergeCells('B1:G1');
  worksheet.mergeCells('B2:G2');
  worksheet.mergeCells('B3:G3');

  worksheet.getCell('B1').value = 'BATAC NATIONAL HIGH SCHOOL';
  worksheet.getCell('B1').font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF1B365D' } };
  worksheet.getCell('B1').alignment = { horizontal: 'center', vertical: 'middle' };

  worksheet.getCell('B2').value = 'OFFICIAL SCHOOL ATTENDANCE SYSTEM';
  worksheet.getCell('B2').font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FF444444' } };
  worksheet.getCell('B2').alignment = { horizontal: 'center', vertical: 'middle' };

  worksheet.getCell('B3').value = 'DAILY ATTENDANCE REPORT LOG';
  worksheet.getCell('B3').font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF777777' } };
  worksheet.getCell('B3').alignment = { horizontal: 'center', vertical: 'middle' };

  // Set Table Row Headers (Row 5)
  const headers = [
    'ID NUMBER', 'STUDENT NAME', 'GRADE & SECTION', 
    'EVENT / SESSION', 'SCAN TYPE', 'STATUS', 'DURATION', 'TIMESTAMP'
  ];

  const headerRow = worksheet.getRow(5);
  headerRow.values = headers;
  headerRow.height = 26;

  headerRow.eachCell((cell) => {
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B365D' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
    };
  });

  // Add School Logo if available
  let logoPath = path.join(__dirname, 'bnhs_logo.jpg');
  if (!fs.existsSync(logoPath)) logoPath = path.join(__dirname, 'bnhs_logo.png');

  if (fs.existsSync(logoPath)) {
    const ext = path.extname(logoPath).toLowerCase() === '.png' ? 'png' : 'jpeg';
    const schoolImage = workbook.addImage({ filename: logoPath, extension: ext });
    worksheet.addImage(schoolImage, {
      tl: { col: 0, row: 0 },
      ext: { width: 70, height: 70 }
    });
  }

  await workbook.xlsx.writeFile(templatePath);
  console.log('[EXCEL TEMPLATE] "template.xlsx" created successfully!');
}

// SCHEMAS & MODELS (DATABASE CODE)
const studentSchema = new mongoose.Schema({
  uid: { type: String, default: '' },
  studentId: { type: String, required: true },
  name: { type: String, required: true },
  yearLevel: { type: String, default: 'Grade 7' },
  section: { type: String, default: 'A' },
  photo: { type: String, default: '' },
  email: { type: String, default: '' },
  phone: { type: String, default: '' },
  assignedEvent: { type: String, default: 'Regular Class' }
});

const attendanceSchema = new mongoose.Schema({
  uid: String,
  name: String,
  studentId: String,
  yearLevel: String,
  section: String,
  photo: String,
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
  systemName: { type: String, default: 'School RFID Attendance System' },
  logoPath: { type: String, default: '' },
  events: { type: [String], default: ['Regular Class', 'Morning Assembly', 'Exam Week', 'School Event'] },
  currentEvent: { type: String, default: 'Regular Class' },
  cutoffTime: { type: String, default: '07:45' },
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

// UPLOADS SETUP (LOGOS & STUDENT PHOTOS)
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname === 'logoFile' ? 'school_logo' + path.extname(file.originalname) : 'student_' + uniqueSuffix + path.extname(file.originalname));
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) cb(null, true);
    else cb(new Error('Only image files are allowed!'), false);
  }
});

// MIDDLEWARES
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/uploads', express.static(uploadsDir));

// NOTIFICATIONS (EMAIL & SMS)
const resend = new Resend(process.env.RESEND_API_KEY || 'YOUR_RESEND_API_KEY');

async function sendEmailNotification(recipientEmail, studentName, scanType, status, eventName, timestamp, duration) {
  if (!recipientEmail) return;
  const durationText = duration ? `<li><strong>Duration:</strong> ${duration}</li>` : '';
  try {
    await resend.emails.send({
      from: 'School Attendance <onboarding@resend.dev>',
      to: recipientEmail,
      subject: `[School Attendance] ${scanType} Notice: ${studentName}`,
      html: `
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e1e8ed; border-radius: 8px;">
          <h2 style="color: #1b365d;">School Attendance Notification (${scanType})</h2>
          <p>Dear Parent / Guardian,</p>
          <p>This is to inform you that <strong>${studentName}</strong> has successfully recorded a <strong>${scanType}</strong>.</p>
          <ul>
            <li><strong>Session / Event:</strong> ${eventName}</li>
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
  let message = `[School Attendance] ${studentName} logged ${scanType} for ${eventName} at ${timestamp}. Status: ${status}.`;
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
    if (!uid) return res.status(400).json({ status: 'error', message: 'No UID provided' });

    const cleanUid = uid.trim().toUpperCase();
    const config = await getConfig();
    
    config.latestUid = cleanUid;
    await config.save();

    const student = await Student.findOne({ uid: cleanUid });
    const now = new Date();

    if (student) {
      const eventName = student.assignedEvent || config.currentEvent || 'Regular Class';
      const startOfDay = new Date(now.setHours(0, 0, 0, 0));
      const scanTime = new Date();

      const lastLog = await Attendance.findOne({
        uid: cleanUid,
        event: eventName,
        rawTimestamp: { $gte: startOfDay }
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
        statusLabel = currentTimeStr > (config.cutoffTime || '07:45') ? 'LATE' : 'ON TIME';
      }

      const record = new Attendance({
        uid: cleanUid,
        name: student.name,
        studentId: student.studentId,
        yearLevel: student.yearLevel || 'Grade 7',
        section: student.section || 'A',
        photo: student.photo || '',
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
        sendEmailNotification(student.email, student.name, scanType, statusLabel, eventName, record.timestamp, duration);
      }
      if (config.enableSms && student.phone) {
        sendSMSNotification(config, student.phone, student.name, scanType, statusLabel, eventName, record.timestamp, duration);
      }

      return res.json({ 
        status: 'success', 
        scanType, 
        isLate: statusLabel === 'LATE', 
        student: {
          name: student.name,
          studentId: student.studentId,
          yearLevel: student.yearLevel,
          section: student.section,
          photo: student.photo,
          scanType,
          status: statusLabel
        },
        message: `${scanType} recorded for ${student.name}` 
      });
    } else {
      return res.json({ status: 'unknown', message: 'RFID Card not registered' });
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
    const attendance = await Attendance.find().sort({ rawTimestamp: -1 }).limit(100);

    res.json({
      latestUid: config.latestUid || '',
      attendance,
      students
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// EXPORT TO EXCEL
app.get('/api/export-excel', async (req, res) => {
  try {
    await ensureExcelTemplateExists();

    const selectedEvent = req.query.event;
    let filter = {};
    if (selectedEvent && selectedEvent !== 'ALL') filter.event = selectedEvent;

    const attendance = await Attendance.find(filter).sort({ rawTimestamp: -1 });

    const templatePath = path.join(__dirname, 'template.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const worksheet = workbook.getWorksheet(1);

    let startRow = 6;

    attendance.forEach((row, index) => {
      const currentRow = worksheet.getRow(startRow + index);
      
      currentRow.getCell(1).value = row.studentId;
      currentRow.getCell(2).value = row.name;
      currentRow.getCell(3).value = `${row.yearLevel} - ${row.section}`;
      currentRow.getCell(4).value = row.event;
      currentRow.getCell(5).value = row.scanType || 'TIME-IN';
      currentRow.getCell(6).value = row.status;
      currentRow.getCell(7).value = row.duration || 'N/A';
      currentRow.getCell(8).value = row.timestamp;

      currentRow.commit();
    });

    const filename = selectedEvent && selectedEvent !== 'ALL' 
      ? `School_Attendance_${selectedEvent.replace(/\s+/g, '_')}.xlsx` 
      : 'School_Attendance_Report.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('[EXCEL EXPORT ERROR]', err.message);
    res.status(500).send('Error generating Excel file: ' + err.message);
  }
});

// SETTINGS & CONFIG ENDPOINTS
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
      config.currentEvent = config.events[0] || 'Regular Class';
    }
    await config.save();
  }
  res.redirect('/');
});

// STUDENT REGISTRATION / UPDATE WITH PHOTO UPLOAD
app.post('/api/register', upload.single('photoFile'), async (req, res) => {
  try {
    const { mongoId, uid, name, studentId, yearLevel, section, assignedEvent, email, phone } = req.body;
    const cleanUid = uid ? uid.trim().toUpperCase() : '';
    let photoPath = '';

    if (req.file) {
      photoPath = `/uploads/${req.file.filename}`;
    }

    if (mongoId) {
      const existingStudent = await Student.findById(mongoId);
      if (existingStudent) {
        existingStudent.uid = cleanUid || existingStudent.uid;
        existingStudent.name = name;
        existingStudent.studentId = studentId;
        existingStudent.yearLevel = yearLevel || 'Grade 7';
        existingStudent.section = section || 'A';
        if (photoPath) existingStudent.photo = photoPath;
        existingStudent.email = email || '';
        existingStudent.phone = phone || '';
        existingStudent.assignedEvent = assignedEvent || 'Regular Class';
        await existingStudent.save();
      }
    } else {
      const newStudent = new Student({
        uid: cleanUid,
        name,
        studentId,
        yearLevel: yearLevel || 'Grade 7',
        section: section || 'A',
        photo: photoPath,
        email: email || '',
        phone: phone || '',
        assignedEvent: assignedEvent || 'Regular Class'
      });
      await newStudent.save();
    }
  } catch (err) {
    console.error('[STUDENT SAVE ERROR]', err.message);
  }
  res.redirect('/');
});

app.post('/api/delete-student', async (req, res) => {
  const { id } = req.body;
  if (id) await Student.findByIdAndDelete(id);
  res.redirect('/');
});

app.post('/api/clear-logs', async (req, res) => {
  await Attendance.deleteMany({});
  res.redirect('/');
});

// PUBLIC STUDENT SELF-REGISTRATION PAGE
app.get('/student-register', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Student Self-Registration</title>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px 0; }
        .card { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 10px 25px rgba(0,0,0,0.08); width: 100%; max-width: 440px; box-sizing: border-box; }
        h2 { text-align: center; color: #1b365d; margin-bottom: 20px; }
        label { font-weight: 600; font-size: 13px; color: #334155; display: block; margin-top: 10px; }
        input, select { width: 100%; padding: 10px; margin-top: 5px; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box; font-size: 14px; }
        .row { display: flex; gap: 10px; }
        .row > div { flex: 1; }
        button { width: 100%; background: #2563eb; color: white; padding: 12px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; font-size: 15px; margin-top: 20px; transition: background 0.2s; }
        button:hover { background: #1d4ed8; }
      </style>
    </head>
    <body>
      <div class="card">
        <h2>Student Registration</h2>
        <form action="/api/register-student-public" method="POST" enctype="multipart/form-data">
          <label>Student ID Number:</label>
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

          <label>Student Photo:</label>
          <input type="file" name="photoFile" accept="image/*" required>

          <label>Parent / Guardian Email:</label>
          <input type="email" name="email" placeholder="parent@gmail.com" required>

          <label>Parent / Guardian Phone (Optional):</label>
          <input type="tel" name="phone" placeholder="09171234567">

          <button type="submit">Submit Registration</button>
        </form>
      </div>
    </body>
    </html>
  `);
});

app.post('/api/register-student-public', upload.single('photoFile'), async (req, res) => {
  try {
    const { name, email, studentId, phone, yearLevel, section } = req.body;
    let photoPath = req.file ? `/uploads/${req.file.filename}` : '';

    const existing = await Student.findOne({ email });
    if (existing) {
      return res.send(`<div style="text-align:center; padding:50px; font-family:Arial;"><h3 style="color:#e74c3c;">Email is already registered!</h3><a href="/student-register">Go Back</a></div>`);
    }

    const newStudent = new Student({
      name,
      email,
      studentId,
      phone: phone || '',
      yearLevel: yearLevel || 'Grade 7',
      section: section || 'A',
      photo: photoPath,
      uid: ''
    });

    await newStudent.save();
    res.send(`
      <div style="text-align:center; padding:50px; font-family:Arial;">
        <h2 style="color:#16a34a;">Registration Successful!</h2>
        <p>Thank you, <strong>${name}</strong> (${yearLevel} - ${section}). Your registration has been submitted. The administrator will link your RFID ID card.</p>
        <a href="/student-register" style="display:inline-block; margin-top:15px; text-decoration:none; color:#2563eb; font-weight:bold;">Register Another Student</a>
      </div>
    `);
  } catch (err) {
    res.status(500).send('Error: ' + err.message);
  }
});

// REDESIGNED ADMIN DASHBOARD ( / )
app.get('/', async (req, res) => {
  const config = await getConfig();
  const eventList = Array.isArray(config.events) ? config.events : ['Regular Class'];
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
    <title>${config.systemName || 'School RFID Attendance System'}</title>
    <style>
      :root {
        --primary: #1b365d;
        --primary-light: #2563eb;
        --success: #16a34a;
        --danger: #dc2626;
        --warning: #d97706;
        --bg: #f8fafc;
        --card-bg: #ffffff;
        --text: #1e293b;
        --border: #e2e8f0;
      }
      body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: var(--bg); color: var(--text); }
      .header-container { display: flex; align-items: center; gap: 15px; margin-bottom: 20px; background: var(--card-bg); padding: 15px 25px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); }
      .header-logo { height: 55px; width: auto; object-fit: contain; border-radius: 6px; }
      h1 { font-size: 1.5rem; color: var(--primary); margin: 0; }
      .container { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
      @media (max-width: 1024px) { .container { grid-template-columns: 1fr; } }
      .card { background: var(--card-bg); padding: 20px 25px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); margin-bottom: 20px; }
      input, select { width: 100%; padding: 10px; margin: 6px 0 12px 0; border: 1px solid var(--border); border-radius: 6px; box-sizing: border-box; font-size: 14px; }
      button, input[type="submit"] { background: var(--primary-light); color: white; padding: 10px 16px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; font-size: 14px; transition: opacity 0.2s; }
      button:hover, input[type="submit"]:hover { opacity: 0.9; }
      .btn-danger { background: var(--danger); }
      .btn-warning { background: var(--warning); color: white; }
      .btn-secondary { background: #475569; }
      table { width: 100%; border-collapse: collapse; margin-top: 10px; background: white; font-size: 14px; }
      th, td { border: 1px solid var(--border); padding: 12px; text-align: left; }
      th { background: #f1f5f9; color: var(--primary); font-weight: 600; }
      .badge-ontime { background: #dcfce7; color: #166534; padding: 4px 10px; border-radius: 20px; font-weight: 600; font-size: 0.75rem; }
      .badge-late { background: #fee2e2; color: #991b1b; padding: 4px 10px; border-radius: 20px; font-weight: 600; font-size: 0.75rem; }
      .badge-type-in { background: #dbeafe; color: #1e40af; padding: 4px 10px; border-radius: 20px; font-weight: 600; font-size: 0.75rem; }
      .badge-type-out { background: #fae8ff; color: #86198f; padding: 4px 10px; border-radius: 20px; font-weight: 600; font-size: 0.75rem; }
      .section-divider { border: 0; height: 1px; background: var(--border); margin: 15px 0; }
      details.settings-card { background: var(--card-bg); border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); padding: 15px 25px; margin-bottom: 20px; }
      details.settings-card summary { font-size: 1.1rem; font-weight: 600; color: var(--primary); cursor: pointer; }
      .hidden-field { display: none; }
      .nav-links { margin-bottom: 20px; background: #eff6ff; padding: 12px 20px; border-radius: 8px; border: 1px solid #bfdbfe; font-size: 14px; }
      .nav-links a { font-weight: 600; color: var(--primary-light); text-decoration: none; }
      
      /* LIVE SCAN POPUP MODAL */
      #scanModal { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); display: none; justify-content: center; align-items: center; z-index: 1000; backdrop-filter: blur(4px); }
      .modal-content { background: white; padding: 35px; border-radius: 16px; width: 420px; text-align: center; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.2); animation: popUp 0.3s ease; }
      @keyframes popUp { 0% { transform: scale(0.8); opacity: 0; } 100% { transform: scale(1); opacity: 1; } }
      .modal-avatar { width: 130px; height: 130px; border-radius: 50%; object-fit: cover; border: 4px solid var(--primary-light); margin: 0 auto 15px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
      .modal-name { font-size: 1.4rem; font-weight: 700; color: var(--primary); margin: 5px 0; }
      .modal-details { color: #64748b; font-size: 0.95rem; margin-bottom: 15px; }
      .modal-status { display: inline-block; padding: 6px 16px; border-radius: 20px; font-weight: 700; font-size: 0.9rem; margin-top: 5px; }
      .student-thumb { width: 45px; height: 45px; border-radius: 50%; object-fit: cover; border: 2px solid var(--border); }
    </style>
  </head>
  <body>

    <!-- LIVE SCAN POPUP & VOICE ANNOUNCEMENT MODAL -->
    <div id="scanModal">
      <div class="modal-content">
        <img id="modalPhoto" src="" alt="Student Photo" class="modal-avatar" onerror="this.src='https://via.placeholder.com/130?text=No+Photo'">
        <div id="modalScanTypeBadge" class="badge-type-in" style="font-size:0.85rem; margin-bottom:8px;">TIME-IN</div>
        <h3 id="modalName" class="modal-name">Student Name</h3>
        <p id="modalDetails" class="modal-details">Grade 7 - Section A | ID: 2026-0001</p>
        <div>
          <span id="modalStatusBadge" class="badge-ontime">ON TIME</span>
        </div>
      </div>
    </div>

    <div id="adminContent">
      <div class="header-container">
        ${logoHtml}
        <h1>${config.systemName || 'School RFID Attendance System'}</h1>
      </div>

      <div class="nav-links">
        <strong>Student Self-Registration Portal:</strong>
        <a href="/student-register" target="_blank">/student-register (Share this link with students/parents)</a>
      </div>

      <div class="container">
        <details class="settings-card">
          <summary>System & Event Configuration</summary>
          
          <div style="margin-top: 15px;">
            <h3>System Name & Logo</h3>
            <form action="/api/update-system-name" method="POST">
              <label><strong>System Title:</strong></label>
              <input type="text" name="systemName" value="${config.systemName || 'School RFID Attendance System'}" required>
              <input type="submit" value="Save Title" style="width: 100%;">
            </form>

            <hr class="section-divider">

            <form action="/api/upload-logo" method="POST" enctype="multipart/form-data">
              <label><strong>School Logo Image:</strong></label>
              <input type="file" name="logoFile" accept="image/*" required>
              <input type="submit" value="Upload School Logo" class="btn-secondary" style="width: 100%; margin-bottom: 10px;">
            </form>

            ${config.logoPath ? `
            <form action="/api/remove-logo" method="POST">
              <button type="submit" class="btn-danger" style="width: 100%;">Remove Logo</button>
            </form>
            ` : ''}
          </div>

          <hr class="section-divider">

          <h3>Attendance Session Management</h3>
          <form action="/api/event-settings" method="POST">
            <label><strong>Active Session / Event:</strong></label>
            <select name="activeEvent">${eventOptions}</select>

            <label><strong>Add New Session / Event:</strong></label>
            <input type="text" name="newEvent" placeholder="e.g. Morning Assembly / Midterm Exam">

            <label><strong>Late Cut-off Time (HH:MM):</strong></label>
            <input type="time" name="cutoffTime" value="${config.cutoffTime || '07:45'}">

            <input type="submit" value="Save Session Settings" style="width: 100%;">
          </form>
        </details>

        <details class="settings-card" id="registrationCard">
          <summary id="formTitle">Register / Edit Student</summary>
          
          <div style="margin-top: 15px;">
            <p style="margin-bottom: 15px; font-size: 14px;">Last Scanned RFID UID: <strong id="scannedUid" style="color: var(--warning);">${config.latestUid || 'None'}</strong></p>
            
            <form action="/api/register" method="POST" id="registerForm" enctype="multipart/form-data">
              <input type="hidden" id="mongoIdInput" name="mongoId">

              <label><strong>RFID Card UID:</strong></label>
              <input type="text" id="uidInput" name="uid" placeholder="Scan or enter card UID">
              
              <label><strong>Student ID Number:</strong></label>
              <input type="text" id="studentIdInput" name="studentId" placeholder="e.g. 2026-1001" required>
              
              <label><strong>Full Name:</strong></label>
              <input type="text" id="nameInput" name="name" placeholder="Juan Dela Cruz" required>

              <div style="display: flex; gap: 10px;">
                <div style="flex: 1;">
                  <label><strong>Grade Level:</strong></label>
                  <select id="yearLevelSelect" name="yearLevel">${gradeOptions}</select>
                </div>
                <div style="flex: 1;">
                  <label><strong>Section:</strong></label>
                  <input type="text" id="sectionInput" name="section" placeholder="e.g. Diamond" required>
                </div>
              </div>

              <label><strong>Student Photo:</strong></label>
              <input type="file" id="photoFileInput" name="photoFile" accept="image/*">

              <div style="display: flex; gap: 10px;">
                <div style="flex: 1;">
                  <label><strong>Parent Email:</strong></label>
                  <input type="email" id="emailInput" name="email" placeholder="parent@gmail.com">
                </div>
                <div style="flex: 1;">
                  <label><strong>Parent Phone:</strong></label>
                  <input type="tel" id="phoneInput" name="phone" placeholder="09171234567">
                </div>
              </div>

              <label><strong>Assigned Session / Event:</strong></label>
              <select id="eventSelect" name="assignedEvent">${eventOptions}</select>

              <button type="button" class="btn-secondary" onclick="useLatestUid()" style="width: 100%; margin-top: 12px; margin-bottom: 10px;">Link Last Scanned Card UID</button>
              <input type="submit" id="submitBtn" value="Save Student Record" style="width: 100%;">
              <button type="button" id="cancelEditBtn" onclick="resetForm()" class="btn-danger hidden-field" style="width: 100%; margin-top: 5px;">Cancel Edit</button>
            </form>
          </div>
        </details>
      </div>

      <div class="card">
        <h2>Live Attendance Log</h2>
        
        <div style="display: flex; gap: 10px; align-items: center; margin: 15px 0; flex-wrap: wrap;">
          <label style="margin: 0;"><strong>Export Excel Report:</strong></label>
          <select id="exportEventSelect" style="width: auto; margin: 0;">
            <option value="ALL">All Sessions</option>
            ${eventOptions}
          </select>
          <button type="button" onclick="downloadExcel()">Download Excel</button>
          <form action="/api/clear-logs" method="POST" style="margin-left: auto;" onsubmit="return confirm('Clear all attendance logs?');">
            <button type="submit" class="btn-danger">Clear All Logs</button>
          </form>
        </div>

        <div style="overflow-x: auto;">
          <table>
            <thead>
              <tr>
                <th>Photo</th>
                <th>Student Name</th>
                <th>ID Number</th>
                <th>Grade & Section</th>
                <th>Session</th>
                <th>Scan Type</th>
                <th>Status</th>
                <th>Duration</th>
                <th>Timestamp</th>
              </tr>
            </thead>
            <tbody id="attendanceTableBody"></tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <h2>Registered Students Database</h2>
        <div style="overflow-x: auto;">
          <table>
            <thead>
              <tr>
                <th>Photo</th>
                <th>ID Number</th>
                <th>Name</th>
                <th>Grade & Section</th>
                <th>Parent Contact</th>
                <th>Session</th>
                <th>Card UID</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody id="studentsTableBody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <script>
      let registeredStudents = [];
      let lastProcessedAttendanceId = null;

      // Text-to-Speech Announcement & Popup Modal
      function announceAndShowScan(record) {
        if (!record) return;

        // 1. Populate Modal Info
        const photoUrl = record.photo || 'https://via.placeholder.com/130?text=No+Photo';
        document.getElementById('modalPhoto').src = photoUrl;
        document.getElementById('modalName').innerText = record.name;
        document.getElementById('modalDetails').innerText = \`\${record.yearLevel} - \${record.section} | ID: \${record.studentId}\`;
        
        const typeBadge = document.getElementById('modalScanTypeBadge');
        typeBadge.innerText = record.scanType;
        typeBadge.className = record.scanType === 'TIME-OUT' ? 'badge-type-out' : 'badge-type-in';

        const statusBadge = document.getElementById('modalStatusBadge');
        statusBadge.innerText = record.status;
        statusBadge.className = record.status === 'LATE' ? 'badge-late' : 'badge-ontime';

        // Show Modal
        const modal = document.getElementById('scanModal');
        modal.style.display = 'flex';

        // Hide after 4 seconds
        setTimeout(() => {
          modal.style.display = 'none';
        }, 4000);

        // 2. Audio Announcement using Web Speech API
        if ('speechSynthesis' in window) {
          window.speechSynthesis.cancel(); // Stop any ongoing speech
          const textToSpeak = \`\${record.scanType} recorded for \${record.name}, \${record.yearLevel} \${record.section}\`;
          const utterance = new SpeechSynthesisUtterance(textToSpeak);
          utterance.rate = 1.0;
          utterance.pitch = 1.0;
          window.speechSynthesis.speak(utterance);
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

        if (student.assignedEvent) document.getElementById('eventSelect').value = student.assignedEvent;
        if (student.yearLevel) document.getElementById('yearLevelSelect').value = student.yearLevel;
        if (student.section) document.getElementById('sectionInput').value = student.section;

        const regCard = document.getElementById('registrationCard');
        regCard.open = true;

        document.getElementById('formTitle').innerText = 'Edit Student (' + student.name + ')';
        document.getElementById('submitBtn').value = 'Update Student Record';
        document.getElementById('cancelEditBtn').classList.remove('hidden-field');
        regCard.scrollIntoView({ behavior: 'smooth' });
      }

      function resetForm() {
        document.getElementById('registerForm').reset();
        document.getElementById('mongoIdInput').value = '';
        document.getElementById('formTitle').innerText = 'Register / Edit Student';
        document.getElementById('submitBtn').value = 'Save Student Record';
        document.getElementById('cancelEditBtn').classList.add('hidden-field');
      }

      async function updateDashboard() {
        try {
          const res = await fetch('/api/live-data');
          const data = await res.json();
          if (data.latestUid) document.getElementById('scannedUid').innerText = data.latestUid;

          registeredStudents = data.students || [];

          // Check for new scan to trigger popup & voice
          if (data.attendance && data.attendance.length > 0) {
            const latestLog = data.attendance[0];
            if (lastProcessedAttendanceId !== latestLog._id) {
              lastProcessedAttendanceId = latestLog._id;
              announceAndShowScan(latestLog);
            }
          }

          const tbody = document.getElementById('attendanceTableBody');
          if (!data.attendance || data.attendance.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;">No attendance records found.</td></tr>';
          } else {
            tbody.innerHTML = data.attendance.map(row => {
              const photoTag = row.photo ? \`<img src="\${row.photo}" class="student-thumb">\` : '<span style="color:#94a3b8; font-size:12px;">No photo</span>';
              const typeBadge = row.scanType === 'TIME-OUT' ? '<span class="badge-type-out">TIME-OUT</span>' : '<span class="badge-type-in">TIME-IN</span>';
              const statusBadge = row.status === 'LATE' ? '<span class="badge-late">LATE</span>' : '<span class="badge-ontime">' + row.status + '</span>';

              return \`
                <tr>
                  <td>\${photoTag}</td>
                  <td><strong>\${row.name}</strong></td>
                  <td>\${row.studentId}</td>
                  <td>\${row.yearLevel} - \${row.section}</td>
                  <td>\${row.event}</td>
                  <td>\${typeBadge}</td>
                  <td>\${statusBadge}</td>
                  <td><strong>\${row.duration || 'N/A'}</strong></td>
                  <td>\${row.timestamp}</td>
                </tr>
              \`;
            }).join('');
          }

          const stBody = document.getElementById('studentsTableBody');
          if (!registeredStudents || registeredStudents.length === 0) {
            stBody.innerHTML = '<tr><td colspan="8" style="text-align:center;">No students registered yet.</td></tr>';
          } else {
            stBody.innerHTML = registeredStudents.map(st => {
              const photoTag = st.photo ? \`<img src="\${st.photo}" class="student-thumb">\` : '<span style="color:#94a3b8; font-size:12px;">No photo</span>';
              const contactInfo = [st.email, st.phone].filter(Boolean).join('<br>') || '<span style="color:#aaa;">None</span>';

              return \`
                <tr>
                  <td>\${photoTag}</td>
                  <td>\${st.studentId}</td>
                  <td><strong>\${st.name}</strong></td>
                  <td>\${st.yearLevel || 'N/A'} - \${st.section || 'N/A'}</td>
                  <td><small>\${contactInfo}</small></td>
                  <td>\${st.assignedEvent || 'Regular Class'}</td>
                  <td>\${st.uid ? '<code>' + st.uid + '</code>' : '<span style="color:var(--warning); font-weight:bold;">No Card Linked</span>'}</td>
                  <td>
                    <button type="button" class="btn-warning" onclick="editStudent('\${st._id}')" style="padding:6px 10px; font-size:12px;">Edit</button>
                    <form action="/api/delete-student" method="POST" style="display:inline;" onsubmit="return confirm('Remove student record?');">
                      <input type="hidden" name="id" value="\${st._id}">
                      <button type="submit" class="btn-danger" style="padding:6px 10px; font-size:12px;">Delete</button>
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

      function downloadExcel() {
        const selected = document.getElementById('exportEventSelect').value;
        window.location.href = '/api/export-excel?event=' + encodeURIComponent(selected);
      }

      updateDashboard();
      setInterval(updateDashboard, 2000);
    </script>
  </body>
  </html>
  `);
});

// START SERVER
app.listen(PORT, async () => {
  await ensureExcelTemplateExists();
  console.log(`School Attendance Server running on port ${PORT}`);
});
