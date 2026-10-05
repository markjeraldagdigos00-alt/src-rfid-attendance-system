const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const https = require('https');
const querystring = require('querystring');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');
const { Resend } = require('resend');

const app = express();
const PORT = process.env.PORT || 3000;

// MONGOOSE DATABASE CONNECTION
const MONGO_URI = process.env.MONGO_URI || 'mongodb+srv://srcadmin:30005BNHS@cluster0.he7jspr.mongodb.net/scholarhub_db?appName=Cluster0';

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

  console.log('[EXCEL TEMPLATE] Generating official school attendance Excel template...');
  
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

  worksheet.mergeCells('C1:F1');
  worksheet.mergeCells('C2:F2');
  worksheet.mergeCells('C3:F3');

  worksheet.getCell('C1').value = 'BATAC NATIONAL HIGH SCHOOL';
  worksheet.getCell('C1').font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF1B365D' } };
  worksheet.getCell('C1').alignment = { horizontal: 'center', vertical: 'middle' };

  worksheet.getCell('C2').value = 'OFFICIAL SCHOOL ATTENDANCE SYSTEM';
  worksheet.getCell('C2').font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FF444444' } };
  worksheet.getCell('C2').alignment = { horizontal: 'center', vertical: 'middle' };

  worksheet.getCell('C3').value = 'DAILY ATTENDANCE REPORT LOG';
  worksheet.getCell('C3').font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FF777777' } };
  worksheet.getCell('C3').alignment = { horizontal: 'center', vertical: 'middle' };

  const headers = [
    'STUDENT ID', 'STUDENT NAME', 'GRADE & SECTION', 
    'EVENT / ACTIVITY', 'SCAN TYPE', 'STATUS', 
    'DURATION', 'TIMESTAMP'
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

  let logoPath = path.join(__dirname, 'bnhs_logo.jpg');
  if (!fs.existsSync(logoPath)) logoPath = path.join(__dirname, 'bnhs_logo.png');

  if (fs.existsSync(logoPath)) {
    const ext = path.extname(logoPath).toLowerCase() === '.png' ? 'png' : 'jpeg';
    const logoImage = workbook.addImage({ filename: logoPath, extension: ext });
    worksheet.addImage(logoImage, { tl: { col: 0, row: 0 }, ext: { width: 70, height: 70 } });
  }

  await workbook.xlsx.writeFile(templatePath);
  console.log('[EXCEL TEMPLATE] "template.xlsx" created successfully!');
}

// SCHEMAS & MODELS
const studentSchema = new mongoose.Schema({
  uid: { type: String, default: '' },
  studentId: { type: String, required: true },
  name: { type: String, required: true },
  gradeLevel: { type: String, default: 'Grade 7' },
  section: { type: String, default: 'Diamond' },
  email: { type: String, default: '' },
  phone: { type: String, default: '' },
  photo: { type: String, default: '' }, // Student Photo URL / Path
  assignedEvent: { type: String, default: 'Daily Attendance' }
});

const attendanceSchema = new mongoose.Schema({
  uid: String,
  name: String,
  studentId: String,
  gradeLevel: String,
  section: String,
  email: String,
  phone: String,
  photo: String,
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
  events: { type: [String], default: ['Daily Attendance', 'Morning Assembly', 'Exam Week', 'School Event'] },
  currentEvent: { type: String, default: 'Daily Attendance' },
  cutoffTime: { type: String, default: '07:45' },
  latestUid: { type: String, default: '' },
  lastScannedStudent: {
    name: { type: String, default: '' },
    studentId: { type: String, default: '' },
    gradeSection: { type: String, default: '' },
    scanType: { type: String, default: '' },
    status: { type: String, default: '' },
    photo: { type: String, default: '' },
    timestamp: { type: String, default: '' }
  },
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

// UPLOADS SETUP
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
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
      from: 'School RFID System <onboarding@resend.dev>',
      to: recipientEmail,
      subject: `[School Attendance] ${scanType} Notice for ${studentName}`,
      html: `
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px; max-width: 600px;">
          <h2 style="color: #1e3a8a;">School Attendance Notification (${scanType})</h2>
          <p>Dear Parent / Guardian,</p>
          <p>This is to formally notify you that <strong>${studentName}</strong> has successfully logged <strong>${scanType}</strong>.</p>
          <ul style="line-height: 1.6;">
            <li><strong>Event / Session:</strong> ${eventName}</li>
            <li><strong>Scan Type:</strong> <span style="color:#2563eb; font-weight:bold;">${scanType}</span></li>
            <li><strong>Status:</strong> <span style="color:${status === 'LATE' ? '#dc2626' : '#16a34a'}; font-weight:bold;">${status}</span></li>
            <li><strong>Timestamp:</strong> ${timestamp}</li>
            ${durationText}
          </ul>
          <p style="color: #64748b; font-size: 12px; margin-top: 20px;">Batac National High School Attendance System</p>
        </div>
      `
    });
  } catch (error) {
    console.error('[EMAIL ERROR]', error.message);
  }
}

function calculateDuration(timeInDate, timeOutDate) {
  const diffMs = timeOutDate - timeInDate;
  if (isNaN(diffMs) || diffMs < 0) return 'N/A';
  const totalMinutes = Math.floor(diffMs / (1000 * 60));
  const hours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;
  return hours > 0 ? `${hours}h ${mins}m` : `${mins}m`;
}

// API: ESP8266 / SCANNER ENDPOINT
app.post('/api/scan', async (req, res) => {
  try {
    const { uid } = req.body;
    if (!uid) return res.status(400).json({ status: 'error', message: 'No UID provided' });

    const cleanUid = uid.trim().toUpperCase();
    const config = await getConfig();
    
    config.latestUid = cleanUid;

    const student = await Student.findOne({ uid: cleanUid });
    const now = new Date();

    if (student) {
      const eventName = student.assignedEvent || config.currentEvent || 'Daily Attendance';
      const startOfDay = new Date(now);
      startOfDay.setHours(0, 0, 0, 0);
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

      const gradeSectionStr = `${student.gradeLevel || 'Grade 7'} - ${student.section || 'Diamond'}`;

      const record = new Attendance({
        uid: cleanUid,
        name: student.name,
        studentId: student.studentId,
        gradeLevel: student.gradeLevel || 'Grade 7',
        section: student.section || 'Diamond',
        email: student.email || '',
        phone: student.phone || '',
        photo: student.photo || '',
        event: eventName,
        scanType,
        status: statusLabel,
        duration: duration || 'N/A',
        timestamp: scanTime.toLocaleString(),
        rawTimestamp: scanTime
      });

      await record.save();

      // Update last scanned student in config for real-time display popup
      config.lastScannedStudent = {
        name: student.name,
        studentId: student.studentId,
        gradeSection: gradeSectionStr,
        scanType: scanType,
        status: statusLabel,
        photo: student.photo || '',
        timestamp: record.timestamp
      };
      await config.save();

      if (config.enableEmail && student.email) {
        sendEmailNotification(student.email, student.name, scanType, statusLabel, eventName, record.timestamp, duration);
      }

      return res.json({ 
        status: 'success', 
        scanType, 
        isLate: statusLabel === 'LATE', 
        student: {
          name: student.name,
          studentId: student.studentId,
          gradeSection: gradeSectionStr,
          photo: student.photo,
          scanType,
          status: statusLabel
        },
        message: `${scanType} recorded for ${student.name}` 
      });
    } else {
      return res.json({ status: 'unknown', message: 'Card not registered' });
    }
  } catch (err) {
    console.error('[SCAN ERROR]', err.message);
    res.status(500).json({ status: 'error', message: 'Server Error' });
  }
});

// REALTIME DATA ENDPOINT
app.get('/api/live-data', async (req, res) => {
  try {
    const config = await getConfig();
    const students = await Student.find().sort({ name: 1 });
    const attendance = await Attendance.find().sort({ rawTimestamp: -1 }).limit(100);

    res.json({
      latestUid: config.latestUid || '',
      lastScannedStudent: config.lastScannedStudent || null,
      attendance: attendance,
      students: students
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
      currentRow.getCell(3).value = `${row.gradeLevel} - ${row.section}`;
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

// SETTINGS & ADMIN ENDPOINTS
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
    config.logoPath = `/uploads/${req.file.filename}`;
    await config.save();
  }
  res.redirect('/');
});

app.post('/api/notification-settings', async (req, res) => {
  const { enableEmail, gmailUser, gmailPass } = req.body;
  const config = await getConfig();

  config.enableEmail = enableEmail === 'on';
  config.gmailUser = gmailUser || 'markjeraldagdigos00@gmail.com';
  if (gmailPass && gmailPass !== '******') config.gmailPass = gmailPass;

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
      config.currentEvent = config.events[0] || 'Daily Attendance';
    }
    await config.save();
  }
  res.redirect('/');
});

// REGISTER / UPDATE STUDENT WITH PHOTO UPLOAD
app.post('/api/register', upload.single('photoFile'), async (req, res) => {
  try {
    const { mongoId, uid, name, studentId, gradeLevel, section, assignedEvent, email, phone } = req.body;
    const cleanUid = uid ? uid.trim().toUpperCase() : '';
    let photoPath = req.file ? `/uploads/${req.file.filename}` : '';

    if (mongoId) {
      const existingStudent = await Student.findById(mongoId);
      if (!photoPath && existingStudent) photoPath = existingStudent.photo;

      await Student.findByIdAndUpdate(mongoId, {
        uid: cleanUid,
        name,
        studentId,
        gradeLevel: gradeLevel || 'Grade 7',
        section: section || 'Diamond',
        email: email || '',
        phone: phone || '',
        photo: photoPath,
        assignedEvent: assignedEvent || 'Daily Attendance'
      });
    } else {
      const newStudent = new Student({
        uid: cleanUid,
        name,
        studentId,
        gradeLevel: gradeLevel || 'Grade 7',
        section: section || 'Diamond',
        email: email || '',
        phone: phone || '',
        photo: photoPath,
        assignedEvent: assignedEvent || 'Daily Attendance'
      });
      await newStudent.save();
    }
  } catch (err) {
    console.error('[SAVE STUDENT ERROR]', err.message);
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

// ADMIN DASHBOARD & LIVE KIOSK VIEW ( / )
app.get('/', async (req, res) => {
  const config = await getConfig();
  const eventList = Array.isArray(config.events) ? config.events : ['Daily Attendance'];
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
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
      :root {
        --primary: #1e3a8a;
        --primary-light: #3b82f6;
        --success: #16a34a;
        --danger: #dc2626;
        --warning: #d97706;
        --bg: #f8fafc;
        --card-bg: #ffffff;
        --text: #1e293b;
        --border: #e2e8f0;
      }
      * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif; }
      body { background: var(--bg); color: var(--text); padding: 20px; }
      .header-container { display: flex; align-items: center; justify-content: space-between; background: var(--card-bg); padding: 20px 30px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); margin-bottom: 25px; }
      .header-left { display: flex; align-items: center; gap: 15px; }
      .header-logo { height: 60px; width: 60px; object-fit: cover; border-radius: 50%; border: 2px solid var(--primary); }
      h1 { font-size: 1.5rem; color: var(--primary); font-weight: 700; }
      .grid-container { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 25px; }
      @media (max-width: 1024px) { .grid-container { grid-template-columns: 1fr; } }
      .card { background: var(--card-bg); padding: 25px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border: 1px solid var(--border); }
      .card h2, .card h3 { color: var(--primary); margin-bottom: 15px; font-size: 1.2rem; border-bottom: 2px solid var(--bg); padding-bottom: 8px; }
      label { display: block; font-weight: 600; font-size: 0.85rem; color: #475569; margin-bottom: 5px; margin-top: 12px; }
      input, select { width: 100%; padding: 10px 14px; border: 1px solid var(--border); border-radius: 8px; font-size: 0.95rem; background: #fff; transition: all 0.2s; }
      input:focus, select:focus { outline: none; border-color: var(--primary-light); box-shadow: 0 0 0 3px rgba(59,130,246,0.1); }
      button, input[type="submit"] { background: var(--primary); color: white; padding: 10px 16px; border: none; border-radius: 8px; font-weight: 600; cursor: pointer; transition: background 0.2s; width: 100%; margin-top: 15px; }
      button:hover, input[type="submit"]:hover { background: #1d4ed8; }
      .btn-danger { background: var(--danger); }
      .btn-danger:hover { background: #b91c1c; }
      .btn-warning { background: var(--warning); width: auto; padding: 6px 12px; font-size: 0.8rem; }
      .btn-secondary { background: #475569; }
      .btn-secondary:hover { background: #334155; }
      
      /* LIVE KIOSK SCANNER MONITOR DISPLAY */
      .scanner-monitor { background: linear-gradient(135deg, #1e3a8a, #3b82f6); color: white; border-radius: 16px; padding: 30px; text-align: center; box-shadow: 0 10px 25px -5px rgba(30,58,138,0.3); margin-bottom: 25px; }
      .scanner-monitor h2 { color: white; border: none; font-size: 1.4rem; margin-bottom: 20px; }
      .scan-display-box { display: flex; align-items: center; justify-content: center; gap: 30px; background: rgba(255,255,255,0.1); backdrop-filter: blur(10px); padding: 25px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.2); }
      .scan-avatar { width: 120px; height: 120px; border-radius: 50%; object-fit: cover; border: 4px solid white; background: #e2e8f0; }
      .scan-info { text-align: left; }
      .scan-info h3 { font-size: 1.8rem; font-weight: 700; color: #fff; margin-bottom: 5px; border: none; padding: 0; }
      .scan-info p { font-size: 1.1rem; opacity: 0.9; margin-bottom: 4px; }
      .badge-scan { display: inline-block; padding: 6px 14px; border-radius: 20px; font-weight: 700; font-size: 0.85rem; margin-top: 8px; }
      .badge-in { background: #22c55e; color: white; }
      .badge-out { background: #a855f7; color: white; }
      .badge-late { background: #ef4444; color: white; }

      table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 0.9rem; }
      th, td { border: 1px solid var(--border); padding: 12px; text-align: left; }
      th { background: #f1f5f9; color: var(--primary); font-weight: 600; }
      tr:nth-child(even) { background: #fafafa; }
      
      .table-container { overflow-x: auto; }
      .hidden { display: none; }
      details.settings-card { background: white; border-radius: 12px; border: 1px solid var(--border); padding: 20px; margin-bottom: 25px; }
      details.settings-card summary { font-size: 1.1rem; font-weight: 600; color: var(--primary); cursor: pointer; }
    </style>
  </head>
  <body>

    <div class="header-container">
      <div class="header-left">
        ${logoHtml}
        <div>
          <h1>${config.systemName || 'School RFID Attendance System'}</h1>
          <p style="color: #64748b; font-size: 0.9rem;">Batac National High School • Real-Time Kiosk & Monitoring</p>
        </div>
      </div>
      <div>
        <span style="font-weight: 600; color: #475569;">Active Session:</span> 
        <strong style="color: var(--primary);">${config.currentEvent}</strong>
      </div>
    </div>

    <!-- LIVE KIOSK MONITOR SCREEN -->
    <div class="scanner-monitor">
      <h2>Live RFID Scanner Kiosk Display</h2>
      <div id="scanMonitorContent">
        <p style="font-size: 1.1rem; opacity: 0.8;">Waiting for RFID card scan... Swipe student ID card on scanner.</p>
      </div>
      <p style="margin-top: 15px; font-size: 0.85rem; opacity: 0.7;">Latest RFID Card UID: <code id="scannedUidDisplay" style="background: rgba(0,0,0,0.2); padding: 2px 6px; border-radius: 4px;">${config.latestUid || 'None'}</code></p>
    </div>

    <div class="grid-container">
      <!-- SYSTEM & EVENT SETTINGS -->
      <details class="settings-card">
        <summary>⚙️ System, Event & Notification Settings</summary>
        <form action="/api/update-system-name" method="POST" style="margin-top: 15px;">
          <label>System Name:</label>
          <input type="text" name="systemName" value="${config.systemName || 'School RFID Attendance System'}" required>
          <input type="submit" value="Update System Name">
        </form>

        <form action="/api/upload-logo" method="POST" enctype="multipart/form-data" style="margin-top: 15px;">
          <label>School Logo Image:</label>
          <input type="file" name="logoFile" accept="image/*" required>
          <input type="submit" value="Upload School Logo" class="btn-secondary">
        </form>

        <form action="/api/event-settings" method="POST" style="margin-top: 15px;">
          <label>Active Event / Session:</label>
          <select name="activeEvent">${eventOptions}</select>

          <label>Add New Event:</label>
          <input type="text" name="newEvent" placeholder="e.g. Intramurals Day 1">

          <label>Late Cut-off Time (HH:MM):</label>
          <input type="time" name="cutoffTime" value="${config.cutoffTime || '07:45'}">

          <input type="submit" value="Save Event Settings">
        </form>
      </details>

      <!-- STUDENT REGISTRATION & ID CARD LINKING -->
      <details class="settings-card" id="registrationCard">
        <summary id="formTitle">👤 Register / Edit Student Profile</summary>
        <form action="/api/register" method="POST" enctype="multipart/form-data" id="registerForm" style="margin-top: 15px;">
          <input type="hidden" id="mongoIdInput" name="mongoId">

          <label>RFID Card UID:</label>
          <div style="display: flex; gap: 8px;">
            <input type="text" id="uidInput" name="uid" placeholder="Scan or type UID">
            <button type="button" class="btn-secondary" onclick="useLatestUid()" style="margin-top:0; width:auto;">Use Last Scanned</button>
          </div>

          <label>Student ID Number:</label>
          <input type="text" id="studentIdInput" name="studentId" placeholder="e.g. 2026-1001" required>

          <label>Full Name:</label>
          <input type="text" id="nameInput" name="name" placeholder="Juan Dela Cruz" required>

          <div style="display: flex; gap: 10px;">
            <div style="flex: 1;">
              <label>Grade Level:</label>
              <select id="gradeLevelSelect" name="gradeLevel">${gradeOptions}</select>
            </div>
            <div style="flex: 1;">
              <label>Section:</label>
              <input type="text" id="sectionInput" name="section" placeholder="Diamond" required>
            </div>
          </div>

          <label>Parent / Guardian Email (For Notifications):</label>
          <input type="email" id="emailInput" name="email" placeholder="parent@gmail.com">

          <label>Student Photo (ID Picture):</label>
          <input type="file" name="photoFile" accept="image/*">

          <label>Assigned Event:</label>
          <select id="eventSelect" name="assignedEvent">${eventOptions}</select>

          <input type="submit" id="submitBtn" value="Save Student Record">
          <button type="button" id="cancelEditBtn" onclick="resetForm()" class="btn-danger hidden" style="margin-top: 8px;">Cancel Edit</button>
        </form>
      </details>
    </div>

    <!-- LIVE ATTENDANCE LOG TABLE -->
    <div class="card" style="margin-bottom: 25px;">
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
        <h2>📋 Live Attendance Log</h2>
        <div style="display: flex; gap: 10px; align-items: center;">
          <select id="exportEventSelect" style="width: auto; margin:0;">
            <option value="ALL">All Events</option>
            ${eventOptions}
          </select>
          <button type="button" onclick="downloadExcel()" style="width:auto; margin:0;">Export Excel</button>
          <form action="/api/clear-logs" method="POST" onsubmit="return confirm('Clear all attendance logs?');" style="margin:0;">
            <button type="submit" class="btn-danger" style="width:auto; margin:0;">Clear Logs</button>
          </form>
        </div>
      </div>
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Photo</th>
              <th>Student Name</th>
              <th>Student ID</th>
              <th>Grade & Section</th>
              <th>Event</th>
              <th>Scan Type</th>
              <th>Status</th>
              <th>Duration</th>
              <th>Timestamp</th>
            </tr>
          </thead>
          <tbody id="attendanceTableBody">
            <tr><td colspan="9" style="text-align:center;">Loading attendance logs...</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- REGISTERED STUDENTS DATABASE TABLE -->
    <div class="card">
      <h2>👥 Registered Students Database</h2>
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Photo</th>
              <th>Student ID</th>
              <th>Full Name</th>
              <th>Grade & Section</th>
              <th>Email</th>
              <th>RFID UID</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody id="studentsTableBody">
            <tr><td colspan="7" style="text-align:center;">Loading students...</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <script>
      let registeredStudents = [];
      let lastAnnouncedTimestamp = '';

      function editStudent(id) {
        const student = registeredStudents.find(s => s._id === id);
        if (!student) return;

        document.getElementById('mongoIdInput').value = student._id;
        document.getElementById('uidInput').value = student.uid || '';
        document.getElementById('studentIdInput').value = student.studentId;
        document.getElementById('nameInput').value = student.name;
        document.getElementById('emailInput').value = student.email || '';
        if (student.gradeLevel) document.getElementById('gradeLevelSelect').value = student.gradeLevel;
        if (student.section) document.getElementById('sectionInput').value = student.section;
        if (student.assignedEvent) document.getElementById('eventSelect').value = student.assignedEvent;

        const regCard = document.getElementById('registrationCard');
        regCard.open = true;
        document.getElementById('formTitle').innerText = '✏️ Edit Student (' + student.name + ')';
        document.getElementById('submitBtn').value = 'Update Student Record';
        document.getElementById('cancelEditBtn').classList.remove('hidden');
        regCard.scrollIntoView({ behavior: 'smooth' });
      }

      function resetForm() {
        document.getElementById('registerForm').reset();
        document.getElementById('mongoIdInput').value = '';
        document.getElementById('formTitle').innerText = '👤 Register / Edit Student Profile';
        document.getElementById('submitBtn').value = 'Save Student Record';
        document.getElementById('cancelEditBtn').classList.add('hidden');
      }

      function useLatestUid() {
        const uid = document.getElementById('scannedUidDisplay').innerText;
        if (uid && uid !== 'None') document.getElementById('uidInput').value = uid;
      }

      function speakAnnouncement(name, scanType, status) {
        if ('speechSynthesis' in window) {
          window.speechSynthesis.cancel(); // Stop previous speech
          const text = \`\${name}, \${scanType === 'TIME-IN' ? 'Time In' : 'Time Out'}. \${status}\`;
          const utterance = new SpeechSynthesisUtterance(text);
          utterance.rate = 1.0;
          window.speechSynthesis.speak(utterance);
        }
      }

      async function updateDashboard() {
        try {
          const res = await fetch('/api/live-data');
          const data = await res.json();

          if (data.latestUid) {
            document.getElementById('scannedUidDisplay').innerText = data.latestUid;
          }

          registeredStudents = data.students || [];

          // Update Live Kiosk Scanner Monitor
          if (data.lastScannedStudent && data.lastScannedStudent.name) {
            const st = data.lastScannedStudent;
            const photoUrl = st.photo ? st.photo : 'https://via.placeholder.com/120?text=No+Photo';
            
            document.getElementById('scanMonitorContent').innerHTML = \`
              <div class="scan-display-box">
                <img src="\${photoUrl}" alt="Student Photo" class="scan-avatar">
                <div class="scan-info">
                  <h3>\${st.name}</h3>
                  <p><strong>ID:</strong> \${st.studentId} | \${st.gradeSection}</p>
                  <p><strong>Time:</strong> \${st.timestamp}</p>
                  <div>
                    <span class="badge-scan \${st.scanType === 'TIME-IN' ? 'badge-in' : 'badge-out'}">\${st.scanType}</span>
                    <span class="badge-scan \${st.status === 'LATE' ? 'badge-late' : 'badge-in'}">\${st.status}</span>
                  </div>
                </div>
              </div>
            \`;

            // Speak announcement if new scan
            if (st.timestamp !== lastAnnouncedTimestamp) {
              lastAnnouncedTimestamp = st.timestamp;
              speakAnnouncement(st.name, st.scanType, st.status);
            }
          }

          // Update Attendance Table
          const tbody = document.getElementById('attendanceTableBody');
          if (!data.attendance || data.attendance.length === 0) {
            tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;">No attendance records found.</td></tr>';
          } else {
            tbody.innerHTML = data.attendance.map(row => {
              const photoImg = row.photo ? \`<img src="\${row.photo}" style="width:40px; height:40px; border-radius:50%; object-fit:cover;">\` : '<span style="color:#aaa;">No Photo</span>';
              const typeBadge = row.scanType === 'TIME-OUT' ? '<span style="background:#a855f7; color:white; padding:4px 8px; border-radius:4px; font-size:0.75rem; font-weight:700;">TIME-OUT</span>' : '<span style="background:#22c55e; color:white; padding:4px 8px; border-radius:4px; font-size:0.75rem; font-weight:700;">TIME-IN</span>';
              const statusBadge = row.status === 'LATE' ? '<span style="background:#ef4444; color:white; padding:4px 8px; border-radius:4px; font-size:0.75rem; font-weight:700;">LATE</span>' : '<span style="background:#3b82f6; color:white; padding:4px 8px; border-radius:4px; font-size:0.75rem; font-weight:700;">' + row.status + '</span>';

              return \`
                <tr>
                  <td>\${photoImg}</td>
                  <td><strong>\${row.name}</strong></td>
                  <td>\${row.studentId}</td>
                  <td>\${row.gradeLevel} - \${row.section}</td>
                  <td>\${row.event}</td>
                  <td>\${typeBadge}</td>
                  <td>\${statusBadge}</td>
                  <td><strong>\${row.duration || 'N/A'}</strong></td>
                  <td>\${row.timestamp}</td>
                </tr>
              \`;
            }).join('');
          }

          // Update Students Database Table
          const stBody = document.getElementById('studentsTableBody');
          if (!registeredStudents || registeredStudents.length === 0) {
            stBody.innerHTML = '<tr><td colspan="7" style="text-align:center;">No students registered yet.</td></tr>';
          } else {
            stBody.innerHTML = registeredStudents.map(st => {
              const photoImg = st.photo ? \`<img src="\${st.photo}" style="width:40px; height:40px; border-radius:50%; object-fit:cover;">\` : '<span style="color:#aaa;">No Photo</span>';
              return \`
                <tr>
                  <td>\${photoImg}</td>
                  <td>\${st.studentId}</td>
                  <td><strong>\${st.name}</strong></td>
                  <td>\${st.gradeLevel || 'Grade 7'} - \${st.section || 'Diamond'}</td>
                  <td>\${st.email || '<span style="color:#aaa;">None</span>'}</td>
                  <td>\${st.uid ? '<code>' + st.uid + '</code>' : '<span style="color:#d97706; font-weight:bold;">Unlinked</span>'}</td>
                  <td>
                    <button type="button" class="btn-warning" onclick="editStudent('\${st._id}')">Edit</button>
                    <form action="/api/delete-student" method="POST" style="display:inline;" onsubmit="return confirm('Delete student?');">
                      <input type="hidden" name="id" value="\${st._id}">
                      <button type="submit" class="btn-danger" style="width:auto; padding:6px 12px; font-size:0.8rem; margin-top:0;">Delete</button>
                    </form>
                  </td>
                </tr>
              \`;
            }).join('');
          }

        } catch (err) {
          console.error('Polling error:', err);
        }
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
  console.log(`[SERVER] School RFID Attendance System running on port ${PORT}`);
});
