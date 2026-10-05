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

  console.log('[EXCEL TEMPLATE] Generating "template.xlsx" with School template...');
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('School Attendance Log');

  worksheet.columns = [
    { key: 'studentId', width: 18 },
    { key: 'name', width: 28 },
    { key: 'gradeSection', width: 20 },
    { key: 'session', width: 18 },
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
    'ID NUMBER', 'STUDENT NAME', 'GRADE & SECTION', 
    'SESSION', 'SCAN TYPE', 'STATUS', 'DURATION', 'TIMESTAMP'
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

  await workbook.xlsx.writeFile(templatePath);
  console.log('[EXCEL TEMPLATE] "template.xlsx" created successfully!');
}

// SCHEMAS & MODELS
const studentSchema = new mongoose.Schema({
  uid: { type: String, default: '' },
  studentId: { type: String, required: true },
  name: { type: String, required: true },
  yearLevel: { type: String, default: 'Grade 7' },
  section: { type: String, default: 'Diamond' },
  photo: { type: String, default: '' },
  email: { type: String, default: '' },
  phone: { type: String, default: '' }
});

const attendanceSchema = new mongoose.Schema({
  uid: String,
  name: String,
  studentId: String,
  yearLevel: String,
  section: String,
  photo: String,
  session: String, // 'Morning' or 'Afternoon'
  scanType: String, // 'TIME-IN' or 'TIME-OUT'
  status: String,   // 'ON TIME', 'LATE', 'COMPLETED'
  duration: String,
  timestamp: String,
  rawTimestamp: { type: Date, default: Date.now }
});

const configSchema = new mongoose.Schema({
  systemName: { type: String, default: 'School RFID Attendance System' },
  logoPath: { type: String, default: '' },
  morningCutoff: { type: String, default: '07:45' },
  afternoonCutoff: { type: String, default: '13:00' },
  latestUid: { type: String, default: '' },
  lastScannedStudent: { type: Object, default: null },
  enableEmail: { type: Boolean, default: false },
  gmailUser: { type: String, default: '' },
  gmailPass: { type: String, default: '' }
});

const Student = mongoose.model('Student', studentSchema);
const Attendance = mongoose.model('Attendance', attendanceSchema);
const Config = mongoose.model('Config', configSchema);

async function getConfig() {
  let config = await Config.findOne();
  if (!config) {
    config = await Config.create({});
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

const resend = new Resend(process.env.RESEND_API_KEY || 'YOUR_RESEND_API_KEY');

async function sendEmailNotification(recipientEmail, studentName, session, scanType, status, timestamp, duration) {
  if (!recipientEmail) return;
  try {
    await resend.emails.send({
      from: 'School Attendance <onboarding@resend.dev>',
      to: recipientEmail,
      subject: `[School Attendance] ${studentName} - ${session} ${scanType}`,
      html: `
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e1e8ed; border-radius: 8px;">
          <h2 style="color: #1b365d;">Attendance Confirmation</h2>
          <p>Hello Parent/Guardian,</p>
          <p>This is to notify you that <strong>${studentName}</strong> has logged <strong>${session} ${scanType}</strong>.</p>
          <ul>
            <li><strong>Status:</strong> <span style="color:${status === 'LATE' ? '#e74c3c' : '#2ecc71'}; font-weight:bold;">${status}</span></li>
            <li><strong>Time:</strong> ${timestamp}</li>
            ${duration ? `<li><strong>Duration:</strong> ${duration}</li>` : ''}
          </ul>
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
    const student = await Student.findOne({ uid: cleanUid });
    const now = new Date();

    if (!student) {
      return res.json({ status: 'unknown', message: 'RFID Card not registered to any student.' });
    }

    const currentHour = now.getHours();
    const session = currentHour < 12 ? 'Morning' : 'Afternoon';
    const cutoff = session === 'Morning' ? (config.morningCutoff || '07:45') : (config.afternoonCutoff || '13:00');

    const startOfDay = new Date(now);
    startOfDay.setHours(0, 0, 0, 0);

    // Check existing logs for today in the same session
    const existingLogs = await Attendance.find({
      uid: cleanUid,
      session: session,
      rawTimestamp: { $gte: startOfDay }
    }).sort({ rawTimestamp: 1 });

    let scanType = 'TIME-IN';
    let statusLabel = 'ON TIME';
    let duration = '';

    if (existingLogs.length === 0) {
      scanType = 'TIME-IN';
      const currentTimeStr = now.toTimeString().slice(0, 5);
      statusLabel = currentTimeStr > cutoff ? 'LATE' : 'ON TIME';
    } else if (existingLogs.length === 1 && existingLogs[0].scanType === 'TIME-IN') {
      scanType = 'TIME-OUT';
      statusLabel = 'COMPLETED';
      duration = calculateDuration(new Date(existingLogs[0].rawTimestamp), now);
    } else {
      scanType = 'TIME-IN';
      statusLabel = 'ON TIME';
    }

    const record = new Attendance({
      uid: cleanUid,
      name: student.name,
      studentId: student.studentId,
      yearLevel: student.yearLevel,
      section: student.section,
      photo: student.photo || '',
      session,
      scanType,
      status: statusLabel,
      duration: duration || 'N/A',
      timestamp: now.toLocaleString(),
      rawTimestamp: now
    });

    await record.save();

    config.latestUid = cleanUid;
    config.lastScannedStudent = {
      name: student.name,
      studentId: student.studentId,
      yearLevel: student.yearLevel,
      section: student.section,
      photo: student.photo || '',
      session,
      scanType,
      status: statusLabel,
      timestamp: record.timestamp
    };
    await config.save();

    if (config.enableEmail && student.email) {
      sendEmailNotification(student.email, student.name, session, scanType, statusLabel, record.timestamp, duration);
    }

    return res.json({ 
      status: 'success', 
      session,
      scanType, 
      student: config.lastScannedStudent,
      message: `${session} ${scanType} recorded for ${student.name}` 
    });

  } catch (err) {
    console.error('[SCAN ERROR]', err);
    res.status(500).json({ status: 'error', message: 'Server Error during scan' });
  }
});

// REALTIME DATA ENDPOINT FOR DASHBOARD & SCANNER SCREEN
app.get('/api/live-data', async (req, res) => {
  try {
    const config = await getConfig();
    const students = await Student.find().sort({ yearLevel: 1, section: 1, name: 1 });
    const attendance = await Attendance.find().sort({ rawTimestamp: -1 }).limit(100);

    res.json({
      latestUid: config.latestUid || '',
      lastScannedStudent: config.lastScannedStudent || null,
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
    const attendance = await Attendance.find().sort({ rawTimestamp: -1 });

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
      currentRow.getCell(4).value = `${row.session} (${row.scanType})`;
      currentRow.getCell(5).value = row.scanType;
      currentRow.getCell(6).value = row.status;
      currentRow.getCell(7).value = row.duration || 'N/A';
      currentRow.getCell(8).value = row.timestamp;
      currentRow.commit();
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="School_Attendance_Report.xlsx"');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    res.status(500).send('Error generating Excel: ' + err.message);
  }
});

// REGISTER / EDIT STUDENT API (WITH PHOTO UPLOAD)
app.post('/api/register-student', upload.single('photoFile'), async (req, res) => {
  try {
    const { mongoId, uid, studentId, name, yearLevel, section, email, phone } = req.body;
    let photoPath = req.file ? `/uploads/${req.file.filename}` : req.body.existingPhoto || '';
    const cleanUid = uid ? uid.trim().toUpperCase() : '';

    if (mongoId) {
      const updateData = { studentId, name, yearLevel, section, email, phone };
      if (cleanUid) updateData.uid = cleanUid;
      if (photoPath) updateData.photo = photoPath;
      await Student.findByIdAndUpdate(mongoId, updateData);
    } else {
      const newStudent = new Student({
        uid: cleanUid,
        studentId,
        name,
        yearLevel: yearLevel || 'Grade 7',
        section: section || 'Diamond',
        photo: photoPath,
        email: email || '',
        phone: phone || ''
      });
      await newStudent.save();
    }
  } catch (err) {
    console.error('[REGISTRATION ERROR]', err.message);
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
  const config = await getConfig();
  config.lastScannedStudent = null;
  config.latestUid = '';
  await config.save();
  res.redirect('/');
});

// ==========================================
// SEPARATE SCANNER KIOSK SCREEN ( /scanner )
// ==========================================
app.get('/scanner', async (req, res) => {
  const config = await getConfig();
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>RFID Entrance Scanner Kiosk</title>
      <style>
        body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; display: flex; flex-direction: column; height: 100vh; overflow: hidden; }
        .header { background: #1e293b; padding: 20px 40px; display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #334155; }
        .header h1 { margin: 0; font-size: 24px; color: #38bdf8; display: flex; align-items: center; gap: 12px; }
        .clock { font-size: 20px; font-weight: bold; color: #cbd5e1; }
        .main-container { flex: 1; display: flex; gap: 30px; padding: 40px; justify-content: center; align-items: center; }
        .scan-prompt-card { background: #1e293b; border: 3px dashed #475569; border-radius: 20px; padding: 60px; text-align: center; width: 450px; box-shadow: 0 10px 25px rgba(0,0,0,0.3); }
        .scan-prompt-card h2 { color: #94a3b8; margin-bottom: 10px; font-size: 28px; }
        .pulse-icon { font-size: 70px; margin: 20px 0; animation: pulse 1.5s infinite; }
        @keyframes pulse { 0% { transform: scale(1); opacity: 1; } 50% { transform: scale(1.1); opacity: 0.7; } 100% { transform: scale(1); opacity: 1; } }
        
        .student-display-card { background: #1e293b; border: 2px solid #38bdf8; border-radius: 20px; padding: 40px; display: flex; gap: 40px; align-items: center; width: 750px; box-shadow: 0 15px 35px rgba(56,189,248,0.2); }
        .student-photo { width: 220px; height: 220px; border-radius: 50%; object-fit: cover; border: 5px solid #38bdf8; background: #334155; box-shadow: 0 8px 20px rgba(0,0,0,0.5); }
        .student-info { flex: 1; }
        .badge { display: inline-block; padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 14px; margin-bottom: 15px; }
        .badge-ontime { background: #22c55e; color: #fff; }
        .badge-late { background: #ef4444; color: #fff; }
        .badge-session { background: #6366f1; color: #fff; margin-left: 8px; }
        .student-info h2 { font-size: 36px; margin: 0 0 10px 0; color: #ffffff; }
        .student-info p { margin: 6px 0; font-size: 18px; color: #cbd5e1; }
        .student-info strong { color: #38bdf8; }
      </style>
    </head>
    <body>
      <div class="header">
        <h1>Batac National High School - RFID Scanner Kiosk</h1>
        <div class="clock" id="liveClock"></div>
      </div>

      <div class="main-container" id="displayArea">
        <div class="scan-prompt-card">
          <div class="pulse-icon">📡</div>
          <h2>Ready to Scan ID</h2>
          <p style="color: #64748b; font-size: 16px;">Please tap your RFID ID card on the scanner device.</p>
        </div>
      </div>

      <script>
        function updateClock() {
          const now = new Date();
          document.getElementById('liveClock').innerText = now.toLocaleDateString() + ' ' + now.toLocaleTimeString();
        }
        setInterval(updateClock, 1000);
        updateClock();

        let lastTimestamp = '';
        let lastSpokenName = '';

        async function pollScanner() {
          try {
            const res = await fetch('/api/live-data');
            const data = await res.json();
            const student = data.lastScannedStudent;

            if (student && student.timestamp !== lastTimestamp) {
              lastTimestamp = student.timestamp;
              
              // Speak student name aloud using Text-to-Speech
              if ('speechSynthesis' in window && student.name !== lastSpokenName) {
                lastSpokenName = student.name;
                const utterance = new SpeechSynthesisUtterance('Welcome, ' + student.name);
                utterance.rate = 1.0;
                window.speechSynthesis.speak(utterance);
              }

              const photoSrc = student.photo ? student.photo : 'https://cdn-icons-png.flaticon.com/512/149/149071.png';
              const statusClass = student.status === 'LATE' ? 'badge-late' : 'badge-ontime';

              document.getElementById('displayArea').innerHTML = \`
                <div class="student-display-card">
                  <img src="\${photoSrc}" alt="Student Photo" class="student-photo" onerror="this.src='https://cdn-icons-png.flaticon.com/512/149/149071.png'">
                  <div class="student-info">
                    <div>
                      <span class="badge \${statusClass}">\${student.status}</span>
                      <span class="badge badge-session">\${student.session} \${student.scanType}</span>
                    </div>
                    <h2>\${student.name}</h2>
                    <p><strong>ID Number:</strong> \${student.studentId}</p>
                    <p><strong>Grade & Section:</strong> \${student.yearLevel} - \${student.section}</p>
                    <p><strong>Time Scanned:</strong> \${student.timestamp}</p>
                  </div>
                </div>
              \`;
            }
          } catch (e) {
            console.error(e);
          }
        }

        setInterval(pollScanner, 1500);
      </script>
    </body>
    </html>
  `);
});

// ==========================================
// ADMIN DASHBOARD ( / )
// ==========================================
app.get('/', async (req, res) => {
  const config = await getConfig();
  let gradeOptions = '';
  for (let i = 7; i <= 12; i++) {
    gradeOptions += `<option value="Grade ${i}">Grade ${i}</option>`;
  }

  res.send(`
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>School Attendance Admin Dashboard</title>
    <style>
      :root { --primary: #1b365d; --accent: #38bdf8; --bg: #f8fafc; --card-bg: #ffffff; --text: #1e293b; --danger: #ef4444; --success: #22c55e; }
      body { font-family: 'Segoe UI', Tahoma, sans-serif; margin: 0; padding: 25px; background: var(--bg); color: var(--text); }
      .header-container { display: flex; justify-content: space-between; align-items: center; background: var(--card-bg); padding: 20px 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); margin-bottom: 25px; }
      h1, h2, h3 { color: var(--primary); margin: 0; }
      .btn { background: #2563eb; color: white; padding: 10px 18px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; text-decoration: none; display: inline-block; }
      .btn:hover { background: #1d4ed8; }
      .btn-danger { background: var(--danger); }
      .btn-success { background: var(--success); }
      .btn-secondary { background: #64748b; }
      .container { display: grid; grid-template-columns: 1fr 2fr; gap: 25px; }
      @media(max-width: 1024px) { .container { grid-template-columns: 1fr; } }
      .card { background: var(--card-bg); padding: 25px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); margin-bottom: 25px; }
      input, select { width: 100%; padding: 10px 12px; margin: 6px 0 16px 0; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box; }
      label { font-weight: 600; font-size: 14px; color: #475569; }
      table { width: 100%; border-collapse: collapse; margin-top: 15px; background: white; border-radius: 8px; overflow: hidden; }
      th, td { border-bottom: 1px solid #e2e8f0; padding: 12px 15px; text-align: left; font-size: 14px; }
      th { background: var(--primary); color: white; font-weight: 600; }
      .badge { padding: 5px 10px; border-radius: 6px; font-weight: bold; font-size: 12px; }
      .badge-ontime { background: #dcfce7; color: #166534; }
      .badge-late { background: #fee2e2; color: #991b1b; }
      .badge-in { background: #e0f2fe; color: #0369a1; }
      .badge-out { background: #f3e8ff; color: #6b21a8; }
      .nav-links { display: flex; gap: 15px; align-items: center; }
      .student-avatar { width: 45px; height: 45px; border-radius: 50%; object-fit: cover; border: 2px solid #cbd5e1; }
      .hidden { display: none; }
    </style>
  </head>
  <body>

    <div class="header-container">
      <div>
        <h1>Batac National High School</h1>
        <p style="margin: 5px 0 0 0; color: #64748b;">School RFID Attendance & Monitoring System</p>
      </div>
      <div class="nav-links">
        <a href="/scanner" target="_blank" class="btn btn-success">📺 Open Kiosk Scanner Screen</a>
        <a href="/api/export-excel" class="btn">📊 Download Excel Report</a>
      </div>
    </div>

    <div class="container">
      <!-- Left Column: Register / Edit Form -->
      <div>
        <div class="card">
          <h2 id="formTitle" style="margin-bottom: 15px;">Register Student</h2>
          <p style="font-size: 13px; color: #64748b; margin-bottom: 15px;">Last Scanned RFID UID: <strong id="scannedUid" style="color: #2563eb;">${config.latestUid || 'None'}</strong></p>
          
          <form action="/api/register-student" method="POST" enctype="multipart/form-data" id="registerForm">
            <input type="hidden" id="mongoIdInput" name="mongoId">
            <input type="hidden" id="existingPhotoInput" name="existingPhoto">

            <label>RFID Card UID:</label>
            <input type="text" id="uidInput" name="uid" placeholder="Tap card or paste UID">

            <label>Student ID Number:</label>
            <input type="text" id="studentIdInput" name="studentId" placeholder="e.g. 2026-1001" required>
            
            <label>Full Name:</label>
            <input type="text" id="nameInput" name="name" placeholder="Juan Dela Cruz" required>

            <div style="display: flex; gap: 10px;">
              <div style="flex: 1;">
                <label>Grade Level:</label>
                <select id="yearLevelSelect" name="yearLevel">${gradeOptions}</select>
              </div>
              <div style="flex: 1;">
                <label>Section:</label>
                <input type="text" id="sectionInput" name="section" placeholder="Diamond" required>
              </div>
            </div>

            <label>Student Photo:</label>
            <input type="file" name="photoFile" accept="image/*">

            <label>Parent Email (for notifications):</label>
            <input type="email" id="emailInput" name="email" placeholder="parent@gmail.com">

            <label>Phone Number (Optional):</label>
            <input type="tel" id="phoneInput" name="phone" placeholder="09171234567">

            <div style="display: flex; gap: 10px; margin-top: 15px;">
              <button type="button" class="btn-secondary" onclick="useLatestUid()" style="flex: 1;">Link Last UID</button>
              <button type="submit" class="btn" style="flex: 1;" id="submitBtn">Save Student</button>
            </div>
            <button type="button" id="cancelEditBtn" onclick="resetForm()" class="btn btn-danger hidden" style="width: 100%; margin-top: 10px;">Cancel Edit</button>
          </form>
        </div>
      </div>

      <!-- Right Column: Registered Students Database & Live Logs -->
      <div>
        <div class="card">
          <h2>Registered Students Database</h2>
          <div style="max-height: 380px; overflow-y: auto;">
            <table>
              <thead>
                <tr>
                  <th>Photo</th>
                  <th>ID Number</th>
                  <th>Name</th>
                  <th>Grade & Section</th>
                  <th>Card UID</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody id="studentsTableBody"></tbody>
            </table>
          </div>
        </div>

        <div class="card">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <h2>Live Attendance Logs</h2>
            <form action="/api/clear-logs" method="POST" onsubmit="return confirm('Clear all attendance logs?');">
              <button type="submit" class="btn btn-danger" style="padding: 6px 12px; font-size: 13px;">Clear Logs</button>
            </form>
          </div>
          <div style="max-height: 400px; overflow-y: auto;">
            <table>
              <thead>
                <tr>
                  <th>Student</th>
                  <th>ID</th>
                  <th>Session</th>
                  <th>Type</th>
                  <th>Status</th>
                  <th>Duration</th>
                  <th>Timestamp</th>
                </tr>
              </thead>
              <tbody id="attendanceTableBody"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <script>
      let registeredStudents = [];

      function editStudent(id) {
        const student = registeredStudents.find(s => s._id === id);
        if (!student) return;

        document.getElementById('mongoIdInput').value = student._id;
        document.getElementById('uidInput').value = student.uid || '';
        document.getElementById('studentIdInput').value = student.studentId;
        document.getElementById('nameInput').value = student.name;
        document.getElementById('yearLevelSelect').value = student.yearLevel || 'Grade 7';
        document.getElementById('sectionInput').value = student.section || 'Diamond';
        document.getElementById('emailInput').value = student.email || '';
        document.getElementById('phoneInput').value = student.phone || '';
        document.getElementById('existingPhotoInput').value = student.photo || '';

        document.getElementById('formTitle').innerText = 'Edit Student (' + student.name + ')';
        document.getElementById('submitBtn').innerText = 'Update Student';
        document.getElementById('cancelEditBtn').classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }

      function resetForm() {
        document.getElementById('registerForm').reset();
        document.getElementById('mongoIdInput').value = '';
        document.getElementById('existingPhotoInput').value = '';
        document.getElementById('formTitle').innerText = 'Register Student';
        document.getElementById('submitBtn').innerText = 'Save Student';
        document.getElementById('cancelEditBtn').classList.add('hidden');
      }

      function useLatestUid() {
        const uid = document.getElementById('scannedUid').innerText;
        if (uid && uid !== 'None') document.getElementById('uidInput').value = uid;
      }

      async function updateDashboard() {
        try {
          const res = await fetch('/api/live-data');
          const data = await res.json();
          if (data.latestUid) document.getElementById('scannedUid').innerText = data.latestUid;
          registeredStudents = data.students || [];

          // Render Students Table
          const stBody = document.getElementById('studentsTableBody');
          if (registeredStudents.length === 0) {
            stBody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#64748b;">No students registered yet.</td></tr>';
          } else {
            stBody.innerHTML = registeredStudents.map(st => {
              const photo = st.photo ? st.photo : 'https://cdn-icons-png.flaticon.com/512/149/149071.png';
              return \`
                <tr>
                  <td><img src="\${photo}" class="student-avatar" onerror="this.src='https://cdn-icons-png.flaticon.com/512/149/149071.png'"></td>
                  <td>\${st.studentId}</td>
                  <td><strong>\${st.name}</strong></td>
                  <td>\${st.yearLevel} - \${st.section}</td>
                  <td><code>\${st.uid || 'Not Linked'}</code></td>
                  <td>
                    <button type="button" class="btn" style="padding: 5px 10px; font-size:12px;" onclick="editStudent('\${st._id}')">Edit</button>
                    <form action="/api/delete-student" method="POST" style="display:inline;" onsubmit="return confirm('Delete student?');">
                      <input type="hidden" name="id" value="\${st._id}">
                      <button type="submit" class="btn btn-danger" style="padding: 5px 10px; font-size:12px;">Del</button>
                    </form>
                  </td>
                </tr>
              \`;
            }).join('');
          }

          // Render Attendance Table
          const attBody = document.getElementById('attendanceTableBody');
          const attendance = data.attendance || [];
          if (attendance.length === 0) {
            attBody.innerHTML = '<tr><td colspan="7" style="text-align:center; color:#64748b;">No attendance logs yet.</td></tr>';
          } else {
            attBody.innerHTML = attendance.map(row => {
              const statusBadge = row.status === 'LATE' ? 'badge-late' : 'badge-ontime';
              const typeBadge = row.scanType === 'TIME-OUT' ? 'badge-out' : 'badge-in';
              const photo = row.photo ? row.photo : 'https://cdn-icons-png.flaticon.com/512/149/149071.png';

              return \`
                <tr>
                  <td>
                    <div style="display: flex; align-items: center; gap: 10px;">
                      <img src="\\${photo}" class="student-avatar" style="width: 35px; height: 35px;" onerror="this.src='https://cdn-icons-png.flaticon.com/512/149/149071.png'">
                      <strong>\\${row.name}</strong>
                    </div>
                  </td>
                  <td>\\${row.studentId}</td>
                  <td><span class="badge" style="background:#e2e8f0; color:#334155;">\\${row.session}</span></td>
                  <td><span class="badge \\${typeBadge}">\\${row.scanType}</span></td>
                  <td><span class="badge \\${statusBadge}">\\${row.status}</span></td>
                  <td><strong>\\${row.duration || 'N/A'}</strong></td>
                  <td>\\${row.timestamp}</td>
                </tr>
              \`;
            }).join('');
          }
        } catch (e) {}
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
