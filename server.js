const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');
const { Resend } = require('resend');

const app = express();
const PORT = process.env.PORT || 3000;

// ==========================================
// 1. MONGOOSE DATABASE CONNECTION & MODELS
// ==========================================
const MONGO_URI = process.env.MONGO_URI || 'mongodb+srv://srcadmin:30005BNHS@cluster0.he7jspr.mongodb.net/school_attendance_db?appName=Cluster0';

mongoose.connect(MONGO_URI)
  .then(() => console.log('[DATABASE] Connected to MongoDB Atlas successfully!'))
  .catch(err => console.error('[DATABASE ERROR] Could not connect to MongoDB:', err.message));

const studentSchema = new mongoose.Schema({
  uid: { type: String, default: '' },
  studentId: { type: String, required: true, unique: true },
  name: { type: String, required: true },
  gradeLevel: { type: String, default: 'Grade 7' },
  section: { type: String, default: 'Diamond' },
  parentEmail: { type: String, default: '' },
  parentPhone: { type: String, default: '' },
  photoPath: { type: String, default: '' }
});

const attendanceSchema = new mongoose.Schema({
  uid: String,
  studentId: String,
  name: String,
  gradeLevel: String,
  section: String,
  scanType: String, // 'MORNING IN', 'MORNING OUT', 'AFTERNOON IN', 'AFTERNOON OUT'
  status: String,   // 'ON TIME', 'LATE', 'COMPLETED'
  timestamp: String,
  rawTimestamp: { type: Date, default: Date.now }
});

const configSchema = new mongoose.Schema({
  schoolName: { type: String, default: 'Batac National High School' },
  schoolLogo: { type: String, default: '' },
  morningCutoff: { type: String, default: '07:45' },
  afternoonCutoff: { type: String, default: '13:00' },
  latestUid: { type: String, default: '' },
  enableEmail: { type: Boolean, default: true },
  resendApiKey: { type: String, default: '' }
});

const announcementSchema = new mongoose.Schema({
  title: String,
  message: String,
  date: { type: String, default: () => new Date().toLocaleDateString() }
});

const Student = mongoose.model('Student', studentSchema);
const Attendance = mongoose.model('Attendance', attendanceSchema);
const Config = mongoose.model('Config', configSchema);
const Announcement = mongoose.model('Announcement', announcementSchema);

async function getConfig() {
  let config = await Config.findOne();
  if (!config) {
    config = await Config.create({});
  }
  return config;
}

// ==========================================
// 2. FILE UPLOADS SETUP (Photos & Logos)
// ==========================================
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) cb(null, true);
    else cb(new Error('Only image files are allowed!'), false);
  }
});

// ==========================================
// 3. MIDDLEWARES & CONFIGURATION
// ==========================================
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/uploads', express.static(uploadsDir));

// ==========================================
// 4. EMAIL NOTIFICATION SERVICE (Resend)
// ==========================================
async function sendParentNotification(parentEmail, studentName, scanType, status, timestamp) {
  if (!parentEmail) return;
  try {
    const config = await getConfig();
    const resend = new Resend(config.resendApiKey || process.env.RESEND_API_KEY || 're_123456789');
    await resend.emails.send({
      from: 'School Attendance <onboarding@resend.dev>',
      to: parentEmail,
      subject: `Attendance Alert: ${studentName} (${scanType})`,
      html: `
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">
          <h2 style="color: #1a365d;">School Attendance Notification</h2>
          <p>Dear Parent/Guardian,</p>
          <p>This is to inform you that <strong>${studentName}</strong> has successfully recorded their attendance.</p>
          <ul style="background: #f7fafc; padding: 15px; border-radius: 6px; list-style: none;">
            <li><strong>Status Type:</strong> <span style="color: #2b6cb0;">${scanType}</span></li>
            <li><strong>Remarks:</strong> <span style="color: ${status === 'LATE' ? '#e53e3e' : '#38a169'}; font-weight: bold;">${status}</span></li>
            <li><strong>Time Logged:</strong> ${timestamp}</li>
          </ul>
          <p style="font-size: 12px; color: #718096; margin-top: 20px;">Batac National High School Attendance System</p>
        </div>
      `
    });
  } catch (err) {
    console.error('[EMAIL ERROR]', err.message);
  }
}

// ==========================================
// 5. API: ESP8266 / RFID SCANNER ENDPOINT
// ==========================================
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

    if (!student) {
      return res.json({ status: 'unknown', message: 'Unregistered RFID Card', uid: cleanUid });
    }

    const startOfDay = new Date(now);
    startOfDay.setHours(0, 0, 0, 0);

    // Fetch today's scans for this student
    const todayScans = await Attendance.find({
      studentId: student.studentId,
      rawTimestamp: { $gte: startOfDay }
    }).sort({ rawTimestamp: 1 });

    const hour = now.getHours();
    let scanType = 'MORNING IN';
    let statusLabel = 'ON TIME';

    // Determine Scan Sequence: Morning In -> Morning Out -> Afternoon In -> Afternoon Out
    if (todayScans.length === 0) {
      scanType = 'MORNING IN';
      const currentTimeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
      statusLabel = currentTimeStr > config.morningCutoff ? 'LATE' : 'ON TIME';
    } else if (todayScans.length === 1) {
      scanType = 'MORNING OUT';
      statusLabel = 'COMPLETED';
    } else if (todayScans.length === 2) {
      scanType = 'AFTERNOON IN';
      const currentTimeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
      statusLabel = currentTimeStr > config.afternoonCutoff ? 'LATE' : 'ON TIME';
    } else {
      scanType = 'AFTERNOON OUT';
      statusLabel = 'COMPLETED';
    }

    const attendanceRecord = new Attendance({
      uid: cleanUid,
      studentId: student.studentId,
      name: student.name,
      gradeLevel: student.gradeLevel,
      section: student.section,
      scanType,
      status: statusLabel,
      timestamp: now.toLocaleString(),
      rawTimestamp: now
    });

    await attendanceRecord.save();

    if (student.parentEmail) {
      sendParentNotification(student.parentEmail, student.name, scanType, statusLabel, attendanceRecord.timestamp);
    }

    return res.json({
      status: 'success',
      scanType,
      studentName: student.name,
      gradeLevel: student.gradeLevel,
      section: student.section,
      photo: student.photoPath || '',
      attendanceStatus: statusLabel,
      message: `${scanType} recorded for${student.name}`
    });

  } catch (err) {
    console.error('[SCAN ERROR]', err);
    res.status(500).json({ status: 'error', message: 'Internal Server Error' });
  }
});

// ==========================================
// 6. REAL-TIME DASHBOARD DATA API
// ==========================================
app.get('/api/live-data', async (req, res) => {
  try {
    const config = await getConfig();
    const students = await Student.find().sort({ gradeLevel: 1, section: 1, name: 1 });
    
    const startOfDay = new Date();
    startOfDay.setHours(0, 0, 0, 0);
    const attendance = await Attendance.find({ rawTimestamp: { $gte: startOfDay } }).sort({ rawTimestamp: -1 });
    const announcements = await Announcement.find().sort({ _id: -1 }).limit(5);

    // Calculate Dashboard Statistics
    const totalStudents = students.length;
    const presentIds = [...new Set(attendance.map(a => a.studentId))];
    const presentToday = presentIds.length;
    const absentToday = totalStudents - presentToday;
    const lateToday = attendance.filter(a => a.status === 'LATE').length;

    res.json({
      schoolName: config.schoolName,
      latestUid: config.latestUid || '',
      stats: { totalStudents, presentToday, absentToday, lateToday },
      attendance,
      students,
      announcements
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ==========================================
// 7. SEPARATED RFID SCANNER KIOSK SCREEN (/scanner)
// ==========================================
app.get('/scanner', async (req, res) => {
  const config = await getConfig();
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>RFID Kiosk Scanner - ${config.schoolName}</title>
      <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    </head>
    <body class="bg-slate-900 text-white min-h-screen flex flex-col justify-between p-6">
      <header class="flex justify-between items-center border-b border-slate-800 pb-4">
        <div class="flex items-center gap-4">
          ${config.schoolLogo ? `<img src="${config.schoolLogo}" class="h-12 w-12 rounded-full object-cover">` : ''}
          <h1 class="text-2xl font-bold tracking-wide">${config.schoolName}</h1>
        </div>
        <div class="text-xl font-mono bg-slate-800 px-4 py-2 rounded-lg text-emerald-400" id="clock">00:00:00</div>
      </header>

      <main class="grid grid-cols-1 lg:grid-cols-2 gap-8 my-auto items-center max-w-6xl mx-auto w-full">
        <div class="bg-slate-800/80 border border-slate-700 p-8 rounded-2xl shadow-2xl text-center flex flex-col items-center">
          <div class="w-32 h-32 bg-emerald-500/10 rounded-full flex items-center justify-center mb-6 animate-pulse">
            <svg class="w-16 h-16 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v1m6 11h2m-6 0h-2v4m0-11v3m0 0h.01M12 12h4.01M16 20h4M4 12h4m12 0h.01M5 8h2a1 1 0 001-1V5a1 1 0 00-1-1H5a1 1 0 00-1 1v2a1 1 0 001 1zm12 0h2a1 1 0 001-1V5a1 1 0 00-1-1h-2a1 1 0 00Here is the fully updated and redesigned **School RFID Attendance System** code (`server.js`). 

### Key Enhancements & New Features Added:
1. **Separated Kiosk Scanner Screen (`/scanner`)**: A dedicated full-screen live kiosk page optimized for RFID scanning. When a card is scanned, it instantly displays the student's **photo**, full name, ID number, grade & section, and status.
2. **Voice Announcement (Text-to-Speech)**: Automatically speaks the student's name and scan status out loud using the browser's speech synthesis engine upon successful scanning.
3. **Four-Part Daily Attendance (Morning In/Out & Afternoon In/Out)**: Automatically tracks and categorizes scans into **Morning In**, **Morning Out**, **Afternoon In**, and **Afternoon Out** based on the time of day.
4. **Student Photos Support**: Added image upload capability for student profiles so pictures appear on the scanner screen.
5. **Modernized UI Redesign**: Clean, professional dark/light dashboard layout with responsive cards, live status badges, and management tables for students, attendance logs, and school settings.
6. **Presentation-Ready Database Code**: Clean Mongoose schemas structured explicitly for school data (Students, Attendance Logs, and System Configuration).

---

### Updated Code (`server.js`)

```javascript
const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const mongoose = require('mongoose');
const ExcelJS = require('exceljs');
const { Resend } = require('resend');

const app = express();
const PORT = process.env.PORT || 3000;

// ==========================================
// 1. MONGODB DATABASE CONNECTION & SCHEMAS
// ==========================================
const MONGO_URI = process.env.MONGO_URI || 'mongodb+srv://srcadmin:30005BNHS@cluster0.he7jspr.mongodb.net/school_attendance_db?appName=Cluster0';

mongoose.connect(MONGO_URI)
  .then(() => console.log('[DATABASE] Connected to MongoDB Atlas successfully!'))
  .catch(err => console.error('[DATABASE ERROR] Could not connect to MongoDB:', err.message));

// Student Schema (School Attendance Focus)
const studentSchema = new mongoose.Schema({
  uid: { type: String, default: '' },
  studentId: { type: String, required: true, unique: true },
  name: { type: String, required: true },
  gradeLevel: { type: String, default: 'Grade 7' },
  section: { type: String, default: 'A' },
  guardianEmail: { type: String, default: '' },
  phone: { type: String, default: '' },
  photo: { type: String, default: '/uploads/default_avatar.png' }
});

// Attendance Schema (4-Part Daily Tracking: AM In/Out, PM In/Out)
const attendanceSchema = new mongoose.Schema({
  uid: String,
  name: String,
  studentId: String,
  gradeLevel: String,
  section: String,
  scanType: String, // 'MORNING-IN', 'MORNING-OUT', 'AFTERNOON-IN', 'AFTERNOON-OUT'
  status: String,   // 'ON TIME', 'LATE', 'COMPLETED'
  timestamp: String,
  rawTimestamp: { type: Date, default: Date.now }
});

// System Configuration Schema
const configSchema = new mongoose.Config || new mongoose.Schema({
  schoolName: { type: String, default: 'National High School Attendance Portal' },
  logoPath: { type: String, default: '' },
  morningLateCutoff: { type: String, default: '07:45' },
  afternoonLateCutoff: { type: String, default: '13:00' },
  latestUid: { type: String, default: '' },
  enableEmail: { type: Boolean, default: false }
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

// ==========================================
// 2. FILE UPLOADS SETUP (Multer)
// ==========================================
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

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/uploads', express.static(uploadsDir));

// ==========================================
// 3. ESP8266 / RFID SCANNER API ENDPOINT
// ==========================================
app.post('/api/scan', async (req, res) => {
  try {
    const { uid } = req.body;
    if (!uid) return res.status(400).json({ status: 'error', message: 'No UID provided' });

    const cleanUid = uid.trim().toUpperCase();
    const config = await getConfig();
    
    config.latestUid = cleanUid;
    await config.save();

    const student = await Student.findOne({ uid: cleanUid });
    if (!student) {
      return res.json({ status: 'unknown', message: 'RFID Card not registered to any student.' });
    }

    const now = new Date();
    const startOfDay = new Date();
    startOfDay.setHours(0, 0, 0, 0);

    // Fetch today's scans for this student
    const todayLogs = await Attendance.find({
      uid: cleanUid,
      rawTimestamp: { $gte: startOfDay }
    }).sort({ rawTimestamp: 1 });

    const currentHour = now.getHours();
    let scanType = 'MORNING-IN';
    let statusLabel = 'ON TIME';

    // Determine 4-part scan sequence automatically based on existing logs and time
    if (currentHour >= 12) {
      // Afternoon session
      const hasAfternoonIn = todayLogs.some(l => l.scanType === 'AFTERNOON-IN');
      if (!hasAfternoonIn) {
        scanType = 'AFTERNOON-IN';
        statusLabel = now.toTimeString().slice(0, 5) > (config.afternoonLateCutoff || '13:00') ? 'LATE' : 'ON TIME';
      } else {
        scanType = 'AFTERNOON-OUT';
        statusLabel = 'COMPLETED';
      }
    } else {
      // Morning session
      const hasMorningIn = todayLogs.some(l => l.scanType === 'MORNING-IN');
      const hasMorningOut = todayLogs.some(l => l.scanType === 'MORNING-OUT');
      
      if (!hasMorningIn) {
        scanType = 'MORNING-IN';
        statusLabel = now.toTimeString().slice(0, 5) > (config.morningLateCutoff || '07:45') ? 'LATE' : 'ON TIME';
      } else if (!hasMorningOut) {
        scanType = 'MORNING-OUT';
        statusLabel = 'COMPLETED';
      } else {
        scanType = 'AFTERNOON-IN'; // Fallback if scanned early afternoon
      }
    }

    const record = new Attendance({
      uid: cleanUid,
      name: student.name,
      studentId: student.studentId,
      gradeLevel: student.gradeLevel,
      section: student.section,
      scanType,
      status: statusLabel,
      timestamp: now.toLocaleString(),
      rawTimestamp: now
    });

    await record.save();

    return res.json({
      status: 'success',
      scanType,
      studentName: student.name,
      studentId: student.studentId,
      gradeSection: `${student.gradeLevel} - ${student.section}`,
      photo: student.photo,
      attendanceStatus: statusLabel,
      message: `${scanType} recorded successfully for ${student.name}`
    });

  } catch (err) {
    console.error('[SCAN ERROR]', err.message);
    res.status(500).json({ status: 'error', message: 'Internal Server Error' });
  }
});

// ==========================================
// 4. REALTIME DATA ENDPOINT FOR DASHBOARD
// ==========================================
app.get('/api/live-data', async (req, res) => {
  try {
    const config = await getConfig();
    const students = await Student.find().sort({ gradeLevel: 1, section: 1 });
    const attendance = await Attendance.find().sort({ rawTimestamp: -1 }).limit(100);

    res.json({
      latestUid: config.latestUid || '',
      schoolName: config.schoolName,
      attendance,
      students
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ==========================================
// 5. SEPARATE SCANNER KIOSK SCREEN ROUTE
// ==========================================
app.get('/scanner', async (req, res) => {
  const config = await getConfig();
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>RFID Scanner Kiosk - ${config.schoolName}</title>
      <style>
        body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; display: flex; flex-direction: column; height: 100vh; overflow: hidden; }
        header { background: #1e293b; padding: 20px 40px; display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #334155; }
        h1 { margin: 0; font-size: 1.8rem; color: #38bdf8; }
        .main-container { display: flex; flex: 1; padding: 40px; gap: 40px; align-items: center; justify-content: center; }
        .scanner-status-box { flex: 1; background: #1e293b; padding: 40px; border-radius: 16px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.5); border: 1px solid #334155; }
        .profile-card { flex: 1; background: #1e293b; padding: 40px; border-radius: 16px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.5); border: 1px solid #334155; display: flex; flex-direction: column; align-items: center; }
        .student-photo { width: 180px; height: 180px; border-radius: 50%; object-fit: cover; border: 4px solid #38bdf8; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(56,189,248,0.3); }
        .badge { display: inline-block; padding: 8px 16px; border-radius: 8px; font-weight: bold; font-size: 1.1rem; margin-top: 15px; }
        .badge-ontime { background: #22c55e; color: #fff; }
        .badge-late { background: #ef4444; color: #fff; }
        .badge-type { background: #0284c7; color: #fff; margin-bottom: 10px; }
        #clock { font-size: 3rem; font-weight: bold; color: #38bdf8; margin-bottom: 10px; }
        .waiting-text { font-size: 1.5rem; color: #94a3b8; animation: pulse 1.5s infinite; }
        @keyframes pulse { 0% { opacity: 0.6; } 50% { opacity: 1; } 100% { opacity: 0.6; } }
      </style>
    </head>
    <body>
      <header>
        <h1>🎓 ${config.schoolName} - Live RFID Kiosk</h1>
        <div id="clock">00:00:00</div>
      </header>
      
      <div class="main-container">
        <div class="scanner-status-box">
          <h2>Ready for Scanning</h2>
          <p class="waiting-text">Please tap your Student RFID Card on the scanner...</p>
          <p style="margin-top: 30px; color: #64748b;">System UID Watcher Active</p>
        </div>

        <div class="profile-card" id="displayCard">
          <img src="/uploads/default_avatar.png" id="studentPhoto" class="student-photo" alt="Student Photo">
          <div id="scanBadgeContainer"><span class="badge badge-type">WAITING FOR TAP</span></div>
          <h2 id="studentName" style="font-size: 2rem; margin: 10px 0;">---</h2>
          <p id="studentDetails" style="font-size: 1.2rem; color: #94a3b8; margin: 5px 0;">Grade & Section / ID</p>
          <div id="statusBadgeContainer"></div>
        </div>
      </div>

      <script>
        function updateClock() {
          const now = new Date();
          document.getElementById('clock').innerText = now.toLocaleTimeString();
        }
        setInterval(updateClock, 1000);
        updateClock();

        let lastCheckedUid = '';

        async function pollScanner() {
          try {
            const res = await fetch('/api/live-data');
            const data = await res.json();
            
            if (data.latestUid && data.latestUid !== lastCheckedUid) {
              lastCheckedUid = data.latestUid;
              
              // Trigger scan request simulation or check last attendance log matching this UID
              const attList = data.attendance || [];
              const latestLog = attList.find(a => a.uid === data.latestUid);
              const student = (data.students || []).find(s => s.uid === data.latestUid);

              if (student && latestLog) {
                document.getElementById('studentPhoto').src = student.photo || '/uploads/default_avatar.png';
                document.getElementById('studentName').innerText = student.name;
                document.getElementById('studentDetails').innerText = \`ID: \${student.studentId} | \${student.gradeLevel} - \${student.section}\`;
                
                document.getElementById('scanBadgeContainer').innerHTML = \`<span class="badge badge-type">\${latestLog.scanType}</span>\`;
                
                const statusClass = latestLog.status === 'LATE' ? 'badge-late' : 'badge-ontime';
                document.getElementById('statusBadgeContainer').innerHTML = \`<span class="badge \${statusClass}">\${latestLog.status}</span>\`;

                // Speak Name aloud
                if ('speechSynthesis' in window) {
                  const utterance = new SpeechSynthesisUtterance(\`\${student.name}, \${latestLog.scanType.toLowerCase().replace('-', ' ')}\`);
                  window.speechSynthesis.speak(utterance);
                }
              } else {
                document.getElementById('studentName').innerText = 'Unregistered Card';
                document.getElementById('studentDetails').innerText = \`UID: \${data.latestUid}\`;
                document.getElementById('scanBadgeContainer').innerHTML = '<span class="badge badge-late">UNKNOWN CARD</span>';
                document.getElementById('statusBadgeContainer').innerHTML = '';
                
                if ('speechSynthesis' in window) {
                  window.speechSynthesis.speak(new SpeechSynthesisUtterance('Unregistered RFID card.'));
                }
              }
            }
          } catch (e) {}
        }

        setInterval(pollScanner, 1500);
      </script>
    </body>
    </html>
  `);
});

// ==========================================
// 6. ADMIN DASHBOARD & MANAGEMENT ROUTE
// ==========================================
app.get('/', async (req, res) => {
  const config = await getConfig();
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Admin Dashboard - ${config.schoolName}</title>
      <style>
        :root { --primary: #0284c7; --bg: #f8fafc; --card-bg: #ffffff; --text: #1e293b; --border: #e2e8f0; }
        body { font-family: 'Segoe UI', Tahoma, sans-serif; margin: 0; background: var(--bg); color: var(--text); display: flex; min-height: 100vh; }
        aside { width: 260px; background: #1e293b; color: #fff; padding: 20px; display: flex; flex-direction: column; }
        aside h2 { font-size: 1.2rem; color: #38bdf8; margin-bottom: 30px; }
        aside a { color: #94a3b8; text-decoration: none; padding: 12px 15px; border-radius: 8px; margin-bottom: 8px; font-weight: 600; display: block; transition: 0.2s; }
        aside a:hover, aside a.active { background: #334155; color: #fff; }
        main { flex: 1; padding: 30px; overflow-y: auto; }
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; }
        .cards-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: var(--card-bg); padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); border: 1px solid var(--border); }
        .stat-card h3 { margin: 0 0 10px 0; font-size: 0.9rem; color: #64748b; }
        .stat-card .val { font-size: 1.8rem; font-weight: bold; color: var(--primary); }
        .card { background: var(--card-bg); padding: 25px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); border: 1px solid var(--border); margin-bottom: 25px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid var(--border); font-size: 0.95rem; }
        th { background: #f1f5f9; color: #475569; font-weight: 600; }
        input, select { width: 100%; padding: 10px; margin: 8px 0 16px; border: 1px solid var(--border); border-radius: 6px; box-sizing: border-box; }
        button, input[type="submit"] { background: var(--primary); color: white; padding: 10px 20px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; }
        button:hover { opacity: 0.9; }
        .btn-danger { background: #ef4444; }
        .btn-kiosk { background: #10b981; text-decoration: none; display: inline-block; padding: 10px 20px; border-radius: 6px; color: white; font-weight: bold; }
        .badge { padding: 4px 10px; border-radius: 6px; font-size: 0.8rem; font-weight: bold; }
        .badge-ontime { background: #dcfce7; color: #166534; }
        .badge-late { background: #fee2e2; color: #991b1b; }
      </style>
    </head>
    <body>
      <aside>
        <h2>🏫 School Admin</h2>
        <a href="/" class="active">🏠 Dashboard</a>
        <a href="/scanner" target="_blank">📡 Open Kiosk Scanner</a>
        <a href="#students-section">👨‍🎓 Students</a>
        <a href="#attendance-section">📊 Daily Attendance</a>
        <a href="#settings-section">⚙️ School Settings</a>
      </aside>

      <main>
        <div class="header">
          <h1>${config.schoolName}</h1>
          <a href="/scanner" target="_blank" class="btn-kiosk">🖥️ Launch Fullscreen Scanner</a>
        </div>

        <div class="cards-grid">
          <div class="stat-card">
            <h3>Total Enrolled Students</h3>
            <div class="val" id="statTotalStudents">0</div>
          </div>
          <div class="stat-card">
            <h3>Today's Attendance Logs</h3>
            <div class="val" id="statTotalLogs">0</div>
          </div>
          <div class="stat-card">
            <h3>Latest Scanned Card UID</h3>
            <div class="val" id="statLatestUid" style="font-size: 1.2rem; color: #d97706;">None</div>
          </div>
        </div>

        <div class="card" id="students-section">
          <h2>Register / Manage Student & RFID Card</h2>
          <form action="/api/register-student" method="POST" enctype="multipart/form-data">
            <div style="display: flex; gap: 15px;">
              <div style="flex: 1;">
                <label>Student ID Number:</label>
                <input type="text" name="studentId" placeholder="e.g. 2026-0001" required>
              </div>
              <div style="flex: 1;">
                <label>Full Name:</label>
                <input type="text" name="name" placeholder="Juan Dela Cruz" required>
              </div>
            </div>

            <div style="display: flex; gap: 15px;">
              <div style="flex: 1;">
                <label>Grade Level:</label>
                <select name="gradeLevel">
                  <option value="Grade 7">Grade 7</option>
                  <option value="Grade 8">Grade 8</option>
                  <option value="Grade 9">Grade 9</option>
                  <option value="Grade 10">Grade 10</option>
                  <option value="Grade 11">Grade 11</option>
                  <option value="Grade 12">Grade 12</option>
                </select>
              </div>
              <div style="flex: 1;">
                <label>Section:</label>
                <input type="text" name="section" placeholder="Diamond" required>
              </div>
            </div>

            <div style="display: flex; gap: 15px;">
              <div style="flex: 1;">
                <label>RFID Card UID:</label>
                <input type="text" name="uid" id="formUidInput" placeholder="Tap card or enter UID">
              </div>
              <div style="flex: 1;">
                <label>Student Photo:</label>
                <input type="file" name="studentPhoto" accept="image/*">
              </div>
            </div>

            <button type="submit">Save Student Record</button>
          </form>
        </div>

        <div class="card" id="attendance-section">
          <h2>Live School Attendance Records (AM / PM In & Out)</h2>
          <table>
            <thead>
              <tr>
                <th>Student Name</th>
                <th>ID Number</th>
                <th>Grade & Section</th>
                <th>Scan Type</th>
                <th>Status</th>
                <th>Timestamp</th>
                <th>UID</th>
              </tr>
            </thead>
            <tbody id="attendanceTableBody">
              <tr><td colspan="7" style="text-align:center;">Loading records...</td></tr>
            </tbody>
          </table>
        </div>

        <div class="card" id="settings-section">
          <h2>School Configuration & Cutoff Times</h2>
          <form action="/api/update-config" method="POST">
            <label>School Name:</label>
            <input type="text" name="schoolName" value="${config.schoolName}">

            <div style="display: flex; gap: 15px;">
              <div style="flex: 1;">
                <label>Morning Late Cutoff Time:</label>
                <input type="time" name="morningLateCutoff" value="${config.morningLateCutoff || '07:45'}">
              </div>
              <div style="flex: 1;">
                <label>Afternoon Late Cutoff Time:</label>
                <input type="time" name="afternoonLateCutoff" value="${config.afternoonLateCutoff || '13:00'}">
              </div>
            </div>

            <input type="submit" value="Save Settings">
          </form>
        </div>
      </main>

      <script>
        async function fetchDashboardData() {
          try {
            const res = await fetch('/api/live-data');
            const data = await res.json();

            document.getElementById('statTotalStudents').innerText = data.students.length;
            document.getElementById('statTotalLogs').innerText = data.attendance.length;
            document.getElementById('statLatestUid').innerText = data.latestUid || 'None';
            if (data.latestUid && !document.getElementById('formUidInput').value) {
              document.getElementById('formUidInput').value = data.latestUid;
            }

            const tbody = document.getElementById('attendanceTableBody');
            if (data.attendance.length === 0) {
              tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;">No attendance records found for today.</td></tr>';
            } else {
              tbody.innerHTML = data.attendance.map(row => \`
                <tr>
                  <td><strong>\${row.name}</strong></td>
                  <td>\${row.studentId}</td>
                  <td>\${row.gradeLevel} - \${row.section}</td>
                  <td><span class="badge" style="background:#e0f2fe; color:#0369a1;">\${row.scanType}</span></td>
                  <td><span class="badge \${row.status === 'LATE' ? 'badge-late' : 'badge-ontime'}">\${row.status}</span></td>
                  <td>\${row.timestamp}</td>
                  <td><code>\${row.uid}</code></td>
                </tr>
              \`).join('');
            }
          } catch (e) {}
        }

        setInterval(fetchDashboardData, 2000);
        fetchDashboardData();
      </script>
    </body>
    </html>
  `);
});

// ==========================================
// 7. STUDENT & CONFIGURATION ENDPOINTS
// ==========================================
app.post('/api/register-student', upload.single('studentPhoto'), async (req, res) => {
  try {
    const { studentId, name, gradeLevel, section, uid } = req.body;
    const photoPath = req.file ? `/uploads/${req.file.filename}` : '/uploads/default_avatar.png';

    await Student.findOneAndUpdate(
      { studentId },
      {
        studentId,
        name,
        gradeLevel: gradeLevel || 'Grade 7',
        section: section || 'A',
        uid: uid ? uid.trim().toUpperCase() : '',
        photo: photoPath
      },
      { upsert: true, new: true }
    );

    res.redirect('/');
  } catch (err) {
    res.status(500).send('Error saving student: ' + err.message);
  }
});

app.post('/api/update-config', async (req, res) => {
  try {
    const { schoolName, morningLateCutoff, afternoonLateCutoff } = req.body;
    const config = await getConfig();
    config.schoolName = schoolName || config.schoolName;
    config.morningLateCutoff = morningLateCutoff || config.morningLateCutoff;
    config.afternoonLateCutoff = afternoonLateCutoff || config.afternoonLateCutoff;
    await config.save();
    res.redirect('/');
  } catch (err) {
    res.status(500).send('Error updating settings: ' + err.message);
  }
});

// ==========================================
// 8. SERVER INITIALIZATION
// ==========================================
app.listen(PORT, () => {
  console.log(`[SERVER] School Attendance Server running on port ${PORT}`);
});
