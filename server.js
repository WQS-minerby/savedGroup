const express = require('express');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const cookieParser = require('cookie-parser');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');

const app = express();
const JWT_SECRET = process.env.JWT_SECRET || 'change-this-secret-before-hosting';
const DATA_DIR = process.env.DATA_DIR || __dirname;
const IS_PRODUCTION = process.env.NODE_ENV === 'production';

if (IS_PRODUCTION && JWT_SECRET === 'change-this-secret-before-hosting') {
  console.warn('JWT_SECRET is not set. Add a strong JWT_SECRET environment variable before real use.');
}

app.use(express.json());
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use((req, res, next) => {
  console.log(`[REQ] ${req.method} ${req.originalUrl}`);
  const origin = req.headers.origin;
  if (origin) {
    res.header('Access-Control-Allow-Origin', origin);
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    res.header('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,PATCH,OPTIONS');
    res.header('Access-Control-Allow-Credentials', 'true');
  }
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

// ─── DATA ──────────────────────────────────────────────────────────────────

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const DATA_FILE = path.join(DATA_DIR, 'saved-data.xlsx');
const JSON_FALLBACK = path.join(DATA_DIR, 'saved-data.json');

let members = [
  { id: '1200280285142014', name: 'NIRINGIYUMUKIZA Eric', phone: '0794707599', shares: 4, email: 'ericniring@gmail.com', role: 'admin' },
  { id: '1200370160383008', name: 'Mukarukundo Philomene', phone: '0795541892', shares: 1, email: 'philomenemukarukundo0@gmail.com', role: 'member' },
  { id: '1200480009530046', name: 'NIYOBUHUNGIRO AMOS', phone: '0790069381', shares: 2, email: 'niyobuhungiroamos034@gmail.com', role: 'member' },
  { id: '1200270047236001', name: 'Twizerimana Francoise', phone: '0785248422', shares: 1, email: 'ftwizerimana41@gmail.com', role: 'member' },
  { id: '1200380198988150', name: 'Byiringiro Bonfils Kevin', phone: '0785377755', shares: 1, email: 'byiringirokevin242@gmail.com', role: 'member' },
  { id: '1200670122691090', name: 'Uwayisaba Agnes', phone: '0796643370', shares: 2, email: 'uwayisabaagnes3@gmail.com', role: 'member' },
  { id: '1200370251061082', name: 'Ishimwe Umulisa Fifi', phone: '0798959121', shares: 3, email: 'ishimweumulisafifi@gmail.com', role: 'member' },
  { id: '1200480176198073', name: 'SHEMA Shalifu', phone: '0728702245', shares: 2, email: 'shemashalifu1@gmail.com', role: 'member' },
  { id: '1200370201209159', name: 'IRADUKUNDA Theopiste', phone: '0794363752', shares: 1, email: 'theopisteiradukunda7@gmail.com', role: 'member' },
  { id: '1200570104099055', name: 'Umwizerwa Christine', phone: '0798226693', shares: 3, email: 'umwizerwatina@gmail.com', role: 'member' },
  { id: '1200670284352162', name: 'Kwizera Divine', phone: '0795625125', shares: 4, email: 'kwizeraleadivine@gmail.com', role: 'member' },
  { id: '1200480139054090', name: 'NGIRUMUKIZA Eric', phone: '0791871436', shares: 2, email: 'ericngirumukiza1@gmail.com', role: 'member' },
  { id: '1200580057166046', name: 'NIYONSHUTI THEONESTE', phone: '0790187782', shares: 2, email: 'theoniyonshuti2022@gmail.com', role: 'member' },
  { id: '1200480204341077', name: 'Mbayeho Happy Selligue', phone: '0790030701', shares: 2, email: 'mbayehohappy@gmail.com', role: 'member' },
  { id: '1200370172602046', name: 'TUYIRAMYE Ancile', phone: '0793032695', shares: 1, email: 'anciletuyiramye@gmail.com', role: 'member' },
  { id: '1200670094358005', name: 'Ingabire Solange', phone: '0787044562', shares: 1, email: 'solangeingabire370@gmail.com', role: 'member' },
  { id: '1200670037120091', name: 'NIWEMAHORO Providence', phone: '0794559573', shares: 2, email: 'providenceniwemahoro10@gmail.com', role: 'member' },
  { id: '1200670169217046', name: 'Muhozi Julia', phone: '0732655344', shares: 4, email: 'juliamuhoozi057@gmail.com', role: 'member' },
];

let weeklyData = {
  '1200280285142014': [2000,2000,2000,2000,2000,2000],
  '1200370160383008': [500,500,500,500,500,500],
  '1200480009530046': [1000,1000,1000,1000,1000,1000],
  '1200270047236001': [500,500,500,500,0,500],
  '1200380198988150': [500,500,500,500,500,500],
  '1200670122691090': [1000,1000,500,500,500,500],
  '1200370251061082': [2000,1500,1500,1500,1500,1500],
  '1200480176198073': [1000,1000,1000,1000,0,1000],
  '1200370201209159': [500,500,500,500,500,500],
  '1200570104099055': [1000,500,1500,1500,1500,1500],
  '1200670284352162': [2000,2000,2000,2000,2000,2000],
  '1200480139054090': [1000,1000,1000,1000,1000,1000],
  '1200580057166046': [1000,1000,1000,1000,1000,1000],
  '1200480204341077': [500,500,500,500,500,500],
  '1200370172602046': [500,500,500,500,500,500],
  '1200670094358005': [0,0,0,500,500,500],
  '1200670037120091': [0,0,0,1000,1000,1000],
  '1200670169217046': [0,0,0,2000,2000,2000],
};

let loans = [
  { id: 'L001', memberId: '1200580057166046', name: 'NIYONSHUTI THEONESTE', amount: 10000, interest: 1000, repaid: 0, date: '2024-01-15', due: '2024-03-15' },
  { id: 'L002', memberId: '1200480009530046', name: 'NIYOBUHUNGIRO AMOS', amount: 8000, interest: 800, repaid: 0, date: '2024-01-15', due: '2024-03-15' },
  { id: 'L003', memberId: '1200480176198073', name: 'SHEMA Shalifu', amount: 10000, interest: 1000, repaid: 0, date: '2024-01-19', due: '2024-03-19' },
  { id: 'L004', memberId: '1200480139054090', name: 'NGIRUMUKIZA Eric', amount: 10000, interest: 1000, repaid: 0, date: '2024-01-18', due: '2024-03-18' },
  { id: 'L005', memberId: '1200670122691090', name: 'UWAYISABA Agnes', amount: 5000, interest: 500, repaid: 0, date: '2024-01-20', due: '2024-03-20' },
];

function parseNumber(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function saveState() {
  const maxWeeks = Math.max(6, ...members.map(m => (weeklyData[m.id] || []).length));
  const memberSheet = [
    ['Member ID', 'Name', 'Phone', 'Email', 'Role', 'Shares', 'Total Saved', 'Password Hash'],
    ...members.map(member => [
      member.id,
      member.name,
      member.phone,
      member.email,
      member.role,
      memberShares(member),
      memberTotalSaved(member),
      member.passwordHash || '',
    ]),
  ];

  const savingHeaders = ['Member ID', 'Name', ...Array.from({ length: maxWeeks }, (_, i) => `Week ${i + 1}`), 'Total Saved'];
  const savingSheet = [
    savingHeaders,
    ...members.map(member => {
      const weeks = Array.from({ length: maxWeeks }, (_, idx) => weeklyData[member.id]?.[idx] || 0);
      return [member.id, member.name, ...weeks, memberTotalSaved(member)];
    }),
  ];

  const loanSheet = [
    ['Loan ID', 'Member ID', 'Name', 'Amount', 'Interest', 'Repaid', 'Date', 'Due Date'],
    ...loans.map(loan => [
      loan.id,
      loan.memberId,
      loan.name,
      loan.amount,
      loan.interest,
      loan.repaid,
      loan.date,
      loan.due,
    ]),
  ];

  const workbook = XLSX.utils.book_new();
  workbook.Props = {
    Title: 'SAVED Group Data',
    Subject: 'SAVED System Storage',
    Author: 'SAVED System',
    CreatedDate: new Date(),
  };
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(memberSheet), 'Members');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(savingSheet), 'Savings');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(loanSheet), 'Loans');
  XLSX.writeFile(workbook, DATA_FILE, { bookType: 'xlsx' });
}

function loadState() {
  if (!fs.existsSync(DATA_FILE)) {
    if (fs.existsSync(JSON_FALLBACK)) {
      try {
        const raw = fs.readFileSync(JSON_FALLBACK, 'utf8');
        const parsed = JSON.parse(raw);
        if (parsed && parsed.members && parsed.weeklyData && parsed.loans) {
          members = parsed.members;
          weeklyData = parsed.weeklyData;
          loans = parsed.loans;
          saveState();
          return;
        }
      } catch (err) {
        console.error('Unable to load fallback JSON.', err);
      }
    }
    saveState();
    return;
  }

  try {
    const workbook = XLSX.readFile(DATA_FILE, { cellDates: true, raw: false });
    const membersSheet = workbook.Sheets['Members'];
    const savingsSheet = workbook.Sheets['Savings'];
    const loansSheet = workbook.Sheets['Loans'];

    if (membersSheet) {
      const rows = XLSX.utils.sheet_to_json(membersSheet, { defval: '' });
      const parsedMembers = rows
        .filter(row => String(row['Member ID'] || '').trim())
        .map(row => ({
          id: String(row['Member ID'] || '').trim(),
          name: String(row['Name'] || '').trim(),
          phone: String(row['Phone'] || '').trim(),
          email: String((row['Email'] || '').toString().trim()).toLowerCase(),
          role: String(row['Role'] || 'member').trim() || 'member',
          shares: parseNumber(row['Shares']),
          passwordHash: String(row['Password Hash'] || '').trim(),
        }));
      if (parsedMembers.length) members = parsedMembers;
    }

    if (savingsSheet) {
      const rows = XLSX.utils.sheet_to_json(savingsSheet, { defval: 0 });
      weeklyData = {};
      rows.forEach(row => {
        const id = String(row['Member ID'] || '').trim();
        if (!id) return;
        const weekKeys = Object.keys(row)
          .filter(key => /^Week\s*\d+/i.test(key))
          .map(key => ({
            key,
            number: Number(key.match(/\d+/)?.[0] || 0),
          }))
          .filter(({ number }) => number >= 1)
          .sort((a, b) => a.number - b.number);
        const weeks = Array.from({ length: Math.max(6, ...weekKeys.map(({ number }) => number)) }, () => 0);
        weekKeys.forEach(({ key, number }) => {
          weeks[number - 1] = parseNumber(row[key]);
        });
        weeklyData[id] = weeks;
      });
    }

    if (loansSheet) {
      const rows = XLSX.utils.sheet_to_json(loansSheet, { defval: '' });
      loans = rows
        .filter(row => String(row['Loan ID'] || '').trim())
        .map(row => ({
          id: String(row['Loan ID'] || '').trim(),
          memberId: String(row['Member ID'] || '').trim(),
          name: String(row['Name'] || '').trim(),
          amount: parseNumber(row['Amount']),
          interest: parseNumber(row['Interest']),
          repaid: parseNumber(row['Repaid']),
          date: String(row['Date'] || '').trim(),
          due: String(row['Due Date'] || '').trim(),
        }));
    }
  } catch (err) {
    console.error('Unable to load saved data from Excel, using defaults.', err);
    saveState();
  }
}

loadState();

// Build users map with hashed passwords (default password = last 4 digits of phone)
const usersDB = {};
let generatedMissingPasswordHashes = false;
members.forEach(m => {
  const defaultPass = m.phone.slice(-4);
  if (!m.passwordHash) {
    m.passwordHash = bcrypt.hashSync(defaultPass, 8);
    generatedMissingPasswordHashes = true;
  }
  usersDB[m.email.toLowerCase()] = {
    ...m,
    passwordHash: m.passwordHash
  };
});
if (generatedMissingPasswordHashes) saveState();

function toAmount(value) {
  const n = Number(value);
  return Number.isFinite(n) && n >= 0 ? Math.round(n) : null;
}

function memberTotalSaved(member) {
  return (weeklyData[member.id] || []).reduce((sum, amount) => sum + amount, 0);
}

function memberShares(member) {
  return Math.floor(memberTotalSaved(member) / 500);
}

function groupSavingsTotal() {
  return members.reduce((sum, m) => sum + (weeklyData[m.id] || []).reduce((a, b) => a + b, 0), 0);
}

function weeklyGrowth() {
  const maxWeeks = Math.max(6, ...Object.values(weeklyData).map(weeks => weeks.length));
  let running = 0;
  return Array.from({ length: maxWeeks }, (_, i) => {
    running += members.reduce((sum, m) => sum + ((weeklyData[m.id] || [])[i] || 0), 0);
    return running;
  });
}

function escapeCsv(value) {
  const text = String(value ?? '');
  return `"${text.replace(/"/g, '""')}"`;
}

function buildCsvRow(fields) {
  return fields.map(escapeCsv).join(',');
}

function addUserLogin(member) {
  const defaultPass = member.phone.slice(-4);
  if (!member.passwordHash) member.passwordHash = bcrypt.hashSync(defaultPass, 8);
  usersDB[member.email.toLowerCase()] = {
    ...member,
    passwordHash: member.passwordHash
  };
}

function nextLoanId() {
  const max = loans.reduce((n, loan) => {
    const match = String(loan.id).match(/\d+/);
    return Math.max(n, match ? Number(match[0]) : 0);
  }, 0);
  return `L${String(max + 1).padStart(3, '0')}`;
}

// ─── AUTH MIDDLEWARE ────────────────────────────────────────────────────────

function authMiddleware(req, res, next) {
  const token = req.cookies.token || (req.headers.authorization || '').replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'Unauthorized' });
  try {
    req.user = jwt.verify(token, JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ error: 'Invalid token' });
  }
}

function adminOnly(req, res, next) {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Admin only' });
  next();
}

// ─── AUTH ROUTES ────────────────────────────────────────────────────────────

app.post('/api/login', (req, res) => {
  const { email, password } = req.body || {};
  if (typeof email !== 'string' || typeof password !== 'string') {
    return res.status(400).json({ error: 'Email and password are required' });
  }
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const user = usersDB[normalizedEmail];
  if (!user || !bcrypt.compareSync(password, user.passwordHash)) {
    return res.status(401).json({ error: 'Invalid email or password' });
  }
  const token = jwt.sign({ id: user.id, name: user.name, email: user.email, role: user.role }, JWT_SECRET, { expiresIn: '7d' });
  res.cookie('token', token, {
    httpOnly: true,
    sameSite: 'lax',
    secure: IS_PRODUCTION,
    maxAge: 7 * 24 * 60 * 60 * 1000,
  });
  res.json({ token, role: user.role, name: user.name, id: user.id });
});

app.post('/api/logout', (req, res) => {
  res.clearCookie('token');
  res.json({ ok: true });
});

app.get('/api/me', authMiddleware, (req, res) => {
  res.json(req.user);
});

app.post('/api/change-password', authMiddleware, (req, res) => {
  const { currentPassword, newPassword } = req.body || {};
  if (typeof currentPassword !== 'string' || typeof newPassword !== 'string') {
    return res.status(400).json({ error: 'Current password and new password are required' });
  }
  if (newPassword.length < 6) {
    return res.status(400).json({ error: 'New password must be at least 6 characters' });
  }

  const member = members.find(m => m.id === req.user.id);
  const user = member ? usersDB[member.email.toLowerCase()] : null;
  if (!member || !user) return res.status(404).json({ error: 'User not found' });
  if (!bcrypt.compareSync(currentPassword, user.passwordHash)) {
    return res.status(401).json({ error: 'Current password is incorrect' });
  }

  member.passwordHash = bcrypt.hashSync(newPassword, 8);
  usersDB[member.email.toLowerCase()] = {
    ...member,
    passwordHash: member.passwordHash,
  };
  saveState();
  res.json({ ok: true });
});

// ─── MEMBER ROUTES ──────────────────────────────────────────────────────────

// Own dashboard data
app.get('/api/my-dashboard', authMiddleware, (req, res) => {
  const m = members.find(x => x.id === req.user.id);
  if (!m) return res.status(404).json({ error: 'Member not found' });
  const weeks = weeklyData[m.id] || [];
  const totalSaved = weeks.reduce((a, b) => a + b, 0);
  const myLoan = loans.find(l => l.memberId === m.id) || null;
  const dynamicShares = memberShares(m);
  const totalShares = members.reduce((sum, member) => sum + memberShares(member), 0);
  res.json({
    member: { ...m, shares: dynamicShares },
    weeks,
    totalSaved,
    loan: myLoan,
    totalShares,
    groupTotal: groupSavingsTotal(),
    groupLoans: loans.reduce((a, l) => a + l.amount, 0),
  });
});

app.put('/api/my-savings', authMiddleware, (req, res) => {
  const member = members.find(x => x.id === req.user.id);
  if (!member) return res.status(404).json({ error: 'Member not found' });
  const week = Number(req.body.week);
  const amount = toAmount(req.body.amount);
  if (!Number.isInteger(week) || week < 1) return res.status(400).json({ error: 'Week must be 1 or higher' });
  if (amount === null) return res.status(400).json({ error: 'Amount must be a positive number or zero' });
  if (!weeklyData[member.id]) weeklyData[member.id] = [];
  while (weeklyData[member.id].length < week) weeklyData[member.id].push(0);
  weeklyData[member.id][week - 1] = amount;
  saveState();
  res.json({ id: member.id, weeks: weeklyData[member.id], total: weeklyData[member.id].reduce((a, b) => a + b, 0) });
});

// ─── ADMIN ROUTES ───────────────────────────────────────────────────────────

app.get('/api/admin/dashboard', authMiddleware, adminOnly, (req, res) => {
  const totSaved = groupSavingsTotal();
  const totLoans = loans.reduce((a, l) => a + l.amount, 0);
  const totInterest = loans.reduce((a, l) => a + l.interest, 0);
  const totRepaid = loans.reduce((a, l) => a + l.repaid, 0);
  const allShares = members.map(m => memberShares(m));
  const shareDistribution = [1, 2, 3, 4].map(count => allShares.filter(sh => sh === count).length);
  res.json({
    totalMembers: members.length,
    totalShares: allShares.reduce((a, b) => a + b, 0),
    totalSaved: totSaved,
    totalLoans: totLoans,
    totalInterest: totInterest,
    totalRepaid: totRepaid,
    outstanding: totLoans + totInterest - totRepaid,
    weeklyGrowth: weeklyGrowth(),
    shareDistribution,
  });
});

app.get('/api/admin/export-records', authMiddleware, adminOnly, (req, res) => {
  const maxWeeks = Math.max(6, ...members.map(m => (weeklyData[m.id] || []).length));
  const header = [
    'Member ID',
    'Name',
    'Phone',
    'Email',
    'Role',
    'Shares',
    'Total Saved',
    ...Array.from({ length: maxWeeks }, (_, i) => `Week ${i + 1}`),
    'Loan Count',
    'Loan Summary (id:amount:interest:repaid:date:due)'
  ];

  const rows = members.map(member => {
    const weeks = Array.from({ length: maxWeeks }, (_, idx) => weeklyData[member.id]?.[idx] || 0);
    const memberLoans = loans.filter(l => l.memberId === member.id);
    const loanSummary = memberLoans.map(l => `${l.id}:${l.amount}:${l.interest}:${l.repaid}:${l.date}:${l.due}`).join(' | ');
    return buildCsvRow([
      member.id,
      member.name,
      member.phone,
      member.email,
      member.role,
      memberShares(member),
      memberTotalSaved(member),
      ...weeks,
      memberLoans.length,
      loanSummary
    ]);
  });

  const csv = [buildCsvRow(header), ...rows].join('\n');
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="saved-group-records.csv"');
  res.send(`\uFEFF${csv}`);
});

app.get('/api/admin/export-records-xlsx', authMiddleware, adminOnly, (req, res) => {
  try {
    saveState();
    const maxWeeks = Math.max(6, ...members.map(m => (weeklyData[m.id] || []).length));

    const memberSheet = [
      ['Member ID', 'Name', 'Phone', 'Email', 'Role', 'Shares', 'Total Saved'],
      ...members.map(member => [
        member.id,
        member.name,
        member.phone,
        member.email,
        member.role,
        memberShares(member),
        memberTotalSaved(member),
      ]),
    ];

    const savingHeaders = ['Member ID', 'Name', ...Array.from({ length: maxWeeks }, (_, i) => `Week ${i + 1}`), 'Total Saved'];
    const savingSheet = [
      savingHeaders,
      ...members.map(member => {
        const weeks = Array.from({ length: maxWeeks }, (_, idx) => weeklyData[member.id]?.[idx] || 0);
        return [member.id, member.name, ...weeks, memberTotalSaved(member)];
      }),
    ];

    const loanSheet = [
      ['Loan ID', 'Member ID', 'Name', 'Amount', 'Interest', 'Repaid', 'Date', 'Due Date'],
      ...loans.map(loan => [
        loan.id,
        loan.memberId,
        loan.name,
        loan.amount,
        loan.interest,
        loan.repaid,
        loan.date,
        loan.due,
      ]),
    ];

    const workbook = XLSX.utils.book_new();
    workbook.Props = {
      Title: 'SAVED Group Data',
      Subject: 'SAVED System Storage',
      Author: 'SAVED System',
      CreatedDate: new Date(),
    };
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(memberSheet), 'Members');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(savingSheet), 'Savings');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(loanSheet), 'Loans');

    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="SAVED_Google_Sheets_Template.xlsx"');
    res.send(buffer);
  } catch (err) {
    console.error('Excel export failed', err);
    res.status(500).json({ error: 'Unable to generate Excel file' });
  }
});

app.get('/api/admin/member-receipt/:memberId', authMiddleware, adminOnly, (req, res) => {
  try {
    const member = members.find(m => m.id === req.params.memberId);
    if (!member) return res.status(404).json({ error: 'Member not found' });

    // Build PDF in memory and stream to response (A6 compact receipt)
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="receipt-${member.id}.pdf"`);

    const doc = new PDFDocument({ size: 'A5', margin: 12, autoFirstPage: false });
    doc.pipe(res);

    const cardX = 8;
    const cardY = 8;
    const bandHeight = 34;
    const footerHeight = 46;
    const infoX = cardX + 8;
    const lineHeight = 9;

    let pageWidth = 0;
    let pageHeight = 0;
    let cardW = 0;
    let cardH = 0;
    let footerY = 0;
    let availableBottom = 0;
    let curY = 0;

    let pageStarted = false;
    const startPage = () => {
      doc.addPage();
      pageStarted = true;
      pageWidth = doc.page.width;
      pageHeight = doc.page.height;
      cardW = pageWidth - 16;
      cardH = pageHeight - 16;
      footerY = cardY + cardH - footerHeight;
      availableBottom = footerY - 8;
      curY = cardY + bandHeight + 6;

      doc.rect(0, 0, pageWidth, pageHeight).fill('#f4f6fb');
      doc.roundedRect(cardX, cardY, cardW, cardH, 6).fill('#ffffff');
      doc.fillColor('#2d6cdf').rect(0, cardY, pageWidth, bandHeight).fill('#2d6cdf');
      doc.fillColor('#ffffff').fontSize(14).font('Helvetica-Bold').text('SAVED', 0, cardY + 5, {
        width: pageWidth,
        align: 'center'
      });
      doc.fontSize(7).font('Helvetica').text('Group Savings & Credit', 0, cardY + 17, {
        width: pageWidth,
        align: 'center'
      });
    };

    const ensureSpace = (height = lineHeight) => {
      if (!pageStarted || curY + height > availableBottom) {
        startPage();
      }
    };

    const writeLine = (text, opts = {}) => {
      const size = opts.size || 7;
      const height = opts.height || (size + 1);
      ensureSpace(height);
      doc.fontSize(size).font(opts.font || 'Helvetica').fillColor(opts.color || '#111').text(text, infoX, curY, {
        width: cardW - 16,
        lineBreak: false
      });
      curY += height;
    };

    writeLine(member.name, { size: 10, font: 'Helvetica-Bold' });
    writeLine(`ID: ${member.id}`, { size: 7, color: '#333' });
    writeLine(`Phone: ${member.phone}`);
    writeLine(`Shares: ${memberShares(member)}  •  Email: ${member.email}`);
    curY += 4;

    const weeks = weeklyData[member.id] || [];
    const weekCount = Math.max(6, weeks.length);
    const weekValues = Array.from({ length: weekCount }, (_, idx) => weeks[idx] || 0);
    const total = weekValues.reduce((a, b) => a + (b || 0), 0);
    const loansForMember = loans.filter(l => l.memberId === member.id);

    writeLine('Savings by week', { size: 8, font: 'Helvetica-Bold', color: '#2d6cdf' });
    const displayCount = weekValues.length;
    const weekFontSize = displayCount > 18 ? 7 : displayCount > 14 ? 7.5 : 8;
    const weekLineHeight = weekFontSize + 2;
    weekValues.forEach((amt, idx) => {
      writeLine(`Week ${idx + 1}: ${amt} RWF`, { size: weekFontSize, height: weekLineHeight });
    });
    writeLine(`Total: ${total} RWF`, { size: 9, font: 'Helvetica-Bold' });

    if (loansForMember.length > 0) {
      writeLine('Loans', { size: 8, font: 'Helvetica-Bold', color: '#2d6cdf' });
      const loan = loansForMember[0];
      writeLine(`${loan.id}: ${loan.amount} RWF, repaid ${loan.repaid} RWF`);
      if (loansForMember.length > 1) {
        writeLine(`+${loansForMember.length - 1} more loan(s) on file`);
      }
    }

    if (curY + lineHeight <= availableBottom) {
      // Add space between the Total line and the summary
      curY += 10;
      doc.fontSize(6.5).fillColor('#555').text('This receipt is a compact summary. See the member dashboard for full details.', infoX, curY, {
        width: cardW - 16,
        height: availableBottom - curY,
        lineBreak: false,
        ellipsis: true
      });

      // Advance current y after the summary and leave extra gap before footer
      curY += 12;
      // Place the footer immediately under the summary instead of fixed bottom
      footerY = curY + 8;
      doc.fillColor('#111').fontSize(8).font('Helvetica-Bold').text('SAVED', cardX + 8, footerY, {
        width: cardW - 16,
        lineBreak: false
      });
      doc.font('Helvetica').fontSize(7).fillColor('#111').text('Group Savings & Credit Platform', cardX + 8, footerY + 12, {
        width: cardW - 16,
        lineBreak: false
      });
      doc.fontSize(7).fillColor('#111').text('© 2026 University of Rwanda | College of Education', cardX + 8, footerY + 20, {
        width: cardW - 16,
        lineBreak: false
      });
      doc.fontSize(7).fillColor('#555').text('Customer Support  •  Terms of Service  •  Privacy Policy', cardX + 8, footerY + 28, {
        width: cardW - 16,
        lineBreak: false
      });
      doc.fontSize(7).text('Support Desk: Eric · +250 794 707 599', cardX + 8, footerY + 36, {
        width: cardW - 16,
        lineBreak: false
      });
    } else {
      // Fallback: if there's no space for the summary, draw footer at its original position
      doc.fillColor('#111').fontSize(8).font('Helvetica-Bold').text('SAVED', cardX + 8, footerY + 6, {
        width: cardW - 16,
        lineBreak: false
      });
      doc.font('Helvetica').fontSize(7).fillColor('#111').text('Group Savings & Credit Platform', cardX + 8, footerY + 15, {
        width: cardW - 16,
        lineBreak: false
      });
      doc.fontSize(7).fillColor('#111').text('© 2026 University of Rwanda | College of Education', cardX + 8, footerY + 23, {
        width: cardW - 16,
        lineBreak: false
      });
      doc.fontSize(7).fillColor('#555').text('Customer Support  •  Terms of Service  •  Privacy Policy', cardX + 8, footerY + 31, {
        width: cardW - 16,
        lineBreak: false
      });
      doc.fontSize(7).text('Support Desk: Eric · +250 794 707 599', cardX + 8, footerY + 38, {
        width: cardW - 16,
        lineBreak: false
      });
    }

    doc.end();
  } catch (err) {
    console.error('Receipt generation failed', err);
    res.status(500).json({ error: 'Unable to generate receipt' });
  }
});

app.get('/api/admin/members', authMiddleware, adminOnly, (req, res) => {
  const enriched = members.map(m => ({
    ...m,
    shares: memberShares(m),
    totalSaved: (weeklyData[m.id] || []).reduce((a, b) => a + b, 0),
    weeks: weeklyData[m.id] || [],
    loan: loans.find(l => l.memberId === m.id) || null,
  }));
  res.json(enriched);
});

app.post('/api/admin/members', authMiddleware, adminOnly, (req, res) => {
  const id = String(req.body.id || '').trim();
  const name = String(req.body.name || '').trim();
  const phone = String(req.body.phone || '').trim();
  const email = String(req.body.email || '').trim().toLowerCase();
  const shares = toAmount(req.body.shares);
  const role = req.body.role === 'admin' ? 'admin' : 'member';
  if (!id || !name || !phone || !email || shares === null || shares < 1) {
    return res.status(400).json({ error: 'Member id, name, phone, email, and shares are required' });
  }
  if (members.some(m => m.id === id)) return res.status(409).json({ error: 'Member id already exists' });
  if (members.some(m => m.email.toLowerCase() === email)) return res.status(409).json({ error: 'Email already exists' });
  const member = { id, name, phone, shares, email, role };
  members.push(member);
  weeklyData[id] = [0,0,0,0,0,0];
  addUserLogin(member);
  saveState();
  res.status(201).json(member);
});

app.put('/api/admin/members/:id', authMiddleware, adminOnly, (req, res) => {
  const member = members.find(m => m.id === req.params.id);
  if (!member) return res.status(404).json({ error: 'Member not found' });
  const name = String(req.body.name || '').trim();
  const phone = String(req.body.phone || '').trim();
  const email = String(req.body.email || '').trim().toLowerCase();
  const shares = toAmount(req.body.shares);
  const role = req.body.role === 'admin' ? 'admin' : 'member';
  if (!name || !phone || !email || shares === null || shares < 1) {
    return res.status(400).json({ error: 'Name, phone, email, and shares are required' });
  }
  if (members.some(m => m.id !== member.id && m.email.toLowerCase() === email)) {
    return res.status(409).json({ error: 'Email already exists' });
  }
  delete usersDB[member.email.toLowerCase()];
  Object.assign(member, { name, phone, shares, email, role });
  addUserLogin(member);
  loans.filter(l => l.memberId === member.id).forEach(l => { l.name = member.name; });
  saveState();
  res.json(member);
});

app.delete('/api/admin/members/:id', authMiddleware, adminOnly, (req, res) => {
  if (req.params.id === req.user.id) return res.status(400).json({ error: 'You cannot remove your own admin account' });
  const index = members.findIndex(m => m.id === req.params.id);
  if (index === -1) return res.status(404).json({ error: 'Member not found' });
  const [member] = members.splice(index, 1);
  delete weeklyData[member.id];
  delete usersDB[member.email.toLowerCase()];
  for (let i = loans.length - 1; i >= 0; i--) {
    if (loans[i].memberId === member.id) loans.splice(i, 1);
  }
  saveState();
  res.json({ ok: true });
});

app.get('/api/admin/loans', authMiddleware, adminOnly, (req, res) => {
  res.json(loans);
});

app.post('/api/admin/loans', authMiddleware, adminOnly, (req, res) => {
  const member = members.find(m => m.id === String(req.body.memberId || '').trim());
  const amount = toAmount(req.body.amount);
  const interest = toAmount(req.body.interest);
  const repaid = toAmount(req.body.repaid || 0);
  if (!member || amount === null || amount < 1 || interest === null || repaid === null) {
    return res.status(400).json({ error: 'Member, amount, interest, and repaid amount are required' });
  }
  const loan = {
    id: nextLoanId(),
    memberId: member.id,
    name: member.name,
    amount,
    interest,
    repaid,
    date: String(req.body.date || new Date().toISOString().slice(0, 10)),
    due: String(req.body.due || ''),
  };
  loans.push(loan);
  saveState();
  res.status(201).json(loan);
});

app.put('/api/admin/loans/:id', authMiddleware, adminOnly, (req, res) => {
  const loan = loans.find(l => l.id === req.params.id);
  if (!loan) return res.status(404).json({ error: 'Loan not found' });
  const member = members.find(m => m.id === String(req.body.memberId || loan.memberId).trim());
  const amount = toAmount(req.body.amount);
  const interest = toAmount(req.body.interest);
  const repaid = toAmount(req.body.repaid || 0);
  if (!member || amount === null || amount < 1 || interest === null || repaid === null) {
    return res.status(400).json({ error: 'Member, amount, interest, and repaid amount are required' });
  }
  Object.assign(loan, {
    memberId: member.id,
    name: member.name,
    amount,
    interest,
    repaid,
    date: String(req.body.date || loan.date),
    due: String(req.body.due || ''),
  });
  saveState();
  res.json(loan);
});

app.delete('/api/admin/loans/:id', authMiddleware, adminOnly, (req, res) => {
  const index = loans.findIndex(l => l.id === req.params.id);
  if (index === -1) return res.status(404).json({ error: 'Loan not found' });
  loans.splice(index, 1);
  saveState();
  res.json({ ok: true });
});

app.get('/api/admin/savings', authMiddleware, adminOnly, (req, res) => {
  const rows = members.map(m => ({
    id: m.id,
    name: m.name,
    shares: memberShares(m),
    weeks: weeklyData[m.id] || [0,0,0,0,0,0],
    total: (weeklyData[m.id] || []).reduce((a, b) => a + b, 0),
  }));
  res.json(rows);
});

app.put('/api/admin/savings/:memberId', authMiddleware, adminOnly, (req, res) => {
  const member = members.find(m => m.id === req.params.memberId);
  const week = Number(req.body.week);
  const amount = toAmount(req.body.amount);
  if (!member) return res.status(404).json({ error: 'Member not found' });
  if (!Number.isInteger(week) || week < 1) return res.status(400).json({ error: 'Week must be 1 or higher' });
  if (amount === null) return res.status(400).json({ error: 'Amount must be a positive number or zero' });
  if (!weeklyData[member.id]) weeklyData[member.id] = [];
  while (weeklyData[member.id].length < week) weeklyData[member.id].push(0);
  weeklyData[member.id][week - 1] = amount;
  saveState();
  res.json({ id: member.id, weeks: weeklyData[member.id], total: weeklyData[member.id].reduce((a, b) => a + b, 0) });
});

// ─── SERVE FRONTEND ─────────────────────────────────────────────────────────

app.use((err, req, res, next) => {
  console.error(err);
  if (res.headersSent) return next(err);
  res.status(500).json({ error: 'Server error' });
});

app.use('/api', (req, res) => {
  console.log(`[API MISS] ${req.method} ${req.originalUrl}`);
  res.status(404).json({ error: 'API route not found' });
});

app.get(/^.*$/, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SAVED System running on http://localhost:${PORT}`));
