'use strict';

const AUTH_TOKEN_KEY = 'saved_auth_token';

// ── HELPERS ──────────────────────────────────────────────────────────────────

const $ = id => document.getElementById(id);
const fmt = n => Number(n).toLocaleString();
const parseAmountInput = value => {
  const normalized = String(value || '').trim().replace(/[,\s]/g, '');
  return normalized === '' ? NaN : Number(normalized);
};
const initials = name => name.trim().split(/\s+/).slice(0,2).map(w=>w[0].toUpperCase()).join('');
const avClass = i => ['av0','av1','av2','av3','av4'][i % 5];
const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
const weekLabels = count => Array.from({ length: Math.max(1, count) }, (_, i) => `Wk ${i + 1}`);
const nextSavingsWeek = weeks => {
  const list = Array.isArray(weeks) ? weeks : [];
  const firstEmptyIndex = list.findIndex(amount => !amount);
  return firstEmptyIndex === -1 ? list.length + 1 : firstEmptyIndex + 1;
};

let currentUser = null;
let charts = {};
let memberRowsClickHandlerSet = false;
let currentPage = null;
let autoRefreshInterval = null;

async function api(path, opts={}) {
  const url = path.startsWith('http') ? path : path;
  const token = localStorage.getItem(AUTH_TOKEN_KEY);
  const headers = {
    'Content-Type':'application/json',
    ...(opts.headers||{}),
  };
  if (token) headers.Authorization = `Bearer ${token}`;

  let r;
  try {
    r = await fetch(url, {
      headers,
      credentials:'include',
      ...opts,
      body: opts.body ? JSON.stringify(opts.body) : undefined,
    });
  } catch {
    throw new Error('Cannot reach the server. Refresh the page and try again.');
  }
  if (!r.ok) {
    const text = await r.text().catch(() => '');
    let payload;
    try { payload = text ? JSON.parse(text) : null; } catch { payload = null; }
    const message = payload?.error || payload?.message || text || `HTTP ${r.status}`;
    throw new Error(message);
  }
  return r.json();
}

// Theme handling: apply saved theme or default to dark
function applyTheme(theme) {
  const body = document.body;
  if (theme === 'light') body.classList.add('light-theme');
  else body.classList.remove('light-theme');
  const btn = document.getElementById('theme-toggle');
  if (btn) btn.textContent = theme === 'light' ? 'Dark theme' : 'Light theme';
}

function toggleTheme() {
  const current = localStorage.getItem('saved_theme') || 'dark';
  const next = current === 'light' ? 'dark' : 'light';
  localStorage.setItem('saved_theme', next);
  applyTheme(next);
  toast(`Theme switched to ${next}`, 'info');
}

// Initialize theme on load
document.addEventListener('DOMContentLoaded', () => {
  const saved = localStorage.getItem('saved_theme') || 'dark';
  applyTheme(saved);
});

async function downloadFile(path) {
  const url = path.startsWith('http') ? path : path;
  const token = localStorage.getItem(AUTH_TOKEN_KEY);
  const headers = {
    ...(token ? { Authorization: `Bearer ${token}` } : {}),
  };

  let r;
  try {
    r = await fetch(url, {
      headers,
      credentials:'include',
    });
  } catch {
    throw new Error('Cannot reach the server. Refresh the page and try again.');
  }
  if (!r.ok) {
    const text = await r.text().catch(() => '');
    let payload;
    try { payload = text ? JSON.parse(text) : null; } catch { payload = null; }
    const message = payload?.error || payload?.message || text || `HTTP ${r.status}`;
    throw new Error(message);
  }
  return r.blob();
}

function showToast(message, opts = {}) {
  const { type = 'success', title = '', timeout = 4000, actionText, action } = opts;
  const root = document.getElementById('toast-root') || document.body;
  const t = document.createElement('div');
  t.className = `toast ${type}`;

  const icon = document.createElement('div');
  icon.className = 'toast-icon';
  icon.innerHTML = type === 'error' ? '⛔' : type === 'info' ? 'ℹ️' : '✓';

  const body = document.createElement('div');
  body.className = 'toast-body';
  if (title) body.innerHTML = `<div class="toast-title">${esc(title)}</div>`;
  const msg = document.createElement('div');
  msg.className = 'toast-msg';
  msg.textContent = message;
  body.appendChild(msg);

  const actions = document.createElement('div');
  actions.className = 'toast-actions';
  if (actionText && typeof action === 'function') {
    const a = document.createElement('button');
    a.className = 'toast-action';
    a.textContent = actionText;
    a.onclick = (e) => { e.stopPropagation(); action(); remove(); };
    actions.appendChild(a);
  }
  const close = document.createElement('button');
  close.className = 'toast-close';
  close.innerHTML = '✕';
  close.onclick = (e) => { e.stopPropagation(); remove(); };
  actions.appendChild(close);
  body.appendChild(actions);

  const progress = document.createElement('div');
  progress.className = 'toast-progress';
  progress.style.width = '100%';
  body.appendChild(progress);

  t.appendChild(icon);
  t.appendChild(body);

  const remove = () => {
    t.classList.remove('show');
    clearTimeout(timer);
    setTimeout(() => t.remove(), 300);
  };

  root.appendChild(t);
  // Allow CSS transitions
  requestAnimationFrame(() => t.classList.add('show'));

  // Animate progress shrink
  requestAnimationFrame(() => { progress.style.transition = `width ${timeout}ms linear`; progress.style.width = '0%'; });

  const timer = setTimeout(() => remove(), timeout + 80);

  // Return control
  return { remove };
}

// Backwards-compat wrapper
function toast(msg, type='success') { return showToast(msg, { type }); }

async function downloadAdminSheet() {
  try {
    const blob = await downloadFile('/api/admin/export-records-xlsx');
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'SAVED_Google_Sheets_Template.xlsx';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast('Excel template download started');
  } catch (e) {
    toast(e.message, 'error');
  }
}

async function downloadMemberReceipt(memberId) {
  try {
    const blob = await downloadFile(`/api/admin/member-receipt/${encodeURIComponent(memberId)}?t=${Date.now()}`);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `receipt-${memberId}.pdf`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast('Receipt download started');
  } catch (e) {
    toast(e.message, 'error');
  }
}

function destroyChart(key) {
  if (charts[key]) { charts[key].destroy(); delete charts[key]; }
}

// ── LOGIN ────────────────────────────────────────────────────────────────────

function renderLogin() {
  document.body.classList.remove('admin-view');
  document.getElementById('app').innerHTML = `
  <div id="login-screen">
    <div class="login-box">
      <div class="login-logo">
        <div class="mark">S</div>
        <h1>SAVED Group</h1>
        <p>Secure savings and credit management</p>
      </div>
      <div class="login-error" id="login-error"></div>
      <div class="field">
        <label>Email address</label>
        <input type="email" id="login-email" placeholder="your@email.com" autocomplete="email"/>
      </div>
      <div class="field">
        <label>Password</label>
        <input type="password" id="login-pass" placeholder="Your password" autocomplete="current-password"/>
        <div class="hint">First login uses the last 4 digits of your registered phone number.</div>
      </div>
      <button class="btn-primary" onclick="doLogin()">Sign In</button>
    </div>
  </div>`;
  document.querySelector('.login-logo .mark').textContent = 'S';
  $('login-email').focus();
  $('login-email').addEventListener('keydown', e => { if(e.key==='Enter') doLogin(); });
  $('login-pass').addEventListener('keydown', e => { if(e.key==='Enter') doLogin(); });
}

async function doLogin() {
  const email = $('login-email').value.trim();
  const pass  = $('login-pass').value;
  const err   = $('login-error');
  err.style.display = 'none';
  try {
    const data = await api('/api/login', { method:'POST', body:{email, password:pass} });
    if (data.token) localStorage.setItem(AUTH_TOKEN_KEY, data.token);
    currentUser = data;
    renderShell();
    if (data.role === 'admin') loadAdminDashboard();
    else loadMyDashboard();
  } catch(e) {
    err.textContent = e.message;
    err.style.display = 'block';
  }
}

async function doLogout() {
  localStorage.removeItem(AUTH_TOKEN_KEY);
  await api('/api/logout', { method:'POST' });
  currentUser = null;
  Object.values(charts).forEach(c=>c.destroy());
  charts = {};
  renderLogin();
}

// ── SHELL ────────────────────────────────────────────────────────────────────

function renderShell() {
  const isAdmin = currentUser.role === 'admin';
  document.body.classList.toggle('admin-view', isAdmin);
  const adminNav = isAdmin ? `
    <div class="nav-group">Management</div>
    <div class="nav-item" data-page="members" onclick="navigate('members')">
      ${ico('users')} Members
    </div>
    <div class="nav-item" data-page="savings" onclick="navigate('savings')">
      ${ico('wallet')} Savings
    </div>
    <div class="nav-item" data-page="loans" onclick="navigate('loans')">
      ${ico('card')} Loans
    </div>
    <div class="nav-item" data-page="analytics" onclick="navigate('analytics')">
      ${ico('chart')} Analytics
    </div>` : '';

  const memberNav = !isAdmin ? `
    <div class="nav-group">My Account</div>
    <div class="nav-item" data-page="my-savings" onclick="navigate('my-savings')">
      ${ico('wallet')} My Savings
    </div>
    <div class="nav-item" data-page="my-loan" onclick="navigate('my-loan')">
      ${ico('card')} My Loan
    </div>` : '';

  document.getElementById('app').innerHTML = `
  <div id="app-shell" class="visible">
    <aside class="sidebar">
      <div class="sidebar-logo">
        <div class="name">SAVED Group</div>
        <div class="sub">Savings & Credit System</div>
      </div>
      <div class="sidebar-user">
        <div class="user-avatar">${initials(currentUser.name)}</div>
        <div>
          <div class="user-name">${currentUser.name.split(' ').slice(0,2).join(' ')}${isAdmin?'<span class="admin-badge">Admin</span>':''}</div>
          <div class="user-role">${isAdmin?'Administrator':'Member'}</div>
        </div>
      </div>
      <nav>
        <div class="nav-group">Overview</div>
        <div class="nav-item active" data-page="${isAdmin?'dashboard':'my-dashboard'}" onclick="navigate('${isAdmin?'dashboard':'my-dashboard'}')">
          ${ico('home')} ${isAdmin?'Dashboard':'My Dashboard'}
        </div>
        ${adminNav}
        ${memberNav}
        <div class="nav-group">Account</div>
        <div class="nav-item" data-page="change-pass" onclick="navigate('change-pass')">
          ${ico('lock')} Change Password
        </div>
      </nav>
      <div class="sidebar-footer">
        <button class="btn-logout" onclick="doLogout()">Sign Out</button>
      </div>
    </aside>
    <main class="main-content" id="main-content">
      <div class="page active" id="page-${isAdmin?'dashboard':'my-dashboard'}">
        <div class="loading"><div class="spinner"></div>Loading…</div>
      </div>
      <div class="page" id="page-members"></div>
      <div class="page" id="page-savings"></div>
      <div class="page" id="page-loans"></div>
      <div class="page" id="page-analytics"></div>
      <div class="page" id="page-my-dashboard"></div>
      <div class="page" id="page-my-savings"></div>
      <div class="page" id="page-my-loan"></div>
      <div class="page" id="page-change-pass"></div>
    </main>
  </div>`;
}

function navigate(page) {
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const pg = $('page-'+page);
  if(pg) pg.classList.add('active');
  document.querySelector(`.nav-item[data-page="${page}"]`)?.classList.add('active');

  currentPage = page;
  
  // Clear old refresh interval
  if (autoRefreshInterval) clearInterval(autoRefreshInterval);
  
  // Start auto-refresh for member pages (every 2 seconds), except my-savings because it contains an input form
  const autoRefreshPages = ['my-dashboard', 'my-loan'];
  if (autoRefreshPages.includes(page) && currentUser?.role !== 'admin') {
    // Refresh immediately when page loads
    if (page === 'my-dashboard') loadMyDashboard().catch(() => {});
    else if (page === 'my-loan') loadMyLoan().catch(() => {});
    
    // Then set up periodic refresh every 2 seconds
    autoRefreshInterval = setInterval(() => {
      if (currentPage === page) {
        if (page === 'my-dashboard') loadMyDashboard().catch(() => {});
        else if (page === 'my-loan') loadMyLoan().catch(() => {});
      }
    }, 2000);
  }

  const isAdmin = currentUser?.role==='admin';
  if (page==='dashboard' && isAdmin) loadAdminDashboard();
  else if (page==='members' && isAdmin) loadMembers();
  else if (page==='savings' && isAdmin) loadSavings();
  else if (page==='loans' && isAdmin) loadLoans();
  else if (page==='analytics' && isAdmin) loadAnalytics();
  else if (page==='my-dashboard' && !isAdmin) loadMyDashboard();
  else if (page==='my-savings' && !isAdmin) loadMySavings();
  else if (page==='my-loan' && !isAdmin) loadMyLoan();
  else if (page==='change-pass') renderChangePass();
}

// ── ICONS ────────────────────────────────────────────────────────────────────

function ico(name) {
  const p = {
    home:`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>`,
    users:`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
    wallet:`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M21 12V7H5a2 2 0 0 1 0-4h14v4"/><path d="M3 5v14a2 2 0 0 0 2 2h16v-5"/><circle cx="18" cy="12" r="1"/></svg>`,
    card:`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="1" y="4" width="22" height="16" rx="2"/><line x1="1" y1="10" x2="23" y2="10"/></svg>`,
    chart:`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>`,
    lock:`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>`,
  };
  return `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7">${p[name]?.replace(/<svg[^>]*>/,'').replace('</svg>','')}</svg>`;
}

// ── ADMIN DASHBOARD ──────────────────────────────────────────────────────────

async function loadAdminDashboard() {
  const pg = $('page-dashboard');
  pg.innerHTML = '<div class="loading"><div class="spinner"></div>Loading…</div>';
  try {
    const d = await api('/api/admin/dashboard');
    const growthData = Array.isArray(d.weeklyGrowth) ? d.weeklyGrowth : [0,0,0,0,0,0];
    const currentWeek = growthData.length;
    const shareDistribution = Array.isArray(d.shareDistribution) ? d.shareDistribution : [0, 0, 0, 0];
    const shareLegendHtml = shareDistribution.map((count, idx) => {
      const colors = ['#4f8eff','#00c9a7','#ffd166','#ff6b6b'];
      return `<span style="display:flex;align-items:center;gap:5px"><span style="width:9px;height:9px;border-radius:2px;background:${colors[idx]};display:inline-block"></span>${idx+1} share${idx===0?'':'s'} (${count})</span>`;
    }).join('');
    pg.innerHTML = `
    <div class="page-header">
      <div>
        <div class="page-title">Dashboard</div>
        <div class="page-sub">SAVED Group — Week ${currentWeek} summary</div>
      </div>
      <button class="btn-small" onclick="downloadAdminSheet()">Download Excel</button>
    </div>
    <div class="stat-grid">
      <div class="stat-card blue">
        <div class="stat-label">Total Members</div>
        <div class="stat-val">${d.totalMembers}</div>
        <div class="stat-sub">${d.totalShares} shares total</div>
      </div>
      <div class="stat-card green">
        <div class="stat-label">Total Savings</div>
        <div class="stat-val">${fmt(d.totalSaved)}</div>
        <div class="stat-sub">RWF collected</div>
      </div>
      <div class="stat-card amber">
        <div class="stat-label">Loans Issued</div>
        <div class="stat-val">${fmt(d.totalLoans)}</div>
        <div class="stat-sub">5 active loans</div>
      </div>
      <div class="stat-card red">
        <div class="stat-label">Outstanding</div>
        <div class="stat-val">${fmt(d.outstanding)}</div>
        <div class="stat-sub">incl. ${fmt(d.totalInterest)} interest</div>
      </div>
    </div>
    <div class="grid-2 mb-24">
      <div class="panel">
        <div class="panel-head"><div><div class="panel-title">Savings Growth</div><div class="panel-sub">Cumulative weekly total (RWF)</div></div></div>
        <div class="panel-body"><div class="chart-wrap" style="height:210px"><canvas id="dash-growth-chart" role="img" aria-label="Weekly savings growth chart"></canvas></div></div>
      </div>
      <div class="panel">
        <div class="panel-head"><div><div class="panel-title">Share Distribution</div><div class="panel-sub">Members by share count</div></div></div>
        <div class="panel-body">
          <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px;font-size:12px;color:var(--text2)">
            ${shareLegendHtml}
          </div>
          <div class="chart-wrap" style="height:170px"><canvas id="dash-share-chart" role="img" aria-label="Share distribution donut chart"></canvas></div>
        </div>
      </div>
    </div>
    <div class="panel">
      <div class="panel-head"><div class="panel-title">Active Loans</div><span class="pill pill-red">0 repaid</span></div>
      <table>
        <thead><tr><th>Member</th><th>Principal</th><th>Interest</th><th>Total Due</th><th>Status</th></tr></thead>
        <tbody id="dash-loan-table"></tbody>
      </table>
    </div>`;

    // Charts
    destroyChart('dash-growth');
    destroyChart('dash-share');
    charts['dash-growth'] = new Chart($('dash-growth-chart'),{
      type:'line',
      data:{labels:weekLabels(growthData.length),
        datasets:[{data: growthData,borderColor:'#00c9a7',backgroundColor:'rgba(0,201,167,.08)',fill:true,tension:.4,pointBackgroundColor:'#00c9a7',pointRadius:4,borderWidth:2}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
        scales:{x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#8892aa',font:{size:11}}},
          y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#8892aa',font:{size:11},callback:v=>fmt(v)}}}}
    });
    charts['dash-share'] = new Chart($('dash-share-chart'),{
      type:'doughnut',
      data:{labels:['1 share','2 shares','3 shares','4 shares'],
        datasets:[{data:d.shareDistribution,backgroundColor:['#4f8eff','#00c9a7','#ffd166','#ff6b6b'],borderWidth:0}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},cutout:'65%'}
    });

    // Loan table
    const loanData = await api('/api/admin/loans');
    const tbody = $('dash-loan-table');
    loanData.forEach(l=>{
      tbody.innerHTML+=`<tr><td>${l.name}</td><td>${fmt(l.amount)} RWF</td><td>${fmt(l.interest)} RWF</td><td style="font-weight:500;color:var(--amber)">${fmt(l.amount+l.interest)} RWF</td><td><span class="pill pill-red">Outstanding</span></td></tr>`;
    });
  } catch (e) {
    pg.innerHTML = `<div class="panel" style="padding:32px;text-align:center;color:var(--text2)"><div style="font-size:18px;margin-bottom:12px">Unable to load dashboard</div><div style="font-size:14px">${e.message}</div></div>`;
    toast(`Dashboard load failed: ${e.message}`, 'error');
  }
}

// ── ADMIN MEMBERS ────────────────────────────────────────────────────────────

async function loadMembers() {
  const pg = $('page-members');
  pg.innerHTML = '<div class="loading"><div class="spinner"></div>Loading…</div>';
  const data = await api('/api/admin/members');
  pg.innerHTML = `
  <div class="page-header"><div class="page-title">Members</div><div class="page-sub">${data.length} registered members</div></div>
  <div class="panel mb-24">
    <div class="panel-head"><div><div class="panel-title">Add / Edit Member</div><div class="panel-sub">Default password is the last 4 digits of phone</div></div></div>
    <div class="panel-body">
      <input type="hidden" id="member-edit-id"/>
      <div class="form-grid">
        <div class="field"><label>Member ID</label><input id="member-id" placeholder="National ID or member code"/></div>
        <div class="field"><label>Name</label><input id="member-name" placeholder="Full name"/></div>
        <div class="field"><label>Phone</label><input id="member-phone" placeholder="0780000000"/></div>
        <div class="field"><label>Email</label><input id="member-email" type="email" placeholder="name@email.com"/></div>
        <div class="field"><label>Shares</label><input id="member-shares" type="number" min="1" value="1"/></div>
        <div class="field"><label>Role</label><select id="member-role"><option value="member">Member</option><option value="admin">Admin</option></select></div>
      </div>
      <div class="action-row">
        <button class="btn-small primary" onclick="saveMember()">Save Member</button>
        <button class="btn-small" onclick="resetMemberForm()">Clear</button>
      </div>
    </div>
  </div>
  <div class="panel">
    <div class="panel-head">
      <div class="panel-title">Member Registry</div>
      <div class="search-wrap">
        <svg class="search-icon" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input class="search-input" placeholder="Search member…" oninput="filterTable(this.value,'member-tbody','member-rows')" id="m-search"/>
      </div>
    </div>
    <table>
      <thead><tr><th>Member</th><th>Phone</th><th>Shares</th><th>Total Saved</th><th>Loan</th><th>Actions</th></tr></thead>
      <tbody id="member-tbody"></tbody>
    </table>
  </div>`;

  window['member-rows'] = data;
  renderMemberRows(data);
  if (!memberRowsClickHandlerSet) {
    const tbody = $('member-tbody');
    if (tbody) {
      tbody.addEventListener('click', e => {
        const btn = e.target.closest('button[data-action="edit-member"]');
        if (!btn) return;
        editMember(btn.dataset.id);
      });
    }
    memberRowsClickHandlerSet = true;
  }
}

function renderMemberRows(data) {
  const tbody = $('member-tbody');
  tbody.innerHTML = '';
  data.forEach((m,i)=>{
    const loan = m.loan;
    tbody.innerHTML+=`<tr>
      <td><div class="member-cell">
        <div class="avatar ${avClass(i)}">${initials(m.name)}</div>
        <div><div class="m-name">${m.name}</div><div class="m-id">${m.id.slice(-8)}</div></div>
      </div></td>
      <td style="font-family:var(--mono);color:var(--text2);font-size:12px">${m.phone}</td>
      <td><div class="shares-bar"><div class="bar-bg"><div class="bar-fill" style="width:${(m.shares/4)*100}%"></div></div><span>${m.shares}</span></div></td>
      <td><span style="font-weight:500">${fmt(m.totalSaved)}</span> <span style="color:var(--text3);font-size:11px">RWF</span></td>
      <td>${loan?`<span class="pill pill-amber">${fmt(loan.amount)} RWF</span>`:`<span class="pill pill-green">No loan</span>`}</td>
      <td><div class="row-actions">
        <button type="button" class="icon-btn" onclick="downloadMemberReceipt('${esc(m.id)}')">Receipt</button>
        <button type="button" class="icon-btn" data-action="edit-member" data-id="${esc(m.id)}">Edit</button>
        <button type="button" class="icon-btn danger" onclick="deleteMember('${m.id}')">Remove</button>
      </div></td>
    </tr>`;
  });
}

function filterTable(q, tbodyId, dataKey) {
  const filtered = (window[dataKey]||[]).filter(m=>m.name.toLowerCase().includes(q.toLowerCase()));
  if (dataKey==='member-rows') renderMemberRows(filtered);
}

function resetMemberForm() {
  ['member-edit-id','member-id','member-name','member-phone','member-email'].forEach(id => { if ($(id)) $(id).value = ''; });
  if ($('member-shares')) $('member-shares').value = 1;
  if ($('member-role')) $('member-role').value = 'member';
  if ($('member-id')) $('member-id').disabled = false;
}

function editMember(id) {
  const m = (window['member-rows'] || []).find(x => x.id === id);
  if (!m) return;
  $('member-edit-id').value = m.id;
  $('member-id').value = m.id;
  $('member-id').disabled = true;
  $('member-name').value = m.name;
  $('member-phone').value = m.phone;
  $('member-email').value = m.email;
  $('member-shares').value = m.shares;
  $('member-role').value = m.role;
  $('member-name')?.focus();
  $('member-id')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

async function saveMember() {
  const editingId = $('member-edit-id').value;
  const body = {
    id: $('member-id').value.trim(),
    name: $('member-name').value.trim(),
    phone: $('member-phone').value.trim(),
    email: $('member-email').value.trim(),
    shares: Number($('member-shares').value),
    role: $('member-role').value,
  };
  try {
    await api(editingId ? `/api/admin/members/${encodeURIComponent(editingId)}` : '/api/admin/members', {
      method: editingId ? 'PUT' : 'POST',
      body
    });
    toast(editingId ? 'Member updated' : 'Member added');
    await loadMembers();
  } catch(e) {
    toast(e.message, 'error');
  }
}

async function deleteMember(id) {
  if (!confirm('Remove this member, savings record, and linked loans?')) return;
  try {
    await api(`/api/admin/members/${encodeURIComponent(id)}`, { method:'DELETE' });
    toast('Member removed');
    await loadMembers();
  } catch(e) {
    toast(e.message, 'error');
  }
}

// ── ADMIN SAVINGS ────────────────────────────────────────────────────────────

async function loadSavings() {
  const pg = $('page-savings');
  pg.innerHTML = '<div class="loading"><div class="spinner"></div>Loading…</div>';
  const data = await api('/api/admin/savings');
  window['savings-rows'] = data;
  const memberOptions = data.map(m => `<option value="${esc(m.id)}">${esc(m.name)}</option>`).join('');
  const maxWeeks = Math.max(6, ...data.map(m => m.weeks.length));
  const gridTemplate = `grid-template-columns:190px repeat(${maxWeeks},1fr) 90px;`;
  let rows = '';
  data.forEach((m,i)=>{
    const weeks = Array.from({ length: maxWeeks }, (_, idx) => m.weeks[idx] || 0);
    let cells = weeks.map((amt, weekIndex)=>
      `<div class="wk-cell ${amt>0?'wk-paid':'wk-miss'} editable" onclick="editSaving('${esc(m.id)}',${weekIndex+1},${amt})" title="Edit week ${weekIndex+1} amount">
        <div class="wk-amt">${amt>0?fmt(amt):'—'}</div>
        <div class="wk-ico">${amt>0?'✓':'✗'}</div>
      </div>`).join('');
    rows+=`<div class="savings-row" style="${gridTemplate}">
      <div class="member-cell">
        <div class="avatar ${avClass(i)}" style="width:28px;height:28px;font-size:10px">${initials(m.name)}</div>
        <div style="font-size:12px;font-weight:500;line-height:1.3">${m.name}</div>
      </div>
      ${cells}
      <div class="total-val">${fmt(m.total)}</div>
    </div>`;
  });
  const totals = Array.from({ length: maxWeeks }, (_, w) => data.reduce((s,m)=>s+(m.weeks[w]||0),0));
  let totCells = totals.map(t=>`<span style="text-align:center;font-size:11px;font-weight:500;color:var(--green)">${fmt(t)}</span>`).join('');
  const headerCells = Array.from({ length: maxWeeks }, (_, i) => `<span>Week ${i+1}</span>`).join('');
  pg.innerHTML = `
  <div class="page-header"><div class="page-title">Savings Tracker</div><div class="page-sub">Per-member contributions across ${maxWeeks} weeks</div></div>
  <div class="panel mb-24">
    <div class="panel-head"><div><div class="panel-title">Update Contribution</div><div class="panel-sub">Set a member amount for any week</div></div></div>
    <div class="panel-body">
      <div class="form-grid compact">
        <div class="field"><label>Member</label><select id="saving-member">${memberOptions}</select></div>
        <div class="field"><label>Week</label><input id="saving-week" type="number" min="1" value="1"/></div>
        <div class="field"><label>Amount</label><input id="saving-amount" type="text" placeholder="Enter amount" autocomplete="off" onfocus="this.select()"/></div>
      </div>
      <div class="action-row"><button class="btn-small primary" onclick="saveSaving()">Save Amount</button></div>
    </div>
  </div>
  <div class="savings-head" style="${gridTemplate}">
    <span>Member</span>${headerCells}<span>Total</span>
  </div>
  ${rows}
  <div class="savings-row" style="${gridTemplate}border-color:rgba(0,201,167,.3);background:rgba(0,201,167,.04)">
    <div style="font-size:12px;font-weight:600;color:var(--green)">GROUP TOTAL</div>
    ${totCells}
    <div style="font-size:14px;font-weight:700;color:var(--green)">${fmt(totals.reduce((a,b)=>a+b,0))}</div>
  </div>`;
}

// ── ADMIN LOANS ──────────────────────────────────────────────────────────────

async function saveSaving() {
  const memberId = $('saving-member').value;
  const week = Number($('saving-week').value);
  const amount = parseAmountInput($('saving-amount').value);
  if (!Number.isInteger(week) || week < 1) return toast('Week must be 1 or higher', 'error');
  if (!Number.isFinite(amount) || amount <= 0) return toast('Amount must be a positive number', 'error');
  const body = { week, amount };
  try {
    await api(`/api/admin/savings/${encodeURIComponent(memberId)}`, { method:'PUT', body });
    toast('Contribution updated');
    await loadSavings();
  } catch(e) {
    toast(e.message, 'error');
  }
}

function editSaving(memberId, week, amount) {
  if ($('saving-member')) $('saving-member').value = memberId;
  if ($('saving-week')) $('saving-week').value = week;
  if ($('saving-amount')) $('saving-amount').value = amount;
  toast(`Edit Week ${week} for member`, 'success');
}

async function loadLoans() {
  const pg = $('page-loans');
  pg.innerHTML = '<div class="loading"><div class="spinner"></div>Loading…</div>';
  const data = await api('/api/admin/loans');
  const members = await api('/api/admin/members');
  window['loan-rows'] = data;
  window['loan-members'] = members;
  const loanMemberOptions = members.map(m => `<option value="${esc(m.id)}">${esc(m.name)}</option>`).join('');
  let cards = '';
  data.forEach(l=>{
    const total = l.amount + l.interest;
    const pct = Math.round((l.repaid/total)*100);
    cards+=`<div class="loan-item">
      <div class="loan-header">
        <div>
          <div class="loan-name">${l.name}</div>
          <div class="loan-id">ID: ${l.memberId.slice(-8)} · Loan ${l.id}</div>
        </div>
        <div class="row-actions">
          <span class="pill ${pct >= 100 ? 'pill-green' : 'pill-red'}">${pct >= 100 ? 'Repaid' : 'Outstanding'}</span>
          <button class="icon-btn" onclick="editLoan('${l.id}')">Edit</button>
          <button class="icon-btn danger" onclick="deleteLoan('${l.id}')">Remove</button>
        </div>
      </div>
      <div class="loan-meta">
        <div class="lm"><label>Principal</label><span>${fmt(l.amount)} RWF</span></div>
        <div class="lm"><label>Interest (10%)</label><span>${fmt(l.interest)} RWF</span></div>
        <div class="lm total"><label>Total Due</label><span>${fmt(total)} RWF</span></div>
        <div class="lm"><label>Repaid</label><span style="color:var(--text3)">${fmt(l.repaid)} RWF</span></div>
      </div>
      <div class="progress-label"><span>Repayment: ${pct}%</span><span>${fmt(l.repaid)} / ${fmt(total)} RWF</span></div>
      <div class="progress-track"><div class="progress-fill" style="width:${pct}%"></div></div>
    </div>`;
  });
  const totPrincipal = data.reduce((a,l)=>a+l.amount,0);
  const totInterest  = data.reduce((a,l)=>a+l.interest,0);
  pg.innerHTML = `
  <div class="page-header"><div class="page-title">Loan Management</div><div class="page-sub">${data.length} active loans</div></div>
  <div class="panel mb-24">
    <div class="panel-head"><div><div class="panel-title">Add / Edit Loan</div><div class="panel-sub">Interest can be entered manually for each loan</div></div></div>
    <div class="panel-body">
      <input type="hidden" id="loan-edit-id"/>
      <div class="form-grid">
        <div class="field"><label>Member</label><select id="loan-member">${loanMemberOptions}</select></div>
        <div class="field"><label>Principal</label><input id="loan-amount" type="number" min="1" step="100" placeholder="10000"/></div>
        <div class="field"><label>Interest</label><input id="loan-interest" type="number" min="0" step="100" placeholder="1000"/></div>
        <div class="field"><label>Repaid</label><input id="loan-repaid" type="number" min="0" step="100" value="0"/></div>
        <div class="field"><label>Date</label><input id="loan-date" type="date"/></div>
        <div class="field"><label>Due Date</label><input id="loan-due" type="date"/></div>
      </div>
      <div class="action-row">
        <button class="btn-small primary" onclick="saveLoan()">Save Loan</button>
        <button class="btn-small" onclick="resetLoanForm()">Clear</button>
      </div>
    </div>
  </div>
  <div class="stat-grid" style="grid-template-columns:repeat(3,1fr)">
    <div class="stat-card amber"><div class="stat-label">Total Principal</div><div class="stat-val">${fmt(totPrincipal)}</div><div class="stat-sub">RWF lent out</div></div>
    <div class="stat-card red"><div class="stat-label">Total Interest</div><div class="stat-val">${fmt(totInterest)}</div><div class="stat-sub">10% rate</div></div>
    <div class="stat-card blue"><div class="stat-label">Total Due</div><div class="stat-val">${fmt(totPrincipal+totInterest)}</div><div class="stat-sub">Outstanding balance</div></div>
  </div>
  ${cards}`;
}

// ── ADMIN ANALYTICS ──────────────────────────────────────────────────────────

function resetLoanForm() {
  ['loan-edit-id','loan-amount','loan-interest','loan-due'].forEach(id => { if ($(id)) $(id).value = ''; });
  if ($('loan-repaid')) $('loan-repaid').value = 0;
  if ($('loan-date')) $('loan-date').value = new Date().toISOString().slice(0, 10);
  if ($('loan-member') && $('loan-member').options.length) $('loan-member').selectedIndex = 0;
}

function editLoan(id) {
  const loan = (window['loan-rows'] || []).find(l => l.id === id);
  if (!loan) return;
  $('loan-edit-id').value = loan.id;
  $('loan-member').value = loan.memberId;
  $('loan-amount').value = loan.amount;
  $('loan-interest').value = loan.interest;
  $('loan-repaid').value = loan.repaid;
  $('loan-date').value = loan.date || '';
  $('loan-due').value = loan.due || '';
}

async function saveLoan() {
  const editingId = $('loan-edit-id').value;
  const body = {
    memberId: $('loan-member').value,
    amount: Number($('loan-amount').value),
    interest: Number($('loan-interest').value),
    repaid: Number($('loan-repaid').value || 0),
    date: $('loan-date').value,
    due: $('loan-due').value,
  };
  try {
    await api(editingId ? `/api/admin/loans/${encodeURIComponent(editingId)}` : '/api/admin/loans', {
      method: editingId ? 'PUT' : 'POST',
      body
    });
    toast(editingId ? 'Loan updated' : 'Loan added');
    await loadLoans();
  } catch(e) {
    toast(e.message, 'error');
  }
}

async function deleteLoan(id) {
  if (!confirm('Remove this loan?')) return;
  try {
    await api(`/api/admin/loans/${encodeURIComponent(id)}`, { method:'DELETE' });
    toast('Loan removed');
    await loadLoans();
  } catch(e) {
    toast(e.message, 'error');
  }
}

async function loadAnalytics() {
  const pg = $('page-analytics');
  pg.innerHTML = '<div class="loading"><div class="spinner"></div>Loading…</div>';
  const data = await api('/api/admin/savings');
  const sorted = [...data].sort((a,b)=>b.total-a.total);

  pg.innerHTML = `
  <div class="page-header"><div class="page-title">Analytics</div><div class="page-sub">Member performance breakdown</div></div>
  <div class="panel mb-24">
    <div class="panel-head"><div class="panel-title">Total Savings per Member</div></div>
    <div class="panel-body"><div class="chart-wrap" style="height:420px"><canvas id="an-member-chart" role="img" aria-label="Member savings bar chart"></canvas></div></div>
  </div>
  <div class="grid-equal mb-24">
    <div class="panel">
      <div class="panel-head"><div class="panel-title">Average Weekly Contribution</div></div>
      <div class="panel-body"><div class="chart-wrap" style="height:250px"><canvas id="an-avg-chart" role="img" aria-label="Average weekly savings per member"></canvas></div></div>
    </div>
    <div class="panel">
      <div class="panel-head"><div class="panel-title">Weekly Group Totals</div></div>
      <div class="panel-body"><div class="chart-wrap" style="height:250px"><canvas id="an-weekly-chart" role="img" aria-label="Group weekly savings totals"></canvas></div></div>
    </div>
  </div>`;

  destroyChart('an-member'); destroyChart('an-avg'); destroyChart('an-weekly');

  charts['an-member'] = new Chart($('an-member-chart'),{
    type:'bar',
    data:{labels:sorted.map(m=>m.name.split(' ').slice(0,2).join(' ')),
      datasets:[{data:sorted.map(m=>m.total),backgroundColor:'rgba(79,142,255,.7)',borderRadius:4,borderWidth:0}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#8892aa',font:{size:11},callback:v=>fmt(v)}},
        y:{grid:{display:false},ticks:{color:'#dde3f0',font:{size:11}}}}}
  });

  const avgs = data.map(m=>{ const nz=m.weeks.filter(x=>x>0); return nz.length?Math.round(nz.reduce((a,b)=>a+b,0)/nz.length):0; });
  charts['an-avg'] = new Chart($('an-avg-chart'),{
    type:'bar',
    data:{labels:data.map(m=>m.name.split(' ')[0]),
      datasets:[{data:avgs,backgroundColor:'rgba(0,201,167,.65)',borderRadius:3,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{grid:{display:false},ticks:{color:'#8892aa',font:{size:10},maxRotation:55}},
        y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#8892aa',font:{size:11}}}}}
  });

  const weekCount = Math.max(6, ...data.map(m => m.weeks.length));
  const weekTotals = Array.from({ length: weekCount }, (_, w) => data.reduce((s,m)=>s+(m.weeks[w]||0),0));
  charts['an-weekly'] = new Chart($('an-weekly-chart'),{
    type:'bar',
    data:{labels:weekLabels(weekCount),
      datasets:[{data:weekTotals,backgroundColor:'rgba(255,209,102,.7)',borderRadius:4,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{grid:{display:false},ticks:{color:'#8892aa',font:{size:11}}},
        y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#8892aa',font:{size:11},callback:v=>fmt(v)}}}}
  });
}

// ── MEMBER DASHBOARD ─────────────────────────────────────────────────────────

async function loadMyDashboard() {
  const pg = $('page-my-dashboard');
  pg.innerHTML = '<div class="loading"><div class="spinner"></div>Loading…</div>';
  const d = await api('/api/my-dashboard');
  const m = d.member;
  const dynamicShares = Math.floor((d.totalSaved || 0) / 500);
  const weeks = d.weeks;
  const wkBoxes = weeks.map((amt,i)=>`
    <div class="week-box ${amt>0?'paid':'miss'}">
      <div class="week-num">Week ${i+1}</div>
      <div class="week-amt">${amt>0?fmt(amt)+'':'—'}</div>
      <div class="week-ico">${amt>0?'✓':'✗'}</div>
    </div>`).join('');
  const loanHtml = d.loan ? `
    <div class="panel mt-20" style="margin-top:20px">
      <div class="panel-head"><div class="panel-title">Your Active Loan</div><span class="pill pill-red">Outstanding</span></div>
      <div class="panel-body">
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px">
          <div><div style="font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:var(--text3)">Principal</div><div style="font-size:18px;font-weight:600;margin-top:4px">${fmt(d.loan.amount)} RWF</div></div>
          <div><div style="font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:var(--text3)">Interest</div><div style="font-size:18px;font-weight:600;margin-top:4px">${fmt(d.loan.interest)} RWF</div></div>
          <div><div style="font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:var(--text3)">Total Due</div><div style="font-size:18px;font-weight:600;margin-top:4px;color:var(--amber)">${fmt(d.loan.amount+d.loan.interest)} RWF</div></div>
          <div><div style="font-size:10px;text-transform:uppercase;letter-spacing:.07em;color:var(--text3)">Repaid</div><div style="font-size:18px;font-weight:600;margin-top:4px;color:var(--text3)">0 RWF</div></div>
        </div>
        <div class="progress-label"><span>Repayment progress</span><span>0%</span></div>
        <div class="progress-track"><div class="progress-fill" style="width:0%"></div></div>
      </div>
    </div>` : `<div class="panel" style="margin-top:20px;padding:20px;text-align:center;color:var(--text2)"><p>✓ You have no active loans.</p></div>`;

  pg.innerHTML = `
  <div class="page-header"><div class="page-title">My Dashboard</div><div class="page-sub">Personal overview</div></div>
  <div class="my-hero">
    <div class="my-hero-avatar">${initials(m.name)}</div>
    <div>
      <div class="my-hero-name">${m.name}</div>
      <div class="my-hero-meta">Member ID: ${m.id} · ${dynamicShares} share${dynamicShares>1?'s':''} · ${m.phone}</div>
    </div>
  </div>
  <div class="my-stat-grid">
    <div class="stat-card green"><div class="stat-label">My Total Savings</div><div class="stat-val">${fmt(d.totalSaved)}</div><div class="stat-sub">RWF contributed</div></div>
    <div class="stat-card blue"><div class="stat-label">My Shares</div><div class="stat-val">${dynamicShares}</div><div class="stat-sub">out of ${d.totalShares} total</div></div>
    <div class="stat-card amber"><div class="stat-label">Group Total</div><div class="stat-val">${fmt(d.groupTotal)}</div><div class="stat-sub">RWF saved collectively</div></div>
  </div>
  <div class="panel mb-24">
    <div class="panel-head"><div class="panel-title">Weekly Savings</div></div>
    <div class="panel-body"><div class="week-grid">${wkBoxes}</div></div>
  </div>
  ${loanHtml}`;
}

async function loadMySavings() {
  const d = await api('/api/my-dashboard');
  const pg = $('page-my-savings');
  const weeks = d.weeks;
  const total = d.totalSaved;
  const wkBoxes = weeks.map((amt,i)=>`
    <div class="week-box ${amt>0?'paid':'miss'}">
      <div class="week-num">Week ${i+1}</div>
      <div class="week-amt">${amt>0?fmt(amt):'—'}</div>
      <div class="week-ico">${amt>0?'✓':'✗'}</div>
    </div>`).join('');
  pg.innerHTML = `
  <div class="page-header"><div class="page-title">My Savings</div><div class="page-sub">Your contribution record</div></div>
  <div class="stat-grid" style="grid-template-columns:repeat(3,1fr)">
    <div class="stat-card green"><div class="stat-label">Total Saved</div><div class="stat-val">${fmt(total)}</div><div class="stat-sub">RWF</div></div>
    <div class="stat-card blue"><div class="stat-label">Weeks Active</div><div class="stat-val">${weeks.filter(x=>x>0).length}</div><div class="stat-sub">of ${weeks.length} weeks</div></div>
    <div class="stat-card amber"><div class="stat-label">Weekly Average</div><div class="stat-val">${fmt(Math.round(weeks.filter(x=>x>0).reduce((a,b)=>a+b,0)/Math.max(1,weeks.filter(x=>x>0).length)))}</div><div class="stat-sub">RWF per active week</div></div>
  </div>
  <div class="panel mb-24">
    <div class="panel-head"><div class="panel-title">Record Your Savings</div><div class="panel-sub">Enter your weekly contribution</div></div>
    <div class="panel-body">
      <div class="form-grid compact">
        <div class="field"><label>Week</label><input id="my-saving-week" type="number" min="1" value="${nextSavingsWeek(weeks)}"/></div>
        <div class="field"><label>Amount</label><input id="my-saving-amount" type="number" min="0" step="1" placeholder="Enter amount" onfocus="this.select()"/></div>
      </div>
      <div class="action-row"><button class="btn-small primary" onclick="saveMySaving()">Submit Saving</button></div>
    </div>
  </div>
  <div class="panel mb-24">
    <div class="panel-head"><div class="panel-title">Weekly Breakdown</div></div>
    <div class="panel-body"><div class="week-grid">${wkBoxes}</div></div>
  </div>
  <div class="panel">
    <div class="panel-head"><div class="panel-title">Savings Chart</div></div>
    <div class="panel-body"><div class="chart-wrap" style="height:200px"><canvas id="my-savings-chart" role="img" aria-label="Your weekly savings chart"></canvas></div></div>
  </div>`;

  const myWeekLabels = weekLabels(weeks.length);
  destroyChart('my-savings');
  charts['my-savings'] = new Chart($('my-savings-chart'),{
    type:'bar',
    data:{labels:myWeekLabels,
      datasets:[{data:weeks,backgroundColor:weeks.map(x=>x>0?'rgba(0,201,167,.7)':'rgba(74,83,112,.4)'),borderRadius:4,borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{grid:{display:false},ticks:{color:'#8892aa',font:{size:11}}},
        y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#8892aa',font:{size:11}}}}}
  });
}

async function saveMySaving() {
  const week = Number($('my-saving-week').value);
  const amount = parseAmountInput($('my-saving-amount').value);
  if (!Number.isInteger(week) || week < 1) return toast('Week must be 1 or higher', 'error');
  if (!Number.isFinite(amount) || amount <= 0) return toast('Amount must be a positive number', 'error');
  const body = { week, amount };
  try {
    await api('/api/my-savings', { method:'PUT', body });
    toast('Savings recorded');
    await loadMySavings();
  } catch(e) {
    toast(e.message, 'error');
  }
}

async function loadMyLoan() {
  const d = await api('/api/my-dashboard');
  const pg = $('page-my-loan');
  if (!d.loan) {
    pg.innerHTML = `
    <div class="page-header"><div class="page-title">My Loan</div></div>
    <div class="panel" style="padding:40px;text-align:center;color:var(--text2)">
      <div style="font-size:32px;margin-bottom:12px">✓</div>
      <div style="font-size:16px;font-weight:500;margin-bottom:6px">No active loan</div>
      <div style="font-size:13px">You currently have no outstanding loan with the group.</div>
    </div>`;
    return;
  }
  const l = d.loan;
  const total = l.amount + l.interest;
  const pct = Math.round((l.repaid/total)*100);
  pg.innerHTML = `
  <div class="page-header"><div class="page-title">My Loan</div><div class="page-sub">Your active loan details</div></div>
  <div class="loan-item">
    <div class="loan-header">
      <div><div class="loan-name">${l.name}</div><div class="loan-id">Loan ID: ${l.id}</div></div>
      <span class="pill pill-red">Outstanding</span>
    </div>
    <div class="loan-meta">
      <div class="lm"><label>Principal</label><span>${fmt(l.amount)} RWF</span></div>
      <div class="lm"><label>Interest (10%)</label><span>${fmt(l.interest)} RWF</span></div>
      <div class="lm total"><label>Total Due</label><span>${fmt(total)} RWF</span></div>
      <div class="lm"><label>Amount Repaid</label><span style="color:var(--text3)">0 RWF</span></div>
    </div>
    <div class="progress-label"><span>Repayment progress: ${pct}%</span><span>${fmt(l.repaid)} / ${fmt(total)} RWF</span></div>
    <div class="progress-track"><div class="progress-fill" style="width:${pct}%"></div></div>
  </div>
  <div class="panel" style="margin-top:16px;padding:16px 20px;font-size:13px;color:var(--text2)">
    <strong style="color:var(--amber)">Reminder:</strong> Please ensure timely repayment to maintain good standing. Contact the group admin (Eric) for any queries.
  </div>`;
}

// ── CHANGE PASSWORD ──────────────────────────────────────────────────────────

function renderChangePass() {
  const pg = $('page-change-pass');
  pg.innerHTML = `
  <div class="page-header"><div class="page-title">Change Password</div><div class="page-sub">Update your login credentials</div></div>
  <div class="change-pass-form">
    <div class="panel" style="max-width:400px">
      <div class="panel-head"><div class="panel-title">New Password</div></div>
      <div class="panel-body">
        <div class="field"><label>Current Password</label><input type="password" id="cp-old" placeholder="Your current password"/></div>
        <div class="field"><label>New Password</label><input type="password" id="cp-new" placeholder="Min. 6 characters"/></div>
        <div class="field"><label>Confirm New Password</label><input type="password" id="cp-confirm" placeholder="Repeat new password"/></div>
        <button class="btn-primary" onclick="doChangePass()">Update Password</button>
      </div>
    </div>
  </div>`;
}

async function doChangePass() {
  const oldPass = $('cp-old').value;
  const newPass = $('cp-new').value;
  const confirm = $('cp-confirm').value;
  if (!oldPass||!newPass||!confirm) { toast('Please fill all fields','error'); return; }
  if (newPass !== confirm) { toast('Passwords do not match','error'); return; }
  if (newPass.length < 6) { toast('Password must be at least 6 characters','error'); return; }
  try {
    await api('/api/change-password', {
      method: 'POST',
      body: { currentPassword: oldPass, newPassword: newPass }
    });
    toast('Password updated successfully','success');
    $('cp-old').value=''; $('cp-new').value=''; $('cp-confirm').value='';
  } catch(e) {
    toast(e.message, 'error');
  }
  return;
}

// ── BOOT ─────────────────────────────────────────────────────────────────────

async function boot() {
  try {
    const user = await api('/api/me');
    currentUser = user;
    renderShell();
    if (user.role==='admin') loadAdminDashboard();
    else loadMyDashboard();
  } catch {
    renderLogin();
  }
}

boot();
