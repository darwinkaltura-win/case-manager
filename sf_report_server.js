// SF Team Report Server — run with:
// "C:\Program Files\sf\client\bin\node.exe" sf_report_server.js

const http = require('http');
const { execFile, exec } = require('child_process');
const url  = require('url');
const fs   = require('fs');
const os   = require('os');
const path = require('path');

const PORT    = 3737;
const NODE    = process.execPath;
const PS_SEND = path.join(__dirname, 'send_sf_report.ps1');
const TASK_NAME = 'SF_Team_Report';

const SF_ORG = 'kaltura';
const TEAM_NAMES = [
  'Russ Lichterman','Darwin Mitra','Fahad Mizi','Alex De Los Santos',
  'Tahmid Hassan','Roxy Hennessy','Rick Rehmann','Zach Hill',
  'Oscar Lagua Espin','Hector Zurita','Agustin Herling','Stivan Tenev',
  'Asad Ali','Julian Lucena Herrera','Renato Pinheiro'
];
const DISPLAY = {
  'Oscar Lagua Espin': 'Oscar Espin',
  'Julian Lucena Herrera': 'Julian Herrera'
};
const OPEN_STATUSES = [
  'New','In Progress','In Work','Awaiting Customer Response','Awaiting CSM',
  'Awaiting Tier 3','Awaiting ADA Team','Awaiting PS','Awaiting Product',
  'Awaiting Vendor','Awaiting R&D','Awaiting Internal','Awaiting FR Review',
  'Awaiting Deployment','Awaiting DevOps','Awaiting Owner Response',
  'Awaiting Response','Awaiting Internal Email','Review Customer Response',
  'Review Internal','Review JIRA Response','Review Akamai Response',
  'Review Customer Response (Reopened)','On Hold','FR In Review',
  'Customer Responded','Will be closed in 48H','Submitted','Resource Requested',
  'Resolved','Solution Provided to Customer',
  'Recommend to Close - Solution Provided',
  'Recommend to Close - No Longer Needed','Approved by Manager'
];
const ACTIONABLE = new Set([
  'New','In Progress','In Work','Customer Responded',
  'Review Customer Response','Review Customer Response (Reopened)',
  'Review Internal','Review JIRA Response','Review Akamai Response',
  'Resolved','FR In Review','Will be closed in 48H',
  'Recommend to Close - Solution Provided',
  'Recommend to Close - No Longer Needed','Approved by Manager'
]);

// ── Salesforce query ─────────────────────────────────────────────────────────
const SF_CMD = 'C:\\Program Files\\sf\\bin\\sf.cmd';

function querySF() {
  return new Promise((resolve, reject) => {
    const nameClause   = TEAM_NAMES.map(n => `'${n}'`).join(',');
    const statusClause = OPEN_STATUSES.map(s => `'${s}'`).join(',');
    const soql = `SELECT CaseNumber, Subject, Status, Priority, Owner.Name, IsEscalated, FLAGS__Case_Flags_Sort__c FROM Case WHERE Owner.Name IN (${nameClause}) AND Status IN (${statusClause}) ORDER BY Owner.Name, CreatedDate DESC LIMIT 2000`;

    // Write SOQL to temp file to avoid shell quoting issues
    const tmpSoql = path.join(os.tmpdir(), 'sf_report_query.soql');
    const tmpCsv  = path.join(os.tmpdir(), 'sf_report_out.csv');
    fs.writeFileSync(tmpSoql, soql, 'utf8');

    execFile('powershell', [
      '-Command',
      `& '${SF_CMD}' data query --target-org ${SF_ORG} --result-format csv --file '${tmpSoql}' --output-file '${tmpCsv}'; exit 0`
    ], { maxBuffer: 20 * 1024 * 1024, timeout: 90000 }, (err, stdout, stderr) => {
      try {
        const raw = fs.readFileSync(tmpCsv, 'utf8');
        const lines = raw.split(/\r?\n/).filter(l => /^\d{8},/.test(l) || /^CaseNumber/.test(l));
        const rows = parseCSV(lines);
        resolve(buildData(rows));
      } catch(e) {
        reject(new Error(stderr || (err && err.message) || e.message));
      }
    });
  });
}

// ── CSV parser ───────────────────────────────────────────────────────────────
function parseCSV(lines) {
  if (!lines.length) return [];
  const headers = lines[0].split(',');
  return lines.slice(1).map(line => {
    const vals = splitCSVLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h.trim()] = (vals[i] || '').trim(); });
    return obj;
  }).filter(r => /^\d{8}/.test(r.CaseNumber));
}

function splitCSVLine(line) {
  const result = []; let cur = ''; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"' && !inQ) { inQ = true; continue; }
    if (c === '"' && inQ) { if (line[i+1] === '"') { cur += '"'; i++; } else { inQ = false; } continue; }
    if (c === ',' && !inQ) { result.push(cur); cur = ''; continue; }
    cur += c;
  }
  result.push(cur);
  return result;
}

// ── Case Handling query ───────────────────────────────────────────────────────
function queryCaseHandling() {
  return new Promise((resolve, reject) => {
    const nameClause = TEAM_NAMES.map(n => `'${n}'`).join(',');
    const soql = `SELECT CreatedBy.Name, Parent.CaseNumber, ParentId, CreatedDate FROM CaseComment WHERE CreatedBy.Name IN (${nameClause}) AND CreatedDate = LAST_N_DAYS:35 ORDER BY CreatedDate DESC LIMIT 5000`;
    const tmpSoql = path.join(os.tmpdir(), 'sf_ch_query.soql');
    const tmpCsv  = path.join(os.tmpdir(), 'sf_ch_out.csv');
    fs.writeFileSync(tmpSoql, soql, 'utf8');
    execFile('powershell', [
      '-Command',
      `& '${SF_CMD}' data query --target-org ${SF_ORG} --result-format csv --file '${tmpSoql}' --output-file '${tmpCsv}'; exit 0`
    ], { maxBuffer: 20 * 1024 * 1024, timeout: 90000 }, (err, stdout, stderr) => {
      try {
        const raw  = fs.readFileSync(tmpCsv, 'utf8');
        const lines = raw.split(/\r?\n/).filter(Boolean);
        resolve(buildCaseHandlingData(parseCSVGeneric(lines)));
      } catch(e) {
        reject(new Error(stderr || (err && err.message) || e.message));
      }
    });
  });
}

function parseCSVGeneric(lines) {
  if (!lines.length) return [];
  const headers = splitCSVLine(lines[0]);
  return lines.slice(1).map(line => {
    const vals = splitCSVLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h.trim()] = (vals[i] || '').trim(); });
    return obj;
  }).filter(r => Object.values(r).some(v => v));
}

function buildCaseHandlingData(rows) {
  const display = name => DISPLAY[name] || name;
  const now = new Date();
  const todayStr = now.toISOString().slice(0, 10);
  const weekStart = new Date(now);
  weekStart.setDate(now.getDate() - ((now.getDay() + 6) % 7));
  weekStart.setHours(0, 0, 0, 0);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  const byPerson = {};
  TEAM_NAMES.forEach(n => { byPerson[n] = { todayCases: new Set(), weekCases: new Set(), monthCases: new Set() }; });

  const dailyCases = {}; // date -> sfName -> Set<caseNum>

  rows.forEach(r => {
    const sfName = r['CreatedBy.Name'];
    if (!byPerson[sfName]) return;
    const caseNum = r['Parent.CaseNumber'] || r.ParentId;
    if (!caseNum) return;
    const d = new Date(r.CreatedDate);
    const dateStr = d.toISOString().slice(0, 10);
    if (dateStr === todayStr) byPerson[sfName].todayCases.add(caseNum);
    if (d >= weekStart)  byPerson[sfName].weekCases.add(caseNum);
    if (d >= monthStart) byPerson[sfName].monthCases.add(caseNum);
    // daily
    if (!dailyCases[dateStr]) dailyCases[dateStr] = {};
    if (!dailyCases[dateStr][sfName]) dailyCases[dateStr][sfName] = new Set();
    dailyCases[dateStr][sfName].add(caseNum);
  });

  // Build last 30 days array
  const daily = [];
  for (let i = 29; i >= 0; i--) {
    const d = new Date(now); d.setDate(now.getDate() - i);
    const dateStr = d.toISOString().slice(0, 10);
    const byName = {};
    TEAM_NAMES.forEach(sfName => {
      byName[display(sfName)] = dailyCases[dateStr]?.[sfName]?.size || 0;
    });
    daily.push({ date: dateStr, byName });
  }

  const persons = TEAM_NAMES.map(sfName => ({
    name: display(sfName), sfName,
    today: byPerson[sfName].todayCases.size,
    week:  byPerson[sfName].weekCases.size,
    month: byPerson[sfName].monthCases.size,
    todayCases: [...byPerson[sfName].todayCases],
    weekCases:  [...byPerson[sfName].weekCases],
    monthCases: [...byPerson[sfName].monthCases]
  }));
  return { persons, daily, fetchedAt: new Date().toISOString() };
}

// ── JPMC Cases query ─────────────────────────────────────────────────────────
function queryJpmcCases() {
  return new Promise((resolve, reject) => {
    const soql = `SELECT Id, CaseNumber, Subject, Status, Priority, Assigned_To__c, Assigned_To__r.Name, Contact.Name, Contact.FirstName, Contact.Email, Description, CreatedDate FROM Case WHERE Status = 'New' AND Assigned_To__c = null AND Subject LIKE '%Video recovery request%' AND Account.Name = 'J.P. Morgan Chase & Co.' ORDER BY CreatedDate DESC LIMIT 100`;
    const tmpSoql = path.join(os.tmpdir(), 'sf_jpmc_query.soql');
    const tmpCsv  = path.join(os.tmpdir(), 'sf_jpmc_out.csv');
    fs.writeFileSync(tmpSoql, soql, 'utf8');
    execFile('powershell', [
      '-Command',
      `& '${SF_CMD}' data query --target-org ${SF_ORG} --result-format csv --file '${tmpSoql}' --output-file '${tmpCsv}'; exit 0`
    ], { maxBuffer: 10 * 1024 * 1024, timeout: 60000 }, (err, stdout, stderr) => {
      try {
        const raw = fs.readFileSync(tmpCsv, 'utf8');
        const lines = raw.split(/\r?\n/).filter(Boolean);
        resolve(parseCSVGeneric(lines));
      } catch(e) {
        reject(new Error(stderr || (err && err.message) || e.message));
      }
    });
  });
}

// ── JPMC Restore Stats query (YTD) ───────────────────────────────────────────
function queryJpmcStats() {
  // Run both queries in parallel
  const queryRestore = () => new Promise((resolve, reject) => {
    const soql = `SELECT Id, CaseNumber, Status, CreatedDate FROM Case WHERE Subject LIKE '%Video recovery request%' AND Account.Name = 'J.P. Morgan Chase & Co.' AND CreatedDate >= 2026-01-01T00:00:00Z ORDER BY CreatedDate ASC LIMIT 2000`;
    const tmpSoql = path.join(os.tmpdir(), 'sf_jpmc_stats.soql');
    const tmpJson = path.join(os.tmpdir(), 'sf_jpmc_stats.json');
    fs.writeFileSync(tmpSoql, soql, 'utf8');
    execFile('powershell', ['-Command',
      `& '${SF_CMD}' data query --target-org ${SF_ORG} --result-format json --file '${tmpSoql}' --output-file '${tmpJson}'; exit 0`
    ], { maxBuffer: 10 * 1024 * 1024, timeout: 60000 }, (err, stdout, stderr) => {
      try {
        const raw = JSON.parse(fs.readFileSync(tmpJson, 'utf8'));
        resolve((raw.result || raw).records || []);
      } catch(e) { reject(new Error(stderr || (err && err.message) || e.message)); }
    });
  });

  const queryTotal = () => new Promise((resolve, reject) => {
    const soql = `SELECT COUNT(Id) total FROM Case WHERE Account.Name = 'J.P. Morgan Chase & Co.' AND CreatedDate >= 2026-01-01T00:00:00Z`;
    const tmpSoql = path.join(os.tmpdir(), 'sf_jpmc_total.soql');
    const tmpJson = path.join(os.tmpdir(), 'sf_jpmc_total.json');
    fs.writeFileSync(tmpSoql, soql, 'utf8');
    execFile('powershell', ['-Command',
      `& '${SF_CMD}' data query --target-org ${SF_ORG} --result-format json --file '${tmpSoql}' --output-file '${tmpJson}'; exit 0`
    ], { maxBuffer: 1 * 1024 * 1024, timeout: 30000 }, (err, stdout, stderr) => {
      try {
        const raw = JSON.parse(fs.readFileSync(tmpJson, 'utf8'));
        const recs = (raw.result || raw).records || [];
        resolve(recs[0] ? (recs[0].total || recs[0].expr0 || 0) : 0);
      } catch(e) { resolve(0); }
    });
  });

  return Promise.all([queryRestore(), queryTotal()]).then(([records, totalAll]) => {
    const daily = {}, weekly = {}, monthly = {};
    records.forEach(r => {
      const d = new Date(r.CreatedDate);
      const dayKey   = d.toISOString().slice(0, 10);
      const monthKey = d.toISOString().slice(0, 7);
      const jan1 = new Date(d.getFullYear(), 0, 1);
      const week = Math.ceil(((d - jan1) / 86400000 + jan1.getDay() + 1) / 7);
      const weekKey = d.getFullYear() + '-W' + String(week).padStart(2, '0');
      daily[dayKey]     = (daily[dayKey]     || 0) + 1;
      weekly[weekKey]   = (weekly[weekKey]   || 0) + 1;
      monthly[monthKey] = (monthly[monthKey] || 0) + 1;
    });
    return { total: records.length, totalAll: Number(totalAll), daily, weekly, monthly };
  });
}

// ── JPMC All Open Cases query (all statuses, assigned + unassigned) ──────────
function queryJpmcNewCases() {
  return new Promise((resolve, reject) => {
    const soql = `SELECT Id, CaseNumber, Subject, Status, Priority, Assigned_To__c, Assigned_To__r.Name, OwnerId, Owner.Name, Contact.Name, Contact.FirstName, Contact.Email, Description, CreatedDate FROM Case WHERE IsClosed = false AND Subject LIKE '%Video recovery request%' AND Account.Name = 'J.P. Morgan Chase & Co.' ORDER BY CreatedDate DESC LIMIT 200`;
    const tmpSoql = path.join(os.tmpdir(), 'sf_jpmc_new_query.soql');
    const tmpJson = path.join(os.tmpdir(), 'sf_jpmc_new_out.json');
    fs.writeFileSync(tmpSoql, soql, 'utf8');
    execFile('powershell', [
      '-Command',
      `& '${SF_CMD}' data query --target-org ${SF_ORG} --result-format json --file '${tmpSoql}' --output-file '${tmpJson}'; exit 0`
    ], { maxBuffer: 10 * 1024 * 1024, timeout: 60000 }, (err, stdout, stderr) => {
      try {
        const raw = JSON.parse(fs.readFileSync(tmpJson, 'utf8'));
        const records = (raw.result || raw).records || raw.result || [];
        // Flatten relationship fields
        const flat = records.map(r => {
          const out = { ...r };
          if (r.Contact) {
            out['Contact.Name']      = r.Contact.Name      || '';
            out['Contact.FirstName'] = r.Contact.FirstName || '';
            out['Contact.Email']     = r.Contact.Email     || '';
          }
          if (r.Assigned_To__r) out['Assigned_To__r.Name'] = r.Assigned_To__r.Name || '';
          if (r.Owner)          out['Owner.Name']           = r.Owner.Name          || '';
          return out;
        });
        resolve(flat);
      } catch(e) {
        reject(new Error(stderr || (err && err.message) || e.message));
      }
    });
  });
}

// ── Get current SF logged-in user ─────────────────────────────────────────────
let _sfCurrentUser = null;
function getSFCurrentUser() {
  if (_sfCurrentUser) return Promise.resolve(_sfCurrentUser);
  return new Promise((resolve) => {
    // Step 1: get username from org display
    execFile('powershell', ['-Command',
      `& '${SF_CMD}' org display --target-org ${SF_ORG} --json; exit 0`
    ], { timeout: 15000 }, (err, stdout) => {
      try {
        const orgInfo = JSON.parse(stdout);
        const username = orgInfo.result.username;
        const token    = orgInfo.result.accessToken;
        const baseUrl  = orgInfo.result.instanceUrl;
        const soql = `SELECT Id,Name,Title FROM User WHERE Username = '${username}' LIMIT 1`;
        const tmpSoql = path.join(os.tmpdir(), 'sf_curuser.soql');
        const tmpJson = path.join(os.tmpdir(), 'sf_curuser.json');
        fs.writeFileSync(tmpSoql, soql, 'utf8');
        execFile('powershell', ['-Command',
          `& '${SF_CMD}' data query --target-org ${SF_ORG} --result-format json --file '${tmpSoql}' --output-file '${tmpJson}'; exit 0`
        ], { timeout: 15000 }, (e2) => {
          try {
            const raw = JSON.parse(fs.readFileSync(tmpJson, 'utf8'));
            const rec = ((raw.result || raw).records || [])[0] || {};
            _sfCurrentUser = { name: rec.Name || username, title: rec.Title || '' };
          } catch { _sfCurrentUser = { name: username, title: '' }; }
          resolve(_sfCurrentUser);
        });
      } catch { resolve({ name: 'Kaltura Support', title: '' }); }
    });
  });
}

// ── Get fresh SF token ────────────────────────────────────────────────────────
function getFreshSFToken() {
  return new Promise((resolve, reject) => {
    execFile('powershell', ['-Command',
      `& '${SF_CMD}' org display --target-org ${SF_ORG} --json; exit 0`
    ], { timeout: 15000 }, (err, stdout, stderr) => {
      try {
        const info = JSON.parse(stdout);
        resolve({ token: info.result.accessToken, instanceUrl: info.result.instanceUrl });
      } catch(e) {
        reject(new Error('Could not get SF token: ' + (stderr || e.message)));
      }
    });
  });
}

// ── Respond to JPMC case (post comment + close) ───────────────────────────────
async function respondToCase(caseId, firstName, commentBodyOverride) {
  const [sfUser, sfOrg] = await Promise.all([getSFCurrentUser(), getFreshSFToken()]);
  return new Promise((resolve, reject) => {
    let commentBody = commentBodyOverride;
    if (!commentBody) {
      const salutation = firstName ? firstName : 'there';
      const sigName = sfUser.name || 'Kaltura Support';
      commentBody = `Hi ${salutation},\n\nThanks for reaching out to Kaltura Customer Care.\n\nI'm happy to confirm that the requested entries have been successfully restored. Please check on your end and let us know if everything looks good.\n\nPlease note that restore requests are handled on a best effort basis.\n\nI will now be marking the case as closed.\n\nShould you notice anything else or need further assistance, feel free to reach out.\n\nBest regards,\n\n${sigName}\nKaltura Customer Care | Kaltura Inc.\nSupport: https://support.kaltura.com\n\nKnowledge Base: https://knowledge.kaltura.com\nWebsite: https://www.kaltura.com\nStatus Alerts: https://status.kaltura.com\n\nGet your support questions answered before login \u2014 try our AI Support Assistant in the bottom left corner!\n\nThe age of Agentic Avatars is here: https://corp.kaltura.com/agentic-avatars/`;
    }

    const token = sfOrg.token;
    const instanceUrl = sfOrg.instanceUrl;

    const commentPayload = JSON.stringify({ ParentId: caseId, CommentBody: commentBody, IsPublished: true });
    const closePayload = JSON.stringify({ Status: 'Closed' });

    const postComment = () => new Promise((res2, rej2) => {
      execFile('powershell', ['-Command', `
$headers = @{ Authorization = 'Bearer ${token}'; 'Content-Type' = 'application/json' }
$body = '${commentPayload.replace(/'/g, "''").replace(/\n/g, '\\n')}'
$r = Invoke-RestMethod -Uri '${instanceUrl}/services/data/v66.0/sobjects/CaseComment' -Method POST -Headers $headers -Body $body -ContentType 'application/json'
Write-Output $r.id
`], { timeout: 30000 }, (err, stdout, stderr) => {
        if (err) return rej2(new Error(stderr || err.message));
        res2(stdout.trim());
      });
    });

    const closeCase = (resolvedAt) => new Promise((res2, rej2) => {
      const closePayload = JSON.stringify({
        Status: 'Closed - Other',
        Closed_Reason__c: 'Issue on Customer Side',
        Root_Cause_OVP__c: 'Configuration Issue (Customer Side)',
        Solution__c: 'CC - Manual task /config change',
        Resolution_Provided__c: true,
        Resolution_End_Date_Time__c: resolvedAt
      }).replace(/'/g, "''");
      execFile('powershell', ['-Command', `
$headers = @{ Authorization = 'Bearer ${token}'; 'Content-Type' = 'application/json' }
$body = '${closePayload}'
Invoke-RestMethod -Uri '${instanceUrl}/services/data/v66.0/sobjects/Case/${caseId}' -Method PATCH -Headers $headers -Body $body -ContentType 'application/json'
Write-Output 'ok'
`], { timeout: 30000 }, (err, stdout, stderr) => {
        if (err) return rej2(new Error(stderr || err.message));
        res2('ok');
      });
    });

    const resolvedAt = new Date().toISOString();
    postComment()
      .then(() => closeCase(resolvedAt))
      .then(() => resolve('ok'))
      .catch(reject);
  });
}

let _teamUserIds = null;
function getTeamUserIds() {
  if (_teamUserIds) return Promise.resolve(_teamUserIds);
  return new Promise((resolve, reject) => {
    const names = TEAM_NAMES.map(n => `'${n}'`).join(',');
    const soql = `SELECT Id, Name FROM User WHERE Name IN (${names}) AND IsActive = true ORDER BY Name`;
    const tmpSoql = path.join(os.tmpdir(), 'sf_users_query.soql');
    const tmpCsv  = path.join(os.tmpdir(), 'sf_users_out.csv');
    fs.writeFileSync(tmpSoql, soql, 'utf8');
    execFile('powershell', [
      '-Command',
      `& '${SF_CMD}' data query --target-org ${SF_ORG} --result-format csv --file '${tmpSoql}' --output-file '${tmpCsv}'; exit 0`
    ], { maxBuffer: 5 * 1024 * 1024, timeout: 30000 }, (err, stdout, stderr) => {
      try {
        const raw = fs.readFileSync(tmpCsv, 'utf8');
        const lines = raw.split(/\r?\n/).filter(Boolean);
        const rows = parseCSVGeneric(lines);
        _teamUserIds = {};
        rows.forEach(r => { _teamUserIds[r.Name] = r.Id; });
        resolve(_teamUserIds);
      } catch(e) {
        reject(new Error(stderr || (err && err.message) || e.message));
      }
    });
  });
}

function assignCaseAssignedTo(caseId, userId) {
  return new Promise((resolve, reject) => {
    execFile('powershell', [
      '-Command',
      `& '${SF_CMD}' data update record --sobject Case --record-id ${caseId} --values "Assigned_To__c=${userId}" --target-org ${SF_ORG}; exit 0`
    ], { maxBuffer: 1024 * 1024, timeout: 30000 }, (err, stdout, stderr) => {
      const out = (stdout + stderr).toLowerCase();
      if (out.includes('successfully updated')) resolve('ok');
      else reject(new Error(stderr || (err && err.message) || 'Update failed'));
    });
  });
}

// ── Build structured data ────────────────────────────────────────────────────
function buildData(rows) {
  const display = name => DISPLAY[name] || name;

  const summary = TEAM_NAMES.map(sfName => {
    const cases = rows.filter(r => r['Owner.Name'] === sfName);
    const esc   = cases.filter(r => r.IsEscalated === 'true');
    const act   = cases.filter(r => ACTIONABLE.has(r.Status));
    const escAct= cases.filter(r => r.IsEscalated === 'true' && ACTIONABLE.has(r.Status));
    const bf    = cases.filter(r => r.FLAGS__Case_Flags_Sort__c && r.FLAGS__Case_Flags_Sort__c.startsWith('L4'));
    return {
      name: display(sfName), sfName,
      open: cases.length, escalated: esc.length,
      actionable: act.length, escActionable: escAct.length,
      blackFlags: bf.length
    };
  });

  const escActionableCases = rows
    .filter(r => r.IsEscalated === 'true' && ACTIONABLE.has(r.Status))
    .map(r => ({ ...r, ownerDisplay: display(r['Owner.Name']) }))
    .sort((a, b) => a.ownerDisplay.localeCompare(b.ownerDisplay));

  const blackFlagCases = rows
    .filter(r => r.FLAGS__Case_Flags_Sort__c && r.FLAGS__Case_Flags_Sort__c.startsWith('L4'))
    .map(r => ({ ...r, ownerDisplay: display(r['Owner.Name']) }))
    .sort((a, b) => a.ownerDisplay.localeCompare(b.ownerDisplay));

  return { summary, escActionableCases, blackFlagCases, fetchedAt: new Date().toISOString() };
}

// ── Server-side data cache ────────────────────────────────────────────────────
let lastData = null;

// ── Build email HTML from cached data ────────────────────────────────────────
function buildEmailHtml(data) {
  const { summary, escActionableCases, blackFlagCases, fetchedAt } = data;
  const runTime = new Date(fetchedAt).toLocaleString();
  const totOpen = summary.reduce((s,r)=>s+r.open,0);
  const totEsc  = summary.reduce((s,r)=>s+r.escalated,0);
  const totAct  = summary.reduce((s,r)=>s+r.actionable,0);
  const totEscA = summary.reduce((s,r)=>s+r.escActionable,0);
  const totBF   = summary.reduce((s,r)=>s+r.blackFlags,0);
  const e = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const n = v => v > 0 ? `<b>${v}</b>` : `<span style="color:#aaa">0</span>`;

  const css = `body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#222}
h2{color:#0078d4;border-bottom:2px solid #0078d4;padding-bottom:4px;margin:0 0 4px}
h3{color:#444;margin:24px 0 8px;font-size:13px}
table{border-collapse:collapse;margin-bottom:20px;min-width:480px}
th{background:#0078d4;color:#fff;padding:6px 12px;text-align:left;font-size:11px}
td{padding:5px 12px;border-bottom:1px solid #e8e8e8;font-size:12px}
.num{text-align:center}.ts{font-size:11px;color:#888;margin-bottom:18px}
.tot td{background:#f0f4ff;font-weight:700;border-top:2px solid #b0c4e8}`;

  const mkRow = r => `<tr><td>${e(r.name)}</td><td class="num">${n(r.open)}</td>
    <td class="num">${n(r.escalated)}</td><td class="num">${n(r.actionable)}</td>
    <td class="num">${n(r.escActionable)}</td><td class="num">${n(r.blackFlags)}</td></tr>`;

  const detailRow = c => `<tr><td><b>${e(c.CaseNumber)}</b></td><td>${e(c.ownerDisplay)}</td>
    <td>${e(c.Subject)}</td><td>${e(c.Status)}</td><td>${e(c.Priority)}</td></tr>`;

  const detailHead = `<tr><th>Case #</th><th>Owner</th><th>Subject</th><th>Status</th><th>Priority</th></tr>`;

  return `<html><head><style>${css}</style></head><body>
<h2>Team Open Cases Report</h2>
<p class="ts">Generated: ${runTime}</p>

<h3>Table 1 — Summary</h3>
<table><tr><th>Name</th><th>Open</th><th>Escalated</th><th>Actionables</th><th>Esc. Actionables</th><th>Black Flags</th></tr>
${summary.map(mkRow).join('')}
<tr class="tot"><td>TOTAL</td><td class="num">${totOpen}</td><td class="num">${totEsc}</td>
<td class="num">${totAct}</td><td class="num">${totEscA}</td><td class="num">${totBF}</td></tr></table>

<h3>Table 2 — Escalated Actionables by Owner</h3>
<table><tr><th>Owner</th><th>Esc. Actionables</th></tr>
${summary.filter(r=>r.escActionable>0).sort((a,b)=>b.escActionable-a.escActionable)
  .map(r=>`<tr><td>${e(r.name)}</td><td class="num"><b>${r.escActionable}</b></td></tr>`).join('')}
<tr class="tot"><td>TOTAL</td><td class="num">${totEscA}</td></tr></table>

<h3>Table 3 — Black Flags by Owner</h3>
<table><tr><th>Owner</th><th>Black Flag Cases</th></tr>
${summary.filter(r=>r.blackFlags>0).sort((a,b)=>b.blackFlags-a.blackFlags)
  .map(r=>`<tr><td>${e(r.name)}</td><td class="num"><b>${r.blackFlags}</b></td></tr>`).join('')}
<tr class="tot"><td>TOTAL</td><td class="num">${totBF}</td></tr></table>

<h3>Table 4 — Escalated Actionables List (${escActionableCases.length} cases)</h3>
<table>${detailHead}${escActionableCases.length ? escActionableCases.map(detailRow).join('') : '<tr><td colspan="5" style="color:#aaa">None</td></tr>'}</table>

<h3>Table 5 — Black Flag Cases (${blackFlagCases.length} cases)</h3>
<table>${detailHead}${blackFlagCases.length ? blackFlagCases.map(detailRow).join('') : '<tr><td colspan="5" style="color:#aaa">None</td></tr>'}</table>
</body></html>`;
}

// ── Open email in Outlook (Display) ─────────────────────────────────────────
function sendEmail(to, subject, htmlBody) {
  return new Promise((resolve, reject) => {
    const tmpHtml = path.join(os.tmpdir(), 'sf_email_body.html');
    fs.writeFileSync(tmpHtml, htmlBody, 'utf8');
    const defaultSubj = `Team Open Cases Report - ${new Date().toLocaleDateString()}`;
    const subj = (subject || defaultSubj).replace(/'/g, "''");
    const toSafe = to.replace(/'/g, "''");
    const tmpPath = tmpHtml.replace(/\\/g, '\\\\');
    const ps = `
$o = New-Object -ComObject Outlook.Application
$m = $o.CreateItem(0)
$m.To = '${toSafe}'
$m.Subject = '${subj}'
$m.HTMLBody = (Get-Content -Raw -Encoding UTF8 '${tmpPath}')
$m.Display()
Write-Host 'opened'`;
    const tmpPs = path.join(os.tmpdir(), 'sf_send_email.ps1');
    fs.writeFileSync(tmpPs, ps, 'utf8');
    // Fire-and-forget: Start-Process launches a new interactive window without blocking
    execFile('powershell', [
      '-Command',
      `Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File \\"${tmpPs.replace(/\\/g, '\\\\')}\\""' -WindowStyle Normal`
    ], { timeout: 10000 }, (err, stdout, stderr) => {
      // Start-Process returns immediately; any error here is a launch failure
      if (err) return reject(stderr || err.message);
      resolve('Email opened in Outlook - click Send to deliver');
    });
  });
}

// ── Schedule (Windows Task Scheduler) ───────────────────────────────────────
function getSchedule() {
  return new Promise(resolve => {
    exec(`schtasks /query /tn "${TASK_NAME}" /fo CSV /nh 2>nul`, (err, stdout) => {
      if (err || !stdout.trim()) return resolve(null);
      const parts = splitCSVLine(stdout.trim());
      resolve({ name: parts[0], nextRun: parts[1], status: parts[2] });
    });
  });
}

function createSchedule(freq, time, days, to) {
  return new Promise((resolve, reject) => {
    const psCmd = `powershell -ExecutionPolicy Bypass -File "${PS_SEND}" -To "${to}"`;
    let cmd = `schtasks /create /f /tn "${TASK_NAME}" /tr "${psCmd}" /sc ${freq} /st ${time}`;
    if (freq === 'WEEKLY' && days) cmd += ` /d ${days}`;
    exec(cmd, (err, stdout, stderr) => {
      if (err) return reject(stderr || err.message);
      resolve('Schedule created: ' + freq + ' at ' + time);
    });
  });
}

function deleteSchedule() {
  return new Promise((resolve, reject) => {
    exec(`schtasks /delete /f /tn "${TASK_NAME}"`, (err, stdout, stderr) => {
      if (err) return reject(stderr || err.message);
      resolve('Schedule deleted');
    });
  });
}

// ── HTML ─────────────────────────────────────────────────────────────────────
const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Team Open Cases Report</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/gridstack@10.3.1/dist/gridstack.min.css"/>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/gridstack@10.3.1/dist/gridstack.all.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;color:#1a2332;font-size:13px}
/* Header */
.header{background:linear-gradient(135deg,#0052cc,#0078d4);color:#fff;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.2)}
.header h1{font-size:18px;font-weight:600;letter-spacing:.3px}
.header .meta{font-size:11px;opacity:.8;margin-top:2px}
.actions{display:flex;gap:10px;align-items:center}
/* Buttons */
.btn{padding:7px 15px;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600;display:flex;align-items:center;gap:6px;transition:all .15s}
.btn-refresh{background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.3)}
.btn-refresh:hover{background:rgba(255,255,255,.25)}
.btn-email{background:#fff;color:#0078d4}
.btn-email:hover{background:#e8f0fe}
.btn-schedule{background:#ff8c00;color:#fff}
.btn-schedule:hover{background:#e07b00}
.btn-primary{background:#0078d4;color:#fff}
.btn-primary:hover{background:#006cbe}
.btn-danger{background:#d13438;color:#fff}
.btn-danger:hover{background:#b02a2e}
.btn-secondary{background:#f3f4f6;color:#444;border:1px solid #ddd}
.btn-secondary:hover{background:#e8e9eb}
.btn:disabled{opacity:.5;cursor:not-allowed}
/* Layout */
.main{padding:20px 24px;max-width:1400px;margin:0 auto}
.grid-2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.grid-full{margin-bottom:16px}
/* Cards */
.card{background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:hidden}
.card-header{padding:12px 16px;border-bottom:1px solid #f0f0f0;display:flex;align-items:center;justify-content:space-between}
.card-title{font-size:13px;font-weight:700;color:#0052cc;display:flex;align-items:center;gap:8px}
.badge{font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px}
.badge-blue{background:#e8f0fe;color:#0052cc}
.badge-orange{background:#fff3e0;color:#e65c00}
.badge-red{background:#fde8e8;color:#b00}
/* Tables */
.tbl-wrap{overflow-x:auto;max-height:420px;overflow-y:auto}
table{width:100%;border-collapse:collapse}
th{background:#f8f9ff;color:#444;font-weight:700;font-size:11px;padding:8px 12px;text-align:left;position:sticky;top:0;z-index:1;border-bottom:2px solid #e0e6ff}
td{padding:7px 12px;border-bottom:1px solid #f2f2f2;font-size:12px}
tr:hover td{background:#f7f9ff}
.num{text-align:center}
.num-zero{text-align:center;color:#ccc}
.sum-row td{background:#f0f4ff;font-weight:700;border-top:2px solid #c0d0f0}
/* Priority colors */
.pri-essential,.pri-critical{color:#b00;font-weight:700}
.pri-high{color:#e65c00;font-weight:600}
.pri-medium{color:#0052cc}
.name-col{font-weight:600}
/* Stat pills */
.pill{display:inline-block;padding:1px 7px;border-radius:10px;font-size:11px;font-weight:700}
.pill-esc{background:#fff3e0;color:#e65c00}
.pill-bf{background:#fde8e8;color:#b00}
.pill-act{background:#e8f0fe;color:#0052cc}
/* Loading */
.loading{display:flex;align-items:center;justify-content:center;padding:48px;flex-direction:column;gap:12px;color:#888}
.spinner{width:32px;height:32px;border:3px solid #e0e6ff;border-top-color:#0078d4;border-radius:50%;animation:spin .8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
/* Toast */
.toast-wrap{position:fixed;top:20px;right:20px;z-index:9999;display:flex;flex-direction:column;gap:8px}
.toast{padding:12px 18px;border-radius:8px;font-size:13px;font-weight:600;display:flex;align-items:center;gap:10px;box-shadow:0 4px 12px rgba(0,0,0,.15);animation:slideIn .2s ease}
.toast-success{background:#e6f4ea;color:#1a7f37;border-left:4px solid #2da44e}
.toast-error{background:#fde8e8;color:#b00;border-left:4px solid #d13438}
.toast-info{background:#e8f0fe;color:#0052cc;border-left:4px solid #0078d4}
@keyframes slideIn{from{transform:translateX(40px);opacity:0}to{transform:translateX(0);opacity:1}}
/* Modal */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:1000;align-items:center;justify-content:center}
.modal-bg.open{display:flex}
.modal{background:#fff;border-radius:12px;padding:24px;width:460px;max-width:95vw;box-shadow:0 8px 32px rgba(0,0,0,.2)}
.modal h3{font-size:16px;font-weight:700;margin-bottom:16px;color:#1a2332;display:flex;align-items:center;gap:8px}
.modal-footer{display:flex;justify-content:flex-end;gap:8px;margin-top:20px}
.form-group{margin-bottom:14px}
label{display:block;font-size:11px;font-weight:700;color:#555;margin-bottom:4px;text-transform:uppercase;letter-spacing:.4px}
input,select{width:100%;padding:8px 12px;border:1px solid #ddd;border-radius:6px;font-size:13px;color:#1a2332;outline:none;transition:border .15s}
input:focus,select:focus{border-color:#0078d4;box-shadow:0 0 0 2px rgba(0,120,212,.15)}
.sched-info{background:#f8f9ff;border-radius:6px;padding:10px 14px;font-size:12px;color:#444;margin-bottom:14px;display:flex;align-items:center;gap:8px}
.days-grid{display:flex;gap:6px;flex-wrap:wrap}
.day-btn{padding:5px 10px;border:1px solid #ddd;border-radius:6px;cursor:pointer;font-size:11px;font-weight:600;background:#fff;color:#555;transition:all .15s}
.day-btn.active{background:#0078d4;color:#fff;border-color:#0078d4}
/* Nav tabs */
.nav-tabs{background:#fff;border-bottom:2px solid #e8ecf0;padding:0 24px;display:flex}
.nav-tab{background:none;border:none;border-bottom:3px solid transparent;margin-bottom:-2px;padding:11px 20px;font-size:13px;font-weight:600;color:#666;cursor:pointer;display:flex;align-items:center;gap:8px;transition:color .15s}
.nav-tab:hover{color:#0052cc}
.nav-tab.active{color:#0052cc;border-bottom-color:#0078d4}
/* Case Handling */
.ch-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:16px}
.ch-stat-card{background:#fff;border-radius:10px;padding:22px;box-shadow:0 1px 4px rgba(0,0,0,.08);text-align:center}
.ch-stat-label{font-size:11px;text-transform:uppercase;letter-spacing:.7px;color:#888;font-weight:700;margin-bottom:8px}
.ch-stat-value{font-size:38px;font-weight:800;letter-spacing:-1.5px;line-height:1;margin-bottom:6px}
.ch-stat-sub{font-size:11px;color:#aaa}
.ch-stat-card.today .ch-stat-value{color:#0078d4}
.ch-stat-card.week  .ch-stat-value{color:#0052cc}
.ch-stat-card.month .ch-stat-value{color:#7c4dff}
.ch-num{display:inline-block;padding:3px 12px;border-radius:12px;font-size:12px;font-weight:700;cursor:pointer;background:#e8f0fe;color:#0078d4;transition:background .15s}
.ch-num:hover{background:#c5d8f8}
.ch-num.ch-week{background:#e8eeff;color:#0052cc}
.ch-num.ch-week:hover{background:#c5d0ff}
.ch-num.ch-month{background:#ede7ff;color:#7c4dff}
.ch-num.ch-month:hover{background:#d5caff}
.ch-case-row{padding:9px 4px;border-bottom:1px solid #f2f2f2;display:flex;align-items:center;gap:8px;font-size:13px}
.ch-case-row:last-child{border-bottom:none}
.ch-period-btn{background:#f3f4f6;color:#555;border:1px solid #ddd;padding:5px 13px;font-size:12px;font-weight:600;border-radius:6px;cursor:pointer;transition:all .15s}
.ch-period-btn:hover{background:#e8ecf5;color:#0052cc}
.ch-period-btn.active{background:#0078d4;color:#fff;border-color:#0078d4}
/* Pin button */
.btn-pin{background:none;border:none;padding:4px 7px;border-radius:5px;cursor:pointer;color:#bbb;font-size:13px;transition:all .15s;line-height:1}
.btn-pin:hover{color:#0078d4;background:#e8f0fe}
.btn-pin.pinned{color:#0078d4}
/* Dashboard */
.dash-empty{text-align:center;padding:80px 24px;color:#aaa}
.dash-empty i{font-size:52px;display:block;margin-bottom:16px;opacity:.25}
.dash-empty p{font-size:14px}
.dash-widget-inner{height:100%;display:flex;flex-direction:column;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.dash-widget-header{padding:10px 14px;border-bottom:1px solid #f0f0f0;display:flex;justify-content:space-between;align-items:center;font-size:12px;font-weight:700;color:#0052cc;flex-shrink:0;cursor:move}
.dash-widget-chart{position:relative;flex:1;min-height:0;padding:10px}
.dash-widget-chart canvas{position:absolute!important;top:10px;left:10px;width:calc(100% - 20px)!important;height:calc(100% - 20px)!important}
.dash-unpin-btn{background:none;border:none;cursor:pointer;color:#bbb;font-size:13px;padding:3px 6px;border-radius:4px;flex-shrink:0}
.dash-unpin-btn:hover{color:#d13438;background:#fde8e8}
/* Gridstack overrides */
.grid-stack{background:transparent}
.grid-stack-item-content{background:transparent;border-radius:10px;overflow:hidden;height:100%}
.grid-stack .grid-stack-placeholder>.placeholder-content{background:rgba(0,120,212,.08);border:2px dashed #0078d4;border-radius:10px}
/* JPMC tab */
.jpmc-filter-bar{background:#f8f9ff;border:1px solid #e0e6ff;border-radius:8px;padding:10px 18px;margin-bottom:16px;font-size:12px;color:#555;display:flex;gap:20px;align-items:center;flex-wrap:wrap}
.jpmc-filter-bar strong{color:#0052cc}
.assign-btn{padding:4px 12px;border-radius:6px;border:1.5px solid #0078d4;background:#fff;color:#0078d4;font-size:11px;font-weight:700;cursor:pointer;transition:all .15s;white-space:nowrap}
.assign-btn:hover{background:#0078d4;color:#fff}
.assign-btn.assigned{background:#e6f4ea;border-color:#2da44e;color:#1a7f37;cursor:default}
.assign-picker{position:absolute;right:0;top:100%;background:#fff;border:1px solid #ddd;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.15);z-index:500;min-width:200px;padding:4px 0;max-height:260px;overflow-y:auto}
.assign-picker-item{padding:8px 16px;font-size:12px;cursor:pointer;color:#1a2332;transition:background .1s}
.assign-picker-item:hover{background:#f0f4ff;color:#0052cc}
.jpmc-pool-row{display:flex;flex-wrap:wrap;gap:6px;align-items:center;padding:8px 14px;background:#f8f9ff;border:1px solid #e0e6ff;border-radius:8px;margin-bottom:12px}
.jpmc-pool-label{display:inline-flex;align-items:center;gap:5px;font-size:12px;cursor:pointer;padding:3px 10px;border-radius:5px;background:#fff;border:1px solid #ddd;transition:background .1s;color:#333;user-select:none}
.jpmc-pool-label:hover{background:#ede9fe;border-color:#b0a0e0}
.jpmc-schedule-wrap{display:inline-flex;align-items:center;gap:5px}
.jpmc-schedule-sel{font-size:11px;border:1.5px solid #c0b0f0;border-radius:6px;padding:3px 7px;color:#4a3a80;background:#f8f4ff;cursor:pointer;outline:none}
.jpmc-schedule-sel:focus{border-color:#6554c0}
.jpmc-autoassign-label{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:700;cursor:pointer;padding:3px 10px;border-radius:6px;border:1.5px solid #0052cc;background:#f0f4ff;color:#0052cc;user-select:none;transition:background .1s}
.jpmc-autoassign-label:hover{background:#dde8ff}
.jpmc-pool-label input{cursor:pointer;accent-color:#6554c0}
/* Collapsible sections */
.collapse-body{transition:max-height .25s ease,opacity .25s ease;overflow:hidden}
.collapse-body.collapsed{max-height:0!important;opacity:0;pointer-events:none}
.btn-collapse{background:none;border:none;cursor:pointer;color:#aaa;font-size:13px;padding:4px 6px;border-radius:5px;transition:all .15s;line-height:1}
.btn-collapse:hover{color:#0078d4;background:#e8f0fe}
.btn-collapse i{transition:transform .25s}
.btn-collapse.collapsed i{transform:rotate(-90deg)}
/* Respond button */
.respond-btn{padding:4px 12px;border-radius:6px;border:1.5px solid #2da44e;background:#fff;color:#1a7f37;font-size:11px;font-weight:700;cursor:pointer;transition:all .15s;white-space:nowrap;display:inline-flex;align-items:center;gap:4px}
.respond-btn:hover{background:#2da44e;color:#fff}
.respond-btn.done{background:#e6f4ea;border-color:#2da44e;color:#1a7f37;cursor:default;opacity:.7}
.respond-btn:disabled{opacity:.5;cursor:not-allowed}
/* New Cases section */
.jpmc-section-divider{margin:20px 0 12px;font-size:12px;font-weight:700;color:#555;display:flex;align-items:center;gap:8px;text-transform:uppercase;letter-spacing:.5px}
.jpmc-section-divider::after{content:'';flex:1;height:1px;background:#e0e6ff}
/* Respond confirm modal */
.respond-preview{background:#f8f9ff;border:1px solid #e0e6ff;border-radius:8px;padding:12px 14px;font-size:12px;color:#333;white-space:pre-wrap;max-height:260px;overflow-y:auto;line-height:1.6;margin-bottom:4px}
</style>
</head>
<body>

<div class="header">
  <div>
    <h1><i class="fa fa-chart-bar"></i> &nbsp;Team Open Cases Report</h1>
    <div class="meta" id="lastUpdated">Loading data...</div>
  </div>
  <div class="actions">
    <button class="btn btn-refresh" onclick="loadData()" id="btnRefresh">
      <i class="fa fa-rotate-right"></i> Refresh
    </button>
    <button class="btn btn-email" onclick="openEmailModal()">
      <i class="fa fa-envelope"></i> Send Email
    </button>
    <button class="btn btn-schedule" onclick="openScheduleModal()">
      <i class="fa fa-clock"></i> Schedule
    </button>
  </div>
</div>

<div class="nav-tabs">
  <button class="nav-tab" id="tab-dashboard" onclick="switchTab('dashboard')">
    <i class="fa fa-gauge"></i> Dashboard
  </button>
  <button class="nav-tab active" id="tab-report" onclick="switchTab('report')">
    <i class="fa fa-chart-bar"></i> Open Cases Report
  </button>
  <button class="nav-tab" id="tab-handling" onclick="switchTab('handling')">
    <i class="fa fa-comments"></i> Case Handling
  </button>
  <button class="nav-tab" id="tab-jpmc" onclick="switchTab('jpmc')">
    <i class="fa fa-building"></i> JPMC Restore Request
  </button>
</div>

<div class="main" id="mainContent">
  <div id="page-report">
    <div class="loading" id="loader">
      <div class="spinner"></div>
      <div>Fetching data from Salesforce...</div>
    </div>
    <div id="reportContent" style="display:none"></div>
  </div>
  <div id="page-handling" style="display:none">
    <div class="loading" id="chLoader" style="display:none">
      <div class="spinner"></div>
      <div>Loading case handling data from Salesforce...</div>
    </div>
    <div id="chContent"></div>
  </div>
  <div id="page-dashboard" style="display:none">
    <div id="dashEmpty" class="dash-empty">
      <i class="fa fa-gauge"></i>
      <p>No charts pinned yet.<br>Click <i class="fa fa-thumbtack"></i> on any chart to add it here.</p>
    </div>
    <div class="grid-stack" id="dashGrid" style="display:none"></div>
  </div>
  <div id="page-jpmc" style="display:none">
    <!-- JPMC Stats Dashboard -->
    <div class="card grid-full" id="jpmcStatsCard" style="margin-bottom:16px">
      <div class="card-header">
        <div style="display:flex;align-items:center;gap:8px">
          <button id="collapse-stats" class="btn-collapse" onclick="toggleJpmcSection('stats')" title="Collapse"><i class="fa fa-chevron-down"></i></button>
          <div class="card-title">
            <i class="fa fa-chart-line" style="color:#0052cc"></i> JPMC Support Tickets — 2026 Overview
            <span class="badge badge-blue" id="statTotalBadge" style="margin-left:8px"></span>
          </div>
        </div>
        <div style="display:flex;gap:6px;align-items:center">
          <button id="statsBtnDay7"  class="ch-period-btn active" onclick="setJpmcStatsPeriod('day7')">7 Days</button>
          <button id="statsBtnDay30" class="ch-period-btn"        onclick="setJpmcStatsPeriod('day30')">30 Days</button>
          <button id="statsBtnDay90" class="ch-period-btn"        onclick="setJpmcStatsPeriod('day90')">90 Days</button>
          <button id="statsBtnWeek"  class="ch-period-btn"        onclick="setJpmcStatsPeriod('weekly')">Weekly</button>
          <button id="statsBtnMonth" class="ch-period-btn"        onclick="setJpmcStatsPeriod('monthly')">Monthly</button>
          <button class="btn btn-secondary" style="font-size:11px;margin-left:6px" onclick="loadJpmcStats()">
            <i class="fa fa-rotate-right"></i>
          </button>
          <button id="pin-jpmc-stats" class="btn-pin" onclick="pinItem('jpmc-stats')" title="Pin to Dashboard"><i class="fa fa-thumbtack"></i></button>
        </div>
      </div>
      <div id="jpmcStatsBody" class="collapse-body" style="max-height:600px">
      <div id="jpmcStatsLoader" style="display:flex;align-items:center;gap:12px;padding:20px 16px">
        <div class="spinner"></div><span style="color:#888;font-size:13px">Loading stats...</span>
      </div>
      <div id="jpmcStatsContent" style="display:none">
        <div style="display:flex;gap:16px;padding:16px 20px 0">
          <div style="flex:0 0 200px;display:flex;flex-direction:column;gap:10px">
            <div class="ch-stat-card today" style="padding:14px 16px">
              <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
                <span class="ch-stat-label" style="margin:0">Restore Requests</span>
                <span id="statRestorePct" style="font-size:13px;font-weight:800;color:#0078d4;background:#e8f0fe;padding:2px 8px;border-radius:10px">—%</span>
              </div>
              <div class="ch-stat-value" id="statTotal" style="font-size:30px">—</div>
              <div class="ch-stat-sub">of 2026 JPMC tickets</div>
            </div>
            <div class="ch-stat-card week" style="padding:14px 16px">
              <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
                <span class="ch-stat-label" style="margin:0">Non-Restore Requests</span>
                <span id="statNonRestorePct" style="font-size:13px;font-weight:800;color:#0052cc;background:#e8eeff;padding:2px 8px;border-radius:10px">—%</span>
              </div>
              <div class="ch-stat-value" id="statTotalAll" style="font-size:30px">—</div>
              <div class="ch-stat-sub">other JPMC tickets 2026</div>
            </div>
          </div>
          <div style="flex:1;min-height:240px;padding-bottom:16px">
            <canvas id="jpmcStatsChart"></canvas>
          </div>
        </div>
      </div>
      </div>
    </div>

    <!-- Section 2: Entry Restore Requests -->
    <div id="jpmcRestoreSection">
    <div class="loading" id="jpmcLoader" style="display:none">
      <div class="spinner"></div>
      <div>Loading JPMC cases from Salesforce...</div>
    </div>
    <div id="jpmcContent"></div>
    </div>

    <!-- Section 3: All Open Cases -->
    <div class="jpmc-section-divider" style="margin:20px 0 12px;cursor:pointer" onclick="toggleJpmcSection('open')">
      <button id="collapse-open" class="btn-collapse" style="margin-right:4px" onclick="event.stopPropagation();toggleJpmcSection('open')"><i class="fa fa-chevron-down"></i></button>
      <i class="fa fa-inbox" style="color:#0078d4"></i> All Open Entry Restore Requests
    </div>
    <div id="jpmcOpenBody" class="collapse-body" style="max-height:2000px">
    <div style="display:flex;gap:16px;align-items:flex-start">
      <div class="card" id="jpmcNewCasesCard" style="flex:1;min-width:0">
        <div class="card-header">
          <div class="card-title"><i class="fa fa-list-check"></i> Open Entry Restore Requests <span class="badge badge-blue" id="jpmcNewBadge"></span></div>
          <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
            <div class="jpmc-schedule-wrap">
              <i class="fa fa-filter" style="color:#0078d4;font-size:12px"></i>
              <select id="jpmcNewStatusFilter" class="jpmc-schedule-sel" onchange="filterJpmcNewCases()">
                <option value="new-all" selected>New &amp; New Assigned</option>
                <option value="New">New</option>
                <option value="New Assigned">New Assigned</option>
                <option value="all">All statuses</option>
                <option value="In Progress">In Progress</option>
                <option value="In Work">In Work</option>
                <option value="Awaiting Customer Response">Awaiting Customer Response</option>
                <option value="Customer Responded">Customer Responded</option>
                <option value="Review Customer Response">Review Customer Response</option>
                <option value="On Hold">On Hold</option>
                <option value="Resolved">Resolved</option>
              </select>
            </div>
            <button class="btn btn-secondary" style="font-size:11px" onclick="loadJpmcNewCases()">
              <i class="fa fa-rotate-right"></i> Refresh
            </button>
          </div>
        </div>
        <div class="loading" id="jpmcNewLoader" style="display:flex;padding:24px">
          <div class="spinner"></div><div style="margin-left:12px">Loading...</div>
        </div>
        <div class="tbl-wrap" id="jpmcNewContent" style="display:none"></div>
      </div>
      <div class="card" id="templateCard" style="width:260px;flex-shrink:0">
        <div class="card-header" style="padding:10px 14px">
          <div class="card-title" style="font-size:12px">
            <i class="fa fa-envelope-open-text" style="color:#0052cc"></i> Response Preview
          </div>
        </div>
        <textarea id="templatePreview" spellcheck="false" style="width:100%;padding:12px 14px;font-size:11px;color:#333;line-height:1.6;height:400px;resize:vertical;border:none;outline:none;background:#fafbff;border-radius:0 0 10px 10px;font-family:Segoe UI,Arial,sans-serif">Hover a case row to preview...</textarea>
      </div>
    </div>
    </div>
  </div>
</div>

<!-- Email Modal -->
<div class="modal-bg" id="emailModal">
  <div class="modal">
    <h3><i class="fa fa-envelope" style="color:#0078d4"></i> Send Report Email</h3>
    <div class="form-group">
      <label>Recipient(s)</label>
      <input type="text" id="emailTo" value="" placeholder="email@kaltura.com">
    </div>
    <div class="form-group">
      <label>Subject</label>
      <input type="text" id="emailSubject" placeholder="Leave blank for default subject">
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('emailModal')">Cancel</button>
      <button class="btn btn-primary" onclick="sendEmail()" id="btnSend">
        <i class="fa fa-envelope-open-text"></i> Open in Outlook
      </button>
    </div>
  </div>
</div>

<!-- Schedule Modal -->
<div class="modal-bg" id="schedModal">
  <div class="modal">
    <h3><i class="fa fa-clock" style="color:#ff8c00"></i> Schedule Email Report</h3>
    <div class="sched-info" id="schedStatus">
      <i class="fa fa-circle-info" style="color:#0078d4"></i>
      <span>Checking schedule...</span>
    </div>
    <div class="form-group">
      <label>Frequency</label>
      <select id="schedFreq" onchange="toggleDays()">
        <option value="DAILY">Daily (weekdays)</option>
        <option value="WEEKLY">Weekly</option>
      </select>
    </div>
    <div class="form-group" id="daysGroup" style="display:none">
      <label>Days</label>
      <div class="days-grid" id="daysGrid">
        <div class="day-btn active" data-day="MON">Mon</div>
        <div class="day-btn" data-day="TUE">Tue</div>
        <div class="day-btn" data-day="WED">Wed</div>
        <div class="day-btn" data-day="THU">Thu</div>
        <div class="day-btn" data-day="FRI">Fri</div>
      </div>
    </div>
    <div class="form-group">
      <label>Time</label>
      <input type="time" id="schedTime" value="09:00">
    </div>
    <div class="form-group">
      <label>Send To</label>
      <input type="text" id="schedEmail" value="">
    </div>
    <div class="modal-footer">
      <button class="btn btn-danger" onclick="deleteSchedule()" id="btnDelSched" style="margin-right:auto">
        <i class="fa fa-trash"></i> Remove
      </button>
      <button class="btn btn-secondary" onclick="closeModal('schedModal')">Cancel</button>
      <button class="btn btn-schedule" onclick="saveSchedule()">
        <i class="fa fa-save"></i> Save Schedule
      </button>
    </div>
  </div>
</div>

<!-- Respond Confirm Modal -->
<div class="modal-bg" id="respondModal">
  <div class="modal" style="width:580px">
    <h3><i class="fa fa-reply" style="color:#2da44e"></i> Send Response &amp; Close Case</h3>
    <div style="font-size:12px;color:#555;margin-bottom:10px">
      Case <strong id="respondCaseNum"></strong> &mdash; <span id="respondContact"></span>
    </div>
    <textarea id="respondPreview" spellcheck="false" style="width:100%;height:320px;padding:10px 12px;font-size:11px;color:#333;line-height:1.6;border:1px solid #d0d7de;border-radius:6px;resize:vertical;font-family:Segoe UI,Arial,sans-serif;background:#fafbff;outline:none"></textarea>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('respondModal')">Cancel</button>
      <button class="btn" style="background:#2da44e;color:#fff" id="respondConfirmBtn" onclick="confirmRespond()">
        <i class="fa fa-paper-plane"></i> Send &amp; Close Case
      </button>
    </div>
  </div>
</div>

<!-- Case Handling Cases Modal -->
<div class="modal-bg" id="chModal">
  <div class="modal">
    <h3><i class="fa fa-ticket-simple" style="color:#0078d4"></i> <span id="chModalTitle">Cases</span></h3>
    <div id="chModalBody" style="max-height:340px;overflow-y:auto;margin-top:4px"></div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="closeModal('chModal')">Close</button>
    </div>
  </div>
</div>

<div class="toast-wrap" id="toasts"></div>

<script>
let cachedData = null;
Chart.register(ChartDataLabels);
let activeCharts = [];

// ── Data loading ─────────────────────────────────────────────────────────────
async function loadData() {
  document.getElementById('loader').style.display = 'flex';
  document.getElementById('reportContent').style.display = 'none';
  document.getElementById('btnRefresh').disabled = true;
  document.getElementById('lastUpdated').textContent = 'Fetching from Salesforce...';
  try {
    const res = await fetch('/api/data');
    if (!res.ok) throw new Error(await res.text());
    cachedData = await res.json();
    renderReport(cachedData);
    const d = new Date(cachedData.fetchedAt);
    document.getElementById('lastUpdated').textContent = 'Last updated: ' + d.toLocaleString();
  } catch(e) {
    toast('Error loading data: ' + e.message, 'error');
    document.getElementById('lastUpdated').textContent = 'Failed to load';
  } finally {
    document.getElementById('loader').style.display = 'none';
    document.getElementById('reportContent').style.display = 'block';
    document.getElementById('btnRefresh').disabled = false;
  }
}

// ── Render ────────────────────────────────────────────────────────────────────
function renderReport(data) {
  const { summary, escActionableCases, blackFlagCases } = data;

  const totOpen = sum(summary,'open'), totEsc = sum(summary,'escalated'),
        totAct  = sum(summary,'actionable'), totEscA = sum(summary,'escActionable'),
        totBF   = sum(summary,'blackFlags');

  // ─ Table 1: Summary ─
  let t1 = \`<table><thead><tr>
    <th>Name</th><th class="num">Open</th><th class="num">Escalated</th>
    <th class="num">Actionables</th><th class="num">Esc. Actionables</th><th class="num">Black Flags</th>
  </tr></thead><tbody>\`;
  summary.forEach(r => {
    t1 += \`<tr>
      <td class="name-col">\${r.name}</td>
      \${nd(r.open)} \${ne(r.escalated,'pill-esc')} \${na(r.actionable,'pill-act')}
      \${ne(r.escActionable,'pill-esc')} \${ne(r.blackFlags,'pill-bf')}
    </tr>\`;
  });
  t1 += \`<tr class="sum-row"><td>TOTAL</td>
    <td class="num">\${totOpen}</td><td class="num">\${totEsc}</td>
    <td class="num">\${totAct}</td><td class="num">\${totEscA}</td><td class="num">\${totBF}</td>
  </tr></tbody></table>\`;

  // ─ Table 2: Escalated Actionables by Owner (count) ─
  const escActOwners = summary.filter(r => r.escActionable > 0).sort((a,b) => b.escActionable - a.escActionable);
  let t2 = '<table><thead><tr><th>Owner</th><th class="num">Esc. Actionables</th></tr></thead><tbody>';
  if (!escActOwners.length) { t2 += '<tr><td colspan="2" style="text-align:center;color:#aaa;padding:20px">None</td></tr>'; }
  else escActOwners.forEach(r => { t2 += \`<tr><td class="name-col">\${r.name}</td><td class="num"><span class="pill pill-esc">\${r.escActionable}</span></td></tr>\`; });
  t2 += \`<tr class="sum-row"><td>TOTAL</td><td class="num">\${totEscA}</td></tr></tbody></table>\`;

  // ─ Table 3: Black Flags by Owner ─
  const bfOwners = summary.filter(r => r.blackFlags > 0).sort((a,b) => b.blackFlags - a.blackFlags);
  let t3 = '<table><thead><tr><th>Owner</th><th class="num">Black Flag Cases</th></tr></thead><tbody>';
  bfOwners.forEach(r => { t3 += \`<tr><td class="name-col">\${r.name}</td><td class="num"><span class="pill pill-bf">\${r.blackFlags}</span></td></tr>\`; });
  t3 += \`<tr class="sum-row"><td>TOTAL</td><td class="num">\${totBF}</td></tr></tbody></table>\`;

  // ─ Table 4: Escalated Actionables ─
  // ─ Table 4 (rendered as t5): Black Flag Cases ─
  let t5 = '<table><thead><tr><th>Case #</th><th>Owner</th><th>Subject</th><th>Status</th><th>Priority</th></tr></thead><tbody>';
  if (!blackFlagCases.length) { t5 += '<tr><td colspan="5" style="text-align:center;color:#aaa;padding:20px">No black flag cases</td></tr>'; }
  else blackFlagCases.forEach(c => { t5 += \`<tr><td><b>\${c.CaseNumber}</b></td><td>\${c.ownerDisplay}</td><td>\${esc(c.Subject)}</td><td>\${c.Status}</td><td class="\${priClass(c.Priority)}">\${c.Priority}</td></tr>\`; });
  t5 += '</tbody></table>';

  // ─ Table 5 (rendered as t6): Escalated Actionables list ─
  let t6 = '<table><thead><tr><th>Case #</th><th>Owner</th><th>Subject</th><th>Status</th><th>Priority</th></tr></thead><tbody>';
  if (!escActionableCases.length) { t6 += '<tr><td colspan="5" style="text-align:center;color:#aaa;padding:20px">No escalated actionable cases</td></tr>'; }
  else escActionableCases.forEach(c => { t6 += \`<tr><td><b>\${c.CaseNumber}</b></td><td>\${c.ownerDisplay}</td><td>\${esc(c.Subject)}</td><td>\${c.Status}</td><td class="\${priClass(c.Priority)}">\${c.Priority}</td></tr>\`; });
  t6 += '</tbody></table>';

  document.getElementById('reportContent').innerHTML = \`
    <div class="grid-full">
      <div class="card">
        <div class="card-header">
          <span class="card-title"><i class="fa fa-table-list"></i> Table 1 &mdash; Summary</span>
          <span class="badge badge-blue">\${totOpen} open cases</span>
        </div>
        <div class="tbl-wrap">\${t1}</div>
      </div>
    </div>
    <div class="grid-2" style="margin-bottom:16px">
      <div class="card">
        <div class="card-header">
          <span class="card-title"><i class="fa fa-chart-bar" style="color:#e65c00"></i> Escalated Actionables by Owner</span>
          <button id="pin-esc-actionables" class="btn-pin" onclick="pinItem('esc-actionables')" title="Pin to Dashboard"><i class="fa fa-thumbtack"></i></button>
        </div>
        <div style="padding:16px 20px"><canvas id="chartEscAct"></canvas></div>
      </div>
      <div class="card">
        <div class="card-header">
          <span class="card-title"><i class="fa fa-chart-bar" style="color:#b00"></i> Black Flags by Owner</span>
          <button id="pin-black-flags" class="btn-pin" onclick="pinItem('black-flags')" title="Pin to Dashboard"><i class="fa fa-thumbtack"></i></button>
        </div>
        <div style="padding:16px 20px"><canvas id="chartBF"></canvas></div>
      </div>
    </div>
    <div class="grid-2">
      <div class="card">
        <div class="card-header">
          <span class="card-title"><i class="fa fa-fire-flame-curved" style="color:#e65c00"></i> Table 2 &mdash; Escalated Actionables by Owner</span>
          <span class="badge badge-orange">\${totEscA} cases</span>
        </div>
        <div class="tbl-wrap">\${t2}</div>
      </div>
      <div class="card">
        <div class="card-header">
          <span class="card-title"><i class="fa fa-flag" style="color:#111"></i> Table 3 &mdash; Black Flags by Owner</span>
          <span class="badge badge-red">\${totBF} flagged</span>
        </div>
        <div class="tbl-wrap">\${t3}</div>
      </div>
    </div>
    <div class="grid-2">
      <div class="card">
        <div class="card-header">
          <span class="card-title"><i class="fa fa-fire-flame-curved" style="color:#e65c00"></i> Table 4 &mdash; Escalated Actionables List</span>
          <span class="badge badge-orange">\${escActionableCases.length} cases</span>
        </div>
        <div class="tbl-wrap">\${t6}</div>
      </div>
      <div class="card">
        <div class="card-header">
          <span class="card-title"><i class="fa fa-skull-crossbones" style="color:#111"></i> Table 5 &mdash; Black Flag Cases</span>
          <span class="badge badge-red">\${blackFlagCases.length} cases</span>
        </div>
        <div class="tbl-wrap">\${t5}</div>
      </div>
    </div>
  \`;

  renderCharts(summary);
}

function renderCharts(summary) {
  activeCharts.forEach(c => c.destroy());
  activeCharts = [];
  // Escalated Actionables chart
  const escData = summary.filter(r => r.escActionable > 0).sort((a,b) => b.escActionable - a.escActionable);
  activeCharts.push(new Chart(document.getElementById('chartEscAct'), {
    type: 'bar',
    data: {
      labels: escData.map(r => r.name),
      datasets: [{ data: escData.map(r => r.escActionable),
        backgroundColor: escData.map(() => 'rgba(230,92,0,0.75)'),
        borderColor: escData.map(() => '#e65c00'),
        borderWidth: 1, borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y', responsive: true, clip: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' ' + ctx.parsed.x + ' cases' } },
        datalabels: { anchor: 'end', align: 'start', color: '#fff', font: { weight: 'bold', size: 12 }, formatter: v => v }
      },
      scales: { x: { beginAtZero: true, ticks: { stepSize: 1 }, grid: { color: '#f0f0f0' } }, y: { grid: { display: false } } }
    }
  }));

  // Black Flags chart
  const bfData = summary.filter(r => r.blackFlags > 0).sort((a,b) => b.blackFlags - a.blackFlags);
  activeCharts.push(new Chart(document.getElementById('chartBF'), {
    type: 'bar',
    data: {
      labels: bfData.map(r => r.name),
      datasets: [{ data: bfData.map(r => r.blackFlags),
        backgroundColor: bfData.map(() => 'rgba(176,0,32,0.75)'),
        borderColor: bfData.map(() => '#b00020'),
        borderWidth: 1, borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y', responsive: true, clip: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' ' + ctx.parsed.x + ' cases' } },
        datalabels: { anchor: 'end', align: 'start', color: '#fff', font: { weight: 'bold', size: 12 }, formatter: v => v }
      },
      scales: { x: { beginAtZero: true, ticks: { stepSize: 1 }, grid: { color: '#f0f0f0' } }, y: { grid: { display: false } } }
    }
  }));
}

function sum(arr, key) { return arr.reduce((s,r) => s + r[key], 0); }
function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function priClass(p) { const l = (p||'').toLowerCase(); return l==='essential'||l==='critical'?'pri-essential':l==='high'?'pri-high':l==='medium'?'pri-medium':''; }
function nd(n) { return n ? \`<td class="num">\${n}</td>\` : \`<td class="num-zero">0</td>\`; }
function ne(n, cls) { return n ? \`<td class="num"><span class="pill \${cls}">\${n}</span></td>\` : \`<td class="num-zero">0</td>\`; }
function na(n, cls) { return n ? \`<td class="num"><span class="pill \${cls}">\${n}</span></td>\` : \`<td class="num-zero">0</td>\`; }

// ── Email modal ───────────────────────────────────────────────────────────────
function openEmailModal() { document.getElementById('emailModal').classList.add('open'); }

async function sendEmail() {
  const to = document.getElementById('emailTo').value.trim();
  const subj = document.getElementById('emailSubject').value.trim();
  if (!to) return toast('Please enter a recipient', 'error');
  const btn = document.getElementById('btnSend');
  btn.disabled = true; btn.innerHTML = '<i class="fa fa-spinner fa-spin"></i> Opening...';
  try {
    const res = await fetch('/api/send-email', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({to, subject: subj}) });
    const txt = await res.text();
    if (!res.ok) throw new Error(txt);
    closeModal('emailModal');
    toast('Email opened in Outlook - click Send to deliver', 'success');
  } catch(e) { toast('Failed: ' + e.message, 'error'); }
  finally { btn.disabled = false; btn.innerHTML = '<i class="fa fa-envelope-open-text"></i> Open in Outlook'; }
}

// ── Schedule modal ────────────────────────────────────────────────────────────
async function openScheduleModal() {
  document.getElementById('schedModal').classList.add('open');
  document.getElementById('schedStatus').innerHTML = '<i class="fa fa-spinner fa-spin" style="color:#0078d4"></i> <span>Checking schedule...</span>';
  try {
    const res = await fetch('/api/schedule');
    const info = await res.json();
    if (info) {
      document.getElementById('schedStatus').innerHTML = \`<i class="fa fa-circle-check" style="color:#2da44e"></i> <span>Active &mdash; Next run: \${info.nextRun}</span>\`;
      document.getElementById('btnDelSched').style.display = '';
    } else {
      document.getElementById('schedStatus').innerHTML = '<i class="fa fa-circle-xmark" style="color:#888"></i> <span>No schedule configured</span>';
      document.getElementById('btnDelSched').style.display = 'none';
    }
  } catch { document.getElementById('schedStatus').innerHTML = '<i class="fa fa-circle-info" style="color:#0078d4"></i> <span>Could not check schedule</span>'; }
}

function toggleDays() {
  const isWeekly = document.getElementById('schedFreq').value === 'WEEKLY';
  document.getElementById('daysGroup').style.display = isWeekly ? '' : 'none';
}

document.querySelectorAll('.day-btn').forEach(btn => {
  btn.addEventListener('click', () => btn.classList.toggle('active'));
});

async function saveSchedule() {
  const freq  = document.getElementById('schedFreq').value;
  const time  = document.getElementById('schedTime').value;
  const email = document.getElementById('schedEmail').value.trim();
  const days  = [...document.querySelectorAll('.day-btn.active')].map(b => b.dataset.day).join(',');
  if (!email) return toast('Please enter an email address', 'error');
  try {
    const res = await fetch('/api/schedule', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({freq, time, days, to: email}) });
    const txt = await res.text();
    if (!res.ok) throw new Error(txt);
    closeModal('schedModal');
    toast('Schedule saved: ' + freq + ' at ' + time, 'success');
  } catch(e) { toast('Failed: ' + e.message, 'error'); }
}

async function deleteSchedule() {
  if (!confirm('Remove the scheduled task?')) return;
  try {
    const res = await fetch('/api/schedule', { method:'DELETE' });
    const txt = await res.text();
    if (!res.ok) throw new Error(txt);
    closeModal('schedModal');
    toast('Schedule removed', 'success');
  } catch(e) { toast('Failed: ' + e.message, 'error'); }
}

// ── Utilities ─────────────────────────────────────────────────────────────────
function closeModal(id) { document.getElementById(id).classList.remove('open'); }
document.querySelectorAll('.modal-bg').forEach(m => { m.addEventListener('click', e => { if (e.target === m) m.classList.remove('open'); }); });

function toast(msg, type='info') {
  const div = document.createElement('div');
  const icon = type==='success'?'circle-check':type==='error'?'circle-xmark':'circle-info';
  div.className = \`toast toast-\${type}\`;
  div.innerHTML = \`<i class="fa fa-\${icon}"></i>\${msg}\`;
  document.getElementById('toasts').appendChild(div);
  setTimeout(() => div.remove(), 4000);
}

// ── Tab switching ─────────────────────────────────────────────────────────────
function switchTab(tab) {
  ['report','handling','dashboard','jpmc'].forEach(t => {
    document.getElementById('page-' + t).style.display = t === tab ? '' : 'none';
    document.getElementById('tab-' + t).classList.toggle('active', t === tab);
  });
  if (tab === 'handling' && !chLoaded) loadCaseHandling();
  if (tab === 'jpmc') {
    // Restore saved auto-refresh timer when entering the tab
    if (!jpmcRefreshTimer) {
      const savedMs = getJpmcRefreshInterval();
      if (savedMs > 0) setJpmcSchedule(savedMs);
    }
    if (!jpmcLoaded) loadJpmcCases();
    if (!jpmcNewLoaded) loadJpmcNewCases();
    if (!_jpmcStats) loadJpmcStats();
  }
  if (tab === 'dashboard') initDashboard();
}

// ── JPMC Section collapse ─────────────────────────────────────────────────────
function toggleJpmcSection(section) {
  const bodyId = { stats: 'jpmcStatsBody', restore: 'jpmcRestoreBody', open: 'jpmcOpenBody' }[section];
  const btnId  = { stats: 'collapse-stats', restore: 'collapse-restore', open: 'collapse-open' }[section];
  const body = document.getElementById(bodyId);
  const btn  = document.getElementById(btnId);
  if (!body) return;
  const isCollapsed = body.classList.toggle('collapsed');
  if (btn) btn.classList.toggle('collapsed', isCollapsed);
  // If expanding stats and chart exists, re-render it
  if (section === 'stats' && !isCollapsed && _jpmcStats) {
    setTimeout(() => renderJpmcStatsChart(_jpmcStatsPeriod), 50);
  }
}

// ── JPMC Stats ────────────────────────────────────────────────────────────────
let _jpmcStats = null;
let _jpmcStatsPeriod = 'day7';
let _jpmcStatsChart = null;

async function loadJpmcStats() {
  const loader  = document.getElementById('jpmcStatsLoader');
  const content = document.getElementById('jpmcStatsContent');
  if (!loader) return;
  loader.style.display = 'flex';
  content.style.display = 'none';
  try {
    const res = await fetch('/api/jpmc-stats');
    if (!res.ok) throw new Error(await res.text());
    _jpmcStats = await res.json();
    const restore    = _jpmcStats.total;
    const totalAll   = _jpmcStats.totalAll || restore;
    const nonRestore = totalAll - restore;
    const restorePct    = totalAll > 0 ? ((restore    / totalAll) * 100).toFixed(1) : '0.0';
    const nonRestorePct = totalAll > 0 ? ((nonRestore / totalAll) * 100).toFixed(1) : '0.0';
    document.getElementById('statTotal').textContent       = restore.toLocaleString();
    document.getElementById('statTotalAll').textContent    = nonRestore.toLocaleString();
    document.getElementById('statRestorePct').textContent  = restorePct + '%';
    document.getElementById('statNonRestorePct').textContent = nonRestorePct + '%';
    document.getElementById('statTotalBadge').textContent  = totalAll.toLocaleString() + ' total';
    renderJpmcStatsChart(_jpmcStatsPeriod);
    loader.style.display = 'none';
    content.style.display = '';
  } catch(e) {
    loader.innerHTML = '<span style="color:#b00;font-size:12px">Failed to load stats: ' + e.message + '</span>';
  }
}

function setJpmcStatsPeriod(period) {
  _jpmcStatsPeriod = period;
  ['Day7','Day30','Day90','Week','Month'].forEach(id => {
    const btn = document.getElementById('statsBtn' + id);
    if (btn) btn.classList.toggle('active', ('day7,day30,day90,weekly,monthly').split(',')[['Day7','Day30','Day90','Week','Month'].indexOf(id)] === period);
  });
  if (_jpmcStats) renderJpmcStatsChart(period);
}

function renderJpmcStatsChart(period) {
  let labels, values;
  if (period === 'day7' || period === 'day30' || period === 'day90') {
    const days = period === 'day7' ? 7 : period === 'day30' ? 30 : 90;
    const allDays = Object.keys(_jpmcStats.daily).sort();
    const cutoff = allDays.slice(-days);
    labels = cutoff;
    values = cutoff.map(k => _jpmcStats.daily[k] || 0);
  } else {
    const data = _jpmcStats[period];
    labels = Object.keys(data).sort();
    values = labels.map(k => data[k]);
  }

  if (_jpmcStatsChart) { _jpmcStatsChart.destroy(); _jpmcStatsChart = null; }
  const ctx = document.getElementById('jpmcStatsChart');
  if (!ctx) return;

  const isDaily = period === 'day7' || period === 'day30' || period === 'day90';
  const colors  = { day7: 'rgba(0,82,204,0.75)', day30: 'rgba(0,82,204,0.75)', day90: 'rgba(0,82,204,0.75)', weekly: 'rgba(0,120,212,0.75)', monthly: 'rgba(124,77,255,0.75)' };
  const borders = { day7: '#0052cc', day30: '#0052cc', day90: '#0052cc', weekly: '#0078d4', monthly: '#7c4dff' };

  _jpmcStatsChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: colors[period],
        borderColor: borders[period],
        borderWidth: 1,
        borderRadius: 3
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => ' ' + c.parsed.y + ' cases' } },
        datalabels: { display: false }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 }, maxRotation: 45 } },
        y: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 10 } }, grid: { color: '#f0f0f0' } }
      }
    }
  });
}

// ── JPMC Restore Request ──────────────────────────────────────────────────────
let jpmcLoaded = false;
let jpmcData = null;
let jpmcTeamUsers = null;
let jpmcAssignCounts = {}; // tracks assignments this session for even distribution (individual assigns)
let jpmcRefreshTimer = null;

function getJpmcRefreshInterval() { return parseInt(localStorage.getItem('jpmcRefreshInterval') || '0'); }
function getJpmcAutoAssign()      { return localStorage.getItem('jpmcAutoAssign') === '1'; }

function setJpmcSchedule(ms) {
  clearJpmcSchedule();
  localStorage.setItem('jpmcRefreshInterval', String(ms));
  if (ms > 0) {
    jpmcRefreshTimer = setInterval(() => { jpmcLoaded = false; loadJpmcCases(); }, ms);
  }
  updateJpmcScheduleUI();
}

function clearJpmcSchedule() {
  if (jpmcRefreshTimer) { clearInterval(jpmcRefreshTimer); jpmcRefreshTimer = null; }
}

function updateJpmcScheduleUI() {
  const sel = document.getElementById('jpmcRefreshSel');
  if (sel) sel.value = String(getJpmcRefreshInterval());
  const aaChk = document.getElementById('jpmcAutoAssignChk');
  if (aaChk) aaChk.checked = getJpmcAutoAssign();
}

async function loadJpmcCases() {
  document.getElementById('jpmcLoader').style.display = 'flex';
  document.getElementById('jpmcContent').innerHTML = '';
  jpmcLoaded = false;
  try {
    const res = await fetch('/api/jpmc-cases');
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    jpmcData = data.cases;
    jpmcTeamUsers = data.teamUsers;
    jpmcAssignCounts = {};
    renderJpmcCases();
    // Restore the active timer selection in the newly rendered UI
    updateJpmcScheduleUI();
    jpmcLoaded = true;
    if (getJpmcAutoAssign() && jpmcData && jpmcData.length > 0) {
      await assignAllJpmcCases();
    }
  } catch(err) {
    document.getElementById('jpmcContent').innerHTML =
      '<div class="loading" style="display:flex"><div style="color:#b00">' + err.message + '</div></div>';
  } finally {
    document.getElementById('jpmcLoader').style.display = 'none';
  }
}

function renderJpmcCases() {
  const cases = jpmcData || [];
  const count = cases.length;
  const esc = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  // ── Build table HTML ──────────────────────────────────────────────────────
  let tableHtml = '';
  if (count === 0) {
    tableHtml = '<div style="padding:48px;text-align:center;color:#aaa"><i class="fa fa-circle-check" style="font-size:36px;display:block;margin-bottom:12px;color:#2da44e"></i>No unassigned cases</div>';
  } else {
    tableHtml = \`<table><thead><tr>
      <th>Case #</th><th>Subject</th><th>Contact</th><th>Created</th><th>Assign To</th>
    </tr></thead><tbody>\`;
    cases.forEach((c, i) => {
      const _cd = (c.CreatedDate || '').replace('+0000', 'Z');
      const _dt = new Date(_cd);
      const created = isNaN(_dt.getTime()) ? (c.CreatedDate || '\u2014') : _dt.toLocaleString();
      const contact = esc(c['Contact.Name'] || c['Contact.Email'] || '—');
      const firstName = esc(c['Contact.FirstName'] || '');
      const rawDesc = (c.Description || '').trim();
      const allMatches = rawDesc.match(/[01]_[a-z0-9]+/gi) || [];
      const kEntry = allMatches.length ? [...new Set(allMatches)].join(', ') : '';
      tableHtml += \`<tr id="jpmc-row-\${i}" data-fname="\${firstName}" data-kid="\${kEntry}" style="cursor:default" onmouseover="previewTemplate(this.dataset.fname,this.dataset.kid)">
        <td><a href="javascript:void(0)" style="color:#0078d4;font-weight:700;text-decoration:none"
             onclick="window.open('https://kaltura.lightning.force.com/lightning/r/Case/\${c.Id}/view')">\${esc(c.CaseNumber)}</a></td>
        <td>\${esc(c.Subject)}</td>
        <td>\${contact}</td>
        <td style="white-space:nowrap;font-size:11px;color:#888">\${created}</td>
        <td style="white-space:nowrap;position:relative">
          <button class="assign-btn" id="assignBtn-\${i}" onclick="toggleAssignPicker(event,\${i},'\${c.Id}')">
            <i class="fa fa-user-plus"></i> Assign <i class="fa fa-chevron-down" style="font-size:9px"></i>
          </button>
          <div class="assign-picker" id="assignPicker-\${i}" style="display:none"></div>
        </td>
      </tr>\`;
    });
    tableHtml += '</tbody></table>';
  }

  let html = \`
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
      <button id="collapse-restore" class="btn-collapse" onclick="toggleJpmcSection('restore')" title="Collapse/Expand"><i class="fa fa-chevron-down"></i></button>
      <span style="font-size:13px;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:.5px">
        <i class="fa fa-film" style="color:#0078d4"></i> Entry Restore Requests
      </span>
    </div>
    <div id="jpmcRestoreBody" class="collapse-body" style="max-height:2000px">
    <div class="jpmc-filter-bar">
      <span><i class="fa fa-filter" style="color:#0078d4"></i></span>
      <span>Account: <strong>J.P. Morgan Chase &amp; Co.</strong></span>
      <span>Status: <strong>New</strong></span>
      <span>Assigned To: <strong>Unassigned</strong></span>
      <span>Subject: <strong>Entry recovery request</strong></span>
    </div>
    <div style="display:flex;gap:16px;align-items:flex-start">
      <div class="card" style="flex:1;min-width:0">
        <div class="card-header">
          <div class="card-title">
            <i class="fa fa-film"></i> Entry Restore Requests
            <span class="badge badge-blue">\${count} case\${count !== 1 ? 's' : ''}</span>
          </div>
          <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
            <button class="btn" style="font-size:11px;background:#0052cc;color:#fff;border-color:#0052cc" onclick="assignAllJpmcCases()" \${count === 0 ? 'disabled' : ''}>
              <i class="fa fa-wand-magic-sparkles"></i> Assign All
            </button>
            <button class="btn btn-secondary" style="font-size:11px" onclick="jpmcLoaded=false;loadJpmcCases()">
              <i class="fa fa-rotate-right"></i> Refresh
            </button>
            <div class="jpmc-schedule-wrap">
              <i class="fa fa-clock" style="color:#6554c0;font-size:12px"></i>
              <select id="jpmcRefreshSel" class="jpmc-schedule-sel" onchange="setJpmcSchedule(parseInt(this.value))">
                <option value="0">No auto-refresh</option>
                <option value="1800000">Every 30 min</option>
                <option value="3600000">Every 1 hr</option>
                <option value="7200000">Every 2 hrs</option>
                <option value="14400000">Every 4 hrs</option>
              </select>
            </div>
            <label class="jpmc-autoassign-label" style="white-space:nowrap">
              <input type="checkbox" id="jpmcAutoAssignChk"
                onchange="localStorage.setItem('jpmcAutoAssign', this.checked ? '1' : '0')">
              <i class="fa fa-wand-magic-sparkles" style="font-size:11px;color:#0052cc"></i> Auto-assign
            </label>
          </div>
        </div>
        <div class="tbl-wrap">\${tableHtml}</div>
      </div>
      \${buildPoolCard()}
    </div>
    </div>
  \`;

  document.getElementById('jpmcContent').innerHTML = html;
  updatePoolCard();
  loadSFUserForTemplate();
}

function buildPoolCard() {
  const users = jpmcTeamUsers || [];
  const options = users.map(function(u) {
    return '<option value="' + u.id + '" data-name="' + u.name + '">' + u.name + '</option>';
  }).join('');
  return \`<div class="card" id="poolCard" style="width:220px;flex-shrink:0">
    <div class="card-header" style="padding:10px 14px">
      <div class="card-title" style="font-size:12px">
        <i class="fa fa-users" style="color:#0052cc"></i> Assignee Pool
      </div>
      <span class="badge badge-blue" id="poolCount">0</span>
    </div>
    <div style="padding:10px 12px;border-bottom:1px solid #f0f0f0;display:flex;gap:6px">
      <select id="poolAddSelect" class="jpmc-schedule-sel" style="flex:1;min-width:0">
        <option value="">Add person...</option>
        \${options}
      </select>
      <button class="btn" style="background:#0078d4;color:#fff;padding:5px 10px;font-size:13px;border-radius:6px;min-width:30px" onclick="poolAddPerson()">
        <i class="fa fa-plus"></i>
      </button>
    </div>
    <div id="poolList" style="padding:6px 0;min-height:40px"></div>
  </div>\`;
}

function updatePoolCard() {
  const poolUsers = jpmcPoolLoad();
  const countEl = document.getElementById('poolCount');
  if (countEl) countEl.textContent = poolUsers.length;
  const listEl = document.getElementById('poolList');
  if (!listEl) return;
  if (!poolUsers.length) {
    listEl.innerHTML = '<div style="padding:10px 14px;font-size:12px;color:#aaa">No assignees added</div>';
    return;
  }
  listEl.innerHTML = poolUsers.map(function(u) {
    return '<div style="display:flex;align-items:center;justify-content:space-between;padding:7px 14px;font-size:12px;border-bottom:1px solid #f5f5f5">' +
      '<span><i class="fa fa-user" style="color:#0078d4;margin-right:6px;font-size:11px"></i>' + u.name + '</span>' +
      '<button data-uid="' + u.id + '" onclick="poolRemovePerson(this.dataset.uid)" style="background:none;border:none;cursor:pointer;color:#bbb;font-size:13px;padding:0 2px;line-height:1" title="Remove">&times;</button>' +
      '</div>';
  }).join('');
}

function poolAddPerson() {
  const sel = document.getElementById('poolAddSelect');
  if (!sel || !sel.value) return;
  localStorage.setItem('jpmcPool_' + sel.value, '1');
  sel.value = '';
  updatePoolCard();
}

function poolRemovePerson(userId) {
  localStorage.setItem('jpmcPool_' + userId, '0');
  updatePoolCard();
}

function jpmcPoolSave() {} // no longer used by checkboxes

function jpmcUpdatePoolCount() {
  const countEl = document.getElementById('poolCount');
  if (countEl) countEl.textContent = jpmcPoolLoad().length;
}

function jpmcPoolLoad() {
  const users = jpmcTeamUsers || [];
  return users.filter(function(u) { return localStorage.getItem('jpmcPool_' + u.id) === '1'; });
}

function buildPoolRow() { return ''; } // legacy no-op


let _sfUser = null;
async function loadSFUserForTemplate() {
  try {
    const res = await fetch('/api/sf-user');
    _sfUser = await res.json();
  } catch(e) {
    _sfUser = { name: 'Kaltura Support', title: 'Customer Support' };
  }
}

function buildTemplateText(firstName, kalturaId) {
  const n = String.fromCharCode(10);
  const name  = (firstName || '[Contact Name]');
  const sigName = _sfUser ? _sfUser.name : '[Your Name]';
  const entryLine = kalturaId && kalturaId !== '-' ? 'Entry ID(s): ' + kalturaId + n + n : '';
  return 'Hi ' + name + ',' + n + n +
    'Thanks for reaching out to Kaltura Customer Care.' + n + n +
    entryLine +
    "I'm happy to confirm that the requested entries have been successfully restored. Please check on your end and let us know if everything looks good." + n + n +
    'Please note that restore requests are handled on a best effort basis.' + n + n +
    'I will now be marking the case as closed.' + n + n +
    'Should you notice anything else or need further assistance, feel free to reach out.' + n + n +
    'Best regards,' + n + n +
    sigName + n +
    'Kaltura Customer Care | Kaltura Inc.' + n +
    'Support: https://support.kaltura.com' + n + n +
    'Knowledge Base: https://knowledge.kaltura.com' + n +
    'Website: https://www.kaltura.com' + n +
    'Status Alerts: https://status.kaltura.com' + n + n +
    'Get your support questions answered before login \u2014 try our AI Support Assistant in the bottom left corner!' + n + n +
    'The age of Agentic Avatars is here: https://corp.kaltura.com/agentic-avatars/';
}

function previewTemplate(firstName, kalturaId) {
  const el = document.getElementById('templatePreview');
  if (!el) return;
  el.value = buildTemplateText(firstName, kalturaId);
}

function pickEvenAssignee() {
  const pool = jpmcPoolLoad();
  if (!pool.length) return null;
  const minCount = Math.min(...pool.map(u => jpmcAssignCounts[u.id] || 0));
  const candidates = pool.filter(u => (jpmcAssignCounts[u.id] || 0) === minCount);
  return candidates[Math.floor(Math.random() * candidates.length)];
}

function toggleAssignPicker(evt, idx, caseId) {
  evt.stopPropagation();
  // Close any other open pickers
  document.querySelectorAll('.assign-picker').forEach(p => {
    if (p.id !== 'assignPicker-' + idx) p.style.display = 'none';
  });
  const picker = document.getElementById('assignPicker-' + idx);
  if (!picker) return;
  if (picker.style.display !== 'none') { picker.style.display = 'none'; return; }

  const pool = jpmcPoolLoad();
  if (!pool.length) {
    toast('Check at least one person under "Assign to" before assigning', 'error');
    return;
  }
  picker.innerHTML = pool.map(u =>
    \`<div class="assign-picker-item" data-uid="\${u.id}" data-uname="\${u.name}" data-idx="\${idx}" data-cid="\${caseId}" onclick="doJpmcPickAssign(event,this)">\${u.name}</div>\`
  ).join('');
  picker.style.display = 'block';
}

// Close all pickers when clicking elsewhere
document.addEventListener('click', () => {
  document.querySelectorAll('.assign-picker').forEach(p => { p.style.display = 'none'; });
});

async function doJpmcPickAssign(evt, el) {
  evt.stopPropagation();
  const idx    = el.dataset.idx;
  const caseId = el.dataset.cid;
  const userId = el.dataset.uid;
  const userName = el.dataset.uname;
  document.getElementById('assignPicker-' + idx).style.display = 'none';
  const btn = document.getElementById('assignBtn-' + idx);
  btn.disabled = true;
  btn.innerHTML = '<i class="fa fa-spinner fa-spin"></i>';
  try {
    const res = await fetch('/api/jpmc-assign', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ caseId, userId })
    });
    if (!res.ok) throw new Error(await res.text());
    jpmcAssignCounts[userId] = (jpmcAssignCounts[userId] || 0) + 1;
    btn.className = 'assign-btn assigned';
    btn.innerHTML = '<i class="fa fa-circle-check"></i> ' + userName;
    toast('Assigned to ' + userName, 'success');
    setTimeout(() => {
      const row = document.getElementById('jpmc-row-' + idx);
      if (row) { row.style.transition = 'opacity .4s'; row.style.opacity = '0.3'; }
    }, 1500);
  } catch(err) {
    btn.disabled = false;
    btn.innerHTML = '<i class="fa fa-user-plus"></i> Assign <i class="fa fa-chevron-down" style="font-size:9px"></i>';
    toast('Error: ' + err.message, 'error');
  }
}

async function assignAllJpmcCases() {
  const pool = jpmcPoolLoad();
  console.log('[AssignAll] pool (' + pool.length + '):', pool.map(function(p){return p.name;}).join(', '));
  if (!pool.length) {
    toast('Check at least one person under "Assign to" before assigning', 'error');
    return;
  }
  const cases = jpmcData || [];
  const pending = cases.map((c, i) => ({ idx: i, caseId: (c.Id || '').trim() }))
    .filter(({ idx, caseId }) => {
      if (!caseId || (caseId.length !== 15 && caseId.length !== 18)) {
        console.warn('[AssignAll] Skipping case at index ' + idx + ': invalid Id "' + caseId + '"');
        return false;
      }
      const btn = document.getElementById('assignBtn-' + idx);
      return btn && !btn.classList.contains('assigned');
    });
  if (!pending.length) { toast('All cases already assigned', 'error'); return; }

  // Build evenly distributed picks: each pool member gets exactly floor(n/k) or ceil(n/k) cases.
  // Use Fisher-Yates shuffle (not sort-random which is biased in V8).
  function fisherYates(arr) {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
  }
  const n = pending.length, k = pool.length;
  const base = Math.floor(n / k), extra = n % k;
  // Randomly decide who gets the extra case(s)
  const orderedPool = fisherYates(pool.slice());
  const pickList = [];
  orderedPool.forEach((u, i) => {
    const quota = base + (i < extra ? 1 : 0);
    for (let j = 0; j < quota; j++) pickList.push(u);
  });
  // Shuffle the pick list so the distribution is random across case positions
  fisherYates(pickList);
  const assignments = pending.map((p, i) => ({ ...p, pick: pickList[i] }));

  // Disable all buttons immediately
  assignments.forEach(({ idx }) => {
    const btn = document.getElementById('assignBtn-' + idx);
    if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fa fa-spinner fa-spin"></i>'; }
  });

  let successCount = 0;
  for (const { idx, caseId, pick } of assignments) {
    const btn = document.getElementById('assignBtn-' + idx);
    try {
      const res = await fetch('/api/jpmc-assign', {
        method: 'POST',
        headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ caseId, userId: pick.id })
      });
      if (!res.ok) throw new Error(await res.text());
      jpmcAssignCounts[pick.id] = (jpmcAssignCounts[pick.id] || 0) + 1;
      if (btn) { btn.className = 'assign-btn assigned'; btn.innerHTML = '<i class="fa fa-circle-check"></i> ' + pick.name; }
      successCount++;
    } catch(err) {
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fa fa-user-plus"></i> Assign'; }
      toast('Error on case ' + (idx + 1) + ': ' + err.message, 'error');
    }
  }
  if (successCount > 0) {
    const dist = pool.map(function(p){ return p.name + ' ×' + (jpmcAssignCounts[p.id]||0); }).join('  |  ');
    toast('Assigned ' + successCount + ' cases — ' + dist, 'success');
    setTimeout(() => {
      document.querySelectorAll('[id^="jpmc-row-"]').forEach(row => {
        row.style.transition = 'opacity .4s'; row.style.opacity = '0.3';
      });
    }, 1500);
  }
}

// ── JPMC New Cases section ────────────────────────────────────────────────────
let _jpmcNewRaw = [];
let jpmcNewLoaded = false;

function filterJpmcNewCases() {
  const sel = document.getElementById('jpmcNewStatusFilter');
  const filter = sel ? sel.value : 'new-all';
  let cases;
  if (filter === 'all') {
    cases = _jpmcNewRaw;
  } else if (filter === 'new-all') {
    cases = _jpmcNewRaw.filter(c => c.Status === 'New' || c.Status === 'New Assigned');
  } else {
    cases = _jpmcNewRaw.filter(c => c.Status === filter);
  }
  const badge = document.getElementById('jpmcNewBadge');
  if (badge) badge.textContent = cases.length + ' case' + (cases.length !== 1 ? 's' : '');
  renderJpmcNewRows(cases);
}

function renderJpmcNewRows(cases) {
  const content = document.getElementById('jpmcNewContent');
  if (!content) return;
  const esc = s => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  if (!cases.length) {
    content.innerHTML = '<div style="padding:40px;text-align:center;color:#aaa"><i class="fa fa-circle-check" style="font-size:32px;display:block;margin-bottom:10px;color:#2da44e"></i>No cases match this filter</div>';
  } else {
    const statusColor = s => {
      const l = (s||'').toLowerCase();
      if (l === 'new') return '#b00020';
      if (l.includes('customer responded') || l.includes('review')) return '#e65c00';
      if (l.includes('awaiting')) return '#0052cc';
      if (l === 'resolved') return '#2da44e';
      return '#555';
    };
    let rows = cases.map((c, i) => {
      const _cd2 = (c.CreatedDate || '').replace('+0000', 'Z');
      const _dt2 = new Date(_cd2);
      const created = isNaN(_dt2.getTime()) ? (c.CreatedDate || '\u2014') : _dt2.toLocaleString();
      const contact = esc(c['Contact.Name'] || c['Contact.Email'] || '—');
      const firstName = esc(c['Contact.FirstName'] || '');
      const rawDesc = (c.Description || '').trim();
      const allMatches = rawDesc.match(/[01]_[a-z0-9]+/gi) || [];
      const kalturaId = allMatches.length ? [...new Set(allMatches)].join(', ') : '-';
      const owner = esc(c['Assigned_To__r.Name'] || c['Owner.Name'] || '-');
      const isUnassigned = !c['Assigned_To__c'] || c['Assigned_To__c'] === '';
      const ownerCell = isUnassigned
        ? '<span style="color:#aaa;font-size:11px">Unassigned</span>'
        : '<span style="color:#1a7f37;font-weight:600">' + owner + '</span>';
      const statusBadge = '<span style="font-size:10px;font-weight:700;color:' + statusColor(c.Status) + ';background:#f8f9ff;border:1px solid currentColor;padding:1px 6px;border-radius:10px;white-space:nowrap">' + esc(c.Status) + '</span>';
      const sfUrl = 'https://kaltura.lightning.force.com/lightning/r/Case/' + c.Id + '/view';
      return '<tr id="jpmcnew-row-' + i + '" data-fname="' + firstName + '" data-kid="' + esc(kalturaId) + '" style="cursor:default" onmouseover="previewTemplate(this.dataset.fname,this.dataset.kid)">' +
        '<td><a href="' + sfUrl + '" target="_blank" style="color:#0078d4;font-weight:700;text-decoration:none">' + esc(c.CaseNumber) + '</a></td>' +
        '<td>' + esc(c.Subject) + '</td>' +
        '<td>' + contact + '</td>' +
        '<td><span style="font-family:monospace;font-size:12px;font-weight:700;color:#0052cc;background:#e8f0fe;padding:2px 8px;border-radius:6px">' + esc(kalturaId) + '</span></td>' +
        '<td>' + ownerCell + '</td>' +
        '<td>' + statusBadge + '</td>' +
        '<td style="white-space:nowrap;font-size:11px;color:#888">' + created + '</td>' +
        '<td style="white-space:nowrap"><button class="respond-btn" id="respondBtn-new-' + i + '"' +
        ' data-cid="' + c.Id + '" data-cnum="' + esc(c.CaseNumber) + '" data-fname="' + firstName + '" data-contact="' + contact + '" data-btnid="respondBtn-new-' + i + '"' +
        ' onclick="openRespondModalFromBtn(this)"><i class="fa fa-reply"></i> Respond</button></td>' +
        '</tr>';
    }).join('');
    content.innerHTML = '<table><thead><tr>' +
      '<th>Case #</th><th>Subject</th><th>Contact</th>' +
      '<th>Kaltura Entry ID</th><th>Owner</th><th>Status</th><th>Created</th><th>Respond</th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table>';
  }
  content.style.display = '';
}

async function loadJpmcNewCases() {
  const loader = document.getElementById('jpmcNewLoader');
  const content = document.getElementById('jpmcNewContent');
  const badge = document.getElementById('jpmcNewBadge');
  if (!loader || !content) return;
  loader.style.display = 'flex';
  content.style.display = 'none';

  try {
    const res = await fetch('/api/jpmc-new-cases');
    if (!res.ok) throw new Error(await res.text());
    const cases = await res.json();
    _jpmcNewRaw = cases;
    jpmcNewLoaded = true;
    filterJpmcNewCases();
    loader.style.display = 'none';
  } catch(err) {
    if (badge) badge.textContent = 'Error';
    loader.style.display = 'none';
    content.innerHTML = '<div style="padding:24px;color:#b00">' + err.message + '</div>';
    content.style.display = '';
  }
}

// ── Respond modal ─────────────────────────────────────────────────────────────
let _respondState = null;
const RESPOND_TEMPLATE = (name) => \`Hi \${name || 'there'},

Thanks for reaching out to Kaltura Customer Care.

I'm happy to confirm that the requested entries have been successfully restored. Please check on your end and let us know if everything looks good.

Please note that restore requests are handled on a best effort basis.

I will now be marking the case as closed.

Should you notice anything else or need further assistance, feel free to reach out.

Best regards,

\${_sfUser ? _sfUser.name : '[Your Name]'}
Kaltura Customer Care | Kaltura Inc.
Support: https://support.kaltura.com

Knowledge Base: https://knowledge.kaltura.com
Website: https://www.kaltura.com
Status Alerts: https://status.kaltura.com

Get your support questions answered before login — try our AI Support Assistant in the bottom left corner!

The age of Agentic Avatars is here: https://corp.kaltura.com/agentic-avatars/\`;

function openRespondModalFromBtn(el) {
  openRespondModal(el.dataset.cid, el.dataset.cnum, el.dataset.fname, el.dataset.contact, el.id);
}

function openRespondModal(caseId, caseNum, firstName, contactName, btnId) {
  _respondState = { caseId, caseNum, firstName, btnId };
  document.getElementById('respondCaseNum').textContent = caseNum;
  document.getElementById('respondContact').textContent = contactName;
  // Use text from the editable template card if available, else build fresh
  const tplCard = document.getElementById('templatePreview');
  const tplText = (tplCard && tplCard.value && tplCard.value !== 'Hover a case row to preview...')
    ? tplCard.value
    : buildTemplateText(firstName, '');
  document.getElementById('respondPreview').value = tplText;
  const btn = document.getElementById('respondConfirmBtn');
  btn.disabled = false;
  btn.innerHTML = '<i class="fa fa-paper-plane"></i> Send & Close Case';
  document.getElementById('respondModal').classList.add('open');
}

async function confirmRespond() {
  if (!_respondState) return;
  const { caseId, btnId } = _respondState;
  const commentBody = document.getElementById('respondPreview').value;
  const confirmBtn = document.getElementById('respondConfirmBtn');
  confirmBtn.disabled = true;
  confirmBtn.innerHTML = '<i class="fa fa-spinner fa-spin"></i> Sending...';

  try {
    const res = await fetch('/api/jpmc-respond', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ caseId, commentBody })
    });
    if (!res.ok) throw new Error(await res.text());
    closeModal('respondModal');
    toast('Response sent and case closed', 'success');
    // Mark the button as done
    const btn = document.getElementById(btnId);
    if (btn) {
      btn.className = 'respond-btn done';
      btn.innerHTML = '<i class="fa fa-circle-check"></i> Sent';
      btn.disabled = true;
      // Fade the row
      const row = btn.closest('tr');
      if (row) { row.style.transition = 'opacity .4s'; row.style.opacity = '0.35'; }
    }
  } catch(err) {
    confirmBtn.disabled = false;
    confirmBtn.innerHTML = '<i class="fa fa-paper-plane"></i> Send & Close Case';
    toast('Error: ' + err.message, 'error');
  }
}

// ── Case Handling ─────────────────────────────────────────────────────────────
let chData = null;
let chLoaded = false;
let chChart = null;

async function loadCaseHandling() {
  document.getElementById('chLoader').style.display = 'flex';
  document.getElementById('chContent').innerHTML = '';
  chLoaded = false;
  try {
    const res = await fetch('/api/case-handling');
    if (!res.ok) throw new Error(await res.text());
    chData = await res.json();
    renderCaseHandling(chData);
    chLoaded = true;
  } catch(e) {
    toast('Error loading case handling: ' + e.message, 'error');
  } finally {
    document.getElementById('chLoader').style.display = 'none';
  }
}

function renderCaseHandling(data) {
  const { persons, fetchedAt } = data;
  const totToday = persons.reduce((s, p) => s + p.today, 0);
  const totWeek  = persons.reduce((s, p) => s + p.week,  0);
  const totMonth = persons.reduce((s, p) => s + p.month, 0);

  let rows = '';
  persons.forEach((p, i) => {
    const cell = (val, period, cls) => val > 0
      ? \`<td class="num"><span class="ch-num \${cls}" onclick="showCases(\${i},'\${period}')">\${val}</span></td>\`
      : \`<td class="num-zero">—</td>\`;
    rows += \`<tr>
      <td class="name-col">\${p.name}</td>
      \${cell(p.today,'today','')}
      \${cell(p.week,'week','ch-week')}
      \${cell(p.month,'month','ch-month')}
    </tr>\`;
  });

  const updated = new Date(fetchedAt).toLocaleTimeString();
  document.getElementById('chContent').innerHTML = \`
    <div class="ch-stats">
      <div class="ch-stat-card today">
        <div class="ch-stat-label">Today</div>
        <div class="ch-stat-value">\${totToday}</div>
        <div class="ch-stat-sub">cases handled</div>
      </div>
      <div class="ch-stat-card week">
        <div class="ch-stat-label">This Week</div>
        <div class="ch-stat-value">\${totWeek}</div>
        <div class="ch-stat-sub">cases handled</div>
      </div>
      <div class="ch-stat-card month">
        <div class="ch-stat-label">This Month</div>
        <div class="ch-stat-value">\${totMonth}</div>
        <div class="ch-stat-sub">cases handled</div>
      </div>
    </div>
    <div class="card grid-full" style="margin-bottom:16px">
      <div class="card-header">
        <span class="card-title"><i class="fa fa-chart-bar" style="color:#0052cc"></i> Cases Handled by Agent</span>
        <div style="display:flex;align-items:center;gap:6px">
          <button id="ch-period-today" class="ch-period-btn active" onclick="setCHPeriod('today')">Today</button>
          <button id="ch-period-week"  class="ch-period-btn"        onclick="setCHPeriod('week')">This Week</button>
          <button id="ch-period-month" class="ch-period-btn"        onclick="setCHPeriod('month')">This Month</button>
          <button id="pin-case-handling" class="btn-pin" onclick="pinItem('case-handling')" title="Pin to Dashboard"><i class="fa fa-thumbtack"></i></button>
        </div>
      </div>
      <div style="padding:16px 20px;height:340px"><canvas id="chartCH"></canvas></div>
    </div>
    <div class="card grid-full">
      <div class="card-header">
        <span class="card-title"><i class="fa fa-comments" style="color:#0078d4"></i> Case Handling by Agent</span>
        <div style="display:flex;align-items:center;gap:10px">
          <span style="font-size:11px;color:#888">Updated \${updated} &mdash; click a number to see cases</span>
          <button class="btn btn-refresh" style="font-size:11px;padding:5px 10px" onclick="chLoaded=false;loadCaseHandling()">
            <i class="fa fa-rotate-right"></i> Refresh
          </button>
        </div>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>Agent</th>
            <th class="num">Today</th>
            <th class="num">This Week</th>
            <th class="num">This Month</th>
          </tr></thead>
          <tbody>
            \${rows}
            <tr class="sum-row">
              <td>Total</td>
              <td class="num">\${totToday}</td>
              <td class="num">\${totWeek}</td>
              <td class="num">\${totMonth}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  \`;
  renderCHChart(persons, 'today');
  updatePinButton('case-handling', dashPinned.has('case-handling'));
}

function renderCHChart(persons, period) {
  const key = period;
  const active = persons.filter(p => p[key] > 0).sort((a, b) => b[key] - a[key]);

  if (chChart) { chChart.destroy(); chChart = null; }
  const ctx = document.getElementById('chartCH').getContext('2d');

  if (!active.length) {
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.fillStyle = '#aaa';
    ctx.font = '13px Segoe UI';
    ctx.textAlign = 'center';
    ctx.fillText('No cases for this period', ctx.canvas.width / 2, 120);
    return;
  }

  const colors = { today: 'rgba(0,120,212,0.8)', week: 'rgba(0,82,204,0.8)', month: 'rgba(124,77,255,0.8)' };
  const borders = { today: '#0078d4', week: '#0052cc', month: '#7c4dff' };

  chChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: active.map(p => p.name),
      datasets: [{ data: active.map(p => p[key]),
        backgroundColor: colors[period],
        borderColor: borders[period],
        borderWidth: 1, borderRadius: 4 }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false, clip: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => ' ' + c.parsed.x + ' cases' } },
        datalabels: { anchor: 'end', align: 'start', color: '#fff', font: { weight: 'bold', size: 12 }, formatter: v => v }
      },
      scales: {
        x: { beginAtZero: true, ticks: { stepSize: 1 }, grid: { color: '#f0f0f0' } },
        y: { grid: { display: false } }
      }
    }
  });
}

let currentCHPeriod = 'today';
function setCHPeriod(period) {
  currentCHPeriod = period;
  ['today','week','month'].forEach(p => {
    const btn = document.getElementById('ch-period-' + p);
    if (btn) btn.classList.toggle('active', p === period);
  });
  if (chData) renderCHChart(chData.persons, period);
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
let dashGrid = null;
let dashInited = false;
let dashPinned = new Set(JSON.parse(localStorage.getItem('sf-dash-pins') || '[]'));

const CHART_DEFS = {
  'esc-actionables': {
    title: 'Escalated Actionables by Owner',
    icon: 'fa-chart-bar', iconColor: '#e65c00',
    color: 'rgba(230,92,0,0.8)', border: '#e65c00',
    getData: () => {
      if (!cachedData) return null;
      return cachedData.summary.filter(r => r.escActionable > 0)
        .sort((a,b) => b.escActionable - a.escActionable)
        .map(r => ({ label: r.name, value: r.escActionable }));
    }
  },
  'black-flags': {
    title: 'Black Flags by Owner',
    icon: 'fa-flag', iconColor: '#111',
    color: 'rgba(176,0,32,0.8)', border: '#b00020',
    getData: () => {
      if (!cachedData) return null;
      return cachedData.summary.filter(r => r.blackFlags > 0)
        .sort((a,b) => b.blackFlags - a.blackFlags)
        .map(r => ({ label: r.name, value: r.blackFlags }));
    }
  },
  'case-handling': {
    title: 'Case Handling by Agent',
    icon: 'fa-comments', iconColor: '#0078d4',
    color: 'rgba(0,120,212,0.8)', border: '#0078d4',
    getData: () => {
      if (!chData) return null;
      const key = currentCHPeriod || 'today';
      return chData.persons.filter(p => p[key] > 0)
        .sort((a,b) => b[key] - a[key])
        .map(p => ({ label: p.name, value: p[key] }));
    }
  },
  'jpmc-stats': {
    title: 'JPMC Support Tickets 2026',
    icon: 'fa-chart-line', iconColor: '#0052cc',
    color: 'rgba(0,82,204,0.8)', border: '#0052cc',
    chartType: 'bar',
    getData: () => {
      if (!_jpmcStats) return null;
      const period = _jpmcStatsPeriod || 'day7';
      let labels, values;
      if (period === 'day7' || period === 'day30' || period === 'day90') {
        const days = period === 'day7' ? 7 : period === 'day30' ? 30 : 90;
        labels = Object.keys(_jpmcStats.daily).sort().slice(-days);
        values = labels.map(k => _jpmcStats.daily[k] || 0);
      } else {
        const data = _jpmcStats[period];
        labels = Object.keys(data).sort();
        values = labels.map(k => data[k]);
      }
      return labels.map((l, i) => ({ label: l, value: values[i] }));
    }
  }
};

function renderHBarChart(canvasId, def) {
  const items = def.getData();
  const existing = Chart.getChart(canvasId);
  if (existing) existing.destroy();
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  if (!items || !items.length) {
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.fillStyle = '#aaa'; ctx.font = '13px Segoe UI'; ctx.textAlign = 'center';
    ctx.fillText('No data', ctx.canvas.width / 2, 80);
    return;
  }
  const isVertical = def.chartType === 'bar';
  new Chart(ctx, {
    type: 'bar',
    data: { labels: items.map(i => i.label), datasets: [{ data: items.map(i => i.value),
      backgroundColor: def.color, borderColor: def.border, borderWidth: 1, borderRadius: 4 }] },
    options: {
      indexAxis: isVertical ? 'x' : 'y', responsive: true, maintainAspectRatio: false, clip: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => ' ' + (isVertical ? c.parsed.y : c.parsed.x) } },
        datalabels: { display: false }
      },
      scales: {
        x: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 10 }, maxRotation: isVertical ? 45 : 0 }, grid: { color: isVertical ? '#f0f0f0' : 'transparent' } },
        y: { beginAtZero: true, grid: { color: isVertical ? '#f0f0f0' : 'transparent' }, ticks: { font: { size: 10 } } }
      }
    }
  });
}

function initDashboard() {
  if (dashInited) return;
  dashInited = true;
  dashGrid = GridStack.init({
    column: 12, cellHeight: 70, margin: 10, animate: true,
    resizable: { handles: 'se,s,e' }
  }, '#dashGrid');
  dashGrid.on('resizestop', (e, el) => {
    const id = el.getAttribute('gs-id');
    if (id) setTimeout(() => renderHBarChart('dash-canvas-' + id, CHART_DEFS[id]), 80);
  });
  dashGrid.on('change', () => {
    const layout = dashGrid.save(false);
    localStorage.setItem('sf-dash-layout', JSON.stringify(layout));
  });
  // Restore previously pinned items
  const layout = JSON.parse(localStorage.getItem('sf-dash-layout') || '[]');
  dashPinned.forEach(id => {
    const saved = layout.find(l => l.id === id) || {};
    _addWidget(id, saved);
  });
  updateDashEmpty();
}

function _addWidget(id, opts) {
  const def = CHART_DEFS[id];
  if (!def) return;
  const canvasId = 'dash-canvas-' + id;
  const html = \`<div class="dash-widget-inner">
    <div class="dash-widget-header">
      <span><i class="fa \${def.icon}" style="color:\${def.iconColor}"></i>&nbsp;\${def.title}</span>
      <button class="dash-unpin-btn" onclick="unpinItem('\${id}')" title="Remove"><i class="fa fa-xmark"></i></button>
    </div>
    <div class="dash-widget-chart"><canvas id="\${canvasId}"></canvas></div>
  </div>\`;
  dashGrid.addWidget({ id, content: html,
    w: opts.w || 6, h: opts.h || 5, x: opts.x, y: opts.y });
  setTimeout(() => renderHBarChart(canvasId, def), 250);
}

function pinItem(id) {
  if (dashPinned.has(id)) { unpinItem(id); return; }
  updatePinButton(id, true);
  const wasInited = dashInited;
  // Add to set AFTER capturing wasInited — so initDashboard (called by switchTab)
  // doesn't include this item, avoiding a double-add on first pin.
  switchTab('dashboard');
  dashPinned.add(id);
  localStorage.setItem('sf-dash-pins', JSON.stringify([...dashPinned]));
  _addWidget(id, {});
  updateDashEmpty();
}

function unpinItem(id) {
  dashPinned.delete(id);
  localStorage.setItem('sf-dash-pins', JSON.stringify([...dashPinned]));
  updatePinButton(id, false);
  if (dashGrid) {
    const el = dashGrid.engine.nodes.find(n => n.id === id)?.el;
    if (el) dashGrid.removeWidget(el);
  }
  updateDashEmpty();
}

function updatePinButton(id, pinned) {
  const btn = document.getElementById('pin-' + id);
  if (!btn) return;
  btn.classList.toggle('pinned', pinned);
  btn.title = pinned ? 'Unpin from Dashboard' : 'Pin to Dashboard';
}

function updateDashEmpty() {
  const hasPins = dashPinned.size > 0;
  document.getElementById('dashEmpty').style.display = hasPins ? 'none' : '';
  document.getElementById('dashGrid').style.display = hasPins ? '' : 'none';
}

function showCases(idx, period) {
  if (!chData) return;
  const p = chData.persons[idx];
  const cases = p[period + 'Cases'];
  const label = { today: 'Today', week: 'This Week', month: 'This Month' }[period];
  document.getElementById('chModalTitle').textContent = p.name + ' \u2014 ' + label;
  document.getElementById('chModalBody').innerHTML = cases.length
    ? cases.map(c => \`<div class="ch-case-row"><i class="fa fa-ticket-simple" style="color:#0078d4"></i> <b>\${c}</b></div>\`).join('')
    : '<p style="color:#aaa;text-align:center;padding:20px">No cases</p>';
  document.getElementById('chModal').classList.add('open');
}

// Restore pin button visual state on load
dashPinned.forEach(id => updatePinButton(id, true));

loadData();
</script>
</body></html>`;

// ── HTTP server ───────────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  const { pathname } = url.parse(req.url);
  const method = req.method;

  const send = (code, body, ct = 'application/json') => {
    res.writeHead(code, { 'Content-Type': ct, 'Access-Control-Allow-Origin': '*' });
    res.end(typeof body === 'string' ? body : JSON.stringify(body));
  };

  const body = () => new Promise(r => { let d = ''; req.on('data', c => d += c); req.on('end', () => r(d)); });

  if (method === 'GET' && pathname === '/') return send(200, HTML, 'text/html');

  if (method === 'GET' && pathname === '/api/data') {
    try {
      lastData = await querySF();
      send(200, lastData);
    }
    catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'POST' && pathname === '/api/send-email') {
    try {
      const { to, subject } = JSON.parse(await body());
      if (!lastData) return send(400, 'No data loaded yet — please refresh first', 'text/plain');
      const html = buildEmailHtml(lastData);
      const msg = await sendEmail(to, subject, html);
      send(200, msg, 'text/plain');
    } catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'GET' && pathname === '/api/schedule') {
    try { send(200, await getSchedule()); }
    catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'POST' && pathname === '/api/schedule') {
    try {
      const { freq, time, days, to } = JSON.parse(await body());
      const msg = await createSchedule(freq, time, days, to);
      send(200, msg, 'text/plain');
    } catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'GET' && pathname === '/api/case-handling') {
    try { send(200, await queryCaseHandling()); }
    catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'DELETE' && pathname === '/api/schedule') {
    try { send(200, await deleteSchedule(), 'text/plain'); }
    catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'GET' && pathname === '/api/jpmc-cases') {
    try {
      const [cases, userIds] = await Promise.all([queryJpmcCases(), getTeamUserIds()]);
      const teamUsers = TEAM_NAMES
        .filter(n => userIds[n])
        .map(n => ({ name: DISPLAY[n] || n, sfName: n, id: userIds[n] }))
        .sort((a, b) => a.name.localeCompare(b.name));
      send(200, { cases, teamUsers });
    } catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'POST' && pathname === '/api/jpmc-assign') {
    try {
      const { caseId, userId } = JSON.parse(await body());
      await assignCaseAssignedTo(caseId, userId);
      send(200, 'ok', 'text/plain');
    } catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'GET' && pathname === '/api/jpmc-stats') {
    try { send(200, await queryJpmcStats()); }
    catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'GET' && pathname === '/api/sf-user') {
    try { send(200, await getSFCurrentUser()); }
    catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'GET' && pathname === '/api/jpmc-new-cases') {
    try {
      send(200, await queryJpmcNewCases());
    } catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  if (method === 'POST' && pathname === '/api/jpmc-respond') {
    try {
      const { caseId, commentBody, firstName } = JSON.parse(await body());
      await respondToCase(caseId, firstName, commentBody);
      send(200, 'ok', 'text/plain');
    } catch(e) { send(500, e.message, 'text/plain'); }
    return;
  }

  send(404, 'Not found', 'text/plain');
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`\n  SF Report running at: http://localhost:${PORT}\n`);
  exec(`start http://localhost:${PORT}`);
});
