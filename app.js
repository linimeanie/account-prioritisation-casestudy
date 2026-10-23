import { prioritise, classifyRows } from './logic.js';
import { DEFAULT_SETTINGS } from './defaults.js';

// ---------- state ----------
const clone = x => JSON.parse(JSON.stringify(x));
const state = {
  data: { accounts: null, signals: null, sdrs: null },
  names: {},
  settings: clone(DEFAULT_SETTINGS),
  baseSettings: clone(DEFAULT_SETTINGS),
  settingsSource: 'built-in defaults',
  result: null,
  prevAssigned: null,
  newlyIn: new Set(),
  tab: 'board',
  expanded: {},
  funnelOpen: null,
  fixOpen: null,
  search: '', colFilters: {}, sort: { key: 'score', dir: 'desc' },
  error: ''
};
const app = document.getElementById('app');
const overlay = document.getElementById('overlay');

// ---------- helpers ----------
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = (n, d = 1) => (n === null || n === undefined || isNaN(n) ? '–' : Number(n).toLocaleString('en-US', { maximumFractionDigits: d, minimumFractionDigits: d }));
const fmt0 = n => Number(n).toLocaleString('en-US');
const ready = () => state.data.accounts && state.data.signals && state.data.sdrs;
const tierOfScore = s => (s >= 20 ? 'A' : s >= 8 ? 'B' : 'C');
// prospects (filled slots with no signal yet) get their own tier
const tier = x => (typeof x === 'object' ? (x.prospect === true || x.prospect === 'yes' ? 'P' : tierOfScore(x.score)) : tierOfScore(x));
// Tiers use food temperature instead of sales jargon ("warm" already means an existing relationship in sales).
// Each one stands for an action: baked potato = call first, cream soup = call + email, sushi = email only.
const tierName = { A: '🥔 Baked potato', B: '🍲 Cream soup', C: '🍣 Sushi', P: '🍞 Bread & butter' };
const tierAction = { A: 'call first', B: 'call, then follow up by email', C: 'one email is enough', P: 'research first, then a cold first touch' };
const tierLegend = () => `<div class="tier-legend"><span><span class="tier A">${tierName.A}</span> hot, score 20+: ${tierAction.A}</span><span><span class="tier B">${tierName.B}</span> warm, 8 to 20: ${tierAction.B}</span><span><span class="tier C">${tierName.C}</span> served cold, under 8: ${tierAction.C}</span><span><span class="tier P">${tierName.P}</span> strong fit, no signal yet: research, then a cold first touch</span></div>`;
const statusKind = st => st === 'Assigned' ? 'good' : st.startsWith('Back on fire') || st.startsWith('In sequence') ? 'info' : st.startsWith('Not routed') ? 'warn' : st.startsWith('Excluded') ? 'bad' : '';

const S = {
  rows: () => state.settings,
  get: (sec, key) => { const r = state.settings.find(r => r.section === sec && String(r.key) === String(key)); return r ? r.value : undefined; },
  base: (sec, key) => { const r = state.baseSettings.find(r => r.section === sec && String(r.key) === String(key)); return r ? r.value : undefined; },
  set: (sec, key, value) => {
    const r = state.settings.find(r => r.section === sec && String(r.key) === String(key));
    if (r) r.value = value; else state.settings.push({ section: sec, key, value });
  },
  del: (sec, key) => { state.settings = state.settings.filter(r => !(r.section === sec && String(r.key) === String(key))); },
  entries: sec => state.settings.filter(r => r.section === sec).map(r => [String(r.key), r.value]),
  changes: () => {
    const k = r => `${r.section}|${r.key}`;
    const a = new Map(state.baseSettings.map(r => [k(r), String(r.value ?? '')]));
    const b = new Map(state.settings.map(r => [k(r), String(r.value ?? '')]));
    let n = 0;
    for (const [key, v] of b) if (a.get(key) !== v) n++;
    for (const key of a.keys()) if (!b.has(key)) n++;
    return n;
  }
};
const realEntries = sec => S.entries(sec).filter(([k]) => k !== 'ACC-0000');

function toast(msg) {
  const t = document.createElement('div');
  t.className = 'toast'; t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 3200);
}

// ---------- files ----------
async function readFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf);
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {});
}
async function addFiles(fileList) {
  state.error = '';
  const unknown = [];
  for (const f of fileList) {
    try {
      const rows = await readFile(f);
      const kind = classifyRows(rows);
      if (!kind) { unknown.push(f.name); continue; }
      if (kind === 'settings') {
        state.settings = rows.map(r => ({ section: r.section, key: r.key, value: r.value ?? '', note: r.note }));
        state.baseSettings = clone(state.settings);
        state.settingsSource = f.name;
      } else state.data[kind] = rows;
      state.names[kind] = f.name;
    } catch (e) { unknown.push(f.name); }
  }
  state.data = { ...state.data }; // new reference, so the saved copy is refreshed
  if (unknown.length) state.error = `Not recognised: ${unknown.join(', ')}. Expected the accounts, signals and sdrs files (Excel or CSV), plus optionally settings.`;
  if (state.result) run(); else render();
}
async function loadSample() {
  try {
    const files = [];
    for (const n of ['accounts', 'signals', 'sdrs']) {
      const r = await fetch(`data/${n}.xlsx`);
      if (!r.ok) throw new Error();
      files.push(new File([await r.blob()], `${n}.xlsx`));
    }
    await addFiles(files);
  } catch { state.error = 'The case data is not bundled with this site. Please drop the three files instead.'; render(); }
}

// ---------- weeks ----------
// The same files are replayed as consecutive Mondays: each week only sees the signals dated up to its
// Monday. Accounts assigned in one week are remembered, so the next week can apply the week-to-week rules.
const isoDay = v => {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number' || /^\d{5}(\.\d+)?$/.test(String(v))) return new Date((Number(v) - 25569) * 86400000).toISOString().slice(0, 10);
  const d = new Date(String(v).trim()); return isNaN(d) ? String(v) : d.toISOString().slice(0, 10);
};
const addDays = (iso, n) => new Date(Date.parse(iso + 'T00:00:00Z') + n * 86400000).toISOString().slice(0, 10);
const WEEKS = 3;
function weekDates() {
  const last = isoDay(S.get('general', 'as_of_date')) || new Date().toISOString().slice(0, 10);
  return Array.from({ length: WEEKS }, (_, i) => addDays(last, -7 * (WEEKS - 1 - i)));
}
function runWeeks() {
  const dates = weekDates(), worked = {};
  state.weeks = dates.map((d, i) => {
    const signals = state.data.signals.filter(x => !x.signal_date || isoDay(x.signal_date) <= d);
    const accounts = state.data.accounts.map(a => isoDay(a.last_touch_date) > d ? { ...a, last_touch_date: '' } : a);
    const settings = state.settings.map(r => r.section === 'general' && r.key === 'as_of_date' ? { ...r, value: d } : r);
    const r = prioritise(accounts, signals, state.data.sdrs, settings, { worked: { ...worked } });
    r.date = d; r.index = i;
    const from = i ? dates[i - 1] : addDays(d, -7);
    r.newSignals = state.data.signals.filter(x => { const sd = isoDay(x.signal_date); return sd > from && sd <= d; });
    r.knownSignals = signals.length;
    if (!i) {
      // the Monday before week 1, so week 1 can also show what moved accounts up or down
      const b = addDays(d, -7);
      r.baseline = prioritise(state.data.accounts.map(a => isoDay(a.last_touch_date) > b ? { ...a, last_touch_date: '' } : a),
        state.data.signals.filter(x => !x.signal_date || isoDay(x.signal_date) <= b), state.data.sdrs,
        state.settings.map(x => x.section === 'general' && x.key === 'as_of_date' ? { ...x, value: b } : x));
      r.baseline.date = b;
    }
    for (const l of r.lists) worked[l.account_id] = { date: d, sdr: l.sdr };
    for (const x of r.scored) if (x.status.startsWith('Back on fire')) worked[x.account_id] = { ...worked[x.account_id], date: d };
    return r;
  });
  state.week = Math.min(state.week ?? 0, WEEKS - 1);
  state.result = state.weeks[state.week];
}
function setWeek(i) { state.week = i; state.result = state.weeks[i]; state.newlyIn = new Set(); render(); }

// ---------- run ----------
let runTimer;
function scheduleRun() { clearTimeout(runTimer); runTimer = setTimeout(run, 220); }
function run() {
  if (!ready()) return render();
  const before = state.result ? new Set(state.result.lists.map(l => l.account_id)) : null;
  const posBefore = state.result ? new Map(state.result.lists.map(l => [l.account_id, `${l.sdr}#${l.rank}`])) : null;
  runWeeks();
  const after = new Set(state.result.lists.map(l => l.account_id));
  if (before) {
    state.newlyIn = new Set([...after].filter(x => !before.has(x)));
    const out = [...before].filter(x => !after.has(x)).length;
    const acc = n => `${n} account${n === 1 ? '' : 's'}`;
    const moved = state.result.lists.filter(l => posBefore.has(l.account_id) && posBefore.get(l.account_id) !== `${l.sdr}#${l.rank}`).length;
    toast(state.newlyIn.size || out ? `Lists updated: ${acc(state.newlyIn.size)} in, ${acc(out)} out, ${moved} moved` : moved ? `Same ${after.size} accounts, ${acc(moved)} changed rank` : 'No change: this setting doesn\'t affect any account on this week\'s lists');
  } else state.newlyIn = new Set();
  render();
  if (document.querySelector('.drawer.settings')) refreshSettingsMarks();
}

// ---------- render ----------
// ---------- survive a page reload ----------
// Kept in sessionStorage: it survives refreshes in the same tab, and is gone when the tab is closed.
const STORE_DATA = 'prio-data-v1', STORE_UI = 'prio-ui-v1';
let savedDataRef = null;
function saveState() {
  try {
    if (savedDataRef !== state.data) { sessionStorage.setItem(STORE_DATA, JSON.stringify({ data: state.data, names: state.names })); savedDataRef = state.data; }
    sessionStorage.setItem(STORE_UI, JSON.stringify({ built: !!state.result, tab: state.tab, week: state.week, settings: state.settings, baseSettings: state.baseSettings,
      settingsSource: state.settingsSource, colFilters: state.colFilters, sort: state.sort, expanded: state.expanded }));
  } catch { /* storage full or blocked: the page still works, it just won't survive a reload */ }
}
function restoreState() {
  try {
    const d = JSON.parse(sessionStorage.getItem(STORE_DATA) || 'null'), u = JSON.parse(sessionStorage.getItem(STORE_UI) || 'null');
    if (d && d.data) { state.data = d.data; state.names = d.names || {}; savedDataRef = state.data; }
    if (u) Object.assign(state, { week: u.week || 0, tab: u.tab || 'board', settings: u.settings || state.settings, baseSettings: u.baseSettings || state.baseSettings,
      settingsSource: u.settingsSource || state.settingsSource, colFilters: u.colFilters || {}, sort: u.sort || state.sort, expanded: u.expanded || {} });
    // settings saved before a new rule existed get the new default rows added (e.g. renewal bands)
    const fill = rows => { const have = new Set(rows.map(r => `${r.section}|${r.key}`)); return [...rows, ...clone(DEFAULT_SETTINGS).filter(r => !have.has(`${r.section}|${r.key}`))]; };
    if (state.settingsSource === 'built-in defaults') { state.settings = fill(state.settings); state.baseSettings = fill(state.baseSettings); }
    return !!(u && u.built && ready());
  } catch { return false; }
}
function clearSaved() { try { sessionStorage.removeItem(STORE_DATA); sessionStorage.removeItem(STORE_UI); } catch {} savedDataRef = null; }

function render() {
  app.innerHTML = state.result ? renderResults() : renderIntro();
  bindCommon();
  saveState();
}

function sdrName(id) {
  const r = (state.data.sdrs || []).find(r => String(r.sdr_id) === String(id).trim());
  return r ? String(r.name).split(' ')[0] : String(id).trim();
}
function territoryText() {
  // before the SDR file is loaded only IDs are known, so show a count instead of cryptic IDs
  return S.entries('territory').map(([reg, ids]) => {
    const list = String(ids || '').split(',').map(x => x.trim()).filter(Boolean);
    if (!list.length) return `${reg} → nobody`;
    return `${reg} → ${state.data.sdrs ? list.map(sdrName).join(' + ') : `${list.length} SDR${list.length > 1 ? 's' : ''}`}`;
  });
}

// A short guide note at the top of each page section. The avatar is a small whale.
const WHALE = '<svg viewBox="0 0 64 64"><circle cx="32" cy="32" r="32" fill="#efe9fd"/><path d="M30 20c-1-4-4-6-6-6m6 6c1-4 4-6 6-6m-6 6v-6" stroke="#a285ef" stroke-width="2.4" fill="none" stroke-linecap="round"/><path d="M9 38c0-8 8-13 19-13 10 0 17 5 19 11l4-5c2-2 5-1 5 1l-2 7 3 5c1 2-1 4-3 3l-6-2c-4 5-11 8-20 8C16 53 9 47 9 38z" fill="#6a3be5"/><path d="M12 42c3 5 9 8 16 8 6 0 11-2 15-5" stroke="#efe9fd" stroke-width="2" fill="none" stroke-linecap="round" opacity=".7"/><circle cx="20" cy="35" r="2.2" fill="#0e022d"/></svg>';
const guideNote = (text, extra = '') => `<div class="panel-note"><div class="note-from"><div class="avatar" aria-hidden="true">${WHALE}</div><div class="bubble"><div class="who"><b>Note</b><span>How to read this</span></div><p>${text}</p>${extra}</div></div></div>`;

function renderIntro() {
  const f = (kind, label, desc, optional) => {
    const ok = kind === 'settings' ? state.settingsSource !== 'built-in defaults' : !!state.data[kind];
    const n = kind === 'settings' ? (ok ? `${state.settings.length} rows` : 'Defaults') : ok ? `${fmt0(state.data[kind].length)} rows` : 'Waiting';
    return `<div class="file ${ok ? 'ok' : ''} ${optional ? 'opt' : ''}"><b>${label}</b><span class="state">${ok ? '✓ ' : ''}${esc(n)}</span><span class="desc">${desc}</span></div>`;
  };
  const changes = S.changes();
  const boosts = S.entries('boost').filter(([, v]) => Number(v) !== 1).length;
  return `
  <div class="wrap">
    <section class="hero">
      <h1>Each SDR's best accounts for the week,<br><em>with the reason behind every pick</em></h1>
      ${guideNote(`Drop in the three case files and the prototype cleans the data, scores all 500 accounts, routes the best 80 to the four SDRs and explains every decision. The signal weights and scoring rules live in Leadership settings on the right: change any of them, before or after the run, and the lists re-rank instantly.`)}
    </section>
    <div class="intro-grid">
      <div class="card pad span-2 stack">
        <div class="step-label">Step 1 · Data</div>
        <h2>Drop the case files</h2>
        <p class="card-lead">accounts, signals and sdrs, in any order and with any file name. Each file is recognised by its columns.</p>
        <label class="drop" id="drop">
          <input type="file" id="fileInput" multiple accept=".xlsx,.xls,.csv" hidden>
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 16V4m0 0l-4 4m4-4l4 4"/><path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>
          <strong>Drop Excel files here or click to choose</strong>
          <span class="muted" style="font-size:13px">.xlsx or .csv · select several at once</span>
        </label>
        <div class="files">
          ${f('accounts', 'Accounts', 'Target accounts')}
          ${f('signals', 'Signals', 'Last 90 days')}
          ${f('sdrs', 'SDRs', 'Team and capacity')}
          ${f('settings', 'Settings', 'Optional', true)}
        </div>
        ${state.error ? `<div class="err">${esc(state.error)}</div>` : ''}
        <div class="cta-row push">
          <button class="btn primary" id="runBtn" ${ready() ? '' : 'disabled'}>Build this week's lists →</button>
          ${state.noSample ? '' : '<button class="btn ghost" id="sampleBtn">Use the case data</button>'}
        </div>
      </div>
      <div class="card pad stack">
        <div class="step-label">Step 2 · Optional</div>
        <h2>Leadership settings</h2>
        <p class="card-lead">Every rule is a setting, not code. Change it now or after the run and watch the lists update.</p>
        <dl class="sum">
          <dt>Weeks</dt><dd>${weekDates().map(dayShort).join(' · ')}</dd>
          <dt>Signal window</dt><dd>${esc(S.get('general', 'lookback_days'))} days, halving every ${esc(S.get('general', 'half_life_days'))}</dd>
          <dt>Signal weights</dt><dd>${S.entries('signal_weight').length} signals, ${S.entries('signal_type_default').length} type defaults</dd>
          <dt>Segment boosts</dt><dd>${boosts ? `${boosts} active` : 'None active'}</dd>
          <dt>Pinned / excluded</dt><dd>${realEntries('pin').length} / ${realEntries('exclude').length} accounts</dd>
          <dt>Territories</dt><dd>${territoryText().map(esc).join('<br>')}</dd>
        </dl>
        <div class="cta-row push">
          <button class="btn" data-open-settings>Review and edit settings</button>
          ${changes ? `<span class="chip warn">${changes} change${changes > 1 ? 's' : ''}</span>` : `<span class="chip">${esc(state.settingsSource === 'built-in defaults' ? 'Defaults' : state.settingsSource)}</span>`}
        </div>
      </div>
      <div class="card how-card"><div class="n">1</div><b>Clean</b><p>Unify country names, fix regions that contradict the country, and drop duplicate, undated or stale signals. Every fix is logged, so nothing changes silently.</p></div>
      <div class="card how-card"><div class="n">2</div><b>Score</b><p>Intent (how strong and fresh the signals are) × fit (devices in the field) × CRM stage (customers and open deals stay with AEs) × leadership boost.</p></div>
      <div class="card how-card"><div class="n">3</div><b>Route</b><p>Best score first, within each SDR's territory, until their weekly capacity is full. Accounts touched in the last two weeks rest, so nobody is over-contacted.</p></div>
    </div>
    <footer>Prototype. All companies and people are fictional. The same logic runs as an <a href="https://github.com/linimeanie/account-prioritisation-casestudy/tree/main/n8n" target="_blank" rel="noopener">n8n workflow</a>; this page is a transparent front end over it.</footer>
  </div>`;
}

function kpis(r) {
  const fixes = r.issues.filter(i => i.area !== 'routing').length;
  const apac = r.scored.filter(s => s.status.startsWith('Not routed: no SDR')).length;
  const reason = r.scored.filter(s => s.signals_used > 0 && s.breakdown.stageMult > 0 && !s.status.startsWith('Excluded')).length;
  const k = (v, l, s) => `<div class="card kpi"><div class="v num">${v}</div><div class="l">${l}</div><div class="s">${s}</div></div>`;
  return `<div class="kpis">
    ${k(r.lists.length, 'accounts assigned', esc(r.team.map(t => `${t.name.split(' ')[0]} ${t.filled}`).join(' · ')))}
    ${k(fmt0(r.summary.accounts), 'accounts scored', `${bridgeRows(r).live} with a live signal SDRs can work`)}
    ${k(`${r.summary.signals_used}<span class="of"> / ${r.summary.signals_in}</span>`, 'signals used', 'rest: duplicate, undated, stale')}
    ${k(fixes, 'data fixes logged', 'each one listed in Data fixes')}
    ${k(apac, 'left for leadership', apac ? 'signals, but no SDR in region' : 'every region is covered')}
  </div>`;
}

function renderResults() {
  const r = state.result;
  const tabs = [
    ['board', 'SDR lists', r.lists.length],
    ['weeks', 'Week over week'],
    ['funnel', 'How 500 became 80'],
    ['explain', 'How the score works'],
    ['fixes', 'Data fixes', r.issues.filter(i => i.area !== 'routing').length],
    ['messages', 'Monday messages'],
    ['all', 'All accounts', r.scored.length]
  ];
  const changes = S.changes();
  return `
  <div class="bar"><div class="wrap">
    <span class="weeksw">${state.weeks.map((w, i) => `<button class="${i === state.week ? 'on' : ''}" data-week="${i}">${weekLabel(i)} · ${dayShort(w.date)}</button>`).join('')}</span>
    <span class="chip">${esc(state.settingsSource)}${changes ? ` + ${changes} change${changes > 1 ? 's' : ''}` : ''}</span>
    <span class="spacer"></span>
    <button class="btn small" data-open-settings>⚙ Leadership settings</button>
    <button class="btn small" id="dlBtn">⬇ Download Excel</button>
    <button class="btn small ghost" id="resetBtn">Start over</button>
  </div></div>
  <div class="wrap">
    ${kpis(r)}
    <nav class="tabs">${tabs.map(([k, l, c]) => `<button class="tab ${state.tab === k ? 'active' : ''}" data-tab="${k}">${l}${c !== undefined ? `<span class="count">${c}</span>` : ''}</button>`).join('')}</nav>
    <section id="panel">${panel(state.tab)}</section>
    <footer>All companies and people are fictional. Scores are relative within this week's run. The food tiers (baked potato ≥ 20, cream soup ≥ 8, sushi below, bread & butter for best-fit prospects) are a reading aid only and don't change the ranking.</footer>
  </div>`;
}

function panel(tab) {
  return { board: pBoard, weeks: pWeeks, funnel: pFunnel, explain: pExplain, fixes: pFixes, messages: pMessages, all: pAll }[tab]();
}

// ---------- SDR board ----------
function accCard(l, s) {
  const isTop = l.rank === 1;
  const first = (l.why_now || '').split('; ')[0];
  return `<div class="acc ${isTop ? 'top' : ''} ${state.newlyIn.has(l.account_id) ? 'new' : ''}" data-acc="${esc(l.account_id)}">
    ${isTop ? '<div class="badge">Start here</div>' : ''}
    <div class="row1"><span class="rk">#${l.rank}</span><span class="nm">${esc(l.company_name)}</span>${l.pinned ? '<span class="chip info">Pinned</span>' : ''}<span class="food" title="${esc(tierName[tier(l)].slice(3))}: ${tierAction[tier(l)]}">${tierName[tier(l)].slice(0, 2)}</span></div>
    <div class="meta-row"><span class="sub">${esc(l.country)} · ${esc(l.industry)}${l.crm_stage === 'Closed lost' ? ' · closed lost' : ''}</span></div>
    <div class="why">${esc(first)}${(l.why_now || '').includes('; ') ? ` <span class="faint">+${l.why_now.split('; ').length - 1} more</span>` : ''}</div>
    ${isTop ? `<div class="angle">${esc(l.suggested_angle)}</div>` : ''}
  </div>`;
}
function pBoard() {
  const r = state.result;
  // from week 2: carried-over accounts that heated up sit above the new ones, same height in every column
  const carried = t => r.scored.filter(x => x.worked && x.worked.sdr === t.name && (x.status.startsWith('In sequence') || x.status.startsWith('Back on fire')));
  const heated = t => carried(t).filter(x => x.status.startsWith('Back on fire')).sort((a, b) => b.score - a.score);
  const maxHeated = state.week ? Math.max(...r.team.map(t => Math.min(4, heated(t).length))) : 0;
  const carriedBlock = t => {
    if (!state.week) return '';
    const h = heated(t), all = carried(t), shownH = h.slice(0, 4);
    return `<div class="carried" style="min-height:${44 + maxHeated * 46}px"><div class="sec-h">Carried over <span>${all.length}</span></div>
      ${shownH.map(x => `<div class="mini" data-acc="${esc(x.account_id)}"><span class="food" title="${esc(tierName[tier(x)].slice(3))}: ${tierAction[tier(x)]}">${tierName[tier(x)].slice(0, 2)}</span><span class="nm">${esc(x.company_name)}</span><span class="why">heated up: ${esc(lower(x.refire[0] || ''))}</span></div>`).join('')}
      <div class="faint more-line">${h.length > 4 ? `+${h.length - 4} more heated up · ` : ''}${all.length - h.length} in your sequence</div></div>
      <div class="sec-h">New this week <span>${r.lists.filter(l => l.sdr === t.name).length}</span></div>`;
  };
  return `${guideNote(state.week ? 'One column per SDR. On top, accounts carried over from earlier weeks that heated up with a new signal: call those again. Below, this week\'s new accounts, best first.' : 'One column per SDR, best account first. Click any account to see how its score was built and why it landed with that rep.', tierLegend())}
  <div class="board">${r.team.map(t => {
    const rows = r.lists.filter(l => l.sdr === t.name);
    const open = state.expanded[t.id];
    const shown = open ? rows : rows.slice(0, 8);
    const pct = t.cap ? Math.round(100 * t.filled / t.cap) : 0;
    return `<div class="col" data-col="${t.id}">
      <div class="col-h">
        <div class="name">${esc(t.name)}</div>
        <div class="meta">${esc(t.territories.join(', ') || 'No territory')} · ${t.tenure} months</div><div class="meta">${esc(t.language)}</div>
        <div class="cap"><i style="width:${pct}%"></i></div>
        <div class="cap-l"><span>${t.filled} of ${t.cap} slots</span><span>${rows.length || state.week ? `🥔 ${rows.filter(x => tier(x) === 'A').length + (state.week ? heated(t).filter(x => tier(x) === 'A').length : 0)} to call first` : ''}</span></div>
      </div>
      ${carriedBlock(t)}
      ${shown.map(l => accCard(l)).join('') || '<div class="note">No qualifying accounts in this territory.</div>'}
      ${rows.length > 8 ? `<button class="btn small ghost more" data-expand="${t.id}">${open ? 'Show fewer' : `Show all ${rows.length}`}</button>` : ''}
    </div>`;
  }).join('')}</div>`;
}

// ---------- funnel ----------
function funnelSteps(res = state.result) {
  const sc = res.scored;
  const by = f => sc.filter(f);
  const lb = S.get('general', 'lookback_days');
  return [
    { k: 'all', label: 'Accounts in the list', desc: 'Every account sourced for the motion', items: sc, color: 'var(--faint)' },
    { k: 'crm', label: 'Already owned by sales or CS', desc: 'Customers and open deals belong to AEs and CS', items: by(s => s.status.startsWith('Excluded: CRM')), minus: true, color: 'var(--bad)' },
    { k: 'lead', label: 'Excluded by leadership', desc: 'Accounts or segments switched off in settings', items: by(s => s.status.startsWith('Excluded by')), minus: true, color: 'var(--bad)' },
    { k: 'seq', label: 'Carried over from earlier weeks', desc: 'Already on a rep\'s list, so they take no new slot', items: by(s => s.status.startsWith('In sequence') || s.status.startsWith('Back on fire')), minus: true, color: 'var(--info)' },
    { k: 'nosig', label: `No signal in ${lb} days`, desc: 'Good companies, but no reason to call this week', items: by(s => s.status.startsWith('No signal')), minus: true, color: 'var(--warn)' },
    { k: 'cool', label: 'Recently touched', desc: `Contacted in the last ${S.get('general', 'recent_touch_days')} days, so they rest`, items: by(s => s.status.startsWith('Cool')), minus: true, color: 'var(--warn)' },
    { k: 'nosdr', label: 'No SDR covers the region', desc: 'A staffing decision, flagged for leadership', items: by(s => s.status.startsWith('Not routed: no SDR')), minus: true, color: 'var(--info)' },
    { k: 'cap', label: 'Waiting for next week', desc: 'Qualified, but reps are full: next week\'s pipeline', items: by(s => s.status.includes('at capacity')), minus: true, color: 'var(--info)' },
    { k: 'ok', label: 'New this week', desc: by(s => s.prospect).length ? `Best accounts within territory and capacity, incl. ${by(s => s.prospect).length} best-fit prospects` : 'Best accounts within territory and capacity', items: by(s => s.status === 'Assigned'), color: 'var(--accent)', final: true }
  ];
}
function pFunnel() {
  const steps = funnelSteps();
  const prev = state.weeks[state.week - 1];
  const before = prev ? Object.fromEntries(funnelSteps(prev).map(x => [x.k, x.items.length])) : null;
  const total = steps[0].items.length || 1;
  const open = steps.find(x => x.k === state.funnelOpen);
  const delta = x => {
    if (!before) return '';
    const d = x.items.length - (before[x.k] || 0);
    return `<span class="delta ${d > 0 ? 'pos' : d < 0 ? 'neg' : ''}">${d === 0 ? '±0' : `${d > 0 ? '+' : '−'}${Math.abs(d)}`} vs last week</span>`;
  };
  const r = state.result;
  const intro = `Every one of the ${total} accounts ends up somewhere, for a stated reason, and the split changes every week. This Monday ${r.knownSignals} signals are known${prev ? `, ${r.newSignals.length} of them new since last week` : ''}. Click a step to see its accounts.`;
  return `${guideNote(intro)}
  <div class="funnel">${steps.filter(x => x.items.length || !x.minus || (before && before[x.k])).map(x => `
    <div class="card fstep ${x.final ? 'final' : ''}" data-funnel="${x.k}">
      <div class="fl"><b>${x.minus ? '− ' : x.final ? '= ' : ''}${x.label}</b><span>${x.desc}</span></div>
      <div class="fbar"><i style="width:${Math.max(1.5, 100 * x.items.length / total)}%;background:${x.color}"></i></div>
      <div class="fc">${x.items.length}${delta(x)}</div>
    </div>`).join('')}</div>
  ${open ? `<div class="card pad" style="margin-top:16px"><h3>${open.label} <span class="faint">(${open.items.length})</span></h3>${accTable(open.items.slice(0, 300))}</div>` : ''}`;
}

function accTable(items) {
  return `<div class="tbl-wrap"><table><thead><tr><th>Company</th><th>Region</th><th>Industry</th><th>CRM stage</th><th class="r">Score</th><th>Status</th></tr></thead><tbody>
  ${items.map(s => `<tr class="click" data-acc="${esc(s.account_id)}"><td><b>${esc(s.company_name)}</b><div class="faint mono">${esc(s.account_id)}</div></td><td>${esc(s.region || '–')}</td><td>${esc(s.industry)}</td><td>${esc(s.crm_stage)}</td><td class="r num">${fmt(s.score)}</td><td><span class="chip ${statusKind(s.status)}">${esc(statusGroup(s))}</span>${s.assigned_to ? ` <span class="faint">${esc(s.assigned_to)} #${s.rank}</span>` : ''}</td></tr>`).join('')}
  </tbody></table></div>`;
}

// ---------- explainer ----------
function pExplain() {
  const types = {};
  for (const s of state.data.signals) (types[s.signal_type] = types[s.signal_type] || new Map()).set(s.signal_detail, (types[s.signal_type].get(s.signal_detail) || 0) + 1);
  const wDetail = Object.fromEntries(S.entries('signal_weight'));
  const wType = Object.fromEntries(S.entries('signal_type_default'));
  const typeOrder = Object.keys(types).sort((a, b) => (Number(wType[b]) || 0) - (Number(wType[a]) || 0));
  const hl = Number(S.get('general', 'half_life_days')) || 30, lb = Number(S.get('general', 'lookback_days')) || 90;
  const pts = Array.from({ length: 31 }, (_, i) => { const d = i * lb / 30; return [d, Math.pow(.5, d / hl)]; });
  const path = pts.map(([d, v], i) => `${i ? 'L' : 'M'}${(d / lb) * 300},${100 - v * 90}`).join(' ');
  const tbl = (sec, label) => `<table><thead><tr><th>${label}</th><th class="r">Multiplier</th></tr></thead><tbody>${S.entries(sec).map(([k, v]) => `<tr><td>${esc(k)}</td><td class="r num">×${esc(v)}</td></tr>`).join('')}</tbody></table>`;
  const devSteps = S.entries('fit_devices').map(([k, v]) => [Number(k), v]).sort((a, b) => b[0] - a[0]);
  return `${guideNote('The score answers one question: which accounts should this rep call this week? Each part below is a setting leadership can change.')}
  <div class="formula">
    <div class="card fx"><div class="t">Intent</div><p>Is the account moving right now? Each signal's weight × its freshness.</p></div><div class="op">×</div>
    <div class="card fx"><div class="t">Fit</div><p>How valuable is a win? Devices in the field, plus provider and size.</p></div><div class="op">×</div>
    <div class="card fx"><div class="t">CRM stage</div><p>Is it an SDR's account at all? Customers and open deals get 0.</p></div><div class="op">×</div>
    <div class="card fx"><div class="t">Boost</div><p>Leadership's lever to push or switch off a segment this week.</p></div><div class="op">=</div>
    <div class="card fx result"><div class="t">Score</div><p>Multiplied, not added: no buying activity means a score of 0.</p></div>
  </div>
  <div class="card ex-wide">
    <h3>1. Signals are weighted by what they say, not just their type</h3>
    <p class="ex-lead">The type labels mislead: a single blog visit is filed under "Website visit (pricing or docs page)", next to three pricing-page visits. So each signal text gets its own weight from 0 to 10. The closer to a buying decision, the higher: a renewal is a switching window, a pricing visit is active evaluation, a blog visit is curiosity.</p>
    <div class="wcols">
    ${typeOrder.map(t => {
      const all = [...types[t].entries()];
      const isRenewal = d => /renews? in ~?\s*\d+\s*months?/i.test(d) || (/renew/i.test(d) && /\bQ[1-4]\s*20\d\d\b/i.test(d));
      const dated = all.filter(([d]) => isRenewal(d));
      const own = all.filter(([d]) => !isRenewal(d) && wDetail[d] !== undefined).sort((a, b) => Number(wDetail[b[0]]) - Number(wDetail[a[0]]));
      const rest = all.filter(([d]) => !isRenewal(d) && wDetail[d] === undefined);
      const bar = (label, w, n) => `<div class="wbar"><span class="wl">${label} <span class="faint">(${n})</span></span><span class="b"><i style="width:${w * 10}%"></i></span><span class="v">${w}</span></div>`;
      return `<div class="wgroup"><div class="grp">${esc(t)}</div>${dated.map(([d, n]) => `<div class="wbar"><span class="wl">${esc(d)} <span class="faint">(${n})</span></span><span class="faint" style="font-size:12px">by renewal date</span><span class="v">5–10</span></div>`).join('')}${own.map(([d, n]) => bar(esc(d), Number(wDetail[d]), n)).join('')}
        ${rest.length ? bar(`${own.length ? 'Other' : 'All'} ${rest.length} variant${rest.length > 1 ? 's' : ''}, e.g. <i>${esc(rest[0][0])}</i>`, Number(wType[t] ?? 3), rest.reduce((x, [, n]) => x + n, 0)) : ''}</div>`;
    }).join('')}
    </div>
  </div>
  <div class="ex-grid">
    <div class="card"><h3>2. Fresh signals count more</h3>
      <p>A signal's weight halves every <b>${hl} days</b> and drops out after <b>${lb} days</b>. A weekly list is about <i>why now</i>: last week's pricing visit beats one from two months ago.</p>
      <p><b>Except renewals:</b> they get more urgent as the date approaches, so they never fade. The renewal date is the signal date + N months ("renews in ~4 months") or the middle of the quarter ("Q1 2027" = 15 Feb 2027). Weight by months to renewal: ${renewalBandsText()}. Once the date has passed, the signal is dropped and flagged.</p>
      <svg viewBox="-12 -14 340 136" class="decay"><line x1="0" y1="100" x2="300" y2="100" stroke="var(--border)"/><path d="${path}" fill="none" stroke="var(--accent)" stroke-width="2.5"/>
        ${[0, hl, Math.min(2 * hl, lb), lb].map(d => `<text x="${(d / lb) * 300}" y="116" font-size="10" fill="var(--faint)" text-anchor="middle">${d}d</text><circle cx="${(d / lb) * 300}" cy="${100 - Math.pow(.5, d / hl) * 90}" r="3" fill="var(--accent)"/><text x="${(d / lb) * 300 + 6}" y="${94 - Math.pow(.5, d / hl) * 90}" font-size="10" fill="var(--muted)">${Math.round(Math.pow(.5, d / hl) * 100)}%</text>`).join('')}
      </svg>
    </div>
    <div class="card"><h3>3. Breadth beats repetition</h3>
      <p>Within one signal type, the strongest counts fully and each repeat counts <b>${Math.round(100 * (S.get('general', 'extra_signal_same_type') ?? .5))}%</b>, so repeating the same activity adds less than a new kind of activity.</p>
      <p>Each additional <i>type</i> adds <b>${Math.round(100 * (S.get('general', 'multi_type_bonus') ?? .15))}%</b>. Pricing visits plus engineering hires plus a new CTO means several teams are moving, which is what a real buying committee looks like.</p>
    </div>
    <div class="card"><h3>4. Fit: devices first</h3>
      <p>Connectivity is billed per connected device, so devices in the field are the best proxy for deal size. Unknown values count as ×1, so missing data never punishes an account.</p>
      <div class="tbl-wrap"><table><thead><tr><th>Devices in the field</th><th class="r">Multiplier</th></tr></thead><tbody>${devSteps.map(([k, v], i) => `<tr><td>${i === 0 ? `${fmt0(k)}+` : `${fmt0(k)} – ${fmt0(devSteps[i - 1][0] - 1)}`}</td><td class="r num">×${esc(v)}</td></tr>`).join('')}</tbody></table></div>
    </div>
    <div class="card"><h3>5. Provider and company size</h3>
      <p>A named competitor or several providers is a plus: they already buy IoT connectivity and can switch. Headcount barely matters, because it says little about device volume.</p>
      <div class="tbl-pair"><div class="tbl-wrap">${tbl('fit_provider', 'Current provider')}</div><div class="tbl-wrap">${tbl('fit_size', 'Employees')}</div></div>
    </div>
    <div class="card"><h3>6. CRM stage decides whose account it is</h3>
      <p>Customers and open opportunities are worked by AEs and CS, so they're set to 0. Closed-lost accounts are kept at a discount: they said no once, but a renewal window can change that.</p>
      <div class="tbl-wrap">${tbl('crm_stage', 'Stage')}</div>
    </div>
    <div class="card"><h3>7. Routing</h3><ul>
      <li><b>Territory first:</b> ${territoryText().map(esc).join('; ')}.</li>
      <li><b>Best score first</b> until each rep reaches their weekly capacity.</li>
      <li><b>Shared territory:</b> the next account goes to the rep with the most free capacity as a share of their total, so both get a fair mix. Ties go to the more experienced rep.</li>
      <li><b>Pinned accounts</b> go first, whatever their score.</li>
      <li><b>Uncovered regions</b> are reported, not forced onto a rep in the wrong time zone.</li></ul></div>
  </div>
  <div class="card ex-wide" style="margin-top:16px"><h3>8. Week to week</h3>
    <div class="ignored">
      <div><b>New this week</b><p>Fresh accounts on a rep's list this Monday, up to their weekly capacity. Capacity is new accounts per week.</p></div>
      <div><b>Carried over</b><p>Accounts a rep already works stay with them for ${esc(S.get('general', 'recycle_after_days') ?? 28)} days. If a new signal makes one hotter (say 🍣 to 🥔), it moves to the top: call again.</p></div>
      <div><b>New signals move accounts</b><p>New signals add points and breadth. Without them every signal fades about 15% a week and drops out after ${esc(S.get('general', 'lookback_days'))} days.</p></div>
      <div><b>Always full capacity</b><p>Free slots first go to accounts waiting in another territory in a language the rep speaks (e.g. Spain → Mateo), then to the best-fit accounts with no signal yet: 🍞 bread & butter.</p></div>
    </div>
  </div>
  <div class="card ex-wide" style="margin-top:16px"><h3>9. Ignored on purpose</h3>
    <div class="ignored">
      <div><b>funding_stage</b><p>Contradicts the dated funding signals: accounts marked "Public" carry a recent "Raised 5M Seed" signal. The dated signal is trusted instead.</p></div>
      <div><b>country</b><p>Only used to fix the region and to match a rep's language when slots are free. No country gets a bonus or a penalty.</p></div>
      <div><b>source</b><p>Kept for transparency, not weighted. The signal text already says how strong the signal is.</p></div>
      <div><b>Machine learning</b><p>No win/loss history to learn from yet. Explicit weights leadership can read, tuned on conversion after a few weeks.</p></div>
    </div>
  </div>`;
}

const renewalBandsText = () => {
  const b = S.entries('renewal_band').map(([k, v]) => [Number(k), v]).sort((x, y) => x[0] - y[0]);
  return b.map(([k, v], i) => `${i ? `${b[i - 1][0]}–${k}` : `≤${k}`} months → <b>${esc(v)}</b>`).join(', ');
};

// ---------- data fixes ----------
function pFixes() {
  const groups = {};
  const label = i => {
    if (i.area === 'accounts' && i.problem.startsWith('Country')) return 'Country names unified';
    if (i.area === 'accounts' && i.problem.startsWith('Region')) return 'Region contradicted the country';
    if (i.area === 'accounts') return 'Region missing';
    if (i.problem.startsWith('Duplicate')) return 'Duplicate signals counted once';
    if (i.problem.startsWith('Older')) return 'Signals older than the lookback window';
    if (i.problem.startsWith('No date')) return 'Signals without a date';
    if (i.problem.startsWith('Dated after')) return 'Signals dated after the run date';
    if (i.problem.startsWith('Account not')) return 'Signals for accounts not in the list';
    return 'Routing';
  };
  const why = {
    'Country names unified': 'USA, US and United States are the same country. Without this, country-level reporting and territory checks split one market into several.',
    'Region contradicted the country': 'UAE and UK accounts were tagged NA. Left alone, they would go to the US reps.',
    'Duplicate signals counted once': 'The same event logged twice would double its weight.',
    'Signals older than the lookback window': 'Too old to say anything about this week.',
    'Signals without a date': 'A signal without a date can\'t be judged as fresh or stale, so it isn\'t trusted.',
    'Signals dated after the run date': 'A date in the future is a data error.',
    'Signals for accounts not in the list': 'The signal points to an account that isn\'t in the target list.',
    'Routing': 'Capacity left unused because the territory had too few qualifying accounts.'
  };
  for (const i of state.result.issues) (groups[label(i)] = groups[label(i)] || []).push(i);
  const names = Object.fromEntries(state.result.scored.map(s => [s.account_id, s.company_name]));
  return `${guideNote('The brief says the data is deliberately messy. Every correction is logged here, so the result can be audited and the source data fixed upstream.')}
  ${Object.entries(groups).map(([g, list]) => `<details class="set" ${state.fixOpen === g ? 'open' : ''} data-fix="${esc(g)}"><summary>${esc(g)} <span class="chip">${list.length}</span></summary><div class="sb"><p class="why">${esc(why[g] || '')}</p>
    <div class="tbl-wrap"><table><thead><tr><th>Item</th><th>Problem</th><th>Action</th></tr></thead><tbody>${list.map(i => `<tr><td>${esc(names[i.item] || i.item)} <span class="faint mono">${names[i.item] ? esc(i.item) : ''}</span></td><td>${esc(i.problem)}</td><td>${esc(i.action)}</td></tr>`).join('')}</tbody></table></div></div></details>`).join('')}`;
}

// ---------- Monday messages ----------
// The Monday brief is a briefing, not the list: hot accounts get a sentence each, warm and light-touch
// accounts are grouped by name, so the message stays short whether a rep has 15 accounts or 25.
const when = d => (d <= 6 ? 'this week' : d <= 13 ? 'last week' : `${Math.round(d / 7)} weeks ago`);
const lower = s => (/^[A-Z][a-z]/.test(s) ? s[0].toLowerCase() + s.slice(1) : s);
const dayMonth = iso => new Date(iso + 'T00:00:00Z').toLocaleDateString('en-GB', { day: 'numeric', month: 'long', timeZone: 'UTC' });
const names = list => list.length <= 1 ? (list[0] || '') : `${list.slice(0, -1).join(', ')} and ${list[list.length - 1]}`;
function reasons(acc, max = 3) {
  const seen = new Set(), out = [];
  for (const sig of acc.breakdown.signals) {
    if (seen.has(sig.detail)) continue;
    seen.add(sig.detail);
    out.push(`${sig.detail}, spotted ${when(sig.age)}`);
    if (out.length >= max) break;
  }
  return out;
}
function brief(t) {
  const first = t.name.split(' ')[0];
  const rows = state.result.lists.filter(l => l.sdr === t.name);
  const byId = Object.fromEntries(state.result.scored.map(s => [s.account_id, s]));
  const week = dayMonth(state.result.summary.week_of);
  if (!rows.length) return { subject: `No new accounts for the week of ${week}`, blocks: [{ p: `Hi ${first},` }, { p: 'No account in your territory qualified this week. Use the time to follow up on last week\'s list.' }] };
  const hot = rows.filter(r => tier(r) === 'A'), warm = rows.filter(r => tier(r) === 'B'), light = rows.filter(r => tier(r) === 'C'), pros = rows.filter(r => tier(r) === 'P');
  const top = rows[0], topAcc = byId[top.account_id];
  // carried over: earlier accounts still on the rep's list, and those that heated up with a new signal
  const mine = state.result.scored.filter(x => x.worked && x.worked.sdr === t.name);
  const seqN = mine.filter(x => x.status.startsWith('In sequence')).length, fire = mine.filter(x => x.status.startsWith('Back on fire'));
  const alsoHot = hot.slice(1, 6), hotOverflow = hot.slice(6);
  const blocks = [
    { p: `Hi ${first},` },
    { p: `Your list for this week is ready: ${rows.length} new accounts${hot.length ? `, ${hot.length} of them baked potatoes to call first` : ''}. Here's where to start.` },
    ...(state.week && (seqN || fire.length) ? [{ h: `Carried over (${seqN + fire.length})` }, { p: fire.length ? `${fire.length === 1 ? 'One has' : `${fire.length} have`} heated up since last week, so call again. The other ${seqN} stay in your sequence.` : `All ${seqN} stay in your sequence: keep going.` }, ...(fire.length ? [{ list: fire.sort((a, b) => b.score - a.score).slice(0, 5).map(x => ({ id: x.account_id, name: `${tierName[tier(x)].slice(0, 2)} ${x.company_name}`, text: lower(x.refire[0] || '') })) }] : []), { h: 'New this week' }] : []),
    { h: `Start with ${top.company_name}`, sub: `${top.country} · ${top.industry}`, id: top.account_id },
    ...(top.prospect ? [{ p: 'No buying signal yet, but the strongest fit in your territory: research first, then a cold first touch.' }] : [{ p: 'Why now:' }, { list: reasons(topAcc).map(text => ({ text })) }]),
    { p: `How to open: ${(top.suggested_angle.split(': ')[1] || top.suggested_angle).replace(/^./, c => c.toLowerCase())}`, muted: true },
    { p: top.loop_in_ae ? `${top.loop_in_ae} owns this account, so loop them in before you call.` : 'No AE on this one yet, so it\'s all yours.', muted: true }
  ];
  if (alsoHot.length) {
    blocks.push({ h: '🥔 More baked potatoes: call these first' });
    blocks.push({ list: alsoHot.map(r => ({ id: r.account_id, name: r.company_name, text: lower(reasons(byId[r.account_id], 1)[0] || '') })) });
    if (hotOverflow.length) blocks.push({ group: hotOverflow, sdr: t.id, lead: 'Plus' });
  }
  if (warm.length) blocks.push({ h: `🍲 Cream soup (${warm.length})` }, { p: 'Warm: worth a call and a follow-up email this week.' }, { group: warm, sdr: t.id });
  if (light.length) blocks.push({ h: `🍣 Sushi (${light.length})` }, { p: 'Served cold: one email from your sequence is enough this week.' }, { group: light, sdr: t.id });
  const prosRest = pros.filter(r => r !== top);
  if (prosRest.length) blocks.push({ h: `🍞 Bread & butter (${prosRest.length})` }, { p: 'No signal yet, but a strong fit: research, then a cold first touch built on device volume and current provider.' }, { group: prosRest, sdr: t.id });
  const ae = rows.filter(r => r.loop_in_ae && r !== top), lost = rows.filter(r => r.crm_stage === 'Closed lost'), pinned = rows.filter(r => r.pinned === 'yes');
  const heads = [];
  if (pinned.length) heads.push(`Leadership asked for ${names(pinned.map(r => r.company_name))} to be worked this week.`);
  const aeGroups = {};
  for (const r of ae) (aeGroups[r.loop_in_ae] = aeGroups[r.loop_in_ae] || []).push(r.company_name);
  for (const [who, list] of Object.entries(aeGroups)) heads.push(`Loop in ${who} before contacting ${names(list)}.`);
  const spill = rows.filter(r => r.spillover);
  if (spill.length) heads.push(`${names(spill.map(r => r.company_name))} ${spill.length > 1 ? 'are' : 'is'} outside your territory and came to you because you speak the language. Give the territory owner a heads-up.`);
  if (lost.length) heads.push(`${names(lost.map(r => r.company_name))} ${lost.length > 1 ? 'were' : 'was'} closed lost before. Acknowledge the history and lead with what has changed.`);
  if (heads.length) blocks.push({ h: 'Heads-up' }, { list: heads.map(text => ({ text })) });
  blocks.push({ p: 'Your full list, with every reason and angle, is in the CRM.', muted: true }, { p: 'Have a great week!' });
  return { subject: `Your ${rows.length} new accounts for the week of ${week}`, blocks };
}
const groupNames = (x, n = 3) => x.group.slice(0, n).map(r => r.company_name);
function briefText(b) {
  return [`Subject: ${b.subject}`, ...b.blocks.map(x => x.group ? `${x.lead ? `${x.lead} ` : ''}${names([...groupNames(x), ...(x.group.length > 3 ? [`${x.group.length - 3} more in your CRM list`] : [])])}.` : x.h ? `\n${x.h}${x.sub ? ` (${x.sub})` : ''}` : x.list ? x.list.map(i => `• ${i.name ? `${i.name}: ` : ''}${i.text}`).join('\n') : x.p)].join('\n');
}
function briefHtml(b) {
  const link = r => `<a href="#" data-acc="${esc(r.account_id)}">${esc(r.company_name)}</a>`;
  return `<div class="subj">${esc(b.subject)}</div>` + b.blocks.map(x => x.group ? `<p class="grpline">${x.lead ? `${x.lead} ` : ''}${x.group.slice(0, 3).map(link).join(', ')}${x.group.length > 3 ? ` and <a href="#" data-goto-list="${esc(x.sdr)}">${x.group.length - 3} more on the list →</a>` : ''}</p>` : x.h ? `<h4>${x.id ? `<a href="#" data-acc="${esc(x.id)}">${esc(x.h)}</a>` : esc(x.h)}${x.sub ? `<span>${esc(x.sub)}</span>` : ''}</h4>`
    : x.list ? `<ul>${x.list.map(i => `<li>${i.name ? `<a href="#" data-acc="${esc(i.id)}"><b>${esc(i.name)}</b></a>: ` : ''}${esc(i.text)}</li>`).join('')}</ul>`
    : `<p${x.muted ? ' class="m"' : ''}>${esc(x.p)}</p>`).join('');
}
function message(t) { return briefText(brief(t)); }
function pMessages() {
  return `${guideNote("Each rep's Monday brief is a briefing, not the list, with the same shape whether they have 15 accounts or 25: where to start, a line per baked potato, a few names for the rest, and heads-ups on AEs and closed-lost history. Every name opens the account. No scores: they mean nothing to a rep.", tierLegend())}
    <div class="msgs">${state.result.team.map(t => `<div class="card msg"><div class="mh"><span>To <b>${esc(t.name)}</b> · Monday 07:00</span><button class="btn small" data-copy="${esc(t.id)}">Copy</button></div><div class="mail">${briefHtml(brief(t))}</div></div>`).join('')}</div>`;
}

// ---------- all accounts ----------
// All accounts: every column can be sorted (click the header) and filtered (the row under it).
// Two words for week to week: "New this week" and "Carried over". Everything else is the food tier.
const statusGroup = s => s.status === 'Assigned' ? 'New this week'
  : s.status.startsWith('In sequence') || s.status.startsWith('Back on fire') ? 'Carried over'
  : s.status.startsWith('Not routed: no SDR') ? 'No SDR for region'
  : s.status.includes('at capacity') ? 'Waiting'
  : s.status.startsWith('No signal') ? 'No signal'
  : s.status.startsWith('Cool') ? 'Recently touched'
  : s.status.startsWith('Excluded: CRM') ? 'Owned by AE / CS'
  : 'Excluded by leadership';
const ALL_COLS = [
  { k: 'company_name', label: 'Company', get: s => s.company_name, filter: 'text' },
  { k: 'region', label: 'Region', get: s => s.region || '–', filter: 'select' },
  { k: 'country', label: 'Country', get: s => s.country || '–', filter: 'select' },
  { k: 'industry', label: 'Industry', get: s => s.industry || '–', filter: 'select' },
  { k: 'crm_stage', label: 'CRM stage', get: s => s.crm_stage || '–', filter: 'select' },
  { k: 'score', label: 'Score', get: s => s.score, num: true, filter: 'tier' },
  { k: 'signals_used', label: 'Signals', get: s => s.signals_used, num: true, filter: 'min' },
  { k: 'status', label: 'Status', get: statusGroup, filter: 'select' },
  { k: 'assigned_to', label: 'SDR', get: s => s.assigned_to || '–', filter: 'select' }
];
function pAll() {
  const f = state.colFilters, q = state.search.toLowerCase();
  const all = state.result.scored;
  const pass = s => ALL_COLS.every(c => {
    const v = f[c.k];
    if (v === undefined || v === '') return true;
    if (c.filter === 'text') return `${s.company_name} ${s.account_id}`.toLowerCase().includes(String(v).toLowerCase());
    if (c.filter === 'tier') return v === 'unassigned' ? s.status !== 'Assigned' : s.status === 'Assigned' && tier(s) === v;
    if (c.filter === 'min') return c.get(s) >= Number(v);
    return String(c.get(s)) === v;
  }) && (!q || `${s.company_name} ${s.account_id} ${s.industry} ${s.country}`.toLowerCase().includes(q));
  const col = ALL_COLS.find(c => c.k === state.sort.key) || ALL_COLS[5];
  const dir = state.sort.dir === 'asc' ? 1 : -1;
  const items = all.filter(pass).sort((a, b) => {
    const x = col.get(a), y = col.get(b);
    return (col.num ? (x - y) : String(x).localeCompare(String(y))) * dir || (b.score - a.score);
  });
  const opts = c => [...new Set(all.map(c.get))].sort((a, b) => String(a).localeCompare(String(b)));
  const fcell = c => {
    const v = f[c.k] ?? '';
    if (c.filter === 'text') return `<input type="search" data-colf="${c.k}" value="${esc(v)}" placeholder="Filter">`;
    if (c.filter === 'tier') return `<select data-colf="${c.k}"><option value="">All</option>${['A', 'B', 'C'].map(t => `<option value="${t}" ${v === t ? 'selected' : ''}>${tierName[t]}</option>`).join('')}<option value="unassigned" ${v === 'unassigned' ? 'selected' : ''}>Not assigned</option></select>`;
    if (c.filter === 'min') return `<select data-colf="${c.k}"><option value="">Any</option>${[1, 2, 3, 4].map(n => `<option value="${n}" ${String(v) === String(n) ? 'selected' : ''}>${n}+</option>`).join('')}</select>`;
    return `<select data-colf="${c.k}"><option value="">All</option>${opts(c).map(o => `<option ${String(v) === String(o) ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
  };
  const active = Object.values(f).filter(v => v !== undefined && v !== '').length;
  return `${guideNote('Every account, one row each, with where it ended up. Click a column header to sort, use the row under it to filter, and click any row to open the account.')}<div class="toolbar"><span class="muted" style="font-size:13px"><b style="color:var(--text)">${items.length}</b> of ${all.length} accounts</span>
    ${active ? `<button class="btn small" id="clearFilters">Clear ${active} filter${active > 1 ? 's' : ''}</button>` : ''}</div>
    <div class="card"><div class="tbl-wrap all-tbl"><table><thead>
      <tr>${ALL_COLS.map(c => `<th class="sortable ${c.num ? 'r' : ''} ${state.sort.key === c.k ? 'on' : ''}" data-sort="${c.k}">${c.label}<span class="arr">${state.sort.key === c.k ? (state.sort.dir === 'asc' ? '↑' : '↓') : '↕'}</span></th>`).join('')}</tr>
      <tr class="frow-h">${ALL_COLS.map(c => `<th>${fcell(c)}</th>`).join('')}</tr>
    </thead><tbody>
    ${items.map(s => `<tr class="click" data-acc="${esc(s.account_id)}">
      <td><b>${esc(s.company_name)}</b><div class="faint mono">${esc(s.account_id)}</div></td>
      <td>${esc(s.region || '–')}</td><td>${esc(s.country || '–')}</td><td>${esc(s.industry)}</td><td>${esc(s.crm_stage)}</td>
      <td class="r num">${fmt(s.score)}${s.status === 'Assigned' ? ` <span title="${esc(tierAction[tier(s)])}">${tierName[tier(s)].slice(0, 2)}</span>` : ''}</td>
      <td class="r num">${s.signals_used}</td>
      <td><span class="chip ${statusKind(s.status)}">${esc(statusGroup(s))}</span></td>
      <td>${s.assigned_to ? `${esc(s.assigned_to.split(' ')[0])} <span class="faint">#${s.rank}</span>` : '<span class="faint">–</span>'}</td></tr>`).join('') || `<tr><td colspan="${ALL_COLS.length}" class="faint" style="text-align:center;padding:24px">No accounts match these filters.</td></tr>`}
    </tbody></table></div></div>`;
}

// ---------- week over week ----------
const weekLabel = i => `Week ${i + 1}`;
const dayShort = iso => new Date(iso + 'T00:00:00Z').toLocaleDateString('en-GB', { day: 'numeric', month: 'short', timeZone: 'UTC' });
// Everything the comparison between week k and the week before needs, with a plain-language reason per account.
function weekCompare(k) {
  const cur = state.weeks[k], prev = state.weeks[k - 1] || cur.baseline;
  if (!prev) return null;
  const P = Object.fromEntries(prev.scored.map(s => [s.account_id, s]));
  const newSig = {};
  for (const x of cur.newSignals) (newSig[String(x.account_id).trim()] = newSig[String(x.account_id).trim()] || []).push(x);
  const strongest = id => (newSig[id] || []).map(x => x.signal_detail).sort((a, b) => (Number(S.get('signal_weight', b)) || 0) - (Number(S.get('signal_weight', a)) || 0))[0];
  // the same four reasons every week, so weeks can be compared
  const why = s => {
    if (s.prospect) return '🍞 Best-fit account, no signal yet: fills a free slot';
    if (s.spillover) return `Same-language rep had free slots: from the ${s.spillover.from} waiting list`;
    if (newSig[s.account_id]) return 'New signal in the last 7 days';
    return 'Earlier signal, still strong enough';
  };
  const entered = cur.lists.map(l => ({ ...l, reason: why(cur.scored.find(s => s.account_id === l.account_id)) }));
  const lastWeek = prev.lists.map(l => ({ ...l, now: cur.scored.find(s => s.account_id === l.account_id) }));
  const fire = cur.scored.filter(s => s.status.startsWith('Back on fire'));
  const inSeq = cur.scored.filter(s => s.status.startsWith('In sequence'));
  // score movement has its own reason: only new signals push a score up; time pulls it down
  const scoreWhy = s => newSig[s.account_id] ? `New signal: ${strongest(s.account_id)}`
    : P[s.account_id].signals_used > s.signals_used ? 'An older signal passed the 90-day window'
    : 'No new activity: signals faded by about 15%';
  const covered = new Set(cur.team.flatMap(t => t.territories));
  const movers = cur.scored.filter(s => !s.worked && P[s.account_id] && covered.has(s.region) && s.breakdown.stageMult > 0 && (s.score > 0 || P[s.account_id].score > 0))
    .map(s => ({ s, p: P[s.account_id], d: s.score - P[s.account_id].score, reason: scoreWhy(s) }));
  return { cur, prev, entered, lastWeek, fire, inSeq, movers, newSig };
}
// One table, same rows every week, so the numbers add up and can be compared side by side.
function bridgeRows(r) {
  // accounts an SDR can actually work this week: a live signal, an SDR-covered region, not resting after a touch
  const live = r.scored.filter(x => x.signals_used > 0 && x.breakdown.stageMult > 0 && !x.status.startsWith('Excluded') && !x.status.includes('no SDR') && !x.status.startsWith('Cool'));
  const freshIds = new Set(r.newSignals.map(x => String(x.account_id).trim()));
  const n = f => live.filter(f).length;
  const own = r.lists.filter(l => !l.prospect && !l.spillover).length, spill = r.lists.filter(l => l.spillover).length, pros = r.lists.filter(l => l.prospect).length;
  const seq = n(x => x.status.startsWith('In sequence')), fire = n(x => x.status.startsWith('Back on fire'));
  const available = live.length - seq - fire;
  return { known: r.knownSignals, fresh: r.newSignals.length, live: live.length, heated: n(x => freshIds.has(x.account_id)), seq, fire, carried: seq + fire,
    available, own, spill, pros, total: r.lists.length, waiting: available - own - spill };
}
function bridgeTable() {
  const B = state.weeks.map(bridgeRows);
  const row = (label, key, cls = '', sign = '') => `<tr class="${cls}"><td>${sign ? `<span class="sign">${sign}</span>` : ''}${label}</td>${B.map((b, i) => `<td class="r num ${i === state.week ? 'on' : ''}">${b[key]}</td>`).join('')}</tr>`;
  return `<div class="card pad bridge"><h3>Week by week at a glance</h3><p class="muted card-sub">Same rows every week, so each column adds up: accounts with a live signal that SDRs can work, minus those already carried over, gives what's available. The reps' 80 new slots are filled from that, then by language, then by best fit.</p>
  <div class="tbl-wrap"><table><thead><tr><th></th>${state.weeks.map((w, i) => `<th class="r ${i === state.week ? 'on' : ''}">${weekLabel(i)} · ${dayShort(w.date)}</th>`).join('')}</tr></thead><tbody>
    ${row('Signals known on Monday', 'known', 'muted-row')}
    ${row('of which new in the last 7 days', 'fresh', 'muted-row sub')}
    ${row('Accounts with a live signal SDRs can work', 'live', 'strong')}
    ${row('of which got a new signal in the last 7 days', 'heated', 'muted-row sub2')}
    ${row('carried over from earlier weeks', 'carried', 'sub', '−')}
    ${row('of which heated up: new signal since their rep got them', 'fire', 'muted-row sub2')}
    ${row('Available for new slots', 'available', 'strong', '=')}
    ${row('from their own territory', 'own', 'sub')}
    ${row('from another territory (same language)', 'spill', 'sub', '+')}
    ${row('🍞 best-fit accounts, no signal yet', 'pros', 'sub', '+')}
    ${row('New this week', 'total', 'strong total', '=')}
    ${row('Waiting for next week', 'waiting', 'muted-row')}
  </tbody></table></div></div>`;
}
function moversHtml(c, prevLabel = 'Last week') {
  const up = c.movers.filter(m => m.d > 0.05).sort((a, b) => b.d - a.d).slice(0, 10);
  const down = c.movers.filter(m => m.d < -0.05).sort((a, b) => a.d - b.d).slice(0, 10);
  const mrow = m => `<tr class="click" data-acc="${esc(m.s.account_id)}"><td><b>${esc(m.s.company_name)}</b><div class="faint">${esc(m.s.region)} · ${esc(m.s.industry)}</div></td><td class="r num">${fmt(m.p.score)}</td><td class="r num">${fmt(m.s.score)}</td><td class="r num ${m.d > 0 ? 'up' : 'down'}">${m.d > 0 ? '▲' : '▼'} ${fmt(Math.abs(m.d))}</td><td>${esc(m.reason)}</td><td><span class="chip ${statusKind(m.s.status)}">${esc(statusGroup(m.s))}</span></td></tr>`;
  const mtable = (rows, empty) => rows.length ? `<div class="tbl-wrap"><table><thead><tr><th>Account</th><th class="r">${prevLabel}</th><th class="r">This week</th><th class="r">Change</th><th>Why</th><th>Now</th></tr></thead><tbody>${rows.map(mrow).join('')}</tbody></table></div>` : `<p class="muted">${empty}</p>`;
  return `<div class="card pad"><h3>▲ Biggest moves up</h3><p class="muted card-sub">Accounts that are not carried over, in regions an SDR covers, ranked by how much their score rose. "Now" shows where each one ended up.</p>${mtable(up, 'No account rose: without new signals, scores can only fade.')}</div>
  <div class="card pad"><h3>▼ Biggest moves down</h3><p class="muted card-sub">Every signal loses about 15% of its weight per week, so accounts without new activity sink.</p>${mtable(down, 'No account fell.')}</div>`;
}
function pWeeks() {
  const k = state.week, cur = state.weeks[k], c = weekCompare(k), prev = c.prev;
  const prevLabel = k ? `${weekLabel(k - 1)} (${dayShort(prev.date)})` : `the Monday before (${dayShort(prev.date)})`;
  const note = guideNote(`Same three files, replayed as consecutive Mondays: each week only sees the signals dated up to that day. Every account is either new this week or carried over, and its food tier says how hot it is. A carried-over account that gets a new signal turns hotter and goes back to the top of its rep's list.`);
  const list = rows => `<ul class="reasons">${rows.map(([n, label]) => `<li><b>${n}</b>${label}</li>`).join('')}</ul>`;
  // 1. what the new accounts look like, by food
  const foods = ['A', 'B', 'C', 'P'].map(t => [cur.lists.filter(l => tier(l) === t).length, `${tierName[t]}: ${tierAction[t]}`]);
  // 2. how they got onto the lists, same four reasons every week
  const rc = {};
  for (const e of c.entered) rc[e.reason] = (rc[e.reason] || 0) + 1;
  const reasonOrder = ['New signal in the last 7 days', 'Earlier signal, still strong enough'];
  const reasons = [...reasonOrder.map(r => [rc[r] || 0, r]), ...Object.entries(rc).filter(([r]) => !reasonOrder.includes(r)).map(([r, n]) => [n, r])];
  // 3. strongest signal type of the new accounts
  const byType = {};
  for (const l of cur.lists) { const a = cur.scored.find(x => x.account_id === l.account_id); const t = a?.breakdown.signals[0]?.type || 'No signal yet (🍞)'; byType[t] = (byType[t] || 0) + 1; }
  const types = Object.entries(byType).sort((x, y) => y[1] - x[1]).map(([t, n]) => [n, esc(t)]);
  // 4. pool coverage
  const cov = cur.summary.coverage.map(v => `<div class="cov ${v.weeks < 1 ? 'low' : ''}"><b>${esc(v.region)}</b><span>${v.backlog} qualified accounts waiting for ${v.cap} weekly slots: ${v.weeks < 1 ? 'less than a week left' : `about ${v.weeks} more week${v.weeks > 1 ? 's' : ''}`}</span></div>`).join('');
  // 5. carried over and heated up
  const heated = c.fire.sort((x, y) => y.score - x.score);
  const heatedHtml = heated.length
    ? `<div class="tbl-wrap"><table><thead><tr><th>Account</th><th>Now</th><th>Rep</th><th>On their list since</th><th>New signal</th></tr></thead><tbody>${heated.map(x => `<tr class="click" data-acc="${esc(x.account_id)}"><td><b>${esc(x.company_name)}</b></td><td><span class="tier ${tier(x)}">${tierName[tier(x)]}</span></td><td>${esc(x.worked.sdr)}</td><td>${dayShort(x.worked.date)}</td><td>${esc(x.refire.slice(0, 2).join('; '))}${x.refire.length > 2 ? ` <span class="faint">+${x.refire.length - 2} more</span>` : ''}</td></tr>`).join('')}</tbody></table></div>`
    : `<p class="muted empty">${k ? 'None this week: no carried-over account got a new signal.' : 'Nothing is carried over yet: this is the first week, so every account on the lists is new.'}</p>`;
  return `<div class="wk">
    ${note}
    <p class="muted cmp-line">Showing <b>${weekLabel(k)} (${dayShort(cur.date)})</b>, compared with ${prevLabel}. Switch weeks at the top of the page.</p>
    ${bridgeTable()}
    <div class="wk-grid three">
      <div class="card pad"><h3>What this week's ${cur.lists.length} new accounts look like</h3>${list(foods)}</div>
      <div class="card pad"><h3>Their strongest signal</h3>${list(types)}</div>
      <div class="card pad"><h3>Signal pool left at this capacity</h3><p class="muted card-sub">How long the accounts still waiting keep the reps busy if no new signals arrive.</p>${cov}</div>
    </div>
    <div class="card pad"><h3>Carried over and heated up</h3><p class="muted card-sub">Accounts a rep already works that got a new signal this week. They stay with the same rep and go to the top of that rep's Monday brief, without taking a new slot.</p>${heatedHtml}</div>
    ${moversHtml(c, k ? weekLabel(k - 1) : dayShort(prev.date))}
  </div>`;
}

function weekDelta(s) {
  const k = state.week, prev = state.weeks[k - 1];
  if (!prev) return '';
  const p = prev.scored.find(x => x.account_id === s.account_id);
  if (!p) return '';
  const fresh = state.weeks[k].newSignals.filter(x => String(x.account_id).trim() === s.account_id);
  return `<div class="sec"><h3>Since last week</h3><dl class="kv">
    <dt>Last week</dt><dd>${fmt(p.score)} · ${esc(statusGroup(p))}${p.assigned_to ? ` (${esc(p.assigned_to)} #${p.rank})` : ''}</dd>
    <dt>This week</dt><dd>${fmt(s.score)} · ${esc(statusGroup(s))}${s.assigned_to ? ` (${esc(s.assigned_to)} #${s.rank})` : ''}</dd>
    <dt>New signals</dt><dd>${fresh.length ? esc(fresh.map(x => x.signal_detail).join('; ')) : '<span class="faint">none, so its signals faded by about 15%</span>'}</dd></dl></div>`;
}

// ---------- account drawer ----------
function openAccount(id) {
  const s = state.result.scored.find(x => x.account_id === id);
  if (!s) return;
  const b = s.breakdown;
  const ex = S.get('exclude', id) !== undefined, pin = S.get('pin', id) !== undefined;
  const dropped = state.result.issues.filter(i => i.area === 'signals' && i.item === id);
  const why = s.status === 'Assigned' && s.spillover
    ? `Routed to <b>${esc(s.assigned_to)}</b> across territories: it waited in the ${esc(s.spillover.from)} queue, ${esc(s.assigned_to.split(' ')[0])} had free slots and speaks ${esc(s.spillover.language)}, the language of ${esc(s.country)}. A good signal account beats a no-signal prospect.`
    : s.status === 'Assigned' && s.prospect
    ? `Assigned to <b>${esc(s.assigned_to)}</b> as a best-fit prospect: no buying signal yet, but one of the strongest-fit accounts in ${esc(s.region)}. There weren't enough signal accounts to fill every slot, so this keeps ${esc(s.assigned_to.split(' ')[0])} at full capacity.`
    : s.status.startsWith('In sequence') ? `Carried over: ${esc(s.worked.sdr)} has had this account on their list since ${esc(dayShort(s.worked.date))}. It stays with them while the sequence runs (${esc(S.get('general', 'recycle_after_days') ?? 28)} days) and doesn't take a new slot.`
    : s.status.startsWith('Back on fire') ? `Carried over and heated up: ${esc(s.worked.sdr)} has had it since ${esc(dayShort(s.worked.date))}, and a new signal arrived since (<b>${esc(s.refire.join('; '))}</b>). It's now ${tierName[tier(s)]}, so it goes to the top of their Monday brief: call again. No new slot needed.`
    : s.status === 'Assigned'
    ? `Assigned to <b>${esc(s.assigned_to)}</b> at rank ${s.rank}${s.pinned ? ' (pinned by leadership)' : ''}. It's in ${esc(s.region)}, which ${esc(s.assigned_to)} covers, and it scored high enough to make their ${state.result.team.find(t => t.name === s.assigned_to)?.cap}-account list.`
    : s.status.includes('at capacity') ? `Qualified, but the ${esc(s.region)} reps were full before reaching it. It would be in next week's pipeline.`
    : s.status.startsWith('Not routed: no SDR') ? `No SDR covers ${esc(s.region)}. Leadership can assign the region to a rep in settings.`
    : s.status.startsWith('Excluded: CRM') ? `${esc(s.crm_stage)} accounts belong to AEs or CS, so an SDR doesn't call them.`
    : s.status.startsWith('No signal') ? 'No buying signal in the lookback window, so there is no reason to call this week.'
    : s.status.startsWith('Cool') ? `Someone touched this account ${b.touchAge} days ago. It rests to avoid over-contacting.`
    : esc(s.status);
  overlay.innerHTML = `<div class="scrim" data-close></div>
  <aside class="drawer" role="dialog" aria-label="${esc(s.company_name)}">
    <div class="dh"><div><div class="faint mono" style="font-size:12px">${esc(s.account_id)}</div><h2>${esc(s.company_name)}</h2>
      <div class="muted" style="font-size:13.5px">${esc(s.country)} · ${esc(s.region)} · ${esc(s.industry)} · ${esc(s.crm_stage)}${s.ae_owner ? ` · AE ${esc(s.ae_owner)}` : ''}</div>
      <div style="margin-top:8px"><span class="chip ${statusKind(s.status)}">${esc(statusGroup(s))}</span>${s.status === 'Assigned' ? ` <span class="tier ${tier(s)}">${tierName[tier(s)]}</span>` : ''}</div></div>
      <button class="btn small ghost x" data-close aria-label="Close">✕</button></div>
    <div class="db">
      <div class="note">${why}</div>
      ${weekDelta(s)}
      <div class="eq">
        <div class="p"><b>${fmt(b.intent)}</b><span>intent</span></div><span class="o">×</span>
        <div class="p"><b>${fmt(b.fit, 2)}</b><span>fit</span></div><span class="o">×</span>
        <div class="p"><b>${fmt(b.stageMult, 2)}</b><span>CRM stage</span></div><span class="o">×</span>
        <div class="p"><b>${fmt(b.boost, 2)}</b><span>boost</span></div><span class="o">=</span>
        <div class="p res"><b>${fmt(s.score)}</b><span>score</span></div>
      </div>
      <div class="sec"><h3>Intent: the signals behind it</h3>
        ${b.signals.length ? `<div class="tbl-wrap"><table><thead><tr><th>Signal</th><th class="r">Weight</th><th class="r">Age</th><th class="r">Fresh</th><th class="r">Counts</th><th class="r">Points</th></tr></thead><tbody>
        ${b.signals.map(x => `<tr><td>${esc(x.detail)}<div class="faint" style="font-size:12px">${esc(x.type)} · ${esc(x.source)} · ${esc(x.date)}${x.renewal ? `<br>Renewal ~${esc(dayShort(x.renewal))} ${esc(x.renewal.slice(0, 4))}, in ${x.monthsToRenewal} months: weight from the renewal window, no fading` : ''}</div></td><td class="r num">${x.weight}</td><td class="r num">${x.age}d</td><td class="r num">${Math.round(x.decay * 100)}%</td><td class="r num">${x.share === 1 ? '100%' : `${Math.round(x.share * 100)}%`}</td><td class="r num"><b>${fmt(x.counted, 2)}</b></td></tr>`).join('')}
        </tbody></table></div>
        <div class="note">Signal points add up to <b>${fmt(b.intentBase, 2)}</b>. ${b.types} signal type${b.types > 1 ? 's' : ''} → breadth ×${fmt(b.breadth, 2)} → intent <b>${fmt(b.intent, 2)}</b>. "Fresh" is the weight left after ${S.get('general', 'half_life_days')}-day halving; "Counts" is 50% for repeats of a type already counted.</div>`
        : '<div class="note">No usable signal in the lookback window.</div>'}
        ${dropped.length ? `<div class="note">Not counted: ${dropped.map(d => esc(d.problem)).join('; ')}.</div>` : ''}
      </div>
      <div class="sec"><h3>Fit</h3><dl class="kv">
        <dt>Devices in field</dt><dd>${b.devices !== null ? fmt0(b.devices) : '<span class="faint">unknown</span>'} → ×${b.devM}</dd>
        <dt>Current provider</dt><dd>${esc(b.provider || 'unknown')} → ×${b.provM}</dd>
        <dt>Employees</dt><dd>${esc(b.size || 'unknown')} → ×${b.sizeM}</dd>
        <dt>Fit</dt><dd><b>×${fmt(b.fit, 2)}</b></dd>
        ${b.boostNotes.length ? `<dt>Boosts</dt><dd>${esc(b.boostNotes.join(', '))}</dd>` : ''}
        <dt>Last touch</dt><dd>${esc(b.last_touch_date || 'never')}${b.touchAge !== null ? ` (${b.touchAge} days ago)` : ''}</dd>
        <dt>Funding stage</dt><dd>${esc(b.funding_stage || '–')} <span class="faint">(not used)</span></dd>
      </dl></div>
      ${s.suggested_angle ? `<div class="sec"><h3>Suggested opening</h3><div class="note">${esc(s.suggested_angle)}</div></div>` : ''}
    </div>
    <div class="df">
      <span class="muted" style="font-size:13px;align-self:center;margin-right:auto">Leadership actions:</span>
      <button class="btn small" data-pin="${esc(id)}">${pin ? 'Unpin' : '📌 Pin to top'}</button>
      <button class="btn small ${ex ? '' : 'danger'}" data-exclude="${esc(id)}">${ex ? 'Include again' : 'Exclude this week'}</button>
    </div>
  </aside>`;
}
function closeOverlay() { overlay.innerHTML = ''; }

// ---------- settings drawer ----------
function numRow(sec, key, label, hint, step = 'any') {
  const v = S.get(sec, key);
  return `<div class="frow"><label>${esc(label)}${hint ? `<div class="hint">${esc(hint)}</div>` : ''}</label><input type="number" step="${step}" data-sec="${esc(sec)}" data-key="${esc(key)}" value="${esc(v ?? '')}"></div>`;
}
function openSettings() {
  const typeOf = {};
  for (const s of state.data.signals || []) typeOf[s.signal_detail] = s.signal_type;
  const groups = {};
  for (const [k] of S.entries('signal_weight')) (groups[typeOf[k] || 'Other'] = groups[typeOf[k] || 'Other'] || []).push(k);
  const team = state.data.sdrs ? state.data.sdrs.map(r => ({ id: String(r.sdr_id), name: r.name, cap: r.weekly_capacity })) : ['SDR-01', 'SDR-02', 'SDR-03', 'SDR-04'].map(id => ({ id, name: id, cap: '' }));
  const regions = [...new Set([...S.entries('territory').map(([k]) => k), ...(state.result ? state.result.scored.map(s => s.region).filter(Boolean) : [])])];
  const names = state.data.accounts ? Object.fromEntries(state.data.accounts.map(a => [String(a.account_id), a.company_name])) : {};
  const list = sec => realEntries(sec).map(([k, v]) => `<span class="chip ${sec === 'pin' ? 'info' : 'bad'}">${esc(names[k] || k)}${v && v !== 'example' ? ` · ${esc(v)}` : ''}<button data-unlist="${sec}" data-id="${esc(k)}" aria-label="Remove">×</button></span>`).join('') || '<span class="faint" style="font-size:13px">None</span>';
  overlay.innerHTML = `<div class="scrim" data-close></div>
  <aside class="drawer settings" role="dialog" aria-label="Leadership settings">
    <div class="dh"><div><h2>Leadership settings</h2><div class="muted" style="font-size:13.5px">Every rule the scoring uses. ${state.result ? 'Changes re-rank the lists immediately.' : 'Changes apply when you build the lists.'} Highlighted fields differ from ${esc(state.settingsSource)}.</div></div>
      <button class="btn small ghost x" data-close aria-label="Close">✕</button></div>
    <div class="db">
      <details class="set" open><summary>Boost or switch off segments</summary><div class="sb">
        <p class="why">The fastest lever: ×1.3 pushes a segment up this week, ×0 removes it. Example: a fleet telematics campaign.</p>
        ${S.entries('boost').map(([k]) => numRow('boost', k, k.replace('industry=', '').replace('region=', 'Region: '), '', '0.1')).join('')}
      </div></details>
      <details class="set"><summary>Exclude or pin single accounts</summary><div class="sb">
        <p class="why">Exclude: a known blocker, legal hold or partner conflict. Pin: a strategic account leadership wants worked this week, whatever its score. You can also do both from any account's detail view.</p>
        <b style="font-size:13px">Pinned</b><div class="pill-list">${list('pin')}</div>
        <div class="add-row"><input list="accList" id="pinInput" placeholder="Company name or account ID"><button class="btn small" data-add="pin">Pin</button></div>
        <b style="font-size:13px;display:block;margin-top:12px">Excluded</b><div class="pill-list">${list('exclude')}</div>
        <div class="add-row"><input list="accList" id="excludeInput" placeholder="Company name or account ID"><button class="btn small" data-add="exclude">Exclude</button></div>
        <datalist id="accList">${state.data.accounts ? state.data.accounts.map(a => `<option value="${esc(a.account_id)}">${esc(a.company_name)}</option>`).join('') : ''}</datalist>
      </div></details>
      <details class="set"><summary>Territories and capacity</summary><div class="sb">
        <p class="why">Who covers which region, and how many new accounts each rep can work this week. Empty capacity uses the SDR file. Try giving APAC to a rep.</p>
        ${regions.map(reg => { const ids = String(S.get('territory', reg) ?? '').split(',').map(x => x.trim()); return `<div class="frow wide"><label>${esc(reg)}</label><div class="checks">${team.map(t => `<label><input type="checkbox" data-terr="${esc(reg)}" value="${esc(t.id)}" ${ids.includes(t.id) ? 'checked' : ''}> ${esc(t.name.split(' ')[0])}</label>`).join('')}</div></div>`; }).join('')}
        <div class="grp">Capacity this week</div>
        ${team.map(t => `<div class="frow"><label>${esc(t.name)}<div class="hint">${t.cap !== '' ? `${t.cap} in the SDR file` : ''}</div></label><input type="number" step="1" min="0" data-sec="capacity" data-key="${esc(t.id)}" value="${esc(S.get('capacity', t.id) ?? '')}" placeholder="${esc(t.cap)}"></div>`).join('')}
      </div></details>
      <details class="set"><summary>Time window</summary><div class="sb">
        <p class="why">How far back signals count, how fast they fade, and how long a touched account rests.</p>
        <div class="frow"><label>Run date of week 3<div class="hint">weeks 1 and 2 are 14 and 7 days earlier</div></label><input type="date" data-sec="general" data-key="as_of_date" value="${esc(S.get('general', 'as_of_date'))}"></div>
        ${numRow('general', 'lookback_days', 'Lookback window (days)', 'older signals are dropped', '1')}
        ${numRow('general', 'half_life_days', 'Half-life (days)', 'weight halves every N days', '1')}
        ${numRow('general', 'recent_touch_days', 'Cool-down after a touch (days)', '', '1')}
        ${numRow('general', 'extra_signal_same_type', 'Repeat of the same type counts', '0.5 = half', '0.05')}
        ${numRow('general', 'multi_type_bonus', 'Bonus per extra signal type', '0.15 = +15%', '0.05')}
        <div class="grp">Week to week</div>
        ${numRow('general', 'recycle_after_days', 'Sequence length (days)', 'after this, an account without reply returns to the pool', '1')}
        <div class="frow"><label>Fill empty slots with best-fit accounts<div class="hint">keeps every rep at full capacity</div></label><input type="checkbox" data-sec="general" data-key="fill_with_fit" ${Number(S.get('general', 'fill_with_fit') ?? 1) ? 'checked' : ''}></div>
        <div class="frow"><label>Language spillover<div class="hint">a rep with free slots takes accounts from another territory's queue in a language they speak</div></label><input type="checkbox" data-sec="general" data-key="language_spillover" ${Number(S.get('general', 'language_spillover') ?? 1) ? 'checked' : ''}></div>
      </div></details>
      <details class="set"><summary>Renewal windows</summary><div class="sb">
        <p class="why">Competitor renewals are scored on months until the renewal, not on signal age. Each row: up to this many months away → this weight.</p>
        ${S.entries('renewal_band').map(([k]) => numRow('renewal_band', k, `Up to ${k} months away`, '', '1')).join('')}
      </div></details>
      <details class="set"><summary>Signal weights (0–10)</summary><div class="sb">
        <p class="why">How close each signal is to a buying decision. Signals not listed use their type default.</p>
        ${Object.entries(groups).map(([gname, keys]) => `<div class="grp">${esc(gname)}</div>${keys.map(k => numRow('signal_weight', k, k, '', '1')).join('')}`).join('')}
        <div class="grp">Type defaults</div>${S.entries('signal_type_default').map(([k]) => numRow('signal_type_default', k, k, '', '1')).join('')}
      </div></details>
      <details class="set"><summary>CRM stage and fit multipliers</summary><div class="sb">
        <p class="why">0 removes a stage entirely. Fit multipliers scale the score by how valuable a win would be.</p>
        <div class="grp">CRM stage</div>${S.entries('crm_stage').map(([k]) => numRow('crm_stage', k, k, '', '0.1')).join('')}
        <div class="grp">Devices in field (at least)</div>${S.entries('fit_devices').map(([k]) => numRow('fit_devices', k, `${fmt0(k)}+ devices`, '', '0.05')).join('')}
        <div class="grp">Current provider</div>${S.entries('fit_provider').map(([k]) => numRow('fit_provider', k, k, '', '0.05')).join('')}
        <div class="grp">Employees</div>${S.entries('fit_size').map(([k]) => numRow('fit_size', k, k, '', '0.05')).join('')}
      </div></details>
      <details class="set"><summary>Opening angles</summary><div class="sb">
        <p class="why">The suggested first line for each account, chosen by its strongest signal type.</p>
        ${S.entries('angle').map(([k, v]) => `<div class="frow wide"><label>${esc(k)}</label><input type="text" data-sec="angle" data-key="${esc(k)}" value="${esc(v)}"></div>`).join('')}
      </div></details>
    </div>
    <div class="df">
      <button class="btn small" id="resetSettings">Reset to ${state.settingsSource === 'built-in defaults' ? 'defaults' : 'uploaded file'}</button>
      <button class="btn small" id="dlSettings">⬇ Download as settings.xlsx</button>
      <span style="flex:1"></span>
      ${state.result ? '<button class="btn small primary" data-close>Done</button>' : `<button class="btn small primary" id="runFromSettings" ${ready() ? '' : 'disabled'}>Build lists</button>`}
    </div>
  </aside>`;
  refreshSettingsMarks();
}
function refreshSettingsMarks() {
  document.querySelectorAll('.drawer.settings input[data-sec]').forEach(el => {
    const b = S.base(el.dataset.sec, el.dataset.key);
    el.classList.toggle('changed', String(b ?? '') !== String(S.get(el.dataset.sec, el.dataset.key) ?? ''));
  });
}
function afterSettingsChange() {
  if (state.result) scheduleRun();
  else { const sum = document.querySelector('.settings-sum'); if (sum) render(); }
}
function resolveAccount(v) {
  v = String(v || '').trim();
  if (!v || !state.data.accounts) return null;
  const a = state.data.accounts.find(a => String(a.account_id) === v || String(a.company_name).toLowerCase() === v.toLowerCase());
  return a ? String(a.account_id) : null;
}
function toggleList(sec, id, on) {
  if (on) S.set(sec, id, sec === 'exclude' ? 'Leadership decision' : 'Leadership priority'); else S.del(sec, id);
  if (sec === 'pin' && on) S.del('exclude', id);
  if (sec === 'exclude' && on) S.del('pin', id);
}

// ---------- downloads ----------
function sheet(rows) { return XLSX.utils.json_to_sheet(rows.length ? rows : [{ note: 'empty' }]); }
function downloadExcel() {
  const r = state.result, wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet([
    { item: 'Week of', value: r.summary.week_of }, { item: 'Accounts assigned', value: r.lists.length },
    ...r.team.map(t => ({ item: t.name, value: `${t.filled} of ${t.cap} (${t.territories.join(', ')})` })),
    { item: 'Signals used', value: `${r.summary.signals_used} of ${r.summary.signals_in}` },
    ...funnelSteps().slice(1).map(s => ({ item: s.label, value: s.items.length })),
    { item: 'Settings', value: `${state.settingsSource}${S.changes() ? ` + ${S.changes()} changes` : ''}` }
  ]), 'Summary');
  for (const t of r.team) XLSX.utils.book_append_sheet(wb, sheet(r.lists.filter(l => l.sdr === t.name).map(({ week_of, sdr, ...x }) => ({ ...x, tier: tierName[tier(x)] }))), t.name.slice(0, 31));
  XLSX.utils.book_append_sheet(wb, sheet(r.lists), 'All assigned');
  const cmp = weekCompare(state.week);
  if (cmp) XLSX.utils.book_append_sheet(wb, sheet([
    ...cmp.entered.map(e => ({ group: 'New on lists', account: e.company_name, sdr: e.sdr, rank: e.rank, why: e.reason })),
    ...cmp.fire.map(x => ({ group: 'Carried over, heated up', account: x.company_name, sdr: x.worked.sdr, rank: '', why: x.refire.join('; ') })),
    ...cmp.movers.sort((a, b) => b.d - a.d).map(m => ({ group: m.d >= 0 ? 'Moved up' : 'Moved down', account: m.s.company_name, sdr: m.s.assigned_to, rank: m.s.rank, why: m.reason, last_week: m.p.score, this_week: m.s.score }))
  ]), 'Week over week');
  XLSX.utils.book_append_sheet(wb, sheet(r.scored.map(({ breakdown, ...x }) => x)), 'All accounts scored');
  XLSX.utils.book_append_sheet(wb, sheet(r.issues), 'Data issues');
  XLSX.utils.book_append_sheet(wb, sheet(state.settings.map(({ section, key, value }) => ({ section, key, value }))), 'Settings used');
  XLSX.writeFile(wb, `sdr_lists_week_${state.week + 1}_${r.summary.week_of}.xlsx`);
}
function downloadSettings() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet(state.settings.map(({ section, key, value, note }) => ({ section, key, value, note: note || '' }))), 'settings');
  XLSX.writeFile(wb, 'settings.xlsx');
}

// ---------- events ----------
function bindCommon() {
  const drop = document.getElementById('drop');
  if (drop) {
    const input = document.getElementById('fileInput');
    input.addEventListener('change', e => addFiles([...e.target.files]));
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => addFiles([...e.dataTransfer.files]));
  }
  const search = document.getElementById('search');
  if (search) search.addEventListener('input', e => { state.search = e.target.value; const pos = e.target.selectionStart; document.getElementById('panel').innerHTML = panel('all'); bindCommon(); const s = document.getElementById('search'); s.focus(); s.setSelectionRange(pos, pos); });
  document.querySelectorAll('[data-colf]').forEach(el => el.addEventListener(el.tagName === 'INPUT' ? 'input' : 'change', e => {
    state.colFilters[el.dataset.colf] = e.target.value;
    const k = el.dataset.colf, pos = e.target.selectionStart;
    document.getElementById('panel').innerHTML = panel('all'); bindCommon(); saveState();
    if (el.tagName === 'INPUT') { const n = document.querySelector(`input[data-colf="${k}"]`); n.focus(); n.setSelectionRange(pos, pos); }
  }));
  document.querySelectorAll('details[data-fix]').forEach(d => d.addEventListener('toggle', () => { if (d.open) state.fixOpen = d.dataset.fix; }));
}

document.addEventListener('click', async e => {
  const t = e.target.closest('button, [data-acc], [data-funnel], [data-close], [data-tab], [data-goto-list], th[data-sort], [data-week]');
  if (!t) return;
  if (t.id === 'runBtn' || t.id === 'runFromSettings') { closeOverlay(); state.tab = 'board'; run(); window.scrollTo(0, 0); return; }
  if (t.id === 'sampleBtn') return loadSample();
  if (t.dataset.sort) { const k = t.dataset.sort; const num = ALL_COLS.find(c => c.k === k)?.num; state.sort = state.sort.key === k ? { key: k, dir: state.sort.dir === 'asc' ? 'desc' : 'asc' } : { key: k, dir: num ? 'desc' : 'asc' }; return render(); }
  if (t.id === 'clearFilters') { state.colFilters = {}; return render(); }
  if (t.id === 'dlBtn') return downloadExcel();
  if (t.id === 'dlSettings') return downloadSettings();
  if (t.id === 'resetBtn') { clearSaved(); Object.assign(state, { data: { accounts: null, signals: null, sdrs: null }, names: {}, result: null, settings: clone(DEFAULT_SETTINGS), baseSettings: clone(DEFAULT_SETTINGS), settingsSource: 'built-in defaults', error: '' }); return render(); }
  if (t.id === 'resetSettings') { state.settings = clone(state.baseSettings); openSettings(); return afterSettingsChange(); }
  if (t.hasAttribute('data-close')) return closeOverlay();
  if (t.hasAttribute('data-open-settings')) return openSettings();
  if (t.dataset.tab) { state.tab = t.dataset.tab; return render(); }
  if (t.dataset.week !== undefined) return setWeek(Number(t.dataset.week));
  if (t.dataset.expand) { state.expanded[t.dataset.expand] = !state.expanded[t.dataset.expand]; return render(); }
  if (t.dataset.funnel) { state.funnelOpen = state.funnelOpen === t.dataset.funnel ? null : t.dataset.funnel; return render(); }
  if (t.dataset.copy) { const tm = state.result.team.find(x => x.id === t.dataset.copy); try { await navigator.clipboard.writeText(message(tm)); toast('Message copied'); } catch { toast('Copy not allowed in this browser'); } return; }
  if (t.dataset.pin || t.dataset.exclude) {
    const sec = t.dataset.pin ? 'pin' : 'exclude', id = t.dataset.pin || t.dataset.exclude;
    toggleList(sec, id, S.get(sec, id) === undefined);
    run(); openAccount(id); return;
  }
  if (t.dataset.unlist) { S.del(t.dataset.unlist, t.dataset.id); openSettings(); return afterSettingsChange(); }
  if (t.dataset.add) {
    const id = resolveAccount(document.getElementById(`${t.dataset.add}Input`).value);
    if (!id) return toast(state.data.accounts ? 'Account not found' : 'Load the accounts file first');
    toggleList(t.dataset.add, id, true); openSettings(); document.querySelectorAll('details.set')[1].open = true; return afterSettingsChange();
  }
  if (t.dataset.gotoList) { e.preventDefault(); state.tab = 'board'; state.expanded[t.dataset.gotoList] = true; render(); document.querySelector(`[data-col="${t.dataset.gotoList}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'start' }); return; }
  if (t.dataset.acc && state.result) { e.preventDefault(); return openAccount(t.dataset.acc); }
});

overlay.addEventListener('change', e => {
  const el = e.target;
  if (el.dataset.terr) {
    const reg = el.dataset.terr;
    const ids = [...overlay.querySelectorAll(`input[data-terr="${CSS.escape(reg)}"]:checked`)].map(x => x.value);
    S.set('territory', reg, ids.join(','));
    return afterSettingsChange();
  }
  if (!el.dataset.sec) return;
  const v = el.type === 'checkbox' ? (el.checked ? 1 : 0) : el.type === 'number' ? (el.value === '' ? '' : Number(el.value)) : el.value;
  S.set(el.dataset.sec, el.dataset.key, v);
  refreshSettingsMarks();
  afterSettingsChange();
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeOverlay(); });

// ---------- day / night switch (remembered per browser) ----------
const SUN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>';
const MOON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>';
const isDark = () => (document.documentElement.dataset.theme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light')) === 'dark';
function paintThemeBtn() { const b = document.getElementById('themeBtn'); if (b) b.innerHTML = isDark() ? `${SUN} Day` : `${MOON} Night`; }
document.getElementById('themeBtn')?.addEventListener('click', () => {
  const next = isDark() ? 'light' : 'dark';
  document.documentElement.dataset.theme = next;
  try { localStorage.setItem('prio-theme', next); } catch {}
  paintThemeBtn();
});
matchMedia('(prefers-color-scheme: dark)').addEventListener?.('change', paintThemeBtn);
paintThemeBtn();

// offer "Use the case data" only when the files are bundled in /data
state.noSample = true;
if (restoreState()) run(); else render();
fetch('data/accounts.xlsx', { method: 'HEAD' }).then(r => { state.noSample = !r.ok; if (!state.result) render(); }).catch(() => {});
