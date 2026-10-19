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
  search: '', statusFilter: 'all',
  error: ''
};
const app = document.getElementById('app');
const overlay = document.getElementById('overlay');

// ---------- helpers ----------
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = (n, d = 1) => (n === null || n === undefined || isNaN(n) ? '–' : Number(n).toLocaleString('en-US', { maximumFractionDigits: d, minimumFractionDigits: d }));
const fmt0 = n => Number(n).toLocaleString('en-US');
const ready = () => state.data.accounts && state.data.signals && state.data.sdrs;
const tier = s => (s >= 20 ? 'A' : s >= 8 ? 'B' : 'C');
const tierName = { A: 'Hot', B: 'Warm', C: 'Light touch' };
const statusKind = st => st === 'Assigned' ? 'good' : st.startsWith('Not routed') ? 'warn' : st.startsWith('Excluded') ? 'bad' : '';

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

// ---------- run ----------
let runTimer;
function scheduleRun() { clearTimeout(runTimer); runTimer = setTimeout(run, 220); }
function run() {
  if (!ready()) return render();
  const before = state.result ? new Set(state.result.lists.map(l => l.account_id)) : null;
  const posBefore = state.result ? new Map(state.result.lists.map(l => [l.account_id, `${l.sdr}#${l.rank}`])) : null;
  state.result = prioritise(state.data.accounts, state.data.signals, state.data.sdrs, state.settings);
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
function render() {
  app.innerHTML = state.result ? renderResults() : renderIntro();
  bindCommon();
}

function renderIntro() {
  const f = (kind, label, desc, optional) => {
    const ok = kind === 'settings' ? state.settingsSource !== 'built-in defaults' : !!state.data[kind];
    const n = kind === 'settings' ? (ok ? `${state.settings.length} rows` : 'Using defaults') : ok ? `${fmt0(state.data[kind].length)} rows` : 'Waiting';
    return `<div class="file ${ok ? 'ok' : ''} ${optional ? 'opt' : ''}"><b>${label}</b><span class="state">${ok ? '✓ ' : ''}${esc(n)}</span><div class="faint" style="font-size:12px">${desc}</div></div>`;
  };
  const changes = S.changes();
  const terr = S.entries('territory').map(([r, ids]) => `${r}: ${ids ? esc(ids) : '<i>nobody</i>'}`).join(' · ');
  return `
  <div class="wrap">
    <section class="hero">
      <h1>Each SDR's best accounts for the week, <em>with the reason behind every pick</em></h1>
      <p class="lead">Drop in the three case files. The prototype cleans the data, scores all 500 accounts on buying signals and fit, routes the best 80 to the four SDRs by territory and capacity, and explains every decision. Leadership can change the rules before or after the run.</p>
    </section>
    <div class="grid-2">
      <div class="card pad">
        <div class="step-label">Step 1 · Data</div>
        <h2>Drop the case files</h2>
        <div class="muted" style="font-size:14px">accounts, signals and sdrs, in any order and with any file name. Each file is recognised by its columns.</div>
        <label class="drop" id="drop">
          <input type="file" id="fileInput" multiple accept=".xlsx,.xls,.csv" hidden>
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 16V4m0 0l-4 4m4-4l4 4"/><path d="M4 16v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>
          <strong>Drop Excel files here or click to choose</strong>
          <span class="muted" style="font-size:13px">.xlsx or .csv · select several at once</span>
        </label>
        <div class="files">
          ${f('accounts', 'Accounts', '500 target accounts')}
          ${f('signals', 'Signals', 'Buying signals, last 90 days')}
          ${f('sdrs', 'SDRs', 'Team, territory, capacity')}
          ${f('settings', 'Settings', 'Optional: leadership rules', true)}
        </div>
        ${state.error ? `<div class="err">${esc(state.error)}</div>` : ''}
        <div class="cta-row">
          <button class="btn primary" id="runBtn" ${ready() ? '' : 'disabled'}>Build this week's lists →</button>
          ${state.noSample ? '' : '<button class="btn ghost" id="sampleBtn">Use the case data</button>'}
        </div>
      </div>
      <div class="card pad settings-sum">
        <div class="step-label">Step 2 · Optional</div>
        <h2>Leadership settings</h2>
        <div class="muted" style="font-size:14px">Every rule is a setting, not code. Change it now or after the run and watch the lists update.</div>
        <ul>
          <li>Run date <b>${esc(S.get('general', 'as_of_date'))}</b>, signals from the last <b>${esc(S.get('general', 'lookback_days'))} days</b>, weight halves every <b>${esc(S.get('general', 'half_life_days'))} days</b></li>
          <li>${S.entries('signal_weight').length} signal weights, ${S.entries('boost').filter(([, v]) => Number(v) !== 1).length} active segment boosts, ${realEntries('exclude').length} excluded and ${realEntries('pin').length} pinned accounts</li>
          <li>Territories: ${terr}</li>
        </ul>
        <div class="cta-row" style="margin-top:0">
          <button class="btn" data-open-settings>Review and edit settings</button>
          ${changes ? `<span class="chip warn">${changes} change${changes > 1 ? 's' : ''}</span>` : `<span class="chip">${esc(state.settingsSource)}</span>`}
        </div>
      </div>
    </div>
    <div class="how">
      <div class="card"><div class="n">1</div><b>Clean</b><p>Unify country names, read "NA" as North America, fix regions that contradict the country, count duplicate signals once, drop undated and stale signals. Every fix is logged.</p></div>
      <div class="card"><div class="n">2</div><b>Score</b><p>Intent × fit × CRM stage × leadership boost. Intent weighs each signal by what it says and how fresh it is. Fit is mostly devices in the field, because emnify earns per device.</p></div>
      <div class="card"><div class="n">3</div><b>Route</b><p>Best score first, within each SDR's territory, until their weekly capacity is full. Customers, open deals and recently touched accounts are held back.</p></div>
    </div>
    <footer>Prototype for the emnify GTM Engineer case study. All companies and people are fictional. The same logic runs as an n8n workflow; this page is a transparent front end over it.</footer>
  </div>`;
}

function kpis(r) {
  const fixes = r.issues.filter(i => i.area !== 'routing').length;
  const apac = r.scored.filter(s => s.status.startsWith('Not routed: no SDR')).length;
  const top = r.scored.find(s => s.status === 'Assigned');
  return `<div class="kpis">
    <div class="card kpi"><div class="v num">${r.lists.length}</div><div class="l">accounts assigned</div><div class="s">${esc(r.team.map(t => `${t.name.split(' ')[0]} ${t.filled}/${t.cap}`).join(' · '))}</div></div>
    <div class="card kpi"><div class="v num">${fmt0(r.summary.accounts)}</div><div class="l">accounts scored</div><div class="s">${r.scored.filter(s => !s.status.startsWith('Excluded') && !s.status.startsWith('No signal') && !s.status.startsWith('Cool')).length} had a reason to call this week</div></div>
    <div class="card kpi"><div class="v num">${r.summary.signals_used}<span class="faint" style="font-size:16px"> / ${r.summary.signals_in}</span></div><div class="l">signals used</div><div class="s">the rest were duplicates, undated or stale</div></div>
    <div class="card kpi"><div class="v num">${fixes}</div><div class="l">data fixes logged</div><div class="s">nothing changed silently</div></div>
    <div class="card kpi"><div class="v num">${apac}</div><div class="l">left for leadership</div><div class="s">${apac ? 'accounts with signals but no SDR in their region' : 'every region is covered'}</div></div>
  </div>`;
}

function renderResults() {
  const r = state.result;
  const tabs = [
    ['board', 'SDR lists', r.lists.length],
    ['funnel', 'How 500 became 80'],
    ['explain', 'How the score works'],
    ['fixes', 'Data fixes', r.issues.filter(i => i.area !== 'routing').length],
    ['messages', 'Monday messages'],
    ['all', 'All accounts', r.scored.length]
  ];
  const changes = S.changes();
  return `
  <div class="bar"><div class="wrap">
    <span class="title">Week of ${esc(r.summary.week_of)}</span>
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
    <footer>All companies and people are fictional. Scores are relative within this week's run. Tiers (Hot ≥ 20, Warm ≥ 8, Light touch below) are a reading aid only and don't change the ranking.</footer>
  </div>`;
}

function panel(tab) {
  return { board: pBoard, funnel: pFunnel, explain: pExplain, fixes: pFixes, messages: pMessages, all: pAll }[tab]();
}

// ---------- SDR board ----------
function accCard(l, s) {
  const isTop = l.rank === 1;
  const first = (l.why_now || '').split('; ')[0];
  return `<div class="acc ${isTop ? 'top' : ''} ${state.newlyIn.has(l.account_id) ? 'new' : ''}" data-acc="${esc(l.account_id)}">
    ${isTop ? '<div class="badge">Start here</div>' : ''}
    <div class="row1"><span class="rk">#${l.rank}</span><span class="nm">${esc(l.company_name)}</span>${l.pinned ? '<span class="chip info">Pinned</span>' : ''}<span class="tier ${tier(l.score)}">${tierName[tier(l.score)]}</span><span class="sc">${fmt(l.score)}</span></div>
    <div class="sub">${esc(l.country)} · ${esc(l.industry)}${l.crm_stage === 'Closed lost' ? ' · closed lost before' : ''}</div>
    <div class="why">${esc(first)}${(l.why_now || '').includes('; ') ? ` <span class="faint">+${l.why_now.split('; ').length - 1} more</span>` : ''}</div>
    ${isTop ? `<div class="angle">${esc(l.suggested_angle)}</div>` : ''}
  </div>`;
}
function pBoard() {
  const r = state.result;
  return `<p class="panel-intro">Each column is what one SDR works this week, best account first. Click any account to see exactly how its score was built and why it landed with this rep.</p>
  <div class="board">${r.team.map(t => {
    const rows = r.lists.filter(l => l.sdr === t.name);
    const open = state.expanded[t.id];
    const shown = open ? rows : rows.slice(0, 8);
    const pct = t.cap ? Math.round(100 * t.filled / t.cap) : 0;
    return `<div class="col">
      <div class="col-h">
        <div class="name">${esc(t.name)}</div>
        <div class="meta">${esc(t.territories.join(', ') || 'No territory')} · ${esc(t.language)} · ${t.tenure} months</div>
        <div class="cap"><i style="width:${pct}%"></i></div>
        <div class="cap-l"><span>${t.filled} of ${t.cap} slots</span><span>${rows.length ? `${rows.filter(x => tier(x.score) === 'A').length} hot` : ''}</span></div>
      </div>
      ${shown.map(l => accCard(l)).join('') || '<div class="note">No qualifying accounts in this territory.</div>'}
      ${rows.length > 8 ? `<button class="btn small ghost more" data-expand="${t.id}">${open ? 'Show fewer' : `Show all ${rows.length}`}</button>` : ''}
    </div>`;
  }).join('')}</div>`;
}

// ---------- funnel ----------
function funnelSteps() {
  const sc = state.result.scored;
  const by = f => sc.filter(f);
  const lb = S.get('general', 'lookback_days');
  return [
    { k: 'all', label: 'Accounts in the list', desc: 'Everything emnify sourced', items: sc, color: 'var(--faint)' },
    { k: 'crm', label: 'Already owned by sales or CS', desc: 'Customers and open opportunities belong to AEs and CS; an SDR call would cause conflict', items: by(s => s.status.startsWith('Excluded: CRM')), minus: true, color: 'var(--bad)' },
    { k: 'lead', label: 'Excluded by leadership', desc: 'Single accounts or whole segments switched off in settings', items: by(s => s.status.startsWith('Excluded by')), minus: true, color: 'var(--bad)' },
    { k: 'nosig', label: `No signal in ${lb} days`, desc: 'Good companies, but nothing says "call this week"', items: by(s => s.status.startsWith('No signal')), minus: true, color: 'var(--warn)' },
    { k: 'cool', label: 'Recently touched', desc: `Contacted in the last ${S.get('general', 'recent_touch_days')} days, so they rest`, items: by(s => s.status.startsWith('Cool')), minus: true, color: 'var(--warn)' },
    { k: 'nosdr', label: 'No SDR covers the region', desc: 'A staffing decision for leadership, flagged rather than forced onto the wrong rep', items: by(s => s.status.startsWith('Not routed: no SDR')), minus: true, color: 'var(--info)' },
    { k: 'cap', label: 'Over this week\'s capacity', desc: 'Qualified, but each rep is full. These are next week\'s pipeline', items: by(s => s.status.includes('at capacity')), minus: true, color: 'var(--info)' },
    { k: 'ok', label: 'Assigned to an SDR', desc: 'The best accounts that fit each rep\'s territory and capacity', items: by(s => s.status === 'Assigned'), color: 'var(--accent)', final: true }
  ];
}
function pFunnel() {
  const steps = funnelSteps();
  const total = steps[0].items.length || 1;
  const open = steps.find(s => s.k === state.funnelOpen);
  return `<p class="panel-intro">Every one of the ${total} accounts ends up somewhere, for a stated reason. Click a step to see the accounts in it.</p>
  <div class="funnel">${steps.filter(s => s.items.length || !s.minus).map(s => `
    <div class="card fstep ${s.final ? 'final' : ''}" data-funnel="${s.k}">
      <div class="fl"><b>${s.minus ? '− ' : s.final ? '= ' : ''}${s.label}</b><span>${s.desc}</span></div>
      <div class="fbar"><i style="width:${Math.max(1.5, 100 * s.items.length / total)}%;background:${s.color}"></i></div>
      <div class="fc">${s.items.length}</div>
    </div>`).join('')}</div>
  ${open ? `<div class="card pad" style="margin-top:16px"><h3>${open.label} <span class="faint">(${open.items.length})</span></h3>${accTable(open.items.slice(0, 300))}</div>` : ''}`;
}

function accTable(items) {
  return `<div class="tbl-wrap"><table><thead><tr><th>Company</th><th>Region</th><th>Industry</th><th>CRM stage</th><th class="r">Score</th><th>Status</th></tr></thead><tbody>
  ${items.map(s => `<tr class="click" data-acc="${esc(s.account_id)}"><td><b>${esc(s.company_name)}</b><div class="faint mono">${esc(s.account_id)}</div></td><td>${esc(s.region || '–')}</td><td>${esc(s.industry)}</td><td>${esc(s.crm_stage)}</td><td class="r num">${fmt(s.score)}</td><td><span class="chip ${statusKind(s.status)}">${esc(s.status.replace('Not routed: ', ''))}</span>${s.assigned_to ? ` <span class="faint">${esc(s.assigned_to)} #${s.rank}</span>` : ''}</td></tr>`).join('')}
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
  return `<p class="panel-intro">The score answers one question: which accounts should this rep call this week? Each part below is a setting leadership can change.</p>
  <div class="formula">
    <div class="card fx"><div class="t">Intent</div><p>Is the account actively moving right now? Built from its signals: weight × freshness.</p></div><div class="op">×</div>
    <div class="card fx"><div class="t">Fit</div><p>How valuable would a win be? Mostly devices in the field, plus provider and size.</p></div><div class="op">×</div>
    <div class="card fx"><div class="t">CRM stage</div><p>Is it an SDR's account at all? Customers and open deals get 0.</p></div><div class="op">×</div>
    <div class="card fx"><div class="t">Boost</div><p>Leadership's lever: push or switch off a segment for this week.</p></div><div class="op">=</div>
    <div class="card fx" style="border-color:var(--accent)"><div class="t">Score</div><p>Multiplied, not added: a huge account with no buying activity scores 0, not "medium".</p></div>
  </div>
  <div class="ex-grid">
    <div class="card"><h3>1. Signals are weighted by what they say, not just their type</h3>
      <p>The type labels mislead: a single blog visit is filed under "Website visit (pricing or docs page)", next to three pricing-page visits. So each signal text gets its own weight from 0 to 10. The closer to a buying decision, the higher: a contract renewal is a switching window, a pricing visit is active evaluation, a blog visit is curiosity.</p>
      ${typeOrder.map(t => {
        const all = [...types[t].entries()];
        const own = all.filter(([d]) => wDetail[d] !== undefined).sort((a, b) => Number(wDetail[b[0]]) - Number(wDetail[a[0]]));
        const rest = all.filter(([d]) => wDetail[d] === undefined);
        const bar = (label, w, n) => `<div class="wbar"><span>${label} <span class="faint">(${n})</span></span><span class="b"><i style="width:${w * 10}%"></i></span><span class="v">${w}</span></div>`;
        return `<div class="grp">${esc(t)}</div>${own.map(([d, n]) => bar(esc(d), Number(wDetail[d]), n)).join('')}
          ${rest.length ? bar(`${own.length ? 'All other' : 'All'} ${rest.length} variant${rest.length > 1 ? 's' : ''}, e.g. <i>${esc(rest[0][0])}</i>`, Number(wType[t] ?? 3), rest.reduce((x, [, n]) => x + n, 0)) : ''}`;
      }).join('')}
      <p class="faint" style="font-size:12.5px;margin-top:10px">Numbers in brackets: how many times the signal appears in the data. Variants without their own weight use their type's default.</p>
    </div>
    <div>
      <div class="card"><h3>2. Fresh signals count more</h3>
        <p>A signal's weight halves every <b>${hl} days</b> and drops out after <b>${lb} days</b>. A weekly list is about <i>why now</i>: last week's pricing visit beats one from two months ago.</p>
        <svg viewBox="0 0 320 120" style="width:100%;max-width:420px;margin-top:6px"><line x1="0" y1="100" x2="300" y2="100" stroke="var(--border)"/><path d="${path}" fill="none" stroke="var(--accent)" stroke-width="2.5"/>
          ${[0, hl, Math.min(2 * hl, lb), lb].map(d => `<text x="${(d / lb) * 300}" y="116" font-size="10" fill="var(--faint)" text-anchor="middle">${d}d</text><circle cx="${(d / lb) * 300}" cy="${100 - Math.pow(.5, d / hl) * 90}" r="3" fill="var(--accent)"/><text x="${(d / lb) * 300 + 6}" y="${96 - Math.pow(.5, d / hl) * 90}" font-size="10" fill="var(--muted)">${Math.round(Math.pow(.5, d / hl) * 100)}%</text>`).join('')}
        </svg>
      </div>
      <div class="card" style="margin-top:16px"><h3>3. Several kinds of activity beat one kind repeated</h3>
        <p>Within one signal type, the strongest counts fully and each repeat counts <b>${Math.round(100 * (S.get('general', 'extra_signal_same_type') ?? .5))}%</b>, so five downloads can't outweigh one renewal. Each additional <i>type</i> adds <b>${Math.round(100 * (S.get('general', 'multi_type_bonus') ?? .15))}%</b>: pricing visits plus engineering hires plus a new CTO means several teams are moving, which is how real buying committees look.</p>
      </div>
      <div class="card" style="margin-top:16px"><h3>4. Fit: devices first</h3>
        <p>emnify earns per connected device, so devices in the field are the best proxy for deal size. A named competitor or several providers is a plus: they already buy IoT connectivity and can switch. Headcount barely matters, because it says little about device volume. Unknown values count as ×1, so missing data never punishes an account.</p>
        <div class="tbl-wrap" style="margin-top:8px"><table><thead><tr><th>Devices in the field</th><th class="r">Multiplier</th></tr></thead><tbody>${devSteps.map(([k, v], i) => `<tr><td>${i === 0 ? `${fmt0(k)}+` : `${fmt0(k)} – ${fmt0(devSteps[i - 1][0] - 1)}`}</td><td class="r num">×${esc(v)}</td></tr>`).join('')}</tbody></table></div>
      </div>
    </div>
    <div class="card"><h3>5. CRM stage decides whose account it is</h3><p>Customers and open opportunities are worked by AEs and CS, so they're set to 0. Closed-lost accounts are kept at a discount: they said no once, but a renewal window can change that.</p><div class="tbl-wrap" style="margin-top:8px">${tbl('crm_stage', 'Stage')}</div></div>
    <div class="card"><h3>6. Provider and size</h3><div class="tbl-wrap">${tbl('fit_provider', 'Current provider')}</div><div class="tbl-wrap" style="margin-top:10px">${tbl('fit_size', 'Employees')}</div></div>
    <div class="card"><h3>7. Routing</h3><ul>
      <li><b>Territory first.</b> ${S.entries('territory').map(([r, ids]) => `${esc(r)} → ${ids ? esc(ids.split(',').map(id => state.result.team.find(t => t.id === id.trim())?.name || id).join(' + ')) : '<i>nobody</i>'}`).join('; ')}.</li>
      <li><b>Best score first</b> until each rep reaches their weekly capacity from the SDR file.</li>
      <li><b>Shared territory:</b> the next account goes to the rep with the most free capacity as a share of their total, so both reps get a fair mix of strong and weaker accounts. Ties go to the more experienced rep.</li>
      <li><b>Pinned accounts</b> go first, whatever their score.</li>
      <li><b>Uncovered regions</b> are reported, not forced onto a rep in the wrong time zone.</li></ul></div>
    <div class="card"><h3>8. Ignored on purpose</h3><ul>
      <li><b>funding_stage:</b> it contradicts the dated funding signals (accounts marked "Public" with a recent "Raised 5M Seed" signal). The dated signal is trusted instead.</li>
      <li><b>language:</b> each territory already matches its SDR's languages.</li>
      <li><b>source:</b> kept for transparency, not weighted. The signal text already says how strong it is.</li>
      <li><b>Machine learning:</b> there is no win/loss history to learn from yet. The weights are explicit judgement that leadership can read and change, to be tuned on conversion data after a few weeks.</li></ul></div>
  </div>`;
}

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
  return `<p class="panel-intro">The brief says the data is deliberately messy. Every correction is logged here, so the result can be audited and the source data fixed upstream.</p>
  ${Object.entries(groups).map(([g, list]) => `<details class="set" ${state.fixOpen === g ? 'open' : ''} data-fix="${esc(g)}"><summary>${esc(g)} <span class="chip">${list.length}</span></summary><div class="sb"><p class="why">${esc(why[g] || '')}</p>
    <div class="tbl-wrap"><table><thead><tr><th>Item</th><th>Problem</th><th>Action</th></tr></thead><tbody>${list.map(i => `<tr><td>${esc(names[i.item] || i.item)} <span class="faint mono">${names[i.item] ? esc(i.item) : ''}</span></td><td>${esc(i.problem)}</td><td>${esc(i.action)}</td></tr>`).join('')}</tbody></table></div></div></details>`).join('')}`;
}

// ---------- Monday messages ----------
function message(t) {
  const rows = state.result.lists.filter(l => l.sdr === t.name);
  if (!rows.length) return `Good morning ${t.name.split(' ')[0]}! No qualifying accounts in your territory this week.`;
  const top = rows[0];
  const lines = rows.slice(0, 10).map(r => `${r.rank}. ${r.company_name} (${r.country}, score ${r.score}): ${r.why_now}`);
  return `Good morning ${t.name.split(' ')[0]}! You have ${rows.length} accounts for the week of ${state.result.summary.week_of}.\n\nStart with ${top.company_name}: ${top.why_now}.\nAngle: ${top.suggested_angle}` + (top.loop_in_ae ? `\nLoop in AE ${top.loop_in_ae}.` : '') + `\n\nTop 10:\n${lines.join('\n')}`;
}
function pMessages() {
  return `<p class="panel-intro">What each rep receives at 07:00 on Monday. In production this is a Slack or email message, with the full list as CRM tasks. Here you can copy it.</p>
  <div class="msgs">${state.result.team.map(t => `<div class="card msg"><div class="mh"><span>To <b>${esc(t.name)}</b> · ${esc(t.territories.join(', '))}</span><button class="btn small" data-copy="${esc(t.id)}">Copy</button></div><pre>${esc(message(t))}</pre></div>`).join('')}</div>`;
}

// ---------- all accounts ----------
function pAll() {
  const q = state.search.toLowerCase();
  const items = state.result.scored.filter(s => (state.statusFilter === 'all' || s.status.startsWith(state.statusFilter)) &&
    (!q || `${s.company_name} ${s.account_id} ${s.industry} ${s.country}`.toLowerCase().includes(q)));
  const opts = [['all', 'All statuses'], ['Assigned', 'Assigned'], ['Not routed', 'Not routed'], ['No signal', 'No signal'], ['Cool-down', 'Cool-down'], ['Excluded', 'Excluded']];
  return `<div class="toolbar"><input type="search" id="search" placeholder="Search company, ID, industry, country" value="${esc(state.search)}">
    <select id="statusFilter">${opts.map(([v, l]) => `<option value="${v}" ${state.statusFilter === v ? 'selected' : ''}>${l}</option>`).join('')}</select>
    <span class="muted" style="font-size:13px">${items.length} accounts</span></div>
    <div class="card">${accTable(items)}</div>`;
}

// ---------- account drawer ----------
function openAccount(id) {
  const s = state.result.scored.find(x => x.account_id === id);
  if (!s) return;
  const b = s.breakdown;
  const ex = S.get('exclude', id) !== undefined, pin = S.get('pin', id) !== undefined;
  const dropped = state.result.issues.filter(i => i.area === 'signals' && i.item === id);
  const why = s.status === 'Assigned'
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
      <div style="margin-top:8px"><span class="chip ${statusKind(s.status)}">${esc(s.status)}</span>${s.status === 'Assigned' ? ` <span class="tier ${tier(s.score)}">${tierName[tier(s.score)]}</span>` : ''}</div></div>
      <button class="btn small ghost x" data-close aria-label="Close">✕</button></div>
    <div class="db">
      <div class="note">${why}</div>
      <div class="eq">
        <div class="p"><b>${fmt(b.intent)}</b><span>intent</span></div><span class="o">×</span>
        <div class="p"><b>${fmt(b.fit, 2)}</b><span>fit</span></div><span class="o">×</span>
        <div class="p"><b>${fmt(b.stageMult, 2)}</b><span>CRM stage</span></div><span class="o">×</span>
        <div class="p"><b>${fmt(b.boost, 2)}</b><span>boost</span></div><span class="o">=</span>
        <div class="p res"><b>${fmt(s.score)}</b><span>score</span></div>
      </div>
      <div class="sec"><h3>Intent: the signals behind it</h3>
        ${b.signals.length ? `<div class="tbl-wrap"><table><thead><tr><th>Signal</th><th class="r">Weight</th><th class="r">Age</th><th class="r">Fresh</th><th class="r">Counts</th><th class="r">Points</th></tr></thead><tbody>
        ${b.signals.map(x => `<tr><td>${esc(x.detail)}<div class="faint" style="font-size:12px">${esc(x.type)} · ${esc(x.source)} · ${esc(x.date)}</div></td><td class="r num">${x.weight}</td><td class="r num">${x.age}d</td><td class="r num">${Math.round(x.decay * 100)}%</td><td class="r num">${x.share === 1 ? '100%' : `${Math.round(x.share * 100)}%`}</td><td class="r num"><b>${fmt(x.counted, 2)}</b></td></tr>`).join('')}
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
        <div class="frow"><label>Run date<div class="hint">signals are aged from this date</div></label><input type="date" data-sec="general" data-key="as_of_date" value="${esc(S.get('general', 'as_of_date'))}"></div>
        ${numRow('general', 'lookback_days', 'Lookback window (days)', 'older signals are dropped', '1')}
        ${numRow('general', 'half_life_days', 'Half-life (days)', 'weight halves every N days', '1')}
        ${numRow('general', 'recent_touch_days', 'Cool-down after a touch (days)', '', '1')}
        ${numRow('general', 'extra_signal_same_type', 'Repeat of the same type counts', '0.5 = half', '0.05')}
        ${numRow('general', 'multi_type_bonus', 'Bonus per extra signal type', '0.15 = +15%', '0.05')}
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
  for (const t of r.team) XLSX.utils.book_append_sheet(wb, sheet(r.lists.filter(l => l.sdr === t.name).map(({ week_of, sdr, ...x }) => ({ ...x, tier: tierName[tier(x.score)] }))), t.name.slice(0, 31));
  XLSX.utils.book_append_sheet(wb, sheet(r.lists), 'All assigned');
  XLSX.utils.book_append_sheet(wb, sheet(r.scored.map(({ breakdown, ...x }) => x)), 'All accounts scored');
  XLSX.utils.book_append_sheet(wb, sheet(r.issues), 'Data issues');
  XLSX.utils.book_append_sheet(wb, sheet(state.settings.map(({ section, key, value }) => ({ section, key, value }))), 'Settings used');
  XLSX.writeFile(wb, `sdr_lists_week_of_${r.summary.week_of}.xlsx`);
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
  const sf = document.getElementById('statusFilter');
  if (sf) sf.addEventListener('change', e => { state.statusFilter = e.target.value; document.getElementById('panel').innerHTML = panel('all'); bindCommon(); });
  document.querySelectorAll('details[data-fix]').forEach(d => d.addEventListener('toggle', () => { if (d.open) state.fixOpen = d.dataset.fix; }));
}

document.addEventListener('click', async e => {
  const t = e.target.closest('button, [data-acc], [data-funnel], [data-close], [data-tab]');
  if (!t) return;
  if (t.id === 'runBtn' || t.id === 'runFromSettings') { closeOverlay(); state.tab = 'board'; run(); window.scrollTo(0, 0); return; }
  if (t.id === 'sampleBtn') return loadSample();
  if (t.id === 'dlBtn') return downloadExcel();
  if (t.id === 'dlSettings') return downloadSettings();
  if (t.id === 'resetBtn') { Object.assign(state, { data: { accounts: null, signals: null, sdrs: null }, names: {}, result: null, settings: clone(DEFAULT_SETTINGS), baseSettings: clone(DEFAULT_SETTINGS), settingsSource: 'built-in defaults', error: '' }); return render(); }
  if (t.id === 'resetSettings') { state.settings = clone(state.baseSettings); openSettings(); return afterSettingsChange(); }
  if (t.hasAttribute('data-close')) return closeOverlay();
  if (t.hasAttribute('data-open-settings')) return openSettings();
  if (t.dataset.tab) { state.tab = t.dataset.tab; return render(); }
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
  if (t.dataset.acc && state.result) return openAccount(t.dataset.acc);
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
  const v = el.type === 'number' ? (el.value === '' ? '' : Number(el.value)) : el.value;
  S.set(el.dataset.sec, el.dataset.key, v);
  refreshSettingsMarks();
  afterSettingsChange();
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeOverlay(); });

// offer "Use the case data" only when the files are bundled in /data
state.noSample = true;
render();
fetch('data/accounts.xlsx', { method: 'HEAD' }).then(r => { state.noSample = !r.ok; if (!state.result) render(); }).catch(() => {});
