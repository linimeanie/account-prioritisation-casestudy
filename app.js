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
const tier = s => (s >= 20 ? 'A' : s >= 8 ? 'B' : 'C');
// Tiers use food temperature instead of sales jargon ("warm" already means an existing relationship in sales).
// Each one stands for an action: baked potato = call first, cream soup = call + email, sushi = email only.
const tierName = { A: '🥔 Baked potato', B: '🍲 Cream soup', C: '🍣 Sushi' };
const tierAction = { A: 'call first', B: 'call, then follow up by email', C: 'one email is enough' };
const tierLegend = () => `<div class="tier-legend"><span><span class="tier A">${tierName.A}</span> hot, score 20+: ${tierAction.A}</span><span><span class="tier B">${tierName.B}</span> warm, 8 to 20: ${tierAction.B}</span><span><span class="tier C">${tierName.C}</span> served cold, under 8: ${tierAction.C}</span></div>`;
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

function sdrName(id) {
  const r = (state.data.sdrs || []).find(r => String(r.sdr_id) === String(id).trim());
  return r ? String(r.name).split(' ')[0] : String(id).trim();
}
function territoryText() {
  return S.entries('territory').map(([reg, ids]) => `${reg} → ${ids ? String(ids).split(',').map(sdrName).join(' + ') : 'nobody'}`);
}

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
      <p class="lead">Drop in the three case files. The prototype cleans the data, scores all 500 accounts, routes the best 80 to the four SDRs and explains every decision. Leadership can change the rules before or after the run.</p>
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
          <dt>Run date</dt><dd>${esc(S.get('general', 'as_of_date'))}</dd>
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
    <footer>Prototype for the emnify GTM Engineer case study. All companies and people are fictional. The same logic runs as an n8n workflow; this page is a transparent front end over it.</footer>
  </div>`;
}

function kpis(r) {
  const fixes = r.issues.filter(i => i.area !== 'routing').length;
  const apac = r.scored.filter(s => s.status.startsWith('Not routed: no SDR')).length;
  const reason = r.scored.filter(s => !s.status.startsWith('Excluded') && !s.status.startsWith('No signal') && !s.status.startsWith('Cool')).length;
  const k = (v, l, s) => `<div class="card kpi"><div class="v num">${v}</div><div class="l">${l}</div><div class="s">${s}</div></div>`;
  return `<div class="kpis">
    ${k(r.lists.length, 'accounts assigned', esc(r.team.map(t => `${t.name.split(' ')[0]} ${t.filled}`).join(' · ')))}
    ${k(fmt0(r.summary.accounts), 'accounts scored', `${reason} with a reason to call`)}
    ${k(`${r.summary.signals_used}<span class="of"> / ${r.summary.signals_in}</span>`, 'signals used', 'rest: duplicate, undated, stale')}
    ${k(fixes, 'data fixes logged', 'each one listed in Data fixes')}
    ${k(apac, 'left for leadership', apac ? 'signals, but no SDR in region' : 'every region is covered')}
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
    <footer>All companies and people are fictional. Scores are relative within this week's run. The food tiers (baked potato ≥ 20, cream soup ≥ 8, sushi below) are a reading aid only and don't change the ranking.</footer>
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
    <div class="row1"><span class="rk">#${l.rank}</span><span class="nm">${esc(l.company_name)}</span>${l.pinned ? '<span class="chip info">Pinned</span>' : ''}<span class="food" title="${esc(tierName[tier(l.score)].slice(3))}: ${tierAction[tier(l.score)]}">${tierName[tier(l.score)].slice(0, 2)}</span></div>
    <div class="meta-row"><span class="sub">${esc(l.country)} · ${esc(l.industry)}${l.crm_stage === 'Closed lost' ? ' · closed lost' : ''}</span></div>
    <div class="why">${esc(first)}${(l.why_now || '').includes('; ') ? ` <span class="faint">+${l.why_now.split('; ').length - 1} more</span>` : ''}</div>
    ${isTop ? `<div class="angle">${esc(l.suggested_angle)}</div>` : ''}
  </div>`;
}
function pBoard() {
  const r = state.result;
  return `<p class="panel-intro">One column per SDR, best account first. Click any account to see how its score was built and why it landed with that rep.</p>
  ${tierLegend()}
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
        <div class="cap-l"><span>${t.filled} of ${t.cap} slots</span><span>${rows.length ? `🥔 ${rows.filter(x => tier(x.score) === 'A').length} to call first` : ''}</span></div>
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
    { k: 'all', label: 'Accounts in the list', desc: 'Every account emnify sourced for the motion', items: sc, color: 'var(--faint)' },
    { k: 'crm', label: 'Already owned by sales or CS', desc: 'Customers and open deals belong to AEs and CS', items: by(s => s.status.startsWith('Excluded: CRM')), minus: true, color: 'var(--bad)' },
    { k: 'lead', label: 'Excluded by leadership', desc: 'Accounts or segments switched off in settings', items: by(s => s.status.startsWith('Excluded by')), minus: true, color: 'var(--bad)' },
    { k: 'nosig', label: `No signal in ${lb} days`, desc: 'Good companies, but no reason to call this week', items: by(s => s.status.startsWith('No signal')), minus: true, color: 'var(--warn)' },
    { k: 'cool', label: 'Recently touched', desc: `Contacted in the last ${S.get('general', 'recent_touch_days')} days, so they rest`, items: by(s => s.status.startsWith('Cool')), minus: true, color: 'var(--warn)' },
    { k: 'nosdr', label: 'No SDR covers the region', desc: 'A staffing decision, flagged for leadership', items: by(s => s.status.startsWith('Not routed: no SDR')), minus: true, color: 'var(--info)' },
    { k: 'cap', label: 'Over this week\'s capacity', desc: 'Qualified, but reps are full: next week\'s pipeline', items: by(s => s.status.includes('at capacity')), minus: true, color: 'var(--info)' },
    { k: 'ok', label: 'Assigned to an SDR', desc: 'Best accounts within territory and capacity', items: by(s => s.status === 'Assigned'), color: 'var(--accent)', final: true }
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
      const own = all.filter(([d]) => wDetail[d] !== undefined).sort((a, b) => Number(wDetail[b[0]]) - Number(wDetail[a[0]]));
      const rest = all.filter(([d]) => wDetail[d] === undefined);
      const bar = (label, w, n) => `<div class="wbar"><span class="wl">${label} <span class="faint">(${n})</span></span><span class="b"><i style="width:${w * 10}%"></i></span><span class="v">${w}</span></div>`;
      return `<div class="wgroup"><div class="grp">${esc(t)}</div>${own.map(([d, n]) => bar(esc(d), Number(wDetail[d]), n)).join('')}
        ${rest.length ? bar(`${own.length ? 'Other' : 'All'} ${rest.length} variant${rest.length > 1 ? 's' : ''}, e.g. <i>${esc(rest[0][0])}</i>`, Number(wType[t] ?? 3), rest.reduce((x, [, n]) => x + n, 0)) : ''}</div>`;
    }).join('')}
    </div>
    <p class="faint" style="font-size:12.5px;margin:12px 0 0">In brackets: how often the signal appears in the data. Variants without their own weight use their type's default.</p>
  </div>
  <div class="ex-grid">
    <div class="card"><h3>2. Fresh signals count more</h3>
      <p>A signal's weight halves every <b>${hl} days</b> and drops out after <b>${lb} days</b>. A weekly list is about <i>why now</i>: last week's pricing visit beats one from two months ago.</p>
      <svg viewBox="-12 -14 340 136" class="decay"><line x1="0" y1="100" x2="300" y2="100" stroke="var(--border)"/><path d="${path}" fill="none" stroke="var(--accent)" stroke-width="2.5"/>
        ${[0, hl, Math.min(2 * hl, lb), lb].map(d => `<text x="${(d / lb) * 300}" y="116" font-size="10" fill="var(--faint)" text-anchor="middle">${d}d</text><circle cx="${(d / lb) * 300}" cy="${100 - Math.pow(.5, d / hl) * 90}" r="3" fill="var(--accent)"/><text x="${(d / lb) * 300 + 6}" y="${94 - Math.pow(.5, d / hl) * 90}" font-size="10" fill="var(--muted)">${Math.round(Math.pow(.5, d / hl) * 100)}%</text>`).join('')}
      </svg>
    </div>
    <div class="card"><h3>3. Breadth beats repetition</h3>
      <p>Within one signal type, the strongest counts fully and each repeat counts <b>${Math.round(100 * (S.get('general', 'extra_signal_same_type') ?? .5))}%</b>, so five downloads can't outweigh one renewal.</p>
      <p>Each additional <i>type</i> adds <b>${Math.round(100 * (S.get('general', 'multi_type_bonus') ?? .15))}%</b>. Pricing visits plus engineering hires plus a new CTO means several teams are moving, which is what a real buying committee looks like.</p>
    </div>
    <div class="card"><h3>4. Fit: devices first</h3>
      <p>emnify earns per connected device, so devices in the field are the best proxy for deal size. Unknown values count as ×1, so missing data never punishes an account.</p>
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
  <div class="card ex-wide" style="margin-top:16px"><h3>8. Ignored on purpose</h3>
    <div class="ignored">
      <div><b>funding_stage</b><p>Contradicts the dated funding signals: accounts marked "Public" carry a recent "Raised 5M Seed" signal. The dated signal is trusted instead.</p></div>
      <div><b>language</b><p>Each territory already matches its SDR's languages, so it adds nothing to routing in this data.</p></div>
      <div><b>source</b><p>Kept for transparency, not weighted. The signal text already says how strong the signal is.</p></div>
      <div><b>Machine learning</b><p>No win/loss history to learn from yet. Explicit weights leadership can read, tuned on conversion after a few weeks.</p></div>
    </div>
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
  const hot = rows.filter(r => tier(r.score) === 'A'), warm = rows.filter(r => tier(r.score) === 'B'), light = rows.filter(r => tier(r.score) === 'C');
  const top = rows[0], topAcc = byId[top.account_id];
  const alsoHot = hot.slice(1, 6), hotOverflow = hot.slice(6);
  const blocks = [
    { p: `Hi ${first},` },
    { p: `Your list for this week is ready: ${rows.length} accounts${hot.length ? `, ${hot.length} of them baked potatoes to call first` : ''}. Here's where to start.` },
    { h: `Start with ${top.company_name}`, sub: `${top.country} · ${top.industry}`, id: top.account_id },
    { p: 'Why now:' },
    { list: reasons(topAcc).map(text => ({ text })) },
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
  const ae = rows.filter(r => r.loop_in_ae && r !== top), lost = rows.filter(r => r.crm_stage === 'Closed lost'), pinned = rows.filter(r => r.pinned === 'yes');
  const heads = [];
  if (pinned.length) heads.push(`Leadership asked for ${names(pinned.map(r => r.company_name))} to be worked this week.`);
  const aeGroups = {};
  for (const r of ae) (aeGroups[r.loop_in_ae] = aeGroups[r.loop_in_ae] || []).push(r.company_name);
  for (const [who, list] of Object.entries(aeGroups)) heads.push(`Loop in ${who} before contacting ${names(list)}.`);
  if (lost.length) heads.push(`${names(lost.map(r => r.company_name))} ${lost.length > 1 ? 'were' : 'was'} closed lost before. Acknowledge the history and lead with what has changed.`);
  if (heads.length) blocks.push({ h: 'Heads-up' }, { list: heads.map(text => ({ text })) });
  blocks.push({ p: 'Your full list, with every reason and angle, is in the CRM.', muted: true }, { p: 'Have a great week!' });
  return { subject: `Your ${rows.length} accounts for the week of ${week}`, blocks };
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
  return `<div class="msg-intro card">
      <b>How the brief is built</b>
      <p>It's a briefing, not the list. Every rep gets the same shape, whether they have 15 accounts or 25: the account to start with and why, a line for each baked potato, a count and a few names for the rest, and heads-ups on AEs, closed-lost history and leadership pins. Every name opens the account. No scores, because they mean nothing to a rep.</p>
      ${tierLegend()}
    </div>
    <div class="msgs">${state.result.team.map(t => `<div class="card msg"><div class="mh"><span>To <b>${esc(t.name)}</b> · Monday 07:00</span><button class="btn small" data-copy="${esc(t.id)}">Copy</button></div><div class="mail">${briefHtml(brief(t))}</div></div>`).join('')}</div>`;
}

// ---------- all accounts ----------
// All accounts: every column can be sorted (click the header) and filtered (the row under it).
const statusGroup = s => s.status === 'Assigned' ? 'Assigned'
  : s.status.startsWith('Not routed: no SDR') ? 'No SDR for region'
  : s.status.includes('at capacity') ? 'Over capacity'
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
    if (c.filter === 'tier') return v === 'unassigned' ? s.status !== 'Assigned' : s.status === 'Assigned' && tier(s.score) === v;
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
  return `<div class="toolbar"><span class="muted" style="font-size:13px"><b style="color:var(--text)">${items.length}</b> of ${all.length} accounts · click a column to sort, use the row below it to filter</span>
    ${active ? `<button class="btn small" id="clearFilters">Clear ${active} filter${active > 1 ? 's' : ''}</button>` : ''}</div>
    <div class="card"><div class="tbl-wrap all-tbl"><table><thead>
      <tr>${ALL_COLS.map(c => `<th class="sortable ${c.num ? 'r' : ''} ${state.sort.key === c.k ? 'on' : ''}" data-sort="${c.k}">${c.label}<span class="arr">${state.sort.key === c.k ? (state.sort.dir === 'asc' ? '↑' : '↓') : '↕'}</span></th>`).join('')}</tr>
      <tr class="frow-h">${ALL_COLS.map(c => `<th>${fcell(c)}</th>`).join('')}</tr>
    </thead><tbody>
    ${items.map(s => `<tr class="click" data-acc="${esc(s.account_id)}">
      <td><b>${esc(s.company_name)}</b><div class="faint mono">${esc(s.account_id)}</div></td>
      <td>${esc(s.region || '–')}</td><td>${esc(s.country || '–')}</td><td>${esc(s.industry)}</td><td>${esc(s.crm_stage)}</td>
      <td class="r num">${fmt(s.score)}${s.status === 'Assigned' ? ` <span title="${esc(tierAction[tier(s.score)])}">${tierName[tier(s.score)].slice(0, 2)}</span>` : ''}</td>
      <td class="r num">${s.signals_used}</td>
      <td><span class="chip ${statusKind(s.status)}">${esc(statusGroup(s))}</span></td>
      <td>${s.assigned_to ? `${esc(s.assigned_to.split(' ')[0])} <span class="faint">#${s.rank}</span>` : '<span class="faint">–</span>'}</td></tr>`).join('') || `<tr><td colspan="${ALL_COLS.length}" class="faint" style="text-align:center;padding:24px">No accounts match these filters.</td></tr>`}
    </tbody></table></div></div>`;
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
  document.querySelectorAll('[data-colf]').forEach(el => el.addEventListener(el.tagName === 'INPUT' ? 'input' : 'change', e => {
    state.colFilters[el.dataset.colf] = e.target.value;
    const k = el.dataset.colf, pos = e.target.selectionStart;
    document.getElementById('panel').innerHTML = panel('all'); bindCommon();
    if (el.tagName === 'INPUT') { const n = document.querySelector(`input[data-colf="${k}"]`); n.focus(); n.setSelectionRange(pos, pos); }
  }));
  document.querySelectorAll('details[data-fix]').forEach(d => d.addEventListener('toggle', () => { if (d.open) state.fixOpen = d.dataset.fix; }));
}

document.addEventListener('click', async e => {
  const t = e.target.closest('button, [data-acc], [data-funnel], [data-close], [data-tab], [data-goto-list], th[data-sort]');
  if (!t) return;
  if (t.id === 'runBtn' || t.id === 'runFromSettings') { closeOverlay(); state.tab = 'board'; run(); window.scrollTo(0, 0); return; }
  if (t.id === 'sampleBtn') return loadSample();
  if (t.dataset.sort) { const k = t.dataset.sort; const num = ALL_COLS.find(c => c.k === k)?.num; state.sort = state.sort.key === k ? { key: k, dir: state.sort.dir === 'asc' ? 'desc' : 'asc' } : { key: k, dir: num ? 'desc' : 'asc' }; return render(); }
  if (t.id === 'clearFilters') { state.colFilters = {}; return render(); }
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
