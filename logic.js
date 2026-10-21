// emnify weekly account prioritisation: clean -> score -> route -> explain.
// Same logic as the "Score and route" node in the n8n workflow. The only addition is
// `breakdown` on each scored account, so the dashboard can show how every score was built.

// ctx.worked: accounts assigned in earlier weeks, { account_id: { date, sdr } }. Empty for a single run.
export function prioritise(accounts, signals, sdrs, settingsRows, ctx = {}) {
  const issues = [];
  const issue = (area, item, problem, action) => issues.push({ area, item, problem, action });
  const str = v => (v === null || v === undefined ? '' : String(v).trim());
  const num = v => (str(v) === '' || isNaN(Number(v)) ? null : Number(v));

  const daysBetween = (a, b) => Math.round((a - b) / 86400000);
  // dates may arrive as text (2026-10-12) or as Excel serial numbers (46307)
  const toDate = v => (typeof v === 'number' || /^\d{5}(\.\d+)?$/.test(str(v)) ? new Date((Number(v) - 25569) * 86400000) : new Date(str(v)));

  // ---------- settings ----------
  const S = {};
  for (const r of settingsRows) {
    const sec = str(r.section), key = str(r.key);
    if (!sec || !key) continue;
    (S[sec] = S[sec] || {})[key] = r.value;
  }
  const g = S.general || {};
  const asOf = str(g.as_of_date) ? toDate(g.as_of_date) : new Date(new Date().toISOString().slice(0, 10));
  const lookback = num(g.lookback_days) ?? 90;
  const halfLife = num(g.half_life_days) ?? 30;
  const touchCooldown = num(g.recent_touch_days) ?? 14;
  const sameTypeShare = num(g.extra_signal_same_type) ?? 0.5;
  const multiTypeBonus = num(g.multi_type_bonus) ?? 0.15;
  const maxReasons = num(g.max_reasons) ?? 3;
  const recycleAfter = num(g.recycle_after_days) ?? 28;
  const fillWithFit = (num(g.fill_with_fit) ?? 1) !== 0;
  const languageSpillover = (num(g.language_spillover) ?? 1) !== 0;
  const worked = ctx.worked || {};
  const tbl = sec => S[sec] || {};

  // ---------- clean accounts ----------
  const COUNTRY = { 'usa': 'United States', 'united states': 'United States', 'us': 'United States',
    'uae': 'United Arab Emirates', 'united arab emirates': 'United Arab Emirates', 'uk': 'United Kingdom', 'united kingdom': 'United Kingdom' };
  const REGION_OF = { 'United States': 'North America', 'Canada': 'North America',
    'United Arab Emirates': 'EMEA', 'United Kingdom': 'EMEA' };
  const titleCase = s => s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());

  const acc = {};
  for (const a of accounts) {
    const id = str(a.account_id);
    if (!id) continue;
    const raw = str(a.country);
    const country = COUNTRY[raw.toLowerCase()] || titleCase(raw);
    if (country !== raw) issue('accounts', id, `Country "${raw}"`, `Normalised to "${country}"`);
    let region = str(a.region);
    if (region === 'NA') region = 'North America';
    if (REGION_OF[country] && REGION_OF[country] !== region) {
      issue('accounts', id, `Region "${str(a.region) || 'blank'}" does not match ${country}`, `Set to ${REGION_OF[country]}`);
      region = REGION_OF[country];
    }
    if (!region) issue('accounts', id, 'Region missing', 'Left unrouted');
    acc[id] = { ...a, account_id: id, country, region, signals: [] };
  }

  // ---------- renewal windows ----------
  // "renews in ~4 months" = signal date + 4 months; "Q1 2027" = the middle of that quarter (15 Feb 2027)
  const renewalDate = (detail, logged) => {
    let m = detail.match(/renews? in ~?\s*(\d+)\s*months?/i);
    if (m) { const d = new Date(logged); d.setUTCMonth(d.getUTCMonth() + Number(m[1])); return d; }
    m = detail.match(/\bQ([1-4])\s*(20\d\d)\b/i);
    if (m && /renew/i.test(detail)) return new Date(Date.UTC(Number(m[2]), (Number(m[1]) - 1) * 3 + 1, 15));
    return null;
  };
  // bands: { months up to: weight }, e.g. 3 → 10, 6 → 8, 12 → 5; beyond the last band, the last band's weight
  const bands = Object.entries(tbl('renewal_band')).map(([k, v]) => [Number(k), Number(v)]).filter(([k, v]) => !isNaN(k) && !isNaN(v)).sort((a, b) => a[0] - b[0]);
  const renewalWeight = months => {
    for (const [upTo, w] of bands) if (months <= upTo) return w;
    return bands.length ? bands[bands.length - 1][1] : 8;
  };

  // ---------- clean signals ----------
  const seen = new Set();
  let used = 0;
  for (const s of signals) {
    const id = str(s.account_id), date = str(s.signal_date);
    const k = [id, date, str(s.signal_type), str(s.signal_detail)].join('|');
    if (seen.has(k)) { issue('signals', id, `Duplicate: ${s.signal_detail} on ${date}`, 'Counted once'); continue; }
    seen.add(k);
    if (!acc[id]) { issue('signals', id, 'Account not in accounts list', 'Signal dropped'); continue; }
    if (!date) { issue('signals', id, `No date: ${s.signal_detail}`, 'Signal dropped (age unknown)'); continue; }
    const age = daysBetween(asOf, toDate(date));
    if (age < 0) { issue('signals', id, `Dated after as_of_date (${date})`, 'Signal dropped'); continue; }
    // renewal signals are scored on time to the renewal, not on signal age: no fading and no 90-day cut
    const renewal = renewalDate(str(s.signal_detail), toDate(date));
    if (renewal) {
      const months = (renewal - asOf) / (86400000 * 30.44);
      if (months < 0) { issue('signals', id, `Renewal date passed (${renewal.toISOString().slice(0, 10)}): ${s.signal_detail}`, 'Signal dropped, flag to the account owner'); continue; }
      const w = renewalWeight(months);
      acc[id].signals.push({ type: str(s.signal_type), detail: str(s.signal_detail), source: str(s.source), date, age,
        weight: w, weightFrom: 'renewal date', renewal: renewal.toISOString().slice(0, 10), monthsToRenewal: Math.round(months * 10) / 10,
        decay: 1, points: w });
      used++;
      continue;
    }
    if (age > lookback) { issue('signals', id, `Older than ${lookback} days (${date})`, 'Signal dropped'); continue; }
    const byDetail = num(tbl('signal_weight')[str(s.signal_detail)]);
    const byType = num(tbl('signal_type_default')[str(s.signal_type)]);
    const w = byDetail ?? byType ?? 3;
    const decay = Math.pow(0.5, age / halfLife);
    acc[id].signals.push({ type: str(s.signal_type), detail: str(s.signal_detail), source: str(s.source), date, age,
      weight: w, weightFrom: byDetail !== null ? 'signal detail' : byType !== null ? 'signal type default' : 'fallback',
      decay, points: w * decay });
    used++;
  }

  // ---------- score ----------
  const deviceSteps = () => Object.entries(tbl('fit_devices')).map(([k, v]) => [Number(k), Number(v)]).sort((x, y) => y[0] - x[0]);
  const devicesMult = d => {
    if (d === null) return 1;
    for (const [min, m] of deviceSteps()) if (d >= min) return m;
    return 1;
  };
  const lookup = (sec, key) => num(tbl(sec)[key]) ?? 1;
  const boosts = Object.entries(tbl('boost')).map(([k, v]) => { const [f, ...rest] = k.split('='); return [f.trim(), rest.join('=').trim(), Number(v)]; });
  const excluded = tbl('exclude'), pinned = tbl('pin');

  const scored = [];
  for (const a of Object.values(acc)) {
    let status = '';
    const stageMult = num(tbl('crm_stage')[str(a.crm_stage)]) ?? 1;
    const touchAge = str(a.last_touch_date) ? daysBetween(asOf, toDate(a.last_touch_date)) : null;
    const isPinned = a.account_id in pinned;

    // intent: strongest signal per type counts fully, repeats of a type count partially
    const byType = {};
    for (const s of a.signals) (byType[s.type] = byType[s.type] || []).push(s);
    let intent = 0;
    for (const list of Object.values(byType)) {
      list.sort((x, y) => y.points - x.points);
      list.forEach((s, i) => { s.share = i === 0 ? 1 : sameTypeShare; s.counted = s.points * s.share; intent += s.counted; });
    }
    const intentBase = intent;
    const types = Object.keys(byType).length;
    const breadth = 1 + multiTypeBonus * Math.max(0, types - 1);
    intent *= breadth;

    const devM = devicesMult(num(a.est_connected_devices));
    const provM = lookup('fit_provider', str(a.current_provider));
    const sizeM = lookup('fit_size', str(a.employee_count));
    const fit = devM * provM * sizeM;
    let boost = 1;
    const boostNotes = [];
    for (const [f, v, m] of boosts) if (str(a[f]) === v && m !== 1) { boost *= m; boostNotes.push(`${f} ${v} x${m}`); }

    const score = intent * fit * stageMult * boost;

    // week to week: an account worked in an earlier week stays with its rep while the sequence runs,
    // unless a new signal arrives (back on fire) or the sequence ended without reply (recycled)
    const w = worked[a.account_id];
    const freshSince = w ? a.signals.filter(x => toDate(x.date) > toDate(w.date)) : [];
    const recycled = !!w && !freshSince.length && daysBetween(asOf, toDate(w.date)) >= recycleAfter;

    if (a.account_id in excluded) status = `Excluded by leadership: ${str(excluded[a.account_id])}`;
    else if (stageMult === 0) status = `Excluded: CRM stage ${a.crm_stage}`;
    else if (boost === 0) status = `Excluded by segment boost (${boostNotes.join(', ')})`;
    else if (w && freshSince.length) status = `Back on fire: new signal since ${w.sdr} started`;
    else if (w && !recycled) status = `In sequence with ${w.sdr} since ${w.date}`;
    else if (touchAge !== null && touchAge <= touchCooldown && !isPinned) status = `Cool-down: touched ${touchAge} days ago`;
    else if (!a.signals.length && !isPinned) status = `No signal in last ${lookback} days`;

    const top = [...a.signals].sort((x, y) => y.points - x.points);
    const ago = d => (d === 0 ? 'today' : d === 1 ? '1 day ago' : `${d} days ago`);
    const reasons = [];
    for (const s of top) {
      if (reasons.length >= maxReasons) break;
      if (!reasons.some(r => r.startsWith(s.detail + ' ('))) reasons.push(`${s.detail} (${ago(s.age)})`);
    }
    const fitBits = [];
    if (num(a.est_connected_devices) !== null) fitBits.push(`~${Number(a.est_connected_devices).toLocaleString('en-US')} devices`);
    if (str(a.current_provider)) fitBits.push(`on ${a.current_provider}`);
    if (str(a.employee_count)) fitBits.push(`${a.employee_count} staff`);
    if (str(a.crm_stage) === 'Closed lost') fitBits.push(`closed lost before (last touch ${a.last_touch_date})`);
    if (boostNotes.length) fitBits.push(`boost: ${boostNotes.join(', ')}`);

    scored.push({
      account_id: a.account_id, company_name: str(a.company_name), country: a.country, region: a.region,
      industry: str(a.industry), crm_stage: str(a.crm_stage), ae_owner: str(a.owner),
      score: Math.round(score * 10) / 10, intent: Math.round(intent * 10) / 10, fit: Math.round(fit * 100) / 100,
      signals_used: a.signals.length, signal_types: types,
      why: reasons.join('; '), fit_notes: fitBits.join(', '),
      suggested_angle: top.length ? str(tbl('angle')[top[0].type]) : '',
      pinned: isPinned, status, assigned_to: '', rank: '', prospect: false, spillover: null,
      worked: w || null, refire: freshSince.map(x => x.detail), recycled,
      breakdown: {
        signals: top, intentBase, types, breadth, intent,
        devices: num(a.est_connected_devices), devM, provider: str(a.current_provider), provM, size: str(a.employee_count), sizeM, fit,
        stageMult, boost, boostNotes, score, touchAge,
        employee_count: str(a.employee_count), funding_stage: str(a.funding_stage), last_touch_date: str(a.last_touch_date)
      }
    });
  }

  // ---------- route ----------
  const reps = {};
  for (const r of sdrs) {
    const id = str(r.sdr_id);
    const cap = num(tbl('capacity')[id]) ?? num(r.weekly_capacity) ?? 0;
    reps[id] = { id, name: str(r.name), region: str(r.region), language: str(r.language), langs: str(r.language).split(',').map(x => x.trim().toLowerCase()).filter(Boolean),
      tenure: num(r.tenure_months) ?? 0, cap, left: cap, list: [] };
  }
  const poolOf = {};
  for (const [region, ids] of Object.entries(tbl('territory'))) poolOf[region] = str(ids).split(',').map(x => x.trim()).filter(x => reps[x]);

  const ranked = scored.filter(s => !s.status)
    .sort((x, y) => (y.pinned - x.pinned) || (y.score - x.score));
  for (const s of ranked) {
    const pool = (poolOf[s.region] || []).map(id => reps[id]);
    if (!pool.length) { s.status = `Not routed: no SDR covers ${s.region || 'unknown region'}`; continue; }
    // the rep with the most free capacity (as a share) takes the next best account; ties go to the more senior rep
    const open = pool.filter(r => r.left > 0)
      .sort((x, y) => (y.left / y.cap - x.left / x.cap) || (y.tenure - x.tenure));
    if (!open.length) { s.status = 'Not routed: territory at capacity this week'; continue; }
    const rep = open[0];
    rep.left--; rep.list.push(s);
    s.assigned_to = rep.name; s.rank = rep.list.length; s.status = 'Assigned';
  }

  // language spillover: a rep with free slots takes qualified signal accounts from another territory's
  // queue when they speak the country's language (English doesn't count: everyone speaks it).
  const LANG = { 'Spain': 'spanish', 'Mexico': 'spanish', 'Colombia': 'spanish', 'Chile': 'spanish', 'Argentina': 'spanish', 'Peru': 'spanish',
    'Brazil': 'portuguese', 'Portugal': 'portuguese', 'Germany': 'german', 'Austria': 'german', 'Switzerland': 'german', 'France': 'french', 'Italy': 'italian' };
  if (languageSpillover) {
    const waiting = scored.filter(s => s.status.startsWith('Not routed')).sort((x, y) => (y.pinned - x.pinned) || (y.score - x.score));
    for (const s of waiting) {
      const lang = LANG[s.country];
      if (!lang) continue;
      const open = Object.values(reps).filter(r => r.left > 0 && r.langs.includes(lang))
        .sort((x, y) => (y.left / y.cap - x.left / x.cap) || (y.tenure - x.tenure));
      if (!open.length) continue;
      const rep = open[0];
      rep.left--; rep.list.push(s);
      s.assigned_to = rep.name; s.rank = rep.list.length; s.status = 'Assigned';
      s.spillover = { language: lang[0].toUpperCase() + lang.slice(1), from: s.region || 'unknown region' };
    }
  }

  // keep every rep at full capacity: when signal accounts run out, the spare slots go to the
  // best-fit accounts with no signal yet (prospecting). Leadership can switch this off.
  if (fillWithFit) {
    const prospects = scored.filter(s => s.status.startsWith('No signal') && s.breakdown.stageMult > 0)
      .sort((x, y) => (y.breakdown.fit * y.breakdown.stageMult) - (x.breakdown.fit * x.breakdown.stageMult) || ((y.breakdown.devices || 0) - (x.breakdown.devices || 0)));
    for (const s of prospects) {
      const open = (poolOf[s.region] || []).map(id => reps[id]).filter(r => r.left > 0)
        .sort((x, y) => (y.left / y.cap - x.left / x.cap) || (y.tenure - x.tenure));
      if (!open.length) continue;
      const rep = open[0];
      rep.left--; rep.list.push(s);
      s.assigned_to = rep.name; s.rank = rep.list.length; s.status = 'Assigned'; s.prospect = true;
      s.suggested_angle = str(tbl('angle').prospect) || 'No signal yet: research the account, then a cold first touch built on device volume and current provider.';
    }
  }

  // how many weeks the qualified backlog covers at current capacity, per territory
  const coverage = Object.entries(poolOf).filter(([, ids]) => ids.length).map(([region, ids]) => {
    const cap = ids.reduce((n, id) => n + reps[id].cap, 0);
    const backlog = scored.filter(s => s.region === region && s.status.includes('at capacity')).length;
    return { region, cap, backlog, weeks: cap ? Math.floor(backlog / cap) : 0 };
  });

  const weekOf = asOf.toISOString().slice(0, 10);
  const lists = [];
  for (const rep of Object.values(reps)) {
    if (rep.left > 0) issue('routing', rep.name, `Only ${rep.list.length} of ${rep.cap} slots filled`, 'Not enough qualifying accounts in territory');
    for (const s of rep.list) lists.push({
      week_of: weekOf, sdr: rep.name, rank: s.rank, account_id: s.account_id, company_name: s.company_name,
      country: s.country, industry: s.industry, crm_stage: s.crm_stage, score: s.score,
      why_now: s.why, fit: s.fit_notes, suggested_angle: s.suggested_angle,
      loop_in_ae: s.ae_owner, pinned: s.pinned ? 'yes' : '', prospect: s.prospect ? 'yes' : '', recycled: s.recycled ? 'yes' : '',
      spillover: s.spillover ? `${s.spillover.language} speaker, from ${s.spillover.from}` : ''
    });
  }
  scored.sort((x, y) => y.score - x.score);
  const summary = {
    week_of: weekOf, accounts: Object.keys(acc).length, signals_in: signals.length, signals_used: used,
    assigned: lists.length, prospects: lists.filter(l => l.prospect).length, coverage, per_sdr: Object.values(reps).map(r => `${r.name}: ${r.list.length}/${r.cap}`).join(', ')
  };
  const team = Object.values(reps).map(r => ({ id: r.id, name: r.name, region: r.region, language: r.language, tenure: r.tenure, cap: r.cap, filled: r.list.length,
    territories: Object.entries(poolOf).filter(([, ids]) => ids.includes(r.id)).map(([reg]) => reg) }));
  return { lists, scored, issues, summary, team };
}

// Files are recognised by their columns, so names and upload order do not matter.
export function classifyRows(rows) {
  if (!rows.length) return null;
  const cols = new Set(Object.keys(rows[0]));
  if (cols.has('company_name')) return 'accounts';
  if (cols.has('signal_type')) return 'signals';
  if (cols.has('sdr_id')) return 'sdrs';
  if (cols.has('section') && cols.has('key')) return 'settings';
  return null;
}
