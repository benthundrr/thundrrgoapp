/* THUNDRR GO HUB — build step
 * Scans ./data for TikTok GO exports and compiles a single data.json the app fetches.
 *
 * Drop these into /data and commit (Netlify rebuilds automatically):
 *   CreatorAnalysis_ManagedCreators_<start>_<end>.xlsx   (authoritative GMV / orders, cumulative for the window)
 *   ContentAnalysis_VideoList_<start>_<end>.xlsx          (per-post history → daily + month-over-month trends)
 *   Creator_Roster_Main_<id>.xlsx                         (contacts + who you manage; biweekly/as-needed)
 *
 * Keep dated files if you want history to accumulate (one CreatorAnalysis per month unlocks GMV month-over-month).
 * Filenames encode the date range — that's how snapshots get keyed.
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const DATA_DIR = path.join(__dirname, 'data');
const OUT = path.join(__dirname, 'data.json');

const norm = s => String(s == null ? '' : s).trim();
const lc = s => norm(s).toLowerCase();
const mLabel = m => m ? `${m.slice(0, 4)}-${m.slice(4, 6)}` : '—';
const num = v => { const n = parseFloat(v); return isFinite(n) ? n : 0; };
const r2 = n => Math.round(n * 100) / 100;
const r4 = n => Math.round(n * 10000) / 10000;

function rangeFromName(f) {
  const m = f.match(/(\d{8})_(\d{8})/);
  return m ? { start: m[1], end: m[2] } : { start: '', end: '' };
}
function sheetRows(file, sheetName) {
  const wb = XLSX.readFile(file);
  const sn = wb.SheetNames.includes(sheetName) ? sheetName : wb.SheetNames[wb.SheetNames.length - 1];
  return XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', raw: true });
}
function ls(re) {
  // Look in /data first, then the repo root. Files can live in either place.
  const dirs = [DATA_DIR, __dirname];
  const seen = new Set(), out = [];
  for (const d of dirs) {
    try {
      if (!fs.existsSync(d) || !fs.statSync(d).isDirectory()) continue;
      for (const f of fs.readdirSync(d)) {
        if (re.test(f) && !seen.has(f)) { seen.add(f); out.push(path.join(d, f)); }
      }
    } catch (e) { /* skip unreadable / non-dir */ }
  }
  return out;
}

/* ---- Creator roster (contacts) — take the most recently modified ---- */
function buildRoster() {
  const files = ls(/^Creator_Roster.*\.xlsx$/i)
    .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
  const roster = {};
  if (!files.length) return roster;
  const rows = sheetRows(files[0], 'creator roster main');
  // header row is the one containing "username"
  let h = rows.findIndex(r => r.some(c => /username/i.test(norm(c))));
  if (h < 0) h = 0;
  const head = rows[h].map(norm);
  const col = re => head.findIndex(c => re.test(c));
  const cUser = col(/username/i), cName = col(/^name$/i), cEmail = col(/email/i),
    cMgr = col(/manager/i), cDisc = col(/discord/i), cPhone = col(/phone/i),
    cStatus = col(/status/i), cCity = col(/^city$/i), cState = col(/state/i),
    cWeb = col(/webinar/i), cDate = col(/^date$/i);
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i]; const u = lc(r[cUser]);
    if (!u) continue;
    roster[u] = {
      username: norm(r[cUser]),
      name: cName >= 0 ? norm(r[cName]) : '',
      email: cEmail >= 0 ? norm(r[cEmail]) : '',
      manager: cMgr >= 0 ? norm(r[cMgr]) : '',
      discord: cDisc >= 0 ? norm(r[cDisc]) : '',
      phone: cPhone >= 0 ? norm(r[cPhone]) : '',
      status: cStatus >= 0 ? norm(r[cStatus]) : '',
      city: cCity >= 0 ? norm(r[cCity]) : '',
      state: cState >= 0 ? norm(r[cState]) : '',
      webinar: cWeb >= 0 ? norm(r[cWeb]) : '',
      date: cDate >= 0 ? norm(r[cDate]) : ''
    };
  }
  return roster;
}

/* ---- CreatorAnalysis → one cumulative snapshot per file (keyed by date range) ---- */
function buildSnapshots(profiles) {
  const files = ls(/^CreatorAnalysis.*\.xlsx$/i).sort();
  const snaps = [];
  for (const file of files) {
    const { start, end } = rangeFromName(path.basename(file));
    const rows = sheetRows(file, 'Data');
    const head = rows[0].map(norm);
    const idx = re => head.findIndex(c => re.test(c));
    const I = {
      id: idx(/Creator ID/i), name: idx(/Creator name/i), level: idx(/Creator level/i),
      city: idx(/Creator city/i), bind: idx(/binding/i), sales: idx(/Sales value/i),
      orders: idx(/^Orders/i), redAmt: idx(/Redemption amount/i), redOrd: idx(/Redeemed orders/i),
      pViews: idx(/Posts with views/i), pSales: idx(/Posts with sales/i), views: idx(/Video views/i),
      ctr: idx(/CTR/i), cvr: idx(/CVR/i), aov: idx(/AOV/i),
      avgV: idx(/Avg.*views/i), avgS: idx(/Avg.*sales/i)
    };
    const byUser = {};
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i]; const u = lc(r[I.id]); if (!u) continue;
      byUser[u] = {
        sales: r2(num(r[I.sales])), orders: num(r[I.orders]),
        redemption: r2(num(r[I.redAmt])), redeemedOrders: num(r[I.redOrd]),
        postsViews: num(r[I.pViews]), postsSales: num(r[I.pSales]), views: num(r[I.views]),
        ctr: r4(num(r[I.ctr])), cvr: r4(num(r[I.cvr])), aov: r2(num(r[I.aov])),
        avgViews: r2(num(r[I.avgV])), avgSales: r2(num(r[I.avgS]))
      };
      profiles[u] = profiles[u] || {};
      profiles[u].name = norm(r[I.name]) || profiles[u].name || norm(r[I.id]);
      profiles[u].level = norm(r[I.level]) || profiles[u].level || '';
      profiles[u].city = norm(r[I.city]) || profiles[u].city || '';
      profiles[u].binding = norm(r[I.bind]) || profiles[u].binding || '';
    }
    snaps.push({ start, end, monthKey: (end || start).slice(0, 6), byUser });
  }
  return snaps.sort((a, b) => (a.end || a.start).localeCompare(b.end || b.start));
}

/* ContentAnalysis → posts (dedup by Post ID across files, keep latest) */
const VALID = p => p.views >= 300; // public/reached proxy (TikTok's true valid count adds POI-match + relevance, not in export)
function indCode(s) {
  s = (s || '').toLowerCase();
  if (s.includes('acc') && s.includes('ttd')) return 'BOTH';
  if (s.includes('accommod')) return 'ACC';
  if (s.includes('things')) return 'TTD';
  return '';
}
function buildPosts(profiles) {
  const files = ls(/^ContentAnalysis.*\.xlsx$/i).sort();
  const map = new Map();
  for (const file of files) {
    const rows = sheetRows(file, 'Data');
    const head = rows[0].map(norm);
    const idx = re => head.findIndex(c => re.test(c));
    const I = {
      pid: idx(/Post ID/i), title: idx(/Post title/i), date: idx(/Post date/i), dur: idx(/Duration/i),
      ind: idx(/Location industry/i),
      loc: idx(/Location name/i), city: idx(/Location city/i), merch: idx(/Merchant name/i),
      cname: idx(/Creator name/i), cid: idx(/Creator ID/i), level: idx(/Creator level/i),
      sales: idx(/Sales value/i), orders: idx(/^Orders/i), views: idx(/Video views/i),
      ctr: idx(/CTR/i), cvr: idx(/CVR/i), comp: idx(/completion/i), like: idx(/Like rate/i),
      cmt: idx(/Comment rate/i), ptype: idx(/Post type/i)
    };
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i]; const u = lc(r[I.cid]); const pid = norm(r[I.pid]);
      if (!u || !pid) continue;
      profiles[u] = profiles[u] || {};
      if (!profiles[u].name) profiles[u].name = norm(r[I.cname]) || norm(r[I.cid]);
      if (!profiles[u].level && I.level >= 0) profiles[u].level = norm(r[I.level]);
      map.set(pid, {
        u, id: pid, date: norm(r[I.date]),
        title: norm(r[I.title]).slice(0, 120),
        ind: indCode(r[I.ind]),
        views: num(r[I.views]), sales: r2(num(r[I.sales])), orders: num(r[I.orders]),
        dur: num(r[I.dur]),
        comp: r4(num(r[I.comp])), like: r4(num(r[I.like])), cmt: r4(num(r[I.cmt])),
        ctr: r4(num(r[I.ctr])), cvr: r4(num(r[I.cvr])),
        loc: norm(r[I.loc]), city: norm(r[I.city]), merch: norm(r[I.merch])
      });
    }
  }
  return [...map.values()].sort((a, b) => a.date.localeCompare(b.date));
}

/* ACC / TTD: valid posts = current month; Paid GMV = cumulative (accrues) */
function buildIndustry(posts, asOf, mk) {
  const acc = { valid: 0, gmv: 0 }, ttd = { valid: 0, gmv: 0 };
  let totalValid = 0;
  for (const p of posts) {
    const isAcc = p.ind === 'ACC' || p.ind === 'BOTH', isTtd = p.ind === 'TTD' || p.ind === 'BOTH';
    if (isAcc) acc.gmv += p.sales;
    if (isTtd) ttd.gmv += p.sales;
    if (VALID(p) && p.date.slice(0, 6) === mk) {
      if (isAcc) acc.valid++;
      if (isTtd) ttd.valid++;
      totalValid++;                 // deduped (counts an ACC,TTD post once)
    }
  }
  acc.gmv = r2(acc.gmv); ttd.gmv = r2(ttd.gmv);
  return { asOf, monthKey: mk, acc, ttd, totalValid };
}

/* Daily series: cumulative Paid GMV (all posts) + cumulative month valid posts */
function buildDaily(posts, mk) {
  const by = {};
  for (const p of posts) {
    const d = by[p.date] = by[p.date] || { ag: 0, tg: 0, av: 0, tv: 0 };
    const isAcc = p.ind === 'ACC' || p.ind === 'BOTH', isTtd = p.ind === 'TTD' || p.ind === 'BOTH';
    if (isAcc) d.ag += p.sales;
    if (isTtd) d.tg += p.sales;
    if (VALID(p) && p.date.slice(0, 6) === mk) { if (isAcc) d.av++; if (isTtd) d.tv++; }
  }
  const dates = Object.keys(by).sort();
  let ag = 0, tg = 0, av = 0, tv = 0; const out = [];
  for (const dt of dates) { const x = by[dt]; ag += x.ag; tg += x.tg; av += x.av; tv += x.tv; out.push({ d: dt, ag: r2(ag), tg: r2(tg), av, tv }); }
  return out;
}

/* Agency benchmarks for the coaching engine */
function buildBenchmarks(posts, snap) {
  const seen = posts.filter(p => p.views > 0);
  const durs = seen.map(p => p.dur).filter(d => d > 0).sort((a, b) => a - b);
  const median = a => a.length ? a[Math.floor(a.length / 2)] : 0;
  const validCount = posts.filter(VALID).length;
  const byUser = snap.byUser || {};
  const sellers = Object.values(byUser).filter(s => s.sales > 0).sort((a, b) => b.sales - a.sales);
  const top = sellers.slice(0, 20);
  const avg = (arr, k) => arr.length ? r2(arr.reduce((s, x) => s + (x[k] || 0), 0) / arr.length) : 0;
  return {
    medianDuration: median(durs),
    sweetLow: 10, sweetHigh: 30, sweetMedian: 21,        // from L2 deck
    volumeFloor: 20, volumeTarget: 100, earnerMedianPosts: 14, // from L2 deck
    validRate: seen.length ? r4(validCount / seen.length) : 0,
    topEarnerAvgViews: Math.round(avg(top, 'avgViews')),
    topEarnerAvgPosts: Math.round(avg(top, 'postsViews')),
    topEarnerAvgAOV: avg(top, 'aov'),
    sellersCount: sellers.length,
    totalCreators: Object.keys(byUser).length
  };
}

function main() {
  const profiles = {};
  const roster = buildRoster();
  const snapshots = buildSnapshots(profiles);
  const posts = buildPosts(profiles);

  const latest = snapshots[snapshots.length - 1] || { start: '', end: '', byUser: {} };
  const mk = (latest.end || (posts.length ? posts[posts.length - 1].date : '')).slice(0, 6);
  const industry = buildIndustry(posts, latest.end, mk);
  const daily = buildDaily(posts, mk);
  const benchmarks = buildBenchmarks(posts, latest);
  const readJSON = f => { try { return JSON.parse(fs.readFileSync(path.join(__dirname, f), 'utf8')); } catch (e) { return null; } };
  const usmap = readJSON('usmap.json');
  const citygeo = readJSON('citygeo.json');
  const out = {
    generatedAt: new Date().toISOString(),
    dataRange: { start: latest.start, end: latest.end },
    counts: {
      profiles: Object.keys(profiles).length,
      roster: Object.keys(roster).length,
      snapshots: snapshots.length,
      posts: posts.length
    },
    industry, daily, benchmarks, usmap, citygeo,
    profiles, roster, snapshots, posts
  };
  fs.writeFileSync(OUT, JSON.stringify(out));
  console.log('[build] data.json',
    `· profiles ${out.counts.profiles} · posts ${out.counts.posts}`,
    `· ${mLabel(mk)} valid: ACC ${industry.acc.valid} / TTD ${industry.ttd.valid} (${industry.totalValid} total)`,
    `· GMV: ACC $${industry.acc.gmv} / TTD $${industry.ttd.gmv}`,
    `· ${(fs.statSync(OUT).size / 1024).toFixed(0)}kb`);
}
main();
