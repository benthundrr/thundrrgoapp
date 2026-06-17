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
  if (!fs.existsSync(DATA_DIR)) return [];
  return fs.readdirSync(DATA_DIR).filter(f => re.test(f)).map(f => path.join(DATA_DIR, f));
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

/* ---- ContentAnalysis → posts (dedup by Post ID across files, keep latest) ---- */
function buildPosts(profiles) {
  const files = ls(/^ContentAnalysis.*\.xlsx$/i).sort();
  const map = new Map();
  for (const file of files) {
    const rows = sheetRows(file, 'Data');
    const head = rows[0].map(norm);
    const idx = re => head.findIndex(c => re.test(c));
    const I = {
      pid: idx(/Post ID/i), title: idx(/Post title/i), date: idx(/Post date/i), dur: idx(/Duration/i),
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
        u, date: norm(r[I.date]),
        title: norm(r[I.title]).slice(0, 120),
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

function main() {
  const profiles = {};
  const roster = buildRoster();
  const snapshots = buildSnapshots(profiles);
  const posts = buildPosts(profiles);

  const latest = snapshots[snapshots.length - 1] || { start: '', end: '' };
  const out = {
    generatedAt: new Date().toISOString(),
    dataRange: { start: latest.start, end: latest.end },
    counts: {
      profiles: Object.keys(profiles).length,
      roster: Object.keys(roster).length,
      snapshots: snapshots.length,
      posts: posts.length
    },
    profiles, roster, snapshots, posts
  };
  fs.writeFileSync(OUT, JSON.stringify(out));
  console.log('[build] data.json written',
    `· profiles ${out.counts.profiles} · roster ${out.counts.roster} · snapshots ${out.counts.snapshots} · posts ${out.counts.posts} · ${(fs.statSync(OUT).size / 1024).toFixed(0)}kb`);
}
main();
