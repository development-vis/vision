"use client";

import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ══════════════════════════════════════════
//  HELPER FUNCTIONS (Date & Scoring)
// ══════════════════════════════════════════
function excelSerialToDate(serial: number) {
  const EXCEL_EPOCH_OFFSET = 25569;
  const daysSinceUnix = serial - EXCEL_EPOCH_OFFSET;
  const utc = new Date(daysSinceUnix * 86400000);
  return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
}

function toDate(v: any) {
  if (v == null || v === '') return null;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate());
  }
  if (typeof v === 'number' && v > 1) {
    return excelSerialToDate(v);
  }
  if (typeof v === 'string') {
    const p = Date.parse(v);
    if (!isNaN(p)) {
      const utc = new Date(p);
      return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
    }
  }
  return null;
}

function fmt(d: Date | null) { 
  return d ? d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) : '—'; 
}

function fmtFull(d: Date | null) { 
  return d ? d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : '—'; 
}

function sc(s: string) {
  const sl = (s || '').toLowerCase();
  if (sl === 'completed' || sl === 'complete') return 'complete';
  if (sl === 'partial' || sl === 'in progress') return 'partial';
  return 'pending';
}

function score(x: any) { 
  const c = sc(x.status); 
  return c === 'complete' ? 1 : c === 'partial' ? 0.5 : 0; 
}

function statsFor(items: any[]) {
  const total = items.length;
  const done = items.filter(x => sc(x.status) === 'complete').length;
  const partial = items.filter(x => sc(x.status) === 'partial').length;
  const weighted = items.reduce((a, x) => a + score(x), 0);
  const pct = total ? Math.round(weighted / total * 100) : 0;
  const donePct = total ? Math.round(done / total * 100) : 0;
  const partialPct = total ? Math.round(partial * 0.5 / total * 100) : 0;
  return { total, done, partial, pct, donePct, partialPct };
}

// ══════════════════════════════════════════
//  MAIN COMPONENT
// ══════════════════════════════════════════
export default function MCHDClientPortal() {
  // Authentication
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  // Data State
  const [towers, setTowers] = useState<any[]>([]);
  const [groups, setGroups] = useState<any[]>([]);
  const [dates, setDates] = useState<any>({ towerStart: null, towerEnd: null, emsStart: null, emsEnd: null });
  
  // UI State
  const [loading, setLoading] = useState(false);
  const [hasData, setHasData] = useState(false);
  const [fileName, setFileName] = useState("MCHD_Master_Workbook.xlsx");
  const [syncTime, setSyncTime] = useState("");
  const [openGroups, setOpenGroups] = useState<Record<number, boolean>>({ 0: true, 1: true, 2: true, 3: true });
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Auto-fetch default file
  useEffect(() => {
    if (isAuthenticated) {
      setLoading(true);
      fetch('/MCHD_Master_Workbook.xlsx')
        .then(async (res) => {
          // Grab the actual Last-Modified date from the server file metadata
          const lastModHeader = res.headers.get('Last-Modified');
          const fileDate = lastModHeader ? new Date(lastModHeader) : new Date();
          
          const buf = await res.arrayBuffer();
          return { buf, fileDate };
        })
        .then(({ buf, fileDate }) => parseData(buf, "MCHD_Master_Workbook.xlsx", fileDate))
        .catch(err => {
          console.error("No default file found. Please upload one.", err);
          setLoading(false);
        });
    }
  }, [isAuthenticated]);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (password === "MCHDpass") {
      setIsAuthenticated(true);
      setError("");
    } else {
      setError("Incorrect password");
    }
  };

  const parseData = (buf: ArrayBuffer, fname: string, fileDate: Date) => {
    try {
      const wb = XLSX.read(buf, { type: 'array' });
      
      // Parse Summary Dates
      const summaryWs = wb.Sheets['Summary Overview'];
      let tStart = null, tEnd = null, eStart = null, eEnd = null;
      if (summaryWs && summaryWs['!ref']) {
        const range = XLSX.utils.decode_range(summaryWs['!ref']);
        for (let r = range.s.r; r <= range.e.r; r++) {
          const aCell = summaryWs[XLSX.utils.encode_cell({ r, c: 0 })];
          const bCell = summaryWs[XLSX.utils.encode_cell({ r, c: 1 })];
          if (!aCell) continue;
          const lbl = String(aCell.v !== undefined ? aCell.v : aCell.w).toLowerCase().trim();
          if (!bCell) continue;
          const rawVal = (bCell.v !== undefined && bCell.v !== null) ? bCell.v : bCell.w;
          if (rawVal == null) continue;
          const date = toDate(rawVal);
          if (!date) continue;

          const hasTower = lbl.includes('tower');
          const hasEMS = lbl.includes('ems') || lbl.includes('station');
          const hasStart = lbl.includes('start');
          const hasEnd = lbl.includes('complet') || lbl.includes('estimat') || lbl.includes('finish') || lbl.includes(' end');

          if (hasTower && hasStart && !hasEnd) tStart = date;
          else if (hasTower && hasEnd) tEnd = date;
          else if (hasEMS && hasStart && !hasEnd) eStart = date;
          else if (hasEMS && hasEnd) eEnd = date;
        }
      }
      setDates({ towerStart: tStart, towerEnd: tEnd, emsStart: eStart, emsEnd: eEnd });

      // Parse Towers
      const tws = wb.Sheets['Tower Sites'];
      const parsedTowers: any[] = [];
      if (tws) {
        XLSX.utils.sheet_to_json(tws, { defval: '', range: 2 }).forEach((r: any) => {
          const name = String(r['Tower Name'] || '').trim();
          const addr = String(r['Address'] || '').trim();
          if (name && addr && !/^Total/i.test(name)) {
            parsedTowers.push({
              name, addr,
              city: String(r['City'] || '').trim(),
              zip: String(r['ZIP'] || '').trim(),
              status: String(r['Upgrade Status'] || '').trim(),
              dateDone: r['Date Completed'] || ''
            });
          }
        });
      }
      setTowers(parsedTowers);

      // Parse Groups
      const parsedGroups: any[] = [];
      [
        { sheet: 'Group 1', area: 'North Conroe / Willis', color: 'g1' },
        { sheet: 'Group 2', area: 'Woodlands / Spring / Shenandoah', color: 'g2' },
        { sheet: 'Group 3', area: 'East / Porter / Splendora', color: 'g3' },
        { sheet: 'Group 4', area: 'West / Magnolia / Montgomery', color: 'g4' },
      ].forEach(gd => {
        const ws = wb.Sheets[gd.sheet];
        if (!ws) return;
        const stations: any[] = [];
        XLSX.utils.sheet_to_json(ws, { defval: '', range: 2 }).forEach((r: any) => {
          const id = String(r['Station'] || '').trim();
          const addr = String(r['Address'] || '').trim();
          if (id.startsWith('St.') && addr) {
            stations.push({
              id, addr,
              city: String(r['City'] || '').trim(),
              zip: String(r['ZIP'] || '').trim(),
              status: String(r['Device Upgrade Status'] || '').trim(),
              dateDone: r['Date Completed'] || '',
              tech: String(r['Tech Initials'] || '').trim()
            });
          }
        });
        if (stations.length) parsedGroups.push({ ...gd, stations });
      });
      setGroups(parsedGroups);

      setHasData(true);
      setFileName(fname);
      
      // Use the actual File Date passed in from the server or user's computer!
      setSyncTime(fileDate.toLocaleString('en-US', { 
        month: 'short', 
        day: 'numeric', 
        year: 'numeric', 
        hour: 'numeric', 
        minute: '2-digit' 
      }));
      setLoading(false);
      
    } catch (err) {
      console.error("Excel parse error:", err);
      setHasData(false);
      setLoading(false);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    setLoading(true);
    
    // Grab the exact "Date Modified" from the file the user is uploading
    const actualFileDate = new Date(file.lastModified);
    
    const reader = new FileReader();
    reader.onload = (ev) => {
      if (ev.target?.result) {
        parseData(ev.target.result as ArrayBuffer, file.name, actualFileDate);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const toggleGrp = (gi: number) => {
    setOpenGroups(prev => ({ ...prev, [gi]: !prev[gi] }));
  };

  // ══════════════════════════════════════════
  //  LOCK SCREEN VIEW
  // ══════════════════════════════════════════
  if (!isAuthenticated) {
    return (
      <div style={{ minHeight: '100vh', width: '100%', backgroundColor: '#000000', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '1rem', fontFamily: 'sans-serif' }}>
        <div style={{ backgroundColor: '#0f172a', border: '1px solid #1e293b', padding: '3rem', borderRadius: '1rem', width: '100%', maxWidth: '500px', textAlign: 'center', boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.5)' }}>
          <div style={{ color: '#4d8ef7', fontWeight: 'bold', letterSpacing: '0.2em', textTransform: 'uppercase', fontSize: '1rem', marginBottom: '1rem' }}>
            Secure Client Portal
          </div>
          <h1 style={{ fontSize: '2.5rem', fontWeight: 900, textTransform: 'uppercase', color: '#e6edf3', margin: '0 0 2.5rem 0' }}>
            MCHD Dashboard
          </h1>
          <form onSubmit={handleLogin} style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>
            <input
              type="password"
              placeholder="Enter Access Password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              style={{ padding: '1.25rem', backgroundColor: '#080d14', border: '1px solid #1e2d3d', borderRadius: '0.5rem', color: '#e6edf3', textAlign: 'center', outline: 'none', fontSize: '1.25rem' }}
            />
            {error && <p style={{ color: '#f87171', fontSize: '1.125rem', fontWeight: 'bold', margin: 0 }}>{error}</p>}
            <button type="submit" style={{ padding: '1.25rem', backgroundColor: '#4d8ef7', color: '#ffffff', fontWeight: 'bold', borderRadius: '0.5rem', border: 'none', cursor: 'pointer', textTransform: 'uppercase', letterSpacing: '0.1em', fontSize: '1.125rem' }}>
              Access Dashboard
            </button>
          </form>
        </div>
      </div>
    );
  }

  // ══════════════════════════════════════════
  //  DASHBOARD CALCULATIONS
  // ══════════════════════════════════════════
  const allStations = groups.flatMap(g => g.stations);
  const ov = statsFor([...towers, ...allStations]);
  const tS = statsFor(towers);
  const sS = statsFor(allStations);

  const estDate = (dates.towerEnd && dates.emsEnd) 
    ? (dates.towerEnd > dates.emsEnd ? dates.towerEnd : dates.emsEnd) 
    : (dates.towerEnd || dates.emsEnd || null);

  const tRange = [dates.towerStart && fmt(dates.towerStart), dates.towerEnd && fmt(dates.towerEnd)].filter(Boolean).join(' – ') || 'Phase 1';
  const sRange = [dates.emsStart && fmt(dates.emsStart), dates.emsEnd && fmt(dates.emsEnd)].filter(Boolean).join(' – ') || 'Phase 2';

  const shortName = (n: string) => n.replace(/\s*Tower\s*$/i, '').replace('Conroe Service Center', 'CSC').replace('East County / Splendora', 'E. County/Splendora').replace('Conroe Service', 'Conroe Svc').trim();

  const renderBadge = (s: string) => {
    const c = sc(s);
    const l = { complete: '● Complete', partial: '◐ Partial', pending: '○ Pending' };
    return (
      <span className={`badge ${c}`}>
        <span className="badge-dot"></span>{l[c]}
      </span>
    );
  };

  // ══════════════════════════════════════════
  //  DASHBOARD VIEW
  // ══════════════════════════════════════════
  return (
    <div className="fixed inset-0 z-[99999] overflow-y-auto mchd-dashboard-wrapper">
      <style dangerouslySetInnerHTML={{ __html: `
        .mchd-dashboard-wrapper { background: #080d14; color: #c9d1d9; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; font-size: 16px; min-height: 100vh; }
        .mchd-dashboard-wrapper * { box-sizing: border-box; margin: 0; padding: 0; }
        .page { max-width: 1400px; margin: 0 auto; padding: 40px 24px; }

        /* HEADER */
        .header { display: flex; align-items: flex-start; justify-content: space-between; margin-bottom: 36px; gap: 24px; flex-wrap: wrap; }
        .org-name { font-size: 14px; font-weight: 700; letter-spacing: 0.16em; text-transform: uppercase; color: #4d8ef7; margin-bottom: 8px; }
        .page-title { font-size: 46px; font-weight: 800; color: #e6edf3; letter-spacing: -0.02em; line-height: 1.1; margin-bottom: 12px; }
        .header-stats { font-size: 16px; color: #6e7d8e; line-height: 1.5; }
        .header-stats b { color: #8b949e; font-weight: 600; }
        
        .header-right { display: flex; flex-direction: column; align-items: flex-end; gap: 12px; flex-shrink: 0; }
        .sync-badge { display: flex; align-items: center; gap: 12px; background: #0d1117; border: 1px solid #1e2d3d; border-radius: 8px; padding: 12px 16px; min-width: 260px; }
        .sync-dot { width: 10px; height: 10px; border-radius: 50%; background: #3b4555; flex-shrink: 0; }
        .sync-dot.live { background: #22c55e; box-shadow: 0 0 10px rgba(34,197,94,0.5); animation: blink 2s infinite; }
        @keyframes blink { 0%,100%{opacity:1} 50%{opacity:0.5} }
        .sync-inner { display: flex; flex-direction: column; gap: 2px; }
        .sync-status { font-weight: 700; color: #e6edf3; font-size: 13px; }
        .sync-file { color: #6e7d8e; font-size: 12px; }
        .upload-btn { display: inline-flex; align-items: center; justify-content: center; gap: 8px; background: #161b22; border: 1px solid #30363d; color: #c9d1d9; padding: 10px 20px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s; white-space: nowrap; width: 100%; }
        .upload-btn:hover { border-color: #4d8ef7; color: #4d8ef7; background: #0d1117; }

        /* SCHEDULE STRIP */
        .sched-strip { display: flex; flex-wrap: wrap; background: #0d1117; border: 1px solid #1a2233; border-radius: 12px; margin-bottom: 32px; overflow: hidden; }
        .sched-cell { flex: 1; min-width: 250px; padding: 16px 24px; display: flex; flex-direction: column; gap: 6px; border-right: 1px solid #1a2233; }
        .sched-cell:last-child { border-right: none; }
        .sched-lbl { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.12em; color: #4d5565; }
        .sched-val { font-size: 16px; font-weight: 700; color: #4d8ef7; }
        .sched-val.green { color: #22c55e; }
        .sched-val.warn { color: #f59e0b; font-style: italic; font-weight: 500; font-size: 14px; }
        .sched-hint { font-size: 12px; color: #2a3a47; margin-top: 2px; }

        /* PROGRESS CARDS */
        .progress-card { background: #0d1117; border: 1px solid #1e2d3d; border-radius: 12px; padding: 28px 32px; margin-bottom: 36px; }
        .progress-card-title { font-size: 14px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: #8b949e; margin-bottom: 24px; }
        .prog-cols { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 32px; }
        .prog-col { padding: 0; }
        @media (min-width: 1024px) {
          .prog-col + .prog-col { border-left: 1px solid #1a2233; padding-left: 32px; }
        }
        .prog-head { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 12px; }
        .prog-lbl { font-size: 13px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: #8b949e; }
        .prog-lbl.towers { color: #a855f7; }
        .prog-lbl.ems { color: #3b82f6; }
        .prog-num { font-size: 16px; font-weight: 600; color: #e6edf3; }
        .prog-track { height: 6px; background: #1a2233; border-radius: 3px; margin-bottom: 10px; position: relative; }
        .prog-bar { height: 6px; border-radius: 3px; transition: width .6s ease; position: absolute; top: 0; left: 0; }
        .prog-bar.bg { background: linear-gradient(90deg,#d97706,#f59e0b); }
        .prog-bar.bg.towers { background: linear-gradient(90deg,#9333ea,#c084fc); }
        .prog-bar.bg.ems { background: linear-gradient(90deg,#2563eb,#60a5fa); }
        .prog-bar.fg { background: linear-gradient(90deg,#16a34a,#22c55e); }
        .prog-bar.fg.towers { background: linear-gradient(90deg,#7c3aed,#a855f7); }
        .prog-bar.fg.ems { background: linear-gradient(90deg,#1d4ed8,#3b82f6); }
        .prog-sub { font-size: 13px; color: #6e7d8e; }

        /* SECTION LABELS */
        .sec-label { font-size: 14px; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase; color: #6e7d8e; margin-bottom: 20px; display: flex; align-items: center; gap: 16px; }
        .sec-label::after { content:''; flex:1; height:2px; background:#1a2233; border-radius: 2px; }

        /* TOWERS */
        .tower-box { background: #0d1117; border: 1px solid #1e2d3d; border-radius: 12px; overflow: hidden; margin-bottom: 36px; }
        .tower-box-hdr { background: linear-gradient(135deg,#130d1f,#1a1030 60%,#160d24); border-bottom: 1px solid #2a1a40; padding: 20px 28px; display: flex; align-items: center; gap: 20px; flex-wrap: wrap; }
        .twr-icon { width: 44px; height: 44px; background: rgba(88,28,135,.3); border-radius: 10px; display: flex; align-items: center; justify-content: center; font-size: 22px; }
        .twr-hdr-title { font-size: 20px; font-weight: 700; color: #e6edf3; }
        .twr-hdr-sub { font-size: 14px; font-weight: 500; color: #a78bfa; margin-top: 4px; }
        .twr-badge { margin-left: auto; background: rgba(88,28,135,.4); border: 1px solid rgba(168,85,247,.3); color: #d8b4fe; font-size: 15px; font-weight: 700; padding: 8px 18px; border-radius: 8px; }
        .twr-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); }
        .twr-cell { padding: 20px 24px; border-right: 1px solid #1a2233; border-bottom: 1px solid #1a2233; display: flex; flex-direction: column; gap: 8px; transition: background .15s; }
        .twr-cell:hover { background: rgba(255,255,255,.02); }
        .twr-cell.complete { background: rgba(34,197,94,.04); }
        .twr-cell.partial { background: rgba(245,158,11,.03); }
        .twr-num { font-size: 12px; font-weight: 700; color: #4d8ef7; text-transform: uppercase; letter-spacing:.12em; }
        .twr-name { font-size: 16px; font-weight: 700; color: #e6edf3; line-height: 1.4; margin-bottom: 4px; }
        .twr-addr { font-size: 14px; color: #6e7d8e; line-height: 1.4; margin-bottom: 6px; }
        .twr-win { display: flex; align-items: center; gap: 6px; font-size: 12px; font-weight: 500; color: #4d8ef7; background: rgba(24,38,55,.7); padding: 4px 10px; border-radius: 6px; width: fit-content; }
        .twr-win-sq { width:6px; height:6px; background:#4d8ef7; border-radius:2px; flex-shrink:0; }
        .twr-done { font-size: 12px; font-weight: 600; color: #22c55e; margin-top: 4px; }

        /* BADGES */
        .badge { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 600; padding: 6px 12px; border-radius: 6px; width: fit-content; }
        .badge.pending { background: rgba(35,35,45,.9); color: #6b7280; border:1px solid rgba(75,85,99,.3); }
        .badge.partial { background: rgba(110,70,0,.35); color: #f59e0b; border:1px solid rgba(245,158,11,.3); }
        .badge.complete { background: rgba(20,83,45,.45); color: #22c55e; border:1px solid rgba(34,197,94,.3); }
        .badge-dot { width:8px; height:8px; border-radius:50%; background:currentColor; }

        /* EMS GROUPS */
        .grp-card { background: #0d1117; border: 1px solid #1e2d3d; border-radius: 12px; overflow: hidden; margin-bottom: 24px; }
        .grp-hdr { display: flex; align-items: center; gap: 16px; padding: 18px 28px; cursor: pointer; user-select: none; transition: background .15s; flex-wrap: wrap; }
        .grp-hdr:hover { background: rgba(255,255,255,.02); }
        .grp-dot { width:16px; height:16px; border-radius:50%; flex-shrink:0; }
        .grp-dot.g1{background:#ef4444} .grp-dot.g2{background:#3b82f6} .grp-dot.g3{background:#22c55e} .grp-dot.g4{background:#f59e0b}
        .grp-name { font-size:18px; font-weight:700; }
        .grp-name.g1{color:#f87171} .grp-name.g2{color:#60a5fa} .grp-name.g3{color:#4ade80} .grp-name.g4{color:#fbbf24}
        .grp-area { font-size:16px; color:#6e7d8e; }
        .grp-meta { margin-left:auto; display:flex; align-items:center; gap:20px; flex-wrap: wrap; }
        .grp-dates { font-size:14px; font-weight: 500; color:#6e7d8e; }
        .grp-done { font-size:14px; font-weight:700; color:#e6edf3; }
        .grp-chev { font-size:12px; color:#4d5565; transition:transform .2s; }
        .grp-hdr.open .grp-chev { transform:rotate(180deg); }
        
        .grp-prog-track { height:4px; background:#1a2233; position:relative; }
        .grp-prog-bar { height:4px; position:absolute; top:0; left:0; transition:width .5s ease; }
        .grp-prog-partial.g1{background:#fca5a5} .grp-prog-partial.g2{background:#93c5fd} .grp-prog-partial.g3{background:#86efac} .grp-prog-partial.g4{background:#fde68a}
        .grp-prog-complete.g1{background:linear-gradient(90deg,#dc2626,#ef4444)}
        .grp-prog-complete.g2{background:linear-gradient(90deg,#2563eb,#3b82f6)}
        .grp-prog-complete.g3{background:linear-gradient(90deg,#16a34a,#22c55e)}
        .grp-prog-complete.g4{background:linear-gradient(90deg,#d97706,#f59e0b)}

        /* TABLE */
        .stn-tbl { display:none; border-top:1px solid #1a2233; overflow-x: auto; }
        .stn-tbl.open { display:block; }
        .stn-hdr-row { display:grid; grid-template-columns: 80px minmax(250px, 1fr) 140px 150px 100px; gap:16px; padding:12px 28px; background:rgba(255,255,255,.025); border-bottom:1px solid #1a2233; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.12em; color:#6e7d8e; min-width: 850px; }
        .stn-row { display:grid; grid-template-columns: 80px minmax(250px, 1fr) 140px 150px 100px; gap:16px; padding:16px 28px; border-bottom:1px solid #111821; align-items:center; transition:background .15s; min-width: 850px; }
        .stn-row:last-child { border-bottom:none; }
        .stn-row:hover { background:rgba(255,255,255,.018); }
        .stn-row.complete { background:rgba(34,197,94,.028); }
        .stn-row.partial { background:rgba(245,158,11,.025); }
        .stn-id { font-size:16px; font-weight:700; color:#4d8ef7; }
        .stn-addr { font-size:15px; font-weight: 500; color:#e6edf3; }
        .stn-city { font-size:13px; color:#6e7d8e; margin-top:2px; }
        .stn-done { font-size:13px; font-weight: 500; color:#6e7d8e; font-family:monospace; }
        .stn-tech { font-size:13px; font-weight: 500; color:#8b949e; font-family:monospace; }

        /* UPLOAD SCREEN */
        .upload-screen { display:flex; flex-direction:column; align-items:center; justify-content:center; min-height:60vh; gap:24px; text-align:center; }
        .drop-zone { width:540px; max-width:100%; border:2px dashed #1e2d3d; border-radius:16px; padding:56px 48px; background:#0d1117; cursor:pointer; transition:all .2s; display:flex; flex-direction:column; align-items:center; gap:16px; margin: 40px auto; }
        .drop-zone:hover, .drop-zone.drag { border-color:#4d8ef7; background:rgba(77,142,247,.05); transform: translateY(-2px); }
        .drop-icon { font-size:48px; }
        .drop-title { font-size:20px; font-weight:700; color:#e6edf3; }
        .drop-sub { font-size:14px; color:#6e7d8e; line-height:1.6; }
        .drop-btn { background:#4d8ef7; color:#fff; border:none; padding:12px 28px; border-radius:8px; font-size:15px; font-weight:700; cursor:pointer; margin-top: 8px; transition: background 0.2s; }
        .drop-btn:hover { background:#3b82f6; }
      `}} />

      <div className="page">
        {/* HEADER */}
        <div className="header">
          <div className="header-left">
            <div className="org-name">Montgomery County Hospital District</div>
            <div className="page-title">Device Upgrade Schedule</div>
            <div className="header-stats">
              {loading ? "Loading Dashboard Data..." : !hasData ? "Load a workbook to view progress" : (
                <>
                  <b>{ov.done} complete{ov.partial ? ` · ${ov.partial} partial` : ''} of {ov.total} sites ({ov.pct}%)</b>
                  &nbsp;·&nbsp; Towers: {tS.done}/{tS.total}
                  &nbsp;·&nbsp; Stations: {sS.done}/{sS.total}
                  &nbsp;·&nbsp; Est. completion: {estDate ? fmtFull(estDate) : '—'}
                </>
              )}
            </div>
          </div>
          <div className="header-right">
            <div className="sync-badge">
              <div className={`sync-dot ${hasData ? 'live' : ''}`}></div>
              <div className="sync-inner">
                <div className="sync-status">{hasData ? `DATA LOADED · Updated ${syncTime}` : "No file loaded"}</div>
                <div className="sync-file">{fileName}</div>
              </div>
            </div>
            <label className="upload-btn" htmlFor="fileInput">↑ Upload Updated XLSX (Local)</label>
            <input type="file" id="fileInput" accept=".xlsx,.xls" style={{ display: 'none' }} ref={fileInputRef} onChange={handleFileUpload} />
          </div>
        </div>

        {/* SCHEDULE STRIP */}
        {hasData && (
          <div className="sched-strip">
            <div className="sched-cell">
              <div className="sched-lbl">🗼 Tower Phase Start</div>
              <div className={`sched-val ${!dates.towerStart ? 'warn' : ''}`}>{dates.towerStart ? fmtFull(dates.towerStart) : 'Not set in workbook'}</div>
              <div className="sched-hint">Summary Overview tab</div>
            </div>
            <div className="sched-cell">
              <div className="sched-lbl">🗼 Tower Est. Completion</div>
              <div className={`sched-val ${dates.towerEnd ? 'green' : 'warn'}`}>{dates.towerEnd ? fmtFull(dates.towerEnd) : 'Not set in workbook'}</div>
              <div className="sched-hint">Summary Overview tab</div>
            </div>
            <div className="sched-cell">
              <div className="sched-lbl">🏥 EMS Phase Start</div>
              <div className={`sched-val ${!dates.emsStart ? 'warn' : ''}`}>{dates.emsStart ? fmtFull(dates.emsStart) : 'Not set in workbook'}</div>
              <div className="sched-hint">Summary Overview tab</div>
            </div>
            <div className="sched-cell">
              <div className="sched-lbl">🏥 EMS Est. Completion</div>
              <div className={`sched-val ${dates.emsEnd ? 'green' : 'warn'}`}>{dates.emsEnd ? fmtFull(dates.emsEnd) : 'Not set in workbook'}</div>
              <div className="sched-hint">Summary Overview tab</div>
            </div>
          </div>
        )}

        {/* UPLOAD SCREEN (If No Data) */}
        {!hasData && !loading && (
          <div className="upload-screen">
            <div className="drop-zone" onClick={() => fileInputRef.current?.click()}>
              <div className="drop-icon">📊</div>
              <div className="drop-title">Load MCHD Master Workbook</div>
              <div className="drop-sub">Drag & drop <strong>MCHD_Master_Workbook.xlsx</strong> here or click to browse.<br />Schedule dates are read automatically from the <strong>Summary Overview</strong> tab.</div>
              <button className="drop-btn" onClick={(e) => { e.stopPropagation(); fileInputRef.current?.click(); }}>Browse Files</button>
            </div>
          </div>
        )}

        {/* DASHBOARD CONTENT */}
        {hasData && (
          <div>
            {/* OVERALL PROGRESS */}
            <div className="progress-card">
              <div className="progress-card-title">Overall Project Progress</div>
              <div className="prog-cols">
                <div className="prog-col">
                  <div className="prog-head">
                    <div className="prog-lbl">Overall</div>
                    <div className="prog-num">{ov.done}{ov.partial ? ` + ${ov.partial}p` : ''} / {ov.total}</div>
                  </div>
                  <div className="prog-track">
                    <div className="prog-bar bg" style={{ width: `${ov.donePct + ov.partialPct}%` }}></div>
                    <div className="prog-bar fg" style={{ width: `${ov.donePct}%` }}></div>
                  </div>
                  <div className="prog-sub">{ov.pct}% Complete{ov.partial ? ' · partial = 50%' : ''}</div>
                </div>
                <div className="prog-col">
                  <div className="prog-head">
                    <div className="prog-lbl towers">Towers</div>
                    <div className="prog-num">{tS.done}{tS.partial ? ` + ${tS.partial}p` : ''} / {tS.total}</div>
                  </div>
                  <div className="prog-track">
                    <div className="prog-bar bg towers" style={{ width: `${tS.donePct + tS.partialPct}%` }}></div>
                    <div className="prog-bar fg towers" style={{ width: `${tS.donePct}%` }}></div>
                  </div>
                  <div className="prog-sub">Phase 1 · {tRange} · {tS.pct}%</div>
                </div>
                <div className="prog-col">
                  <div className="prog-head">
                    <div className="prog-lbl ems">EMS Stations</div>
                    <div className="prog-num">{sS.done}{sS.partial ? ` + ${sS.partial}p` : ''} / {sS.total}</div>
                  </div>
                  <div className="prog-track">
                    <div className="prog-bar bg ems" style={{ width: `${sS.donePct + sS.partialPct}%` }}></div>
                    <div className="prog-bar fg ems" style={{ width: `${sS.donePct}%` }}></div>
                  </div>
                  <div className="prog-sub">Phase 2 · {sRange} · {sS.pct}%</div>
                </div>
              </div>
            </div>

            {/* TOWERS SECTION */}
            <div className="sec-label">{tRange ? `Phase 1 — Tower Sites · ${tRange}` : 'Phase 1 — Tower Sites'}</div>
            <div className="tower-box">
              <div className="tower-box-hdr">
                <div className="twr-icon">🗼</div>
                <div>
                  <div className="twr-hdr-title">Tower Phase</div>
                  <div className="twr-hdr-sub">{tRange ? `${tRange} · ${towers.length} Locations` : `${towers.length} Locations`}</div>
                </div>
                <div className="twr-badge">
                  {tS.partial ? `${tS.done} done · ${tS.partial} partial` : `${tS.done} / ${towers.length}`}
                </div>
              </div>
              <div className="twr-grid">
                {towers.map((t, i) => {
                  const c = sc(t.status);
                  const dd = toDate(t.dateDone);
                  return (
                    <div key={i} className={`twr-cell ${c}`}>
                      <div className="twr-num">TWR {i + 1}</div>
                      <div className="twr-name">{shortName(t.name)}</div>
                      <div className="twr-addr">{t.addr}, {t.city}</div>
                      {tRange && <div className="twr-win"><span className="twr-win-sq"></span>{tRange}</div>}
                      {renderBadge(t.status)}
                      {dd && <div className="twr-done">Done: {fmtFull(dd)}</div>}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* STATIONS SECTION */}
            <div className="sec-label">{sRange ? `Phase 2 — EMS Stations · ${sRange}` : 'Phase 2 — EMS Stations'}</div>
            <div>
              {groups.map((g, gi) => {
                const { total, done, partial, pct, donePct, partialPct } = statsFor(g.stations);
                const label = `${done}/${total} Done${partial ? ' · ' + partial + ' Partial' : ''}  ${pct}%`;
                
                let grpDateStr = '';
                if (dates.emsStart && dates.emsEnd && groups.length > 0) {
                  const seg = (dates.emsEnd.getTime() - dates.emsStart.getTime()) / groups.length;
                  const gS = new Date(dates.emsStart.getTime() + gi * seg);
                  const gE = new Date(dates.emsStart.getTime() + (gi + 1) * seg);
                  grpDateStr = `${fmt(gS)} – ${fmt(gE)}`;
                } else if (dates.emsStart) {
                  grpDateStr = `Starting ${fmt(dates.emsStart)}`;
                }

                const isOpen = openGroups[gi];

                return (
                  <div key={gi} className="grp-card">
                    <div className={`grp-hdr ${isOpen ? 'open' : ''}`} onClick={() => toggleGrp(gi)}>
                      <div className={`grp-dot ${g.color}`}></div>
                      <div className={`grp-name ${g.color}`}>{g.sheet}</div>
                      <div className="grp-area">— {g.area}</div>
                      <div className="grp-meta">
                        <div className="grp-dates">{grpDateStr}</div>
                        <div className="grp-done">{label}</div>
                        <div className="grp-chev">▲</div>
                      </div>
                    </div>
                    <div className="grp-prog-track">
                      <div className={`grp-prog-bar grp-prog-partial ${g.color}`} style={{ width: `${donePct + partialPct}%` }}></div>
                      <div className={`grp-prog-bar grp-prog-complete ${g.color}`} style={{ width: `${donePct}%` }}></div>
                    </div>
                    <div className={`stn-tbl ${isOpen ? 'open' : ''}`}>
                      <div className="stn-hdr-row">
                        <div>Station</div><div>Address</div><div>Status</div><div>Date Completed</div><div>Tech</div>
                      </div>
                      {g.stations.map((s: any) => {
                        const c = sc(s.status);
                        const dd = toDate(s.dateDone);
                        return (
                          <div key={s.id} className={`stn-row ${c}`}>
                            <div className="stn-id">{s.id}</div>
                            <div>
                              <div className="stn-addr">{s.addr}</div>
                              <div className="stn-city">{s.city}, TX {s.zip}</div>
                            </div>
                            {renderBadge(s.status)}
                            <div className="stn-done">{dd ? fmtFull(dd) : '—'}</div>
                            <div className="stn-tech">{s.tech || '—'}</div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}