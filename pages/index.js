import { useState, useEffect, useCallback } from 'react';

// ── Colours ───────────────────────────────────────────────────────
const C = {
  // House palette
  stone: '#2C2925',
  warm: '#6B5E52',
  sand: '#B5A494',
  parchment:'#F5F0EA',
  linen: '#EDE7DE',
  white: '#ffffff',
  // Functional
  green: '#4A6741',
  lightGreen:'#DCE8DA',
  red: '#7B3B3B',
  lightRed: '#F2E0E0',
  orange: '#8B5E3C',
  ltOrange: '#F2E6D9',
  // Aliases for legacy refs
  navy: '#2C2925',
  darkBlue: '#2C2925',
  midBlue: '#6B5E52',
  lightBlue:'#EDE7DE',
  vlBlue: '#F5F0EA',
  grey: '#6B5E52',
  ltGrey: '#EDE7DE',
  brdr: '#B5A494',
  purple: '#5C4A5A',
  ltPurple: '#EDE7DE',
  teal: '#3D5A58',
  ltTeal: '#DCE8E7',
  gold: '#7A6A3A',
  ltGold: '#F2EDD9',
};

// Muted desaturated expiry bucket colours
const BKT = {
  '2026': { bg:'#7B3B3B', lt:'#F2E0E0' },
  '2027': { bg:'#8B5E3C', lt:'#F2E6D9' },
  '2028': { bg:'#4A6741', lt:'#DCE8DA' },
  '2029': { bg:'#3D5A6B', lt:'#D9E6F0' },
  '2030': { bg:'#5C4A5A', lt:'#EDE7ED' },
  '2031+':{ bg:'#3D5A58', lt:'#DCE8E7' },
  'Vacant':{ bg:'#8F3A3A', lt:'#F1DEDE' },
};

// ── Building configs ─────────────────────────────────────────────
const BUILDINGS = {
  castlereagh: {
    id: 'castlereagh',
    file: '/data/17_castlereagh.xlsx',
    name: '17 Castlereagh Street',
    suburb: 'Sydney CBD',
    shortName: 'Castlereagh',
    accent: C.stone,
    expiryCol: 'Z',
    psmCol: 'AA',
    nameCol: 'AB',
    reviewTypeCol: 'P',
    nextReviewCol: 'Q',
    netCol: 'N',
    dataStart: 3,
    dataEnd: 27,
    floorOrder: ['L12','L11','L10','L09','L08','L07','L06','L05','L04','L03','L02','L01','GF','LWR GND'],
    critDatesStart: 67,
    critDatesEnd: 78,
  },
  elizabeth: {
    id: 'elizabeth',
    file: '/data/1_elizabeth_plaza.xlsx',
    name: '1 Elizabeth Plaza',
    suburb: 'North Sydney',
    shortName: 'Elizabeth',
    accent: C.warm,
    expiryCol: 'AD',
    psmCol: 'AE',
    nameCol: 'AF',
    reviewTypeCol: 'S',
    nextReviewCol: 'T',
    netCol: 'Q',
    dataStart: 3,
    dataEnd: 24,
    floorOrder: ['L12','L11','L10','L09','L08','L07','L06','L05','L04','L03'],
    critDatesStart: 48,
    critDatesEnd: 59,
  },
};

// ── SheetJS helpers ──────────────────────────────────────────────
function cv(ws, addr) {
  const cell = ws[addr]; if (!cell) return null;
  return cell.v !== undefined ? cell.v : null;
}
function cd(ws, addr) {
  const cell = ws[addr]; if (!cell) return null;
  if (cell.t === 'd') return new Date(cell.v);
  if (cell.t === 'n' && cell.v > 40000 && cell.v < 80000)
    return new Date(Math.round((cell.v - 25569) * 86400000));
  return null;
}
function fmtCcy(v) { return v > 0 ? '$' + Math.round(v).toLocaleString('en-AU') : '—'; }
function fmtPct(v, dec=1) { return v != null ? (v*100).toFixed(dec)+'%' : '—'; }
function fmtNum(v, dec=1) { return v != null ? v.toLocaleString('en-AU',{maximumFractionDigits:dec,minimumFractionDigits:dec}) : '—'; }
function fmtDate(d) {
  if (!d) return '—';
  try { return d.toLocaleDateString('en-AU',{day:'2-digit',month:'short',year:'numeric'}); }
  catch { return '—'; }
}
function fmtPSM(v) { return v>0 ? '$'+Math.round(v).toLocaleString('en-AU') : '—'; }
function bucketKey(d) {
  if (!d) return 'Vacant';
  const y = d.getFullYear();
  if (y <= 2026) return '2026';
  if (y <= 2030) return String(y);
  return '2031+';
}

// ── Parser ───────────────────────────────────────────────────────
function parseBuilding(wb, cfg) {
  const id = wb.Sheets['Input Data']; if (!id) return null;
  const today = new Date(); today.setHours(0,0,0,0);
  const suites = [];

  for (let r = cfg.dataStart; r <= cfg.dataEnd; r++) {
    const floor = cv(id, `B${r}`);
    const suiteNum = cv(id, `C${r}`);
    const nla = parseFloat(cv(id, `F${r}`)) || 0;
    const expCell = id[`${cfg.expiryCol}${r}`];
    const psm = parseFloat(cv(id, `${cfg.psmCol}${r}`)) || 0;
    const dispName = cv(id, `${cfg.nameCol}${r}`) || suiteNum || `Row ${r}`;

    if (!floor && !suiteNum && !nla) continue;

    let expiry = null;
    if (expCell) {
      if (expCell.t === 'd') expiry = new Date(expCell.v);
      else if (expCell.t === 'n' && expCell.v > 40000)
        expiry = new Date(Math.round((expCell.v - 25569) * 86400000));
    }

    const active = expiry && expiry > today && (psm > 0 || cv(id, `${cfg.nameCol}${r}`) != null);
    const expired = expiry && expiry <= today;
    const vacant = !expiry || psm === 0;
    const remYrs = active ? (expiry - today) / (365.25 * 86400000) : 0;
    const grossPA = active ? psm * nla : 0;
    const mktRaw = parseFloat(cv(id, `X${r}`)) || 0;
    const netPA = active ? (parseFloat(cv(id, `${cfg.netCol}${r}`)) || grossPA) : 0;

    suites.push({
      r, floor, suiteNum: String(suiteNum || ''), nla, expiry, psm, dispName,
      active, expired, vacant, remYrs, grossPA, netPA, mktPSM: mktRaw,
    });
  }

  const active = suites.filter(s => s.active);
  const totalNLA = suites.reduce((s,x) => s+x.nla, 0);
  const occNLA = active.reduce((s,x) => s+x.nla, 0);
  const vacNLA = totalNLA - occNLA;
  const occ = totalNLA > 0 ? occNLA/totalNLA : 0;
  const grossInc = active.reduce((s,x) => s+x.grossPA, 0);
  const netInc = active.reduce((s,x) => s+x.netPA, 0);
  const wNLA_n = active.reduce((s,x) => s+x.nla*x.remYrs, 0);
  const wNLA_d = active.reduce((s,x) => s+x.nla, 0);
  const wInc_n = active.reduce((s,x) => s+x.grossPA*x.remYrs, 0);
  const wInc_d = active.reduce((s,x) => s+x.grossPA, 0);
  const waleNLA = wNLA_d > 0 ? wNLA_n/wNLA_d : 0;
  const waleInc = wInc_d > 0 ? wInc_n/wInc_d : 0;

  const buckets = {'2026':[],'2027':[],'2028':[],'2029':[],'2030':[],'2031+':[],'Vacant':[]};
  for (const s of suites) {
    if (!s.active) { buckets['Vacant'].push(s); continue; }
    buckets[bucketKey(s.expiry)].push(s);
  }

  const mktRents = {};
  const mr = wb.Sheets['Market Rents 26-27'];
  if (mr) {
    for (let r = 4; r <= 30; r++) {
      const sn = cv(mr, `B${r}`); const psm = parseFloat(cv(mr, `F${r}`)) || 0;
      if (sn && psm) mktRents[String(sn)] = psm;
    }
  }
  for (const s of suites) {
    if (mktRents[s.suiteNum]) s.mktPSM = mktRents[s.suiteNum];
  }

  const floorMap = {};
  for (const s of suites) {
    if (!s.floor) continue;
    if (!floorMap[s.floor]) floorMap[s.floor] = { floor:s.floor, nla:0, income:0, tenants:[], minExpiry:null };
    floorMap[s.floor].nla += s.nla;
    if (s.active) {
      floorMap[s.floor].income += s.grossPA;
      floorMap[s.floor].tenants.push(`${s.dispName} (${s.suiteNum})`);
      if (!floorMap[s.floor].minExpiry || s.expiry < floorMap[s.floor].minExpiry)
        floorMap[s.floor].minExpiry = s.expiry;
    }
  }
  const byFloor = cfg.floorOrder.map(f => floorMap[f]).filter(Boolean);

  const criticalDates = [];
  const d18mo = new Date(today); d18mo.setMonth(d18mo.getMonth()+18);
  for (const s of suites) {
    if (!s.active) continue;
    if (s.expiry && s.expiry <= d18mo) {
      const days = Math.round((s.expiry - today)/86400000);
      criticalDates.push({ suite:s.suiteNum, tenant:s.dispName, event:'Lease Expiry',
        date:s.expiry, days, note:`${s.nla.toFixed(0)} sqm · ${fmtPSM(s.psm)} psm` });
    }
    const rv = id[`${cfg.nextReviewCol}${s.r}`];
    if (rv) {
      let rvDate = null;
      if (rv.t==='d') rvDate = new Date(rv.v);
      else if (rv.t==='n' && rv.v>40000) rvDate = new Date(Math.round((rv.v-25569)*86400000));
      const d12mo = new Date(today); d12mo.setMonth(d12mo.getMonth()+12);
      if (rvDate && rvDate >= today && rvDate <= d12mo) {
        const days = Math.round((rvDate - today)/86400000);
        const rt = cv(id, `${cfg.reviewTypeCol}${s.r}`) || '';
        criticalDates.push({ suite:s.suiteNum, tenant:s.dispName, event:`Rent Review (${rt})`,
          date:rvDate, days, note:`Review date: ${fmtDate(rvDate)}` });
      }
    }
  }
  criticalDates.sort((a,b) => a.date - b.date);

  let valuation = null;
  const vt = wb.Sheets['Valuations'];
  if (vt) {
    const v1a = parseFloat(cv(vt,'C6')) || 0;
    const v1d = cd(vt,'C7'), v1v = cv(vt,'C8'), v1cr = parseFloat(cv(vt,'C10')) || 0;
    const v2a = parseFloat(cv(vt,'D6')) || 0, v2d = cd(vt,'D7');
    const v3a = parseFloat(cv(vt,'E6')) || 0;
    if (v1a > 0) valuation = {
      v1:{ amount:v1a, date:v1d, valuer:v1v, capRate:v1cr },
      v2:{ amount:v2a, date:v2d }, v3:{ amount:v3a },
      movement: v1a&&v2a ? v1a-v2a : null,
      movPct: v1a&&v2a ? (v1a-v2a)/v2a : null,
    };
  }

  return { suites, active, totalNLA, occNLA, vacNLA, occ, grossInc, netInc,
    waleNLA, waleInc, buckets, byFloor, criticalDates, valuation, cfg };
}

// ── PIN Page ─────────────────────────────────────────────────────
function PinPage({ onSuccess }) {
  const [pin, setPin] = useState('');
  const [err, setErr] = useState(false);
  const correct = process.env.NEXT_PUBLIC_DASHBOARD_PIN || '2025';

  function submit(e) {
    e.preventDefault();
    if (pin === correct) { sessionStorage.setItem('dash_auth','1'); onSuccess(); }
    else { setErr(true); setPin(''); }
  }

  return (
    <div style={{ minHeight:'100vh', display:'flex', alignItems:'center', justifyContent:'center',
      background:`linear-gradient(135deg, ${C.stone} 0%, ${C.warm} 100%)` }}>
      <div style={{ background:C.parchment, borderRadius:12, padding:'48px 40px', width:360,
        boxShadow:'0 20px 60px rgba(0,0,0,0.35)', textAlign:'center', fontFamily:'Georgia, serif' }}>
        <div style={{ fontSize:40, marginBottom:16 }}>🏢</div>
        <div style={{ color:C.sand, fontSize:11, fontWeight:600, letterSpacing:2,
          textTransform:'uppercase', marginBottom:8 }}>Goldberg Family Office</div>
        <h1 style={{ color:C.stone, fontSize:20, fontWeight:700, marginBottom:4 }}>Portfolio Dashboard</h1>
        <p style={{ color:C.warm, fontSize:12, marginBottom:8 }}>17 Castlereagh Street, Sydney CBD</p>
        <p style={{ color:C.warm, fontSize:12, marginBottom:28 }}>1 Elizabeth Plaza, North Sydney</p>
        <form onSubmit={submit}>
          <input type="password" inputMode="numeric" placeholder="Enter PIN" value={pin}
            onChange={e => { setPin(e.target.value); setErr(false); }} maxLength={8} autoFocus
            style={{ width:'100%', padding:'12px 16px', fontSize:22, textAlign:'center',
              letterSpacing:8, border:`2px solid ${err?C.red:C.sand}`, borderRadius:8,
              marginBottom:err?8:16, outline:'none', background:err?'#f9f0f0':C.white,
              fontFamily:'Georgia, serif' }}/>
          {err && <p style={{ color:C.red, fontSize:12, marginBottom:12 }}>Incorrect PIN</p>}
          <button type="submit" style={{ width:'100%', padding:'12px', background:C.stone,
            color:C.parchment, border:'none', borderRadius:8, fontSize:14, fontWeight:600,
            cursor:'pointer', fontFamily:'Georgia, serif', letterSpacing:1 }}>
            Enter Dashboard
          </button>
        </form>
        <p style={{ color:C.sand, fontSize:11, marginTop:20 }}>Authorised access only</p>
      </div>
    </div>
  );
}

// ── Top Nav ──────────────────────────────────────────────────────
function TopNav({ view, setView, loaded }) {
  const groups = [
    { label:'', tabs:['portfolio'] },
    { label:'17 Castlereagh St, Sydney CBD', tabs:['cs_dashboard','cs_stack'], color:C.stone },
    { label:'1 Elizabeth Plaza, North Sydney', tabs:['ep_dashboard','ep_stack'], color:C.warm },
  ];

  return (
    <div style={{ background:C.stone, fontFamily:'Georgia, serif' }}>
      {/* GFO Header */}
      <div style={{ maxWidth:1400, margin:'0 auto', padding:'10px 24px 6px', display:'flex',
        justifyContent:'space-between', alignItems:'baseline',
        borderBottom:'1px solid rgba(255,255,255,0.1)' }}>
        <span style={{ color:C.parchment, fontSize:13, fontWeight:700, letterSpacing:2,
          fontFamily:'Georgia, serif', textTransform:'uppercase' }}>Goldberg Family Office</span>
        <span style={{ color:C.sand, fontSize:11, fontFamily:'Georgia, serif' }}>Prepared by Jake Goldberg</span>
      </div>
      {/* Building group labels */}
      <div style={{ maxWidth:1400, margin:'0 auto', display:'flex',
        borderBottom:`1px solid rgba(255,255,255,0.1)`, padding:'0 24px' }}>
        <div style={{ width:120 }}/>
        {groups.slice(1).map(g => (
          <div key={g.label} style={{ flex:1, padding:'6px 8px',
            borderLeft:'1px solid rgba(255,255,255,0.1)', borderRight:'1px solid rgba(255,255,255,0.1)' }}>
            <span style={{ color:C.sand, fontSize:10, fontWeight:600 }}>{g.label}</span>
          </div>
        ))}
      </div>
      {/* Tabs */}
      <div style={{ maxWidth:1400, margin:'0 auto', display:'flex',
        alignItems:'stretch', padding:'0 24px' }}>
        <button onClick={() => setView('portfolio')}
          style={{ padding:'14px 20px', border:'none', background:'transparent', cursor:'pointer',
            fontSize:13, fontWeight:700, fontFamily:'Georgia, serif',
            color: view==='portfolio'?C.parchment:C.sand,
            borderBottom: view==='portfolio'?`3px solid ${C.sand}`:'3px solid transparent',
            transition:'all 0.15s', whiteSpace:'nowrap', width:120, flexShrink:0 }}>
          Portfolio
        </button>
        <div style={{ flex:1, display:'flex', borderLeft:'1px solid rgba(255,255,255,0.1)',
          borderRight:'1px solid rgba(255,255,255,0.1)' }}>
          {['cs_dashboard','cs_stack'].map(id => {
            const label = id.includes('stack') ? 'Stack Plan' : 'Dashboard';
            return (
              <button key={id} onClick={() => setView(id)}
                style={{
