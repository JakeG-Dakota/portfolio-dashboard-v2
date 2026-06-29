import { useState, useEffect, useCallback } from 'react';
// ── Colours ───────────────────────────────────────────────────────
const C = {
  navy:'#1a3a5c', darkBlue:'#204e7a', midBlue:'#2e75b6',
  lightBlue:'#daeaf7', vlBlue:'#eef5fb', white:'#ffffff',
  grey:'#595959', ltGrey:'#f2f2f2', brdr:'#d0d8e4',
  green:'#1e5631', lightGreen:'#e2efda',
  red:'#7b0000', lightRed:'#ffe5e5',
  orange:'#c55a11', ltOrange:'#fce4d6',
  purple:'#5c2d91', ltPurple:'#ede7f6',
  teal:'#006064', ltTeal:'#e0f7fa',
  gold:'#b8860b', ltGold:'#fff8dc',
};
const BKT = {
  '2026': { bg:'#c0392b', lt:'#fadbd8' },
  '2027': { bg:'#d35400', lt:'#fdebd0' },
  '2028': { bg:'#1e8449', lt:'#d5f5e3' },
  '2029': { bg:'#1a5276', lt:'#d6eaf8' },
  '2030': { bg:'#6c3483', lt:'#e8daef' },
  '2031+':{ bg:'#0e6655', lt:'#d1f2eb' },
  'Vacant':{ bg:'#808080', lt:'#ececec' },
};

// ── Building configs ─────────────────────────────────────────────
const BUILDINGS = {
  castlereagh: {
    id: 'castlereagh',
    file: '/data/17_castlereagh.xlsx',
    name: '17 Castlereagh Street',
    suburb: 'Sydney CBD',
    shortName: 'Castlereagh',
    accent: C.darkBlue,
    expiryCol: 'Z',
    psmCol:    'AA',
    nameCol:   'AB',
    reviewTypeCol: 'P',
    nextReviewCol: 'Q',
    dataStart: 3,
    dataEnd:   27,
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
    accent: '#1F6B75',
    expiryCol: 'AD',
    psmCol:    'AE',
    nameCol:   'AF',
    reviewTypeCol: 'S',
    nextReviewCol: 'T',
    dataStart: 3,
    dataEnd:   23,
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
  const id  = wb.Sheets['Input Data']; if (!id) return null;
  const today = new Date(); today.setHours(0,0,0,0);
  const suites = [];

  for (let r = cfg.dataStart; r <= cfg.dataEnd; r++) {
    const floor    = cv(id, `B${r}`);
    const suiteNum = cv(id, `C${r}`);
    const nla      = parseFloat(cv(id, `F${r}`)) || 0;
    const expCell  = id[`${cfg.expiryCol}${r}`];
    const psm      = parseFloat(cv(id, `${cfg.psmCol}${r}`)) || 0;
    const dispName = cv(id, `${cfg.nameCol}${r}`) || suiteNum || `Row ${r}`;

    if (!floor && !suiteNum && !nla) continue;

    let expiry = null;
    if (expCell) {
      if (expCell.t === 'd') expiry = new Date(expCell.v);
      else if (expCell.t === 'n' && expCell.v > 40000)
        expiry = new Date(Math.round((expCell.v - 25569) * 86400000));
    }

    const active  = expiry && expiry > today && psm > 0;
    const expired = expiry && expiry <= today;
    const vacant  = !expiry || psm === 0;
    const remYrs  = active ? (expiry - today) / (365.25 * 86400000) : 0;
    const grossPA = active ? psm * nla : 0;

    const mktRaw = parseFloat(cv(id, `X${r}`)) || 0;

    suites.push({
      r, floor, suiteNum: String(suiteNum || ''), nla, expiry, psm, dispName,
      active, expired, vacant, remYrs, grossPA, mktPSM: mktRaw,
    });
  }

  const active   = suites.filter(s => s.active);
  const totalNLA = suites.reduce((s,x) => s+x.nla, 0);
  const occNLA   = active.reduce((s,x) => s+x.nla, 0);
  const vacNLA   = totalNLA - occNLA;
  const occ      = totalNLA > 0 ? occNLA/totalNLA : 0;
  const grossInc = active.reduce((s,x) => s+x.grossPA, 0);
  const wNLA_n   = active.reduce((s,x) => s+x.nla*x.remYrs, 0);
  const wNLA_d   = active.reduce((s,x) => s+x.nla, 0);
  const wInc_n   = active.reduce((s,x) => s+x.grossPA*x.remYrs, 0);
  const wInc_d   = active.reduce((s,x) => s+x.grossPA, 0);
  const waleNLA  = wNLA_d > 0 ? wNLA_n/wNLA_d : 0;
  const waleInc  = wInc_d > 0 ? wInc_n/wInc_d : 0;

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

  return { suites, active, totalNLA, occNLA, vacNLA, occ, grossInc,
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
      background:`linear-gradient(135deg, ${C.navy} 0%, #1F6B75 100%)` }}>
      <div style={{ background:C.white, borderRadius:12, padding:'48px 40px', width:360,
        boxShadow:'0 20px 60px rgba(0,0,0,0.35)', textAlign:'center' }}>
        <div style={{ fontSize:40, marginBottom:16 }}>🏢</div>
        <h1 style={{ color:C.navy, fontSize:18, fontWeight:700, marginBottom:4 }}>Portfolio Dashboard</h1>
        <p style={{ color:C.grey, fontSize:12, marginBottom:8 }}>17 Castlereagh Street, Sydney CBD</p>
        <p style={{ color:C.grey, fontSize:12, marginBottom:28 }}>1 Elizabeth Plaza, North Sydney</p>
        <form onSubmit={submit}>
          <input type="password" inputMode="numeric" placeholder="Enter PIN" value={pin}
            onChange={e => { setPin(e.target.value); setErr(false); }} maxLength={8} autoFocus
            style={{ width:'100%', padding:'12px 16px', fontSize:22, textAlign:'center',
              letterSpacing:8, border:`2px solid ${err?C.red:C.brdr}`, borderRadius:8,
              marginBottom:err?8:16, outline:'none', background:err?'#fff5f5':C.white }}/>
          {err && <p style={{ color:C.red, fontSize:12, marginBottom:12 }}>Incorrect PIN</p>}
          <button type="submit" style={{ width:'100%', padding:'12px', background:C.darkBlue,
            color:C.white, border:'none', borderRadius:8, fontSize:14, fontWeight:600, cursor:'pointer' }}>
            Enter Dashboard
          </button>
        </form>
        <p style={{ color:'#bbb', fontSize:11, marginTop:20 }}>Authorised access only</p>
      </div>
    </div>
  );
}

// ── Top Nav ──────────────────────────────────────────────────────
function TopNav({ view, setView, loaded }) {
  const tabs = [
    { id:'portfolio', label:'📊 Portfolio', group:null },
    { id:'cs_dashboard', label:'Dashboard', group:'17 Castlereagh' },
    { id:'cs_stack',    label:'Stack Plan', group:'17 Castlereagh' },
    { id:'ep_dashboard', label:'Dashboard', group:'1 Elizabeth Plaza' },
    { id:'ep_stack',    label:'Stack Plan', group:'1 Elizabeth Plaza' },
  ];

  const groups = [
    { label:'', tabs:['portfolio'] },
    { label:'17 Castlereagh St, Sydney CBD', tabs:['cs_dashboard','cs_stack'], color:C.darkBlue },
    { label:'1 Elizabeth Plaza, North Sydney', tabs:['ep_dashboard','ep_stack'], color:'#1F6B75' },
  ];

  return (
    <div style={{ background:C.navy }}>
      {/* GFO Header */}
      <div style={{ maxWidth:1400, margin:'0 auto', padding:'10px 24px 6px', display:'flex',
        justifyContent:'space-between', alignItems:'baseline',
        borderBottom:'1px solid rgba(255,255,255,0.1)' }}>
        <span style={{ color:'rgba(255,255,255,0.9)', fontSize:13, fontWeight:700,
          letterSpacing:1, fontFamily:'Georgia, serif' }}>Goldberg Family Office</span>
        <span style={{ color:'rgba(255,255,255,0.45)', fontSize:11,
          fontFamily:'Georgia, serif' }}>Prepared by Jake Goldberg</span>
      </div>
      {/* Building group labels */}
      <div style={{ maxWidth:1400, margin:'0 auto', display:'flex',
        borderBottom:`1px solid rgba(255,255,255,0.1)`, padding:'0 24px' }}>
        <div style={{ width:120 }}/>
        {groups.slice(1).map(g => (
          <div key={g.label} style={{ flex:1, padding:'6px 8px',
            borderLeft:'1px solid rgba(255,255,255,0.1)', borderRight:'1px solid rgba(255,255,255,0.1)' }}>
            <span style={{ color:'rgba(255,255,255,0.5)', fontSize:10, fontWeight:600 }}>{g.label}</span>
          </div>
        ))}
      </div>

      {/* Tabs */}
      <div style={{ maxWidth:1400, margin:'0 auto', display:'flex',
        alignItems:'stretch', padding:'0 24px' }}>
        {/* Portfolio tab */}
        <button onClick={() => setView('portfolio')}
          style={{ padding:'14px 20px', border:'none', background:'transparent', cursor:'pointer',
            fontSize:13, fontWeight:700, color: view==='portfolio'?C.white:'rgba(255,255,255,0.55)',
            borderBottom: view==='portfolio'?`3px solid ${C.white}`:'3px solid transparent',
            transition:'all 0.15s', whiteSpace:'nowrap', width:120, flexShrink:0 }}>
          📊 Portfolio
        </button>
        {/* Castlereagh tabs */}
        <div style={{ flex:1, display:'flex', borderLeft:'1px solid rgba(255,255,255,0.1)',
          borderRight:'1px solid rgba(255,255,255,0.1)' }}>
          {['cs_dashboard','cs_stack'].map(id => {
            const label = id.includes('stack') ? 'Stack Plan' : 'Dashboard';
            return (
              <button key={id} onClick={() => setView(id)}
                style={{ flex:1, padding:'14px 16px', border:'none', background:'transparent',
                  cursor:'pointer', fontSize:13, fontWeight:600,
                  color: view===id?C.white:'rgba(255,255,255,0.55)',
                  borderBottom: view===id?`3px solid ${C.darkBlue==='#204e7a'?'#6ea8fe':'#6ea8fe'}`:'3px solid transparent',
                  transition:'all 0.15s' }}>
                {label} {!loaded.castlereagh && id==='cs_dashboard' ? '⏳' : ''}
              </button>
            );
          })}
        </div>
        {/* Elizabeth tabs */}
        <div style={{ flex:1, display:'flex', borderLeft:'1px solid rgba(255,255,255,0.1)' }}>
          {['ep_dashboard','ep_stack'].map(id => {
            const label = id.includes('stack') ? 'Stack Plan' : 'Dashboard';
            return (
              <button key={id} onClick={() => setView(id)}
                style={{ flex:1, padding:'14px 16px', border:'none', background:'transparent',
                  cursor:'pointer', fontSize:13, fontWeight:600,
                  color: view===id?C.white:'rgba(255,255,255,0.55)',
                  borderBottom: view===id?'3px solid #4ecdc4':'3px solid transparent',
                  transition:'all 0.15s' }}>
                {label} {!loaded.elizabeth && id==='ep_dashboard' ? '⏳' : ''}
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ── Portfolio overview ───────────────────────────────────────────
function PortfolioView({ buildings }) {
  const cs = buildings.castlereagh;
  const ep = buildings.elizabeth;
  const portfolioNLA   = (cs?.totalNLA||0) + (ep?.totalNLA||0);
  const portfolioInc   = (cs?.grossInc||0) + (ep?.grossInc||0);
  const portfolioOccNLA= (cs?.occNLA||0) + (ep?.occNLA||0);
  const portfolioOcc   = portfolioNLA>0 ? portfolioOccNLA/portfolioNLA : 0;
  const waleNLA_n = ((cs?.waleNLA||0)*(cs?.occNLA||0)) + ((ep?.waleNLA||0)*(ep?.occNLA||0));
  const waleNLA_d = (cs?.occNLA||0) + (ep?.occNLA||0);
  const portfolioWALE  = waleNLA_d > 0 ? waleNLA_n/waleNLA_d : 0;

  return (
    <div>
      {/* Portfolio KPIs */}
      <div style={{ background:C.navy, margin:'-24px -24px 24px', padding:'16px 24px' }}>
        <div style={{ maxWidth:1400-48, display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:12 }}>
          {[
            ['PORTFOLIO NLA', `${fmtNum(portfolioNLA,0)} sqm`, C.midBlue],
            ['PORTFOLIO OCCUPANCY', fmtPct(portfolioOcc), portfolioOcc>=0.90?C.green:C.orange],
            ['PORTFOLIO GROSS INCOME PA', fmtCcy(portfolioInc), C.midBlue],
            ['PORTFOLIO WALE (NLA)', `${portfolioWALE.toFixed(2)} yrs`, C.teal],
          ].map(([l,v,a]) => (
            <div key={l} style={{ background:'rgba(255,255,255,0.08)', borderRadius:8,
              padding:'12px 16px', borderLeft:`3px solid ${a}` }}>
              <div style={{ color:'rgba(255,255,255,0.55)', fontSize:10, fontWeight:600, marginBottom:4 }}>{l}</div>
              <div style={{ color:C.white, fontSize:20, fontWeight:700 }}>{v}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Side by side comparison */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:20, marginBottom:20 }}>
        {[
          { data:cs, cfg:BUILDINGS.castlereagh, accent:C.darkBlue },
          { data:ep, cfg:BUILDINGS.elizabeth,   accent:'#1F6B75' },
        ].map(({ data, cfg, accent }) => (
          <Card key={cfg.id} title={cfg.name} accent={accent}>
            {!data ? (
              <div style={{ padding:24, textAlign:'center', color:C.grey }}>
                <div style={{ fontSize:24, marginBottom:8 }}>⏳</div>
                <div>Loading workbook...</div>
              </div>
            ) : (
              <>
                <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:8, marginBottom:16 }}>
                  {[
                    ['Total NLA', `${fmtNum(data.totalNLA,0)} sqm`],
                    ['Occupancy', fmtPct(data.occ)],
                    ['Gross Income PA', fmtCcy(data.grossInc)],
                    ['WALE (NLA)', `${data.waleNLA.toFixed(2)} yrs`],
                    ['WALE (Income)', `${data.waleInc.toFixed(2)} yrs`],
                    ['Vacant NLA', `${fmtNum(data.vacNLA,0)} sqm`],
                  ].map(([l,v]) => (
                    <div key={l} style={{ background:C.vlBlue, borderRadius:6, padding:'10px 12px' }}>
                      <div style={{ fontSize:10, color:C.grey, marginBottom:3 }}>{l}</div>
                      <div style={{ fontSize:15, fontWeight:700, color:C.navy }}>{v}</div>
                    </div>
                  ))}
                </div>
                {/* Mini expiry bar */}
                <div style={{ fontSize:11, color:C.grey, marginBottom:6, fontWeight:600 }}>
                  Expiry Profile
                </div>
                <div style={{ display:'flex', height:20, borderRadius:4, overflow:'hidden', gap:1 }}>
                  {['2026','2027','2028','2029','2030','2031+'].map(yr => {
                    const n = data.buckets[yr]?.reduce((s,x)=>s+x.nla,0)||0;
                    const pct = data.totalNLA>0 ? (n/data.totalNLA)*100 : 0;
                    return pct > 1 ? (
                      <div key={yr} style={{ flex:pct, background:BKT[yr].bg, display:'flex',
                        alignItems:'center', justifyContent:'center', minWidth:0 }}
                        title={`${yr}: ${fmtNum(n,0)} sqm`}>
                        <span style={{ fontSize:9, color:'white', fontWeight:700, overflow:'hidden',
                          whiteSpace:'nowrap' }}>{pct>6?yr:''}</span>
                      </div>
                    ) : null;
                  })}
                  <div style={{ flex: (data.vacNLA/data.totalNLA)*100,
                    background:BKT['Vacant'].bg, minWidth:0 }}
                    title={`Vacant: ${fmtNum(data.vacNLA,0)} sqm`}/>
                </div>
              </>
            )}
          </Card>
        ))}
      </div>

      {/* Portfolio expiry table */}
      <Card title="Combined Expiry Profile — Both Assets" accent={C.navy}>
        <table>
          <thead><tr>
            <th>Expiry Year</th>
            <th style={{background:'#2c5e9e'}}>Castlereagh — Income</th>
            <th style={{background:'#2c5e9e'}}>NLA (sqm)</th>
            <th style={{background:'#1a5c5c'}}># Leases</th>
            <th style={{background:'#1a5c5c'}}>Elizabeth — Income</th>
            <th style={{background:'#1a5c5c'}}>NLA (sqm)</th>
            <th style={{background:'#1a5c5c'}}># Leases</th>
            <th>Portfolio Total Inc</th>
            <th>Portfolio %</th>
          </tr></thead>
          <tbody>
            {['2026','2027','2028','2029','2030','2031+','Vacant'].map(yr => {
              const csS = cs?.buckets[yr]||[], epS = ep?.buckets[yr]||[];
              const csInc = csS.reduce((s,x)=>s+x.grossPA,0);
              const csNLA = csS.reduce((s,x)=>s+x.nla,0);
              const epInc = epS.reduce((s,x)=>s+x.grossPA,0);
              const epNLA = epS.reduce((s,x)=>s+x.nla,0);
              const totInc = csInc + epInc;
              return (
                <tr key={yr}>
                  <td><span style={{ background:BKT[yr].bg, color:'white',
                    padding:'2px 8px', borderRadius:4, fontSize:11, fontWeight:700 }}>{yr}</span></td>
                  <td style={{ textAlign:'right' }}>{fmtCcy(csInc)}</td>
                  <td style={{ textAlign:'right' }}>{fmtNum(csNLA,0)}</td>
                  <td style={{ textAlign:'center' }}>{csS.length}</td>
                  <td style={{ textAlign:'right' }}>{fmtCcy(epInc)}</td>
                  <td style={{ textAlign:'right' }}>{fmtNum(epNLA,0)}</td>
                  <td style={{ textAlign:'center' }}>{epS.length}</td>
                  <td style={{ textAlign:'right', fontWeight:700 }}>{fmtCcy(totInc)}</td>
                  <td style={{ textAlign:'right' }}>{portfolioInc>0?fmtPct(totInc/portfolioInc):'—'}</td>
                </tr>
              );
            })}
            <tr style={{ fontWeight:700, background:C.navy }}>
              {['TOTAL','','','','','','',fmtCcy(portfolioInc),'100%'].map((v,i)=>(
                <td key={i} style={{ color:'white', textAlign:i>=7?'right':'left' }}>{v}</td>
              ))}
            </tr>
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ── Building Dashboard ───────────────────────────────────────────
function BuildingDashboard({ data, cfg }) {
  const [tab, setTab] = useState('overview');
  const subTabs = [
    {id:'overview',  label:'Overview'},
    {id:'leases',    label:'Leases'},
    {id:'expiry',    label:'Expiry Profile'},
    {id:'market',    label:'Market Rents'},
    {id:'floors',    label:'By Floor'},
    {id:'critical',  label:'Critical Dates'},
    {id:'valuation', label:'Valuation'},
  ];

  const accent = cfg.id==='castlereagh' ? C.darkBlue : '#1F6B75';
  const totalInc = data.active.reduce((s,x)=>s+x.grossPA,0);

  return (
    <div>
      {/* Building KPI bar */}
      <div style={{ background:accent, margin:'-24px -24px 0', padding:'12px 24px' }}>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(6,1fr)', gap:10, maxWidth:1352 }}>
          {[
            ['TOTAL NLA', `${fmtNum(data.totalNLA,0)} sqm`],
            ['OCCUPANCY', fmtPct(data.occ)],
            ['VACANT NLA', `${fmtNum(data.vacNLA,0)} sqm`],
            ['GROSS INCOME PA', fmtCcy(data.grossInc)],
            ['WALE (NLA)', `${data.waleNLA.toFixed(2)} yrs`],
            ['WALE (Income)', `${data.waleInc.toFixed(2)} yrs`],
          ].map(([l,v]) => (
            <div key={l} style={{ background:'rgba(255,255,255,0.1)', borderRadius:6, padding:'10px 12px' }}>
              <div style={{ color:'rgba(255,255,255,0.6)', fontSize:10, fontWeight:600, marginBottom:3 }}>{l}</div>
              <div style={{ color:'white', fontSize:16, fontWeight:700 }}>{v}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Sub-tab nav */}
      <div style={{ background:C.white, borderBottom:`1px solid ${C.brdr}`,
        margin:'0 -24px', padding:'0 24px', marginBottom:20 }}>
        <div style={{ display:'flex', gap:0 }}>
          {subTabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)}
              style={{ padding:'12px 16px', border:'none', background:'transparent',
                fontSize:12, fontWeight:600, cursor:'pointer', whiteSpace:'nowrap',
                color: tab===t.id ? accent : C.grey,
                borderBottom: tab===t.id ? `2px solid ${accent}` : '2px solid transparent',
                transition:'all 0.15s' }}>
              {t.label}
            </button>
          ))}
        </div>
      </div>

      {tab==='overview'  && <OverviewSection data={data} totalInc={totalInc} accent={accent} />}
      {tab==='leases'    && <LeasesSection data={data} />}
      {tab==='expiry'    && <ExpirySection data={data} totalInc={totalInc} />}
      {tab==='market'    && <MarketSection data={data} />}
      {tab==='floors'    && <FloorsSection data={data} totalInc={totalInc} />}
      {tab==='critical'  && <CriticalSection data={data} />}
      {tab==='valuation' && <ValuationSection data={data} />}
    </div>
  );
}

// ── Overview ─────────────────────────────────────────────────────
function OverviewSection({ data, totalInc, accent }) {
  return (
    <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:20 }}>
      <Card title="Lease Expiry Profile" accent={accent}>
        <table>
          <thead><tr><th>Year</th><th>Gross Inc PA</th><th>Inc %</th><th>NLA</th><th>Area %</th><th>#</th></tr></thead>
          <tbody>
            {Object.entries(BKT).map(([yr,{bg}]) => {
              const s = data.buckets[yr]||[];
              const inc = s.reduce((t,x)=>t+x.grossPA,0);
              const nla = s.reduce((t,x)=>t+x.nla,0);
              return (
                <tr key={yr}>
                  <td><span style={{background:bg,color:'white',padding:'1px 6px',borderRadius:3,fontSize:11,fontWeight:700}}>{yr}</span></td>
                  <td style={{textAlign:'right',fontWeight:600}}>{fmtCcy(inc)}</td>
                  <td style={{textAlign:'right'}}>{totalInc>0?fmtPct(inc/totalInc):'—'}</td>
                  <td style={{textAlign:'right'}}>{fmtNum(nla,0)}</td>
                  <td style={{textAlign:'right'}}>{data.totalNLA>0?fmtPct(nla/data.totalNLA):'—'}</td>
                  <td style={{textAlign:'center'}}>{s.length}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>

      <Card title="Top Tenants by Income" accent={accent}>
        <table>
          <thead><tr><th>Tenant</th><th>Suite</th><th>Floor</th><th>Income PA</th><th>%</th></tr></thead>
          <tbody>
            {[...data.active].sort((a,b)=>b.grossPA-a.grossPA).slice(0,10).map(s => (
              <tr key={s.r}>
                <td style={{fontWeight:600}}>{s.dispName}</td>
                <td>{s.suiteNum}</td>
                <td>{s.floor}</td>
                <td style={{textAlign:'right',fontWeight:600}}>{fmtCcy(s.grossPA)}</td>
                <td style={{textAlign:'right'}}>{totalInc>0?fmtPct(s.grossPA/totalInc):'—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>

      <Card title="WALE Summary" accent={accent}>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12,padding:'4px 0'}}>
          {[
            ['WALE by NLA',`${data.waleNLA.toFixed(2)} years`,'Area-weighted remaining term'],
            ['WALE by Income',`${data.waleInc.toFixed(2)} years`,'Income-weighted remaining term'],
            ['Occupied NLA',`${fmtNum(data.occNLA,0)} sqm`,fmtPct(data.occ)+' occupancy'],
            ['Vacant NLA',`${fmtNum(data.vacNLA,0)} sqm`,fmtPct(data.vacNLA/data.totalNLA)+' vacancy'],
          ].map(([l,v,s])=>(
            <div key={l} style={{background:C.vlBlue,borderRadius:6,padding:'12px',borderLeft:`3px solid ${accent}`}}>
              <div style={{fontSize:10,color:C.grey,marginBottom:3}}>{l}</div>
              <div style={{fontSize:17,fontWeight:700,color:C.navy,marginBottom:2}}>{v}</div>
              <div style={{fontSize:11,color:C.grey}}>{s}</div>
            </div>
          ))}
        </div>
      </Card>

      <Card title="Upcoming Critical Dates (Next 6 Months)" accent={accent}>
        {data.criticalDates.filter(d=>d.days<=180).length===0 ? (
          <p style={{color:C.grey,padding:12,textAlign:'center',fontSize:13}}>No events in next 6 months</p>
        ) : (
          <table>
            <thead><tr><th>Suite</th><th>Tenant</th><th>Event</th><th>Date</th><th>Days</th></tr></thead>
            <tbody>
              {data.criticalDates.filter(d=>d.days<=180).slice(0,8).map((d,i)=>(
                <tr key={i}>
                  <td style={{fontWeight:700}}>{d.suite}</td>
                  <td>{d.tenant}</td>
                  <td>{d.event}</td>
                  <td>{fmtDate(d.date)}</td>
                  <td style={{textAlign:'center',fontWeight:700,
                    color:d.days<=60?C.red:d.days<=120?C.orange:C.grey}}>
                    {d.days}d
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

// ── Leases ───────────────────────────────────────────────────────
function LeasesSection({ data }) {
  const [sortCol, setSortCol] = useState('expiry');
  const [sortDir, setSortDir] = useState(1);
  const sorted = [...data.suites].sort((a,b) => {
    const av=a[sortCol]??'', bv=b[sortCol]??'';
    return av<bv?-sortDir:av>bv?sortDir:0;
  });
  function Th({col,label}) {
    const on=sortCol===col;
    return <th style={{cursor:'pointer',background:on?C.navy:C.darkBlue}}
      onClick={()=>{setSortCol(col);setSortDir(s=>s.col===col?-s.dir:1);}}>
      {label}{on?(sortDir===1?'↑':'↓'):''}
    </th>;
  }
  return (
    <Card title={`Lease Register — ${data.suites.length} suites`}>
      <div style={{overflowX:'auto'}}>
        <table>
          <thead><tr>
            <Th col="floor" label="Floor"/><Th col="suiteNum" label="Suite"/>
            <Th col="dispName" label="Tenant"/><Th col="nla" label="NLA"/>
            <Th col="psm" label="Face PSM"/><Th col="grossPA" label="Gross PA"/>
            <Th col="expiry" label="Expiry"/><Th col="remYrs" label="Remaining"/>
          </tr></thead>
          <tbody>
            {sorted.map(s=>{
              const bg=s.expired||s.vacant?'#fff0f0':s.active&&s.remYrs<1?'#fff8f0':'inherit';
              return <tr key={s.r} style={{background:bg}}>
                <td>{s.floor}</td>
                <td style={{fontWeight:700}}>{s.suiteNum}</td>
                <td style={{fontWeight:600,color:s.vacant||s.expired?C.red:'inherit'}}>{s.dispName}</td>
                <td style={{textAlign:'right'}}>{fmtNum(s.nla,0)}</td>
                <td style={{textAlign:'right'}}>{s.active?fmtPSM(s.psm):'—'}</td>
                <td style={{textAlign:'right',fontWeight:600}}>{s.active?fmtCcy(s.grossPA):'—'}</td>
                <td style={{color:s.expired?C.red:s.active&&s.remYrs<1?C.orange:'inherit'}}>
                  {s.expiry?fmtDate(s.expiry):s.vacant?'VACANT':'—'}
                </td>
                <td style={{textAlign:'right',color:s.remYrs<1?C.red:s.remYrs<2?C.orange:'inherit'}}>
                  {s.active?`${s.remYrs.toFixed(2)} yrs`:'—'}
                </td>
              </tr>;
            })}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

// ── Expiry Profile ───────────────────────────────────────────────
function ExpirySection({ data, totalInc }) {
  return (
    <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:16}}>
      {Object.entries(BKT).map(([yr,{bg,lt}])=>{
        const suites=data.buckets[yr]||[];
        const inc=suites.reduce((s,x)=>s+x.grossPA,0);
        const nla=suites.reduce((s,x)=>s+x.nla,0);
        return (
          <div key={yr} style={{background:C.white,borderRadius:8,overflow:'hidden',
            boxShadow:'0 2px 8px rgba(0,0,0,0.08)'}}>
            <div style={{background:bg,color:'white',padding:'10px 14px',
              display:'flex',justifyContent:'space-between',alignItems:'center'}}>
              <span style={{fontSize:15,fontWeight:700}}>{yr}</span>
              <span style={{fontSize:12}}>{suites.length} suite{suites.length!==1?'s':''}</span>
            </div>
            <div style={{padding:'10px 14px',background:lt,display:'grid',
              gridTemplateColumns:'1fr 1fr',gap:8,fontSize:12}}>
              <div><div style={{color:C.grey,fontSize:10}}>Gross Income PA</div><div style={{fontWeight:700}}>{fmtCcy(inc)}</div></div>
              <div><div style={{color:C.grey,fontSize:10}}>% of Portfolio</div><div style={{fontWeight:700}}>{totalInc>0?fmtPct(inc/totalInc):'—'}</div></div>
              <div><div style={{color:C.grey,fontSize:10}}>NLA (sqm)</div><div style={{fontWeight:700}}>{fmtNum(nla,0)}</div></div>
              <div><div style={{color:C.grey,fontSize:10}}>Area %</div><div style={{fontWeight:700}}>{data.totalNLA>0?fmtPct(nla/data.totalNLA):'—'}</div></div>
            </div>
            {suites.map(s=>(
              <div key={s.r} style={{display:'flex',justifyContent:'space-between',
                padding:'5px 14px',borderTop:`1px solid ${lt}`,fontSize:11}}>
                <span style={{fontWeight:600}}>{s.dispName}</span>
                <span style={{color:C.grey}}>{s.suiteNum} · {s.expiry?fmtDate(s.expiry):'Vacant'}</span>
              </div>
            ))}
          </div>
        );
      })}
    </div>
  );
}

// ── Market Rents ─────────────────────────────────────────────────
function MarketSection({ data }) {
  const suites = data.active.filter(s=>s.psm>0);
  return (
    <Card title="Passing Rent vs Market Rent">
      <table>
        <thead><tr><th>Suite</th><th>Tenant</th><th>Floor</th><th>NLA</th>
          <th>Face PSM</th><th>Market PSM</th><th>Var ($)</th><th>Var (%)</th><th>Position</th></tr></thead>
        <tbody>
          {suites.map(s=>{
            const m=s.mktPSM||0, f=s.psm;
            const va=m>0?f-m:null, vp=m>0?(f-m)/m:null;
            const pos=vp==null?'—':vp>0.05?'▲ Above Mkt':vp<-0.05?'▼ Below Mkt':'≈ At Market';
            const bg=vp==null?'inherit':vp>0.05?'#fff0f0':vp<-0.05?'#f0fff4':C.vlBlue;
            return <tr key={s.r} style={{background:bg}}>
              <td style={{fontWeight:700}}>{s.suiteNum}</td>
              <td>{s.dispName}</td><td>{s.floor}</td>
              <td style={{textAlign:'right'}}>{fmtNum(s.nla,0)}</td>
              <td style={{textAlign:'right'}}>{fmtPSM(f)}</td>
              <td style={{textAlign:'right'}}>{m>0?fmtPSM(m):'—'}</td>
              <td style={{textAlign:'right',color:va==null?'inherit':va>0?C.red:C.green}}>{va!=null?fmtPSM(va):'—'}</td>
              <td style={{textAlign:'right',color:vp==null?'inherit':vp>0?C.red:C.green}}>{vp!=null?fmtPct(vp):'—'}</td>
              <td style={{fontWeight:600,color:pos.includes('Above')?C.red:pos.includes('Below')?C.green:C.grey}}>{pos}</td>
            </tr>;
          })}
        </tbody>
      </table>
    </Card>
  );
}

// ── By Floor ─────────────────────────────────────────────────────
function FloorsSection({ data, totalInc }) {
  return (
    <Card title="Income Analysis — By Floor">
      <table>
        <thead><tr><th>Floor</th><th>Active Tenants</th><th>Total NLA</th>
          <th>Gross Income PA</th><th>Income %</th><th>Earliest Expiry</th></tr></thead>
        <tbody>
          {data.byFloor.map(f=>(
            <tr key={f.floor}>
              <td style={{fontWeight:700}}>{f.floor}</td>
              <td style={{fontSize:11}}>{f.tenants.length>0?f.tenants.join(', '):
                <span style={{color:C.red}}>Vacant</span>}</td>
              <td style={{textAlign:'right'}}>{fmtNum(f.nla,0)}</td>
              <td style={{textAlign:'right',fontWeight:600}}>{f.income>0?fmtCcy(f.income):'—'}</td>
              <td style={{textAlign:'right'}}>{f.income>0&&totalInc>0?fmtPct(f.income/totalInc):'—'}</td>
              <td style={{color:f.minExpiry&&f.minExpiry<new Date(new Date().setMonth(new Date().getMonth()+12))?C.orange:'inherit'}}>
                {f.minExpiry?fmtDate(f.minExpiry):'—'}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </Card>
  );
}

// ── Critical Dates ───────────────────────────────────────────────
function CriticalSection({ data }) {
  return (
    <Card title="Critical Dates — Next 18 Months (Lease Expiries & Rent Reviews)">
      {data.criticalDates.length===0 ? (
        <p style={{color:C.grey,padding:16,textAlign:'center'}}>No events in next 18 months</p>
      ) : (
        <table>
          <thead><tr><th>Suite</th><th>Tenant</th><th>Event</th><th>Date</th><th>Days Away</th><th>Note</th></tr></thead>
          <tbody>
            {data.criticalDates.map((d,i)=>(
              <tr key={i}>
                <td style={{fontWeight:700}}>{d.suite}</td><td>{d.tenant}</td><td>{d.event}</td>
                <td>{fmtDate(d.date)}</td>
                <td style={{textAlign:'center',fontWeight:700,
                  color:d.days<0?C.red:d.days<=60?C.red:d.days<=120?C.orange:C.grey}}>
                  {d.days<0?`${Math.abs(d.days)}d overdue`:`${d.days}d`}
                </td>
                <td style={{fontSize:11,color:C.grey}}>{d.note}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </Card>
  );
}

// ── Valuation ─────────────────────────────────────────────────────
function ValuationSection({ data }) {
  const { valuation } = data;
  return (
    <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:20}}>
      <Card title="Property Valuation">
        {!valuation ? (
          <div style={{padding:24,textAlign:'center',color:C.grey}}>
            <div style={{fontSize:36,marginBottom:12}}>📋</div>
            <div style={{fontWeight:600,marginBottom:8}}>No valuation data entered yet</div>
            <div style={{fontSize:12}}>Enter data in the Valuations tab of your Excel file,
              then replace the file in GitHub.</div>
          </div>
        ) : (
          <div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:10,marginBottom:14}}>
              {[['Latest (Val 1)',valuation.v1],['Val 2',valuation.v2],['Val 3',valuation.v3]].map(([l,v],i)=>(
                <div key={i} style={{background:i===0?C.lightGreen:C.vlBlue,borderRadius:6,padding:'10px 12px',
                  border:`1px solid ${i===0?C.green:'#c0d8f0'}`}}>
                  <div style={{color:C.grey,fontSize:10,marginBottom:3}}>{l}</div>
                  <div style={{fontSize:16,fontWeight:700,color:C.navy}}>{v.amount?fmtCcy(v.amount):'—'}</div>
                  {v.date&&<div style={{fontSize:11,color:C.grey,marginTop:3}}>{fmtDate(v.date)}</div>}
                </div>
              ))}
            </div>
            {valuation.v1.valuer&&<div style={{fontSize:12,color:C.grey,marginBottom:6}}><b>Valuer:</b> {valuation.v1.valuer}</div>}
            {valuation.v1.capRate>0&&<div style={{fontSize:12,color:C.grey,marginBottom:6}}><b>Cap Rate:</b> {fmtPct(valuation.v1.capRate)}</div>}
            {valuation.movement!=null&&(
              <div style={{background:valuation.movement>=0?C.lightGreen:'#fff0f0',
                borderRadius:6,padding:'10px 14px',marginTop:10}}>
                <div style={{fontSize:11,color:C.grey,marginBottom:3}}>Movement Val 1 vs Val 2</div>
                <div style={{fontSize:15,fontWeight:700,color:valuation.movement>=0?C.green:C.red}}>
                  {valuation.movement>=0?'+':''}{fmtCcy(valuation.movement)} ({fmtPct(valuation.movPct)})
                </div>
              </div>
            )}
          </div>
        )}
      </Card>
      <Card title="How to Update">
        <div style={{fontSize:12,color:C.grey,lineHeight:1.9,padding:'4px 0'}}>
          <p>1. Open your Excel file and go to the <b>Valuations</b> tab</p>
          <p>2. Enter figures in the yellow input cells</p>
          <p>3. <b>Save the file in Excel</b> (caches formula results)</p>
          <p>4. Copy file to <code>public/data/</code> in your GitHub repo</p>
          <p>5. Commit and push in GitHub Desktop</p>
          <p>6. Dashboard updates in ~60 seconds</p>
        </div>
      </Card>
    </div>
  );
}

// ── Stack Plan ────────────────────────────────────────────────────
function StackPlan({ data, cfg }) {
  const [hover, setHover] = useState(null);
  const accent = cfg.id==='castlereagh' ? C.darkBlue : '#1F6B75';
  const totalInc = data.active.reduce((s,x)=>s+x.grossPA,0);

  return (
    <div>
      {/* Header stats */}
      <div style={{background:accent,margin:'-24px -24px 20px',padding:'12px 24px'}}>
        <div style={{display:'flex',gap:24,flexWrap:'wrap'}}>
          {[
            ['Total NLA',`${fmtNum(data.totalNLA,0)} sqm`],
            ['Occupancy',fmtPct(data.occ)],
            ['Gross Income',fmtCcy(data.grossInc)+' pa'],
            ['WALE (NLA)',`${data.waleNLA.toFixed(2)} yrs`],
            ['Active Leases',`${data.active.length} suites`],
            ['Vacant NLA',`${fmtNum(data.vacNLA,0)} sqm`],
          ].map(([l,v])=>(
            <div key={l} style={{borderLeft:`2px solid rgba(255,255,255,0.3)`,paddingLeft:12}}>
              <div style={{color:'rgba(255,255,255,0.6)',fontSize:10}}>{l}</div>
              <div style={{color:'white',fontSize:14,fontWeight:700}}>{v}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{display:'grid',gridTemplateColumns:'1fr 280px',gap:20,alignItems:'start'}}>
        {/* Stack */}
        <div>
          <div style={{background:C.white,borderRadius:8,overflow:'hidden',
            boxShadow:'0 2px 8px rgba(0,0,0,0.08)'}}>
            <div style={{background:C.navy,color:'white',padding:'10px 16px',
              fontSize:13,fontWeight:700}}>
              Building Stack — {cfg.name} | {cfg.suburb}
            </div>
            <div style={{padding:'12px 16px',display:'flex',flexDirection:'column',gap:4}}>
              {/* Roof cap */}
              <div style={{height:6,background:'#bbb',borderRadius:4,marginBottom:4}}/>

              {cfg.floorOrder.map(fl=>{
                const flSuites = data.suites.filter(s=>s.floor===fl);
                if(!flSuites.length) return null;
                const flNLA = flSuites.reduce((s,x)=>s+x.nla,0);

                return (
                  <div key={fl} style={{display:'flex',gap:4,alignItems:'stretch',minHeight:52}}>
                    {/* Floor label */}
                    <div style={{width:64,flexShrink:0,background:C.navy,color:'white',
                      borderRadius:4,display:'flex',flexDirection:'column',
                      alignItems:'center',justifyContent:'center',fontSize:11,fontWeight:700,padding:'2px 0'}}>
                      <span>{fl}</span>
                      <span style={{fontSize:9,opacity:0.7,marginTop:1}}>{fmtNum(flNLA,0)}</span>
                    </div>

                    {/* Suite blocks */}
                    <div style={{flex:1,display:'flex',gap:3,alignItems:'stretch'}}>
                      {flSuites.map(s=>{
                        const pct = flNLA>0 ? s.nla/flNLA : 1/flSuites.length;
                        const bk = !s.active ? 'Vacant' : bucketKey(s.expiry);
                        const bg = BKT[bk]?.bg || '#ccc';
                        const lt = BKT[bk]?.lt || '#eee';
                        const isHov = hover===`${fl}-${s.suiteNum}`;

                        return (
                          <div key={s.r}
                            onMouseEnter={()=>setHover(`${fl}-${s.suiteNum}`)}
                            onMouseLeave={()=>setHover(null)}
                            title={`${s.dispName}\n${s.suiteNum} | ${fmtNum(s.nla,0)} sqm | ${s.active?fmtPSM(s.psm)+' psm | '+fmtDate(s.expiry):s.vacant?'VACANT':'Expired'}`}
                            style={{
                              flex: pct,
                              background: isHov ? (s.active?lt:bg) : bg,
                              borderRadius:4, minWidth:20, overflow:'hidden',
                              cursor:'default', position:'relative',
                              border: isHov?`2px solid ${C.white}`:'2px solid transparent',
                              transition:'all 0.1s',
                              display:'flex',flexDirection:'column',
                              alignItems:'center',justifyContent:'center',
                              padding:'2px 4px',
                            }}>
                            <span style={{color:isHov?(s.active?C.navy:'white'):'white',
                              fontSize:9,fontWeight:700,textAlign:'center',
                              lineHeight:1.2,overflow:'hidden',
                              textOverflow:'ellipsis',whiteSpace:'nowrap',
                              maxWidth:'100%'}}>
                              {s.suiteNum}
                            </span>
                            {pct > 0.12 && (
                              <span style={{color:isHov?(s.active?C.grey:'rgba(255,255,255,0.8)'):'rgba(255,255,255,0.8)',
                                fontSize:8,textAlign:'center',overflow:'hidden',
                                textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:'100%'}}>
                                {s.dispName.split(' ').slice(0,2).join(' ')}
                              </span>
                            )}
                            {pct > 0.15 && (
                              <span style={{color:isHov?(s.active?C.grey:'rgba(255,255,255,0.7)'):'rgba(255,255,255,0.7)',
                                fontSize:8,textAlign:'center'}}>
                                {fmtNum(s.nla,0)} sqm
                              </span>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}

              {/* Base */}
              <div style={{height:6,background:'#bbb',borderRadius:4,marginTop:4}}/>
            </div>
          </div>

          {/* Legend */}
          <div style={{background:C.white,borderRadius:8,marginTop:12,padding:'12px 16px',
            boxShadow:'0 2px 8px rgba(0,0,0,0.08)'}}>
            <div style={{fontSize:11,fontWeight:700,color:C.navy,marginBottom:8}}>
              Legend — Expiry Year
            </div>
            <div style={{display:'flex',flexWrap:'wrap',gap:10}}>
              {Object.entries(BKT).map(([yr,{bg}])=>(
                <div key={yr} style={{display:'flex',alignItems:'center',gap:5}}>
                  <div style={{width:16,height:16,background:bg,borderRadius:3,flexShrink:0}}/>
                  <span style={{fontSize:11,color:C.grey}}>{yr}</span>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Suite detail panel */}
        <div>
          <div style={{background:C.white,borderRadius:8,boxShadow:'0 2px 8px rgba(0,0,0,0.08)',
            position:'sticky',top:12}}>
            <div style={{background:C.navy,color:'white',padding:'10px 14px',
              fontSize:12,fontWeight:700,borderRadius:'8px 8px 0 0'}}>
              Suite Directory
            </div>
            <div style={{maxHeight:600,overflowY:'auto'}}>
              {cfg.floorOrder.map(fl=>{
                const flSuites=data.suites.filter(s=>s.floor===fl);
                if(!flSuites.length) return null;
                return (
                  <div key={fl}>
                    <div style={{background:C.vlBlue,padding:'5px 12px',
                      fontSize:10,fontWeight:700,color:C.navy,
                      borderTop:`1px solid ${C.brdr}`}}>{fl}</div>
                    {flSuites.map(s=>{
                      const bk = !s.active?'Vacant':bucketKey(s.expiry);
                      return (
                        <div key={s.r} style={{padding:'6px 12px',
                          borderTop:`1px solid #f0f4f8`,fontSize:11,
                          background:hover===`${fl}-${s.suiteNum}`?C.lightBlue:'white'}}>
                          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:4}}>
                            <span style={{fontWeight:700,color:C.navy}}>{s.suiteNum}</span>
                            <span style={{background:BKT[bk].bg,color:'white',
                              fontSize:9,padding:'1px 5px',borderRadius:3,flexShrink:0}}>{bk}</span>
                          </div>
                          <div style={{color:s.active?C.grey:C.red,fontSize:10,marginTop:1}}>{s.dispName}</div>
                          <div style={{color:'#aaa',fontSize:10}}>
                            {s.active?`${fmtNum(s.nla,0)} sqm · ${fmtPSM(s.psm)} · ${fmtDate(s.expiry)}`:
                             s.vacant?`${fmtNum(s.nla,0)} sqm · VACANT`:`${fmtNum(s.nla,0)} sqm · EXPIRED`}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Card ─────────────────────────────────────────────────────────
function Card({ title, children, accent=C.navy }) {
  return (
    <div style={{background:C.white,borderRadius:8,overflow:'hidden',
      boxShadow:'0 2px 8px rgba(0,0,0,0.08)'}}>
      <div style={{background:accent,color:'white',padding:'10px 16px',fontSize:13,fontWeight:700}}>
        {title}
      </div>
      <div style={{padding:16,overflowX:'auto'}}>{children}</div>
    </div>
  );
}

// ── Loading / Error ──────────────────────────────────────────────
function Spinner() {
  return (
    <div style={{display:'flex',flexDirection:'column',alignItems:'center',
      justifyContent:'center',padding:80,gap:16}}>
      <div style={{width:40,height:40,border:`4px solid ${C.lightBlue}`,
        borderTop:`4px solid ${C.darkBlue}`,borderRadius:'50%',
        animation:'spin 0.8s linear infinite'}}/>
      <div style={{color:C.grey,fontSize:13}}>Loading workbook data...</div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────
export default function App() {
  const [auth,    setAuth]    = useState(false);
  const [view,    setView]    = useState('portfolio');
  const [bldData, setBldData] = useState({ castlereagh:null, elizabeth:null });
  const [errors,  setErrors]  = useState({});
  const [loaded,  setLoaded]  = useState({ castlereagh:false, elizabeth:false });

  useEffect(() => {
    if (typeof window !== 'undefined' && sessionStorage.getItem('dash_auth')==='1')
      setAuth(true);
  }, []);

  const loadBuilding = useCallback(async (key) => {
    const cfg = BUILDINGS[key];
    try {
      const xlsxMod = await import('xlsx');
      const XLSX = xlsxMod.default || xlsxMod;
      const res = await fetch(cfg.file+'?t='+Date.now());
      if (!res.ok) throw new Error(
        `Could not load ${cfg.file} (${res.status}). ` +
        `Place the file in public/data/ in your GitHub repo.`
      );
      const buf = await res.arrayBuffer();
      const wb  = XLSX.read(buf, { type:'array', cellDates:true, cellFormula:false });
      const data = parseBuilding(wb, cfg);
      if (!data) throw new Error(`Input Data sheet not found in ${cfg.name} workbook.`);
      setBldData(prev => ({ ...prev, [key]:data }));
      setLoaded(prev  => ({ ...prev, [key]:true }));
    } catch(e) {
      console.error(key, e);
      setErrors(prev => ({ ...prev, [key]:e.message }));
      setLoaded(prev  => ({ ...prev, [key]:true }));
    }
  }, []);

  useEffect(() => {
    if (!auth) return;
    loadBuilding('castlereagh');
    loadBuilding('elizabeth');
  }, [auth, loadBuilding]);

  if (!auth) return <PinPage onSuccess={() => setAuth(true)} />;

  const cs = bldData.castlereagh;
  const ep = bldData.elizabeth;

  return (
    <div style={{ minHeight:'100vh' }}>
      <TopNav view={view} setView={setView} loaded={loaded} />
      <div style={{ maxWidth:1400, margin:'0 auto', padding:24 }}>

        {view==='portfolio' && (
          <PortfolioView buildings={{ castlereagh:cs, elizabeth:ep }} />
        )}

        {view==='cs_dashboard' && (
          errors.castlereagh ? (
            <Card title="Error loading 17 Castlereagh">
              <div style={{padding:16,color:C.red}}>{errors.castlereagh}</div>
            </Card>
          ) : !cs ? <Spinner /> : (
            <BuildingDashboard data={cs} cfg={BUILDINGS.castlereagh} />
          )
        )}

        {view==='cs_stack' && (
          errors.castlereagh ? (
            <Card title="Error loading 17 Castlereagh">
              <div style={{padding:16,color:C.red}}>{errors.castlereagh}</div>
            </Card>
          ) : !cs ? <Spinner /> : (
            <StackPlan data={cs} cfg={BUILDINGS.castlereagh} />
          )
        )}

        {view==='ep_dashboard' && (
          errors.elizabeth ? (
            <Card title="Error loading 1 Elizabeth Plaza">
              <div style={{padding:16,color:C.red}}>{errors.elizabeth}</div>
            </Card>
          ) : !ep ? <Spinner /> : (
            <BuildingDashboard data={ep} cfg={BUILDINGS.elizabeth} />
          )
        )}

        {view==='ep_stack' && (
          errors.elizabeth ? (
            <Card title="Error loading 1 Elizabeth Plaza">
              <div style={{padding:16,color:C.red}}>{errors.elizabeth}</div>
            </Card>
          ) : !ep ? <Spinner /> : (
            <StackPlan data={ep} cfg={BUILDINGS.elizabeth} />
          )
        )}
      </div>
    </div>
  );
}
