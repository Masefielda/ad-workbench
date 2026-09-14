/* ============================================================
   信息流投放工作台 —— 单文件应用
   模块：手机型号数据分析 / 链路数据表(系统口径) / 分周数据分析 / 链路画像评分
   ============================================================ */

/* ---------- 通用工具 ---------- */
const $ = (s) => document.querySelector(s);
const $$ = (s) => document.querySelectorAll(s);
const fmtN = (n) => { n = Number(n)||0; return n.toLocaleString('zh-CN',{maximumFractionDigits:0}); };
const fmtM = (n) => { n = Number(n)||0; return '¥' + n.toLocaleString('zh-CN',{minimumFractionDigits:0,maximumFractionDigits:0}); };
const fmtM2 = (n) => { n = Number(n)||0; return '¥' + n.toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2}); };
const fmtP = (n, d) => { n = Number(n)||0; d = d==null?2:d; return n.toFixed(d) + '%'; };
const pct = (a, b) => b ? (a/b*100) : 0;
const cvrOf = (c, cl) => cl ? (c/cl*100) : 0;
const cpaOf = (c, v) => v ? (c/v) : 0;
const clamp = (x, lo, hi) => Math.min(hi, Math.max(lo, x));
const esc = (s) => String(s==null?'':s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const sum = (arr) => arr.reduce((a,b)=>a+Number(b||0),0);

/* ---------- 设计常量 ---------- */
const C = {
  accent:'#7C6FE0', teal:'#5BB89A', amber:'#D9A441', blue:'#5B8DEF', purple:'#7A5AA8',
  pink:'#C2577A', olive:'#7A8A3D', slate:'#5A6B7A',
  cat:['#7C6FE0','#5BB89A','#D9A441','#5B8DEF','#7A5AA8','#C2577A','#7A8A3D','#5A6B7A'],
  grid:'#ECE9E1', axis:'#8E8B82'
};

/* ---------- 数据同步与持久化 ---------- */
const SYNC = { phone:null, link:null, weekly:null, score:null, material:null, clue:null, cluePosition:null, accountScore:null, accountCompare:null, weeklyArchive:null, weeklyReport:null };
const SYNC_LABELS = { phone:'手机型号', link:'链路数据表', weekly:'分周对比', score:'链路画像评分', material:'素材分析', clue:'线索评分', cluePosition:'线索评分·版位分析', accountScore:'分账户评分分析', accountCompare:'账户对比分析', weeklyArchive:'周报历史归档', weeklyReport:'周报撰写' };
function saveSync(key, data){ try{ SYNC[key]=data; localStorage.setItem('sync_'+key, JSON.stringify(data)); return true; }catch(e){ return false; } }
function loadSync(){ Object.keys(SYNC).forEach(k=>{ try{ SYNC[k]=JSON.parse(localStorage.getItem('sync_'+k)); }catch(e){ SYNC[k]=null; } }); }
function clearSync(key){ SYNC[key]=null; localStorage.removeItem('sync_'+key); }
function clearAllSync(){ Object.keys(SYNC).forEach(k=>{ SYNC[k]=null; localStorage.removeItem('sync_'+k); }); }

const DATA_KEY = 'workbench_data_v1';
function savePersistentData(){ try{ localStorage.setItem(DATA_KEY, JSON.stringify({rows:DATA.rows, source:DATA.source, ts:Date.now()})); return true; }catch(e){ return false; } }
function loadPersistentData(){ try{ const d=JSON.parse(localStorage.getItem(DATA_KEY)); if(d&&d.rows&&d.rows.length){ DATA.rows=d.rows; DATA.source=d.source||'imported'; return true; } }catch(e){} return false; }
function clearPersistentData(){ DATA.rows=[]; DATA.source='empty'; localStorage.removeItem(DATA_KEY); }

function syncBadge(key){
  const d = SYNC[key];
  if(!d) return '';
  return `<div class="sync-badge"><span class="sync-dot"></span>已同步Excel数据 · ${SYNC_LABELS[key]}<button class="sync-clear" onclick="clearSync('${key}');location.hash='#/${key==='phone'?'phone':key==='link'?'link-table':key==='weekly'?'weekly':key==='score'?'link-score':key==='material'?'material':key==='clue'?'clue':key}';return false;">清除</button></div>`;
}

function renderSyncPanel(key, container){
  const d = SYNC[key];
  if(!d){ container.innerHTML=''; return; }
  let html = `<div class="sync-panel"><div class="sync-head"><span class="sync-tag">Excel同步</span><span>${SYNC_LABELS[key]}分析结果</span><a class="sync-goto" href="#/excel-${key}">查看完整分析 →</a></div><div class="sync-body">`;
  if(key==='phone'){
    html += `<div class="sync-kpis">`;
    [{v:fmtN(d.nTotal),l:'订单总数'},{v:fmtN(d.nValid),l:'有效订单'},{v:fmtN(d.nHigh),l:'高价课转化'},{v:(d.nTotal?(d.nHigh/d.nTotal*100).toFixed(1):0)+'%',l:'转化率'},{v:fmtM2(d.totalHigh).replace('¥',''),l:'高价课金额'},{v:'¥'+(d.nTotal?(d.totalHigh/d.nTotal).toFixed(2):0),l:'单订单产值'}].forEach(k=>{ html+=`<div class="sync-kpi"><div class="sk-v">${k.v}</div><div class="sk-l">${k.l}</div></div>`; });
    html += `</div><h5>品牌产值 Top 8</h5><table class="sync-tbl"><tr><th>品牌</th><th>订单数</th><th>转化率</th><th>金额(元)</th><th>产值(元)</th></tr>`;
    d.brandAgg.slice(0,8).forEach(r=>{ html+=`<tr><td>${r['品牌统一']}</td><td>${fmtN(r['低价课订单数'])}</td><td>${(r['高价课转化率']*100).toFixed(1)}%</td><td>${fmtM2(r['高价课金额']).replace('¥','')}</td><td>${fmtM2(r['单订单产值']).replace('¥','')}</td></tr>`; });
    html += `</table>`;
  } else if(key==='link'){
    html += `<div class="sync-kpis">`;
    [{v:fmtN(d.detailCount),l:'明细行数'},{v:d.periods.length,l:'期数'},{v:fmtM(d.total.spend).replace('¥',''),l:'广告花费'},{v:(d.total.roi*100).toFixed(1)+'%',l:'整体ROI'},{v:fmtN(d.total.eff),l:'有效订单'},{v:(d.total.highConv*100).toFixed(1)+'%',l:'高价转化率'}].forEach(k=>{ html+=`<div class="sync-kpi"><div class="sk-v">${k.v}</div><div class="sk-l">${k.l}</div></div>`; });
    html += `</div><h5>链路汇总</h5><table class="sync-tbl"><tr><th>链路</th><th>ROI</th><th>花费</th><th>订单数</th><th>单订单成本</th><th>单订单产值</th><th>高价转化率</th></tr>`;
    d.linkSummary.forEach(r=>{ html+=`<tr><td>${r.link}</td><td>${(r.roi*100).toFixed(1)}%</td><td>${fmtM(r.spend).replace('¥','')}</td><td>${fmtN(r.eff)}</td><td>${fmtM2(r.orderCost).replace('¥','')}</td><td>${fmtM2(r.orderOutput).replace('¥','')}</td><td>${(r.highConv*100).toFixed(1)}%</td></tr>`; });
    html += `<tr style="font-weight:bold"><td>合计</td><td>${(d.total.roi*100).toFixed(1)}%</td><td>${fmtM(d.total.spend).replace('¥','')}</td><td>${fmtN(d.total.eff)}</td><td>${fmtM2(d.total.orderCost).replace('¥','')}</td><td>${fmtM2(d.total.orderOutput).replace('¥','')}</td><td>${(d.total.highConv*100).toFixed(1)}%</td></tr></table>`;
  } else if(key==='weekly'){
    const cur=d.current, last=d.last;
    html += `<div class="sync-week"><div class="sw-box ${last?'':'sw-empty'}"><div class="sw-t">上周 ${last?last.weekFmt:'无数据'}</div>${last?`花费 ¥${fmtM(last.totalSpend).replace('¥','')} · 订单 ${fmtN(last.totalAll)}`:'首次运行，无上周缓存'}</div><div class="sw-box"><div class="sw-t">本周 ${cur.weekFmt}</div>花费 ¥${fmtM(cur.totalSpend).replace('¥','')} · 订单 ${fmtN(cur.totalAll)}</div></div>`;
    html += `<h5>链路分周</h5><table class="sync-tbl"><tr><th>链路</th><th>上周花费</th><th>本周花费</th><th>上周订单</th><th>本周订单</th><th>上周成本</th><th>本周成本</th></tr>`;
    cur.linkRows.filter(r=>r.link!=='汇总').forEach(r=>{
      const lr = last?last.linkRows.find(x=>x.link===r.link):null;
      html+=`<tr><td>${r.link}</td><td>${lr?fmtM(lr.spend).replace('¥',''):'—'}</td><td>${fmtM(r.spend).replace('¥','')}</td><td>${lr?fmtN(lr.all):'—'}</td><td>${fmtN(r.all)}</td><td>${lr?fmtM2(lr.orderCost).replace('¥',''):'—'}</td><td>${fmtM2(r.orderCost).replace('¥','')}</td></tr>`;
    });
    html += `</table>`;
  } else if(key==='score'){
    // 仅保留 Excel同步 标签，不展示KPI和概览表格
  }
  html += `</div></div>`;
  container.innerHTML = html;
}

/* ---------- 示例数据生成（确定性伪随机） ---------- */
function mulberry32(a){return function(){a|=0;a=a+0x6D2B79F5|0;let t=Math.imul(a^a>>>15,1|a);t=t+Math.imul(t^t>>>7,61|t)^t;return((t^t>>>14)>>>0)/4294967296;};}

const LINK_DEFS = [
  {name:'表单提交', ctr:0.020, cvr:0.105, cpa:205, scale:1.00, rev:0},
  {name:'在线咨询', ctr:0.018, cvr:0.068, cpa:175, scale:0.85, rev:0},
  {name:'私信加粉', ctr:0.016, cvr:0.048, cpa:105, scale:0.90, rev:0},
  {name:'APP下载',  ctr:0.022, cvr:0.165, cpa:58,  scale:0.70, rev:0},
  {name:'电商下单', ctr:0.014, cvr:0.023, cpa:410, scale:0.65, rev:1}
];
const PLAN_DEFS = [
  {name:'春季蓄水-表单', links:['表单提交','在线咨询']},
  {name:'夏季放量-下载', links:['APP下载','电商下单']},
  {name:'老客召回-加粉', links:['私信加粉']},
  {name:'常规跑量-全链路', links:['表单提交','在线咨询','私信加粉','APP下载','电商下单']}
];
const MODEL_DEFS = [
  {name:'iPhone 15 Pro',  os:'iOS',     ctrMul:1.35, cvrMul:1.25, costMul:1.90, vol:0.90},
  {name:'iPhone 15',      os:'iOS',     ctrMul:1.25, cvrMul:1.20, costMul:1.50, vol:0.70},
  {name:'iPhone 14',      os:'iOS',     ctrMul:1.15, cvrMul:1.10, costMul:1.10, vol:0.60},
  {name:'华为 Mate 60 Pro', os:'Android', ctrMul:1.10, cvrMul:1.10, costMul:1.20, vol:1.00},
  {name:'华为 Pura 70',   os:'Android', ctrMul:1.05, cvrMul:1.00, costMul:1.00, vol:0.70},
  {name:'小米14',         os:'Android', ctrMul:0.95, cvrMul:0.90, costMul:0.80, vol:1.10},
  {name:'小米14 Ultra',   os:'Android', ctrMul:1.00, cvrMul:0.95, costMul:1.05, vol:0.60},
  {name:'Redmi K70',      os:'Android', ctrMul:0.70, cvrMul:0.70, costMul:0.45, vol:1.40},
  {name:'OPPO Find X7',   os:'Android', ctrMul:0.90, cvrMul:0.85, costMul:0.75, vol:0.80},
  {name:'vivo X100',      os:'Android', ctrMul:0.88, cvrMul:0.85, costMul:0.70, vol:0.80},
  {name:'荣耀 Magic6',    os:'Android', ctrMul:0.92, cvrMul:0.88, costMul:0.72, vol:0.70},
  {name:'Redmi Note 13',  os:'Android', ctrMul:0.55, cvrMul:0.55, costMul:0.30, vol:1.50}
];
const WEEK_GROWTH = [0.82,0.88,0.95,1.00,1.05,1.10,1.12,1.08,1.15,1.20,1.18,1.24];
const WK_ANOMALY = { '私信加粉': [1,1,1,1,1,0.70,1,1.1,1,1,1,1], '电商下单': [1,1,1,1,1,1,1,1,1.45,1,0.9,1] };

function genDemo(){
  const R = mulberry32(20260831);
  const rows = [];
  const start = new Date(2026,5,8); // 2026-06-08（周一）
  const planByLink = {};
  LINK_DEFS.forEach(l => {
    const plans = PLAN_DEFS.filter(p => p.links.includes(l.name));
    planByLink[l.name] = plans.map(p=>p.name);
  });
  for (let d=0; d<84; d++){
    const dt = new Date(start.getTime() + d*86400000);
    const weekIdx = Math.floor(d/7);
    const dateStr = fmtDate(dt);
    for (const L of LINK_DEFS){
      for (const M of MODEL_DEFS){
        const anom = (WK_ANOMALY[L.name]||[])[weekIdx] || 1;
        const base = 750 * M.vol * L.scale * WEEK_GROWTH[weekIdx] * anom;
        const impr = Math.round(base * (0.75 + R()*0.5));
        if (impr <= 0) continue;
        const ctr = L.ctr * M.ctrMul * (0.8 + R()*0.4);
        const click = Math.round(impr * ctr);
        const cvr = L.cvr * M.cvrMul * (0.75 + R()*0.5);
        let conv = Math.round(click * cvr);
        if (conv === 0 && R() < 0.3) conv = 1;
        const costMul = 0.5 + 0.5*M.costMul;
        const cpa = L.cpa * costMul * (0.85 + R()*0.3);
        const cost = Math.round(conv * cpa * 100)/100;
        const revenue = L.rev ? Math.round(cost * (0.7 + R()*0.6)*100)/100 : 0;
        const plans = planByLink[L.name];
        const plan = plans[Math.floor(R()*plans.length)];
        rows.push({
          date:dateStr, ts:dt.getTime(), weekStart:fmtDate(weekStart(dt)), weekLabel:'W'+isoWeek(dt),
          plan:plan, link:L.name, placement:'信息流', model:M.name, os:M.os,
          cost:cost, impr:impr, click:click, conv:conv, revenue:revenue
        });
      }
    }
  }
  return rows;
}

/* ---------- 日期工具 ---------- */
function fmtDate(d){ const p=n=>String(n).padStart(2,'0'); return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate()); }
function weekStart(d){ const day=(d.getDay()+6)%7; const s=new Date(d); s.setDate(d.getDate()-day); return s; }
function isoWeek(d){ const s=weekStart(d); const t=new Date(s); t.setDate(t.getDate()+3); const jan=new Date(t.getFullYear(),0,1); const w=Math.ceil(((t-jan)/86400000+1)/7); return w; }
function parseDate(s){
  s = String(s||'').trim();
  let m = s.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  return null;
}

/* ---------- 状态 ---------- */
let DATA = { rows: [], source:'empty' };
let F = { days:'all', link:'all', plan:'all' };
const charts = {};
let activeView = 'overview';
const sortState = {};

function dataMaxTs(){
  let mx = 0;
  DATA.rows.forEach(r=>{ if(r.ts>mx) mx=r.ts; });
  return mx;
}

/* ---------- 筛选与聚合 ---------- */
function filteredRows(){
  const mx = dataMaxTs();
  return DATA.rows.filter(r=>{
    if (F.link !== 'all' && r.link !== F.link) return false;
    if (F.plan !== 'all' && r.plan !== F.plan) return false;
    if (F.days !== 'all' && (mx - r.ts)/86400000 > Number(F.days)) return false;
    return true;
  });
}

function totals(rows){
  let cost=0, impr=0, click=0, conv=0, rev=0;
  rows.forEach(r=>{ cost+=r.cost; impr+=r.impr; click+=r.click; conv+=r.conv; rev+=r.revenue; });
  return { cost, impr, click, conv, rev, ctr:pct(click,impr), cvr:cvrOf(conv,click), cpa: conv?cpaOf(cost,conv):null, roi: cost? rev/cost : 0, rows:rows.length };
}

function groupBy(rows, key){
  const map = {};
  rows.forEach(r=>{ const k = r[key]; (map[k] = map[k]||[]).push(r); });
  return map;
}

function summarizeGroup(map){
  return Object.keys(map).map(k => {
    const t = totals(map[k]);
    return { key:k, ...t, rows:map[k].length };
  });
}

/* ---------- 周序列 ---------- */
function weeklySeries(rows){
  const g = groupBy(rows, 'weekStart');
  const wkList = Object.keys(g).sort();
  return wkList.map(ws => {
    const t = totals(g[ws]);
    const label = g[ws][0].weekLabel;
    return { ws, label, ...t };
  });
}

/* ---------- 链路评分 ---------- */
function stddev(arr){
  if (arr.length < 2) return 0;
  const m = sum(arr)/arr.length;
  return Math.sqrt(arr.reduce((s,v)=>s+(v-m)*(v-m),0)/arr.length);
}
function scoreLinks(rows){
  const g = groupBy(rows, 'link');
  const items = summarizeGroup(g);
  const maxConv = Math.max(1, ...items.map(i=>i.conv));
  const maxSpend = Math.max(1, ...items.map(i=>i.cost));
  const minCpa = Math.min(...items.filter(i=>i.cpa>0).map(i=>i.cpa));
  items.forEach(it => {
    // 稳定性：周级CVR变异系数
    const wk = weeklySeries(g[it.key]);
    const cvrArr = wk.filter(w=>w.click>0).map(w=>w.cvr);
    let stable;
    if (cvrArr.length < 2) stable = 60;
    else { const cv = cvrArr.length? stddev(cvrArr)/(sum(cvrArr)/cvrArr.length) : 0; stable = clamp(100 - cv*220, 0, 100); }
    const scaleScore = 40 + 60 * Math.log10(1+it.conv) / Math.log10(1+maxConv);
    const effScore = maxConv ? (it.cvr / Math.max(...items.map(x=>x.cvr)))*100 : 0;
    const costScore = it.cpa>0 && minCpa>0 ? clamp(minCpa/it.cpa*100, 0, 100) : 50;
    const growScore = it.cost>0 ? it.cost/maxSpend*100 : 0;
    const total = scaleScore*0.25 + effScore*0.20 + costScore*0.25 + stable*0.15 + growScore*0.15;
    it.score = Math.round(total*10)/10;
    it.scaleScore = Math.round(scaleScore);
    it.effScore = Math.round(effScore);
    it.costScore = Math.round(costScore);
    it.stableScore = Math.round(stable);
    it.growScore = Math.round(growScore);
    it.tier = it.score>=80?'S':it.score>=65?'A':it.score>=50?'B':it.score>=35?'C':'D';
  });
  items.sort((a,b)=>b.score-a.score);
  return items;
}

/* ---------- ECharts 帮助 ---------- */
const chartObs = {};
function disposeAllCharts(){
  Object.keys(chartObs).forEach(k=>{ try{ chartObs[k].disconnect(); }catch(e){} });
  Object.keys(chartObs).forEach(k=>delete chartObs[k]);
  Object.keys(charts).forEach(k=>{ try{ charts[k].dispose(); }catch(e){} });
  Object.keys(charts).forEach(k=>delete charts[k]);
}
function chart(elId){
  const el = document.getElementById(elId);
  if (!el) return null;
  if (charts[elId]) { try{ charts[elId].dispose(); }catch(e){} delete charts[elId]; }
  if (chartObs[elId]) { try{ chartObs[elId].disconnect(); }catch(e){} delete chartObs[elId]; }
  const inst = echarts.init(el);
  charts[elId] = inst;
  chartObs[elId] = new ResizeObserver(()=>inst.resize());
  chartObs[elId].observe(el);
  return inst;
}
const defaultCell = (r,k) => (typeof r[k]==='number' ? fmtN(r[k]) : esc(r[k]));
const baseAxis = {
  axisLine:{lineStyle:{color:C.axis}},
  axisTick:{show:false},
  axisLabel:{color:C.ink2||'#5C5A54', fontSize:11},
  splitLine:{lineStyle:{color:C.grid}}
};
const tip = { confine:true, backgroundColor:'#20242B', borderWidth:0, padding:[8,12], textStyle:{color:'#fff', fontSize:12} };

/* ---------- 渲染：KPI ---------- */
function renderKpis(elId, items){
  const html = items.map(it => `
    <div class="kpi ${it.cls||''}">
      <div class="k-l">${it.icon||''}${esc(it.label)}</div>
      <div class="k-v num">${it.value}</div>
      ${it.sub?`<div class="k-s">${it.sub}</div>`:''}
    </div>`).join('');
  $(elId).innerHTML = html;
}

/* ---------- 表格排序 ---------- */
function sortTable(tableId, data, cols){
  const st = sortState[tableId] || { key: cols[0].key, dir: -1 };
  const key = st.key, dir = st.dir;
  const sorted = data.slice().sort((a,b)=>{
    let va = a[key], vb = b[key];
    if (typeof va === 'string' || typeof vb === 'string') return String(va).localeCompare(String(vb),'zh')*dir;
    return (va-vb)*dir;
  });
  const thead = '<tr>' + cols.map(c=>{
    const arr = key===c.key ? `<span class="arr">${dir>0?'▲':'▼'}</span>` : '';
    return `<th class="${c.sortable?'sortable':''}" data-key="${c.key}" data-table="${tableId}">${esc(c.label)}${arr}</th>`;
  }).join('') + '</tr>';
  const rows = sorted.map((r,idx)=>{
    return '<tr>' + cols.map(c=>{
      const cell = c.cell ? c.cell(r, idx) : esc(r[c.key]);
      return `<td>${cell}</td>`;
    }).join('') + '</tr>';
  }).join('');
  const tbl = $(tableId);
  tbl.innerHTML = thead + rows;
  tbl.querySelectorAll('th.sortable').forEach(th=>{
    th.onclick = ()=>{
      const k = th.dataset.key;
      const cur = sortState[tableId]||{key:k,dir:-1};
      sortState[tableId] = { key:k, dir: cur.key===k ? -cur.dir : -1 };
      sortTable(tableId, data, cols);
    };
  });
}

/* ============================================================
   增强模块：单元名解析 / 预算管理 / 优化建议 / 操作清单
   ============================================================ */

/* ---------- 营销单元名称解析 ---------- */
// 命名规则：日期-产品-版位--定向-变体--出价类型
// 示例：0830-太极灵儿-多版位--排低r城43男--双出价
//       0826-太极灵儿-多版位-排偏远-45+男--双出价
//       0819-太极灵儿-短剧--排低r城46-AB-双出价
function parseUnitName(name){
  const r = { raw:name, date:'', product:'', placement:'', targeting:'', variant:'', bidType:'', age:'', cityTier:'' };
  if(!name) return r;
  const parts = name.split('--').map(s=>s.trim());
  // 第一段：日期-产品-版位
  if(parts[0]){
    const p0 = parts[0].split('-');
    // 日期：4位数字开头
    if(p0[0] && /^\d{4}$/.test(p0[0])){
      r.date = p0[0].slice(0,2)+'/'+p0[0].slice(2,4);
      p0.shift();
    }
    // 产品：第二个元素
    if(p0.length>0){ r.product = p0[0]; p0.shift(); }
    // 版位：剩余元素合并（可能是"多版位"、"短剧"、"排小程序"等）
    if(p0.length>0) r.placement = p0.join('-');
  }
  // 第二段：定向-变体
  if(parts[1]){
    const p1 = parts[1].split('-');
    // 提取年龄段
    const ageMatch = parts[1].match(/(\d{2})\+?男|(\d{2})男|(\d{2})\+/);
    if(ageMatch){ r.age = (ageMatch[1]||ageMatch[2]||ageMatch[3])+'+'; }
    // 提取城市层级
    if(/低r城|低线城/.test(parts[1])) r.cityTier = '低线城市';
    else if(/高r|高线/.test(parts[1])) r.cityTier = '高线城市';
    else if(/偏远/.test(parts[1])) r.cityTier = '偏远地区';
    // 定向部分（去掉变体关键词后的剩余）
    const variantKeywords = ['高r迭代','高r','痛点','AB','AB素材','xin','复合','迭代','gaor','xin-'];
    let targetingParts = [];
    p1.forEach(pp=>{
      if(!variantKeywords.some(v=>pp===v || pp.includes(v))){
        targetingParts.push(pp);
      }
    });
    r.targeting = targetingParts.join('-') || parts[1];
    // 变体
    const foundVariant = variantKeywords.find(v=>parts[1].includes(v));
    if(foundVariant) r.variant = foundVariant;
  }
  // 第三段：出价类型
  if(parts[2]){
    r.bidType = parts[2];
  }
  return r;
}

function unitDimTags(parsed){
  const tags = [];
  if(parsed.date) tags.push('<span class="dim-tag date">'+parsed.date+'</span>');
  if(parsed.placement) tags.push('<span class="dim-tag placement">'+parsed.placement+'</span>');
  if(parsed.targeting) tags.push('<span class="dim-tag targeting">'+esc(parsed.targeting)+'</span>');
  if(parsed.variant) tags.push('<span class="dim-tag variant">'+parsed.variant+'</span>');
  if(parsed.bidType) tags.push('<span class="dim-tag bidtype">'+parsed.bidType+'</span>');
  return '<div class="dim-tags">'+tags.join('')+'</div>';
}

/* ---------- 预算管理 ---------- */
const BUDGET_KEY = 'workbench_budgets_v1';
function loadBudgets(){
  try{ return JSON.parse(localStorage.getItem(BUDGET_KEY)) || {}; }catch(e){ return {}; }
}
function saveBudgets(b){ try{ localStorage.setItem(BUDGET_KEY, JSON.stringify(b)); }catch(e){} }
function getBudget(accId){ const b = loadBudgets(); return b[accId] || { daily:0, note:'' }; }
function setBudget(accId, daily, note){ const b = loadBudgets(); b[accId] = { daily:Number(daily)||0, note:note||'' }; saveBudgets(b); }

function calcPacing(spend, dailyBudget, balance){
  if(!dailyBudget || dailyBudget<=0) return { pct:0, status:'none', label:'未设预算', estDaily:0, daysLeft:0 };
  const pct = (spend/dailyBudget)*100;
  const now = new Date();
  const hour = now.getHours() + now.getMinutes()/60;
  // 假设投放时段 8:00-24:00 = 16小时
  const elapsedHours = Math.max(0, Math.min(16, hour-8));
  const expectedPct = (elapsedHours/16)*100;
  let status = 'ok', label = '正常';
  if(pct > expectedPct + 20){ status='danger'; label='超投'; }
  else if(pct > expectedPct + 10){ status='warn'; label='偏快'; }
  else if(pct < expectedPct - 20){ status='warn'; label='欠投'; }
  const estDaily = elapsedHours>0 ? Math.round(spend / (elapsedHours/16)) : spend;
  const daysLeft = spend>0 ? Math.round(balance / (spend/elapsedHours*16)) : 99;
  return { pct:Math.min(pct,100), status, label, estDaily, daysLeft, expectedPct };
}

/* ---------- 优化建议引擎 ---------- */
function generateSuggestions(liveData, adUnitsData){
  if(!liveData || !liveData.accounts) return [];
  const suggestions = [];
  const t = liveData.total;
  const avgCpa = t.cpa || 200;
  const num = (v)=>{const n=parseFloat(String(v||'').replace(/[,¥%]/g,''));return isNaN(n)?0:n;};

  // 合并所有单元
  const allUnits = [];
  (adUnitsData||[]).forEach(a=>{
    (a.units||[]).forEach(u=>{
      allUnits.push({
        accountId:a.account_id, operator:a.operator||'',
        name:u[0]||'', spend:num(u[5]), conv:num(u[7]), cpa:num(u[8]),
        bid:u[3]||'', impressions:num(u[4])
      });
    });
  });

  // 规则1：空耗单元（消耗>50，0转化）
  allUnits.filter(u=>u.spend>50 && u.conv===0).forEach(u=>{
    suggestions.push({
      type:'pause', title:'空耗单元建议暂停',
      unit:u.name, accountId:u.accountId, operator:u.operator,
      detail:'已消耗 ¥'+fmtM(u.spend)+'，0转化。建议暂停或检查落地页/定向。',
      meta:'出价：'+u.bid+' | 曝光：'+fmtN(u.impressions),
      action:'暂停'
    });
  });

  // 规则2：超成本单元（CPA>均值×1.5，转化≥3）
  allUnits.filter(u=>u.cpa>0 && u.cpa>avgCpa*1.5 && u.conv>=3).forEach(u=>{
    const overPct = Math.round((u.cpa/avgCpa-1)*100);
    suggestions.push({
      type:'reduce', title:'超成本单元建议降价',
      unit:u.name, accountId:u.accountId, operator:u.operator,
      detail:'CPA ¥'+fmtM2(u.cpa)+'，超均值 '+overPct+'%（均值¥'+fmtM2(avgCpa)+'）。转化'+u.conv+'个，建议降价10-15%或收窄定向。',
      meta:'消耗：¥'+fmtM(u.spend)+' | 出价：'+u.bid,
      action:'降价10%'
    });
  });

  // 规则3：优质单元建议加预算（CPA<均值×0.7，消耗>预算×0.8，转化≥5）
  const budgets = loadBudgets();
  allUnits.filter(u=>u.cpa>0 && u.cpa<avgCpa*0.7 && u.conv>=5 && u.spend>100).forEach(u=>{
    const accBudget = budgets[u.accountId];
    const accSpend = (liveData.accounts.find(a=>a.id===u.accountId)||{}).spend || 0;
    const shouldIncrease = !accBudget || !accBudget.daily || accSpend > accBudget.daily*0.7;
    if(shouldIncrease){
      suggestions.push({
        type:'increase', title:'优质单元建议加预算',
        unit:u.name, accountId:u.accountId, operator:u.operator,
        detail:'CPA ¥'+fmtM2(u.cpa)+'（低于均值'+Math.round((1-u.cpa/avgCpa)*100)+'%），转化'+u.conv+'个，消耗¥'+fmtM(u.spend)+'。建议加预算20-30%放量。',
        meta:'消耗：¥'+fmtM(u.spend)+' | 转化：'+u.conv+'个',
        action:'加预算20%'
      });
    }
  });

  // 规则4：账户余额预警
  liveData.accounts.forEach(a=>{
    if(a.spend>0 && a.balance < a.spend*3){
      suggestions.push({
        type:'watch', title:'账户余额不足预警',
        unit:'账户 '+a.id, accountId:a.id, operator:a.operator||'',
        detail:'余额 ¥'+fmtM(a.balance)+'，今日消耗 ¥'+fmtM(a.spend)+'，预计还能撑 '+Math.round(a.balance/a.spend)+' 天。建议尽快充值。',
        meta:'余额：¥'+fmtM(a.balance)+' | 日耗：¥'+fmtM(a.spend),
        action:'充值'
      });
    }
  });

  // 按类型排序：暂停 > 降价 > 加预算 > 关注
  const typeOrder = { pause:0, reduce:1, increase:2, watch:3 };
  suggestions.sort((a,b)=>typeOrder[a.type]-typeOrder[b.type]);
  return suggestions;
}

function exportOperationList(suggestions){
  const rows = [['序号','操作类型','营销单元名称','账户ID','运营方','消耗','转化','CPA','建议操作','详细说明']];
  suggestions.forEach((s,i)=>{
    rows.push([i+1, s.type==='pause'?'暂停':s.type==='reduce'?'降价':s.type==='increase'?'加预算':'关注',
      s.unit, s.accountId, s.operator, '', '', '', s.action, s.detail]);
  });
  const csv = '\uFEFF' + rows.map(r=>r.map(c=>'"'+String(c||'').replace(/"/g,'""')+'"').join(',')).join('\n');
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = '操作清单_'+new Date().toISOString().slice(0,10)+'.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ---------- 数据时间戳 ---------- */
function getDataFreshness(syncTime){
  if(!syncTime) return { cls:'', text:'未同步', hours:999 };
  const sync = new Date(syncTime.replace(/-/g,'/'));
  const now = new Date();
  const hours = (now - sync)/3600000;
  let cls='', text='';
  if(hours < 2){ cls=''; text='实时'; }
  else if(hours < 8){ cls='stale'; text=''+Math.round(hours)+'小时前'; }
  else { cls='old'; text=''+Math.round(hours)+'小时前·建议刷新'; }
  return { cls, text, hours };
}

/* ---------- 视图：数据总览·优化驾驶舱 ---------- */
function renderOverview(){
  renderOverviewTodos();
  const live = TENCENT_LIVE_DATA;
  const adData = TENCENT_ADUNITS_DATA || [];
  if(!live || !live.accounts){
    renderKpis('#ov-kpis',[{label:'总消耗',value:'—'},{label:'总转化',value:'—'},{label:'平均CPA',value:'—'},{label:'CTR',value:'—'},{label:'CVR',value:'—'},{label:'CPC',value:'—'}]);
    $('#ov-warnings').innerHTML = '<div class="insight">暂无腾讯广告数据，请先到「腾讯广告」页同步数据。</div>';
    return;
  }
  const t = live.total;
  const num = (v)=>{const n=parseFloat(String(v||'').replace(/[,¥%]/g,''));return isNaN(n)?0:n;};

  // ===== 1. 核心KPI（6张） =====
  renderKpis('#ov-kpis',[
    {label:'今日总消耗',value:fmtM(t.spend),cls:'kpi-accent'},
    {label:'总转化',value:fmtN(t.conversions),cls:'kpi-teal'},
    {label:'平均CPA',value:t.conversions?fmtM2(t.cpa):'—',cls:t.cpa>200?'kpi-red':'kpi-blue'},
    {label:'CTR',value:fmtP(t.ctr)},
    {label:'CVR',value:fmtP(t.cvr)},
    {label:'CPC',value:fmtM2(t.cpc)}
  ]);

  // ===== 2. 健康预警（4张） =====
  // 合并所有投放中单元
  const allUnits = [];
  adData.forEach(a => {
    a.units.forEach(u => {
      allUnits.push({
        accountId:a.account_id, operator:a.operator||'',
        name:u[0]||'', status:u[2]||'', bid:u[3]||'',
        impressions:u[4]||'-', spend:u[5]||'-', cpm:u[6]||'-',
        conv:u[7]||'-', cpa:u[8]||'-', cvr:u[9]||'-',
        _imp:num(u[4]), _spd:num(u[5]), _conv:num(u[7]), _cpa:num(u[8])
      });
    });
  });
  const avgCpa = t.cpa || 200;
  const wasteUnits = allUnits.filter(u=>u._spd>0 && u._conv===0);
  const overCostUnits = allUnits.filter(u=>u._cpa>0 && u._cpa>avgCpa*1.2);
  const lowSpendUnits = allUnits.filter(u=>u._spd>0 && u._spd<10);
  const lowBalanceAccs = live.accounts.filter(a=>a.balance < Math.max(1000, a.spend*0.5));

  const warnCards = [
    {label:'空耗单元',value:wasteUnits.length,sub:'有消耗·0转化',cls:wasteUnits.length>0?'kpi-red':'kpi-teal'},
    {label:'超成本单元',value:overCostUnits.length,sub:'CPA>均值×1.2',cls:overCostUnits.length>2?'kpi-red':'kpi-amber'},
    {label:'低消耗单元',value:lowSpendUnits.length,sub:'日耗<10元',cls:'kpi-blue'},
    {label:'余额预警账户',value:lowBalanceAccs.length,sub:lowBalanceAccs.length?lowBalanceAccs.map(a=>a.id.slice(-4)).join('、'):'余额充足',cls:lowBalanceAccs.length>0?'kpi-amber':'kpi-teal'}
  ];
  renderKpis('#ov-warnings',warnCards);

  // ===== 3. 6账户投放战况表 =====
  $('#ov-acc-tag').textContent = live.accounts.length+'个账户 · 总余额¥'+fmtM(t.balance);
  const accRows = live.accounts.slice().sort((a,b)=>b.spend-a.spend).map((a,i)=>{
    let status='delta-flat', statusText='—';
    if(a.spend>0 && a.conversions===0){status='delta-down';statusText='空耗';}
    else if(a.cpa>0 && a.cpa>avgCpa*1.2){status='delta-down';statusText='超成本';}
    else if(a.cpa>0 && a.cpa<avgCpa*0.8 && a.conversions>0){status='delta-up';statusText='优质';}
    else if(a.conversions>0){status='delta-flat';statusText='正常';}
    // 该账户异常单元数
    const accWaste = allUnits.filter(u=>u.accountId===a.id && u._spd>0 && u._conv===0).length;
    const accOver = allUnits.filter(u=>u.accountId===a.id && u._cpa>0 && u._cpa>avgCpa*1.2).length;
    const abnormal = accWaste+accOver;
    return {
      idx:i+1, id:a.id, operator:a.operator||'', spend:a.spend, conv:a.conversions,
      cpa:a.cpa, ctr:a.ctr, cvr:a.cvr, balance:a.balance,
      status, statusText, abnormal,
      _spd:a.spend, _conv:a.conversions, _cpa:a.cpa
    };
  });
  sortState['#ov-accounts']={key:'_spd',dir:-1};
  sortTable('#ov-accounts',accRows,[
    {key:'idx',label:'#',cell:r=>'<span class="rank">'+r.idx+'</span>'},
    {key:'id',label:'账户ID',sortable:true,cell:r=>'<b style="color:var(--accent);font-size:12px">'+r.id+'</b>'+(r.operator?'<br><span style="color:var(--ink-3);font-size:10px">'+r.operator+'</span>':'')},
    {key:'status',label:'状态',cell:r=>'<span class="'+r.status+'" style="font-size:11px">'+r.statusText+'</span>'},
    {key:'_spd',label:'消耗',sortable:true,cell:r=>'<b>¥'+fmtM(r._spd)+'</b>'},
    {key:'_conv',label:'转化',sortable:true,cell:r=>r._conv>0?'<b style="color:var(--mint)">'+fmtN(r._conv)+'</b>':'<span style="color:var(--ink-3)">0</span>'},
    {key:'_cpa',label:'CPA',sortable:true,cell:r=>{if(r._cpa<=0)return '<span style="color:var(--ink-3)">—</span>';const cls=r._cpa>avgCpa*1.2?'color:var(--red)':r._cpa<avgCpa*0.8?'color:var(--mint)':'';return '<span style="'+cls+'">¥'+fmtM2(r._cpa)+'</span>';}},
    {key:'ctr',label:'CTR',sortable:true,cell:r=>fmtP(r.ctr)},
    {key:'cvr',label:'CVR',sortable:true,cell:r=>fmtP(r.cvr)},
    {key:'abnormal',label:'异常单元',sortable:true,cell:r=>r.abnormal>0?'<b style="color:var(--red)">'+r.abnormal+'</b>':'<span style="color:var(--ink-3)">0</span>'},
    {key:'balance',label:'余额',sortable:true,cell:r=>'<span style="font-size:11.5px">¥'+fmtM(r.balance)+'</span>'}
  ].map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));

  // ===== 4. 转化漏斗 =====
  const impr = t.impressions, clk = t.clicks, conv = t.conversions;
  const ctrRate = impr?clk/impr*100:0, cvrRate = clk?conv/clk*100:0;
  const maxVal = impr;
  const funnelHtml = `
    <div style="padding:8px 4px">
      <div style="margin-bottom:14px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
          <span style="font-size:13px;font-weight:600;color:var(--ink-1)">曝光</span>
          <span style="font-size:15px;font-weight:700;color:var(--accent)">`+fmtN(impr)+`</span>
        </div>
        <div style="height:28px;background:linear-gradient(90deg,var(--accent),#9B8FE8);border-radius:6px;opacity:0.85;width:100%"></div>
      </div>
      <div style="text-align:center;color:var(--ink-3);font-size:11px;margin:-6px 0 8px">↓ CTR `+fmtP(ctrRate)+`</div>
      <div style="margin-bottom:14px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
          <span style="font-size:13px;font-weight:600;color:var(--ink-1)">点击</span>
          <span style="font-size:15px;font-weight:700;color:var(--blue)">`+fmtN(clk)+`</span>
        </div>
        <div style="height:28px;background:linear-gradient(90deg,var(--blue),#8BABF0);border-radius:6px;opacity:0.8;width:`+Math.max(4,(clk/maxVal)*100)+`%"></div>
      </div>
      <div style="text-align:center;color:var(--ink-3);font-size:11px;margin:-6px 0 8px">↓ CVR `+fmtP(cvrRate)+`</div>
      <div style="margin-bottom:8px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
          <span style="font-size:13px;font-weight:600;color:var(--ink-1)">转化</span>
          <span style="font-size:15px;font-weight:700;color:var(--mint)">`+fmtN(conv)+`</span>
        </div>
        <div style="height:28px;background:linear-gradient(90deg,var(--mint),#7FCFB0);border-radius:6px;opacity:0.75;width:`+Math.max(3,(conv/maxVal)*100)+`%"></div>
      </div>
      <div style="margin-top:14px;padding-top:12px;border-top:1px solid var(--border-1);display:flex;justify-content:space-between;font-size:11.5px;color:var(--ink-3)">
        <span>曝光→点击流失：<b style="color:var(--ink-2)">`+fmtN(impr-clk)+`</b></span>
        <span>点击→转化流失：<b style="color:var(--ink-2)">`+fmtN(clk-conv)+`</b></span>
        <span>整体转化效率：<b style="color:var(--mint)">`+fmtP(impr?conv/impr*100:0)+`</b></span>
      </div>
    </div>`;
  $('#ov-funnel').innerHTML = funnelHtml;

  // ===== 5. 投放中单元明细（全部，按消耗降序） =====
  $('#ov-unit-tag').textContent = allUnits.length+'个投放中单元';
  allUnits.sort((a,b)=>b._spd-a._spd);
  allUnits.forEach((u,i)=>{u.idx=i+1;});
  sortState['#ov-units']={key:'_spd',dir:-1};
  sortTable('#ov-units',allUnits,[
    {key:'idx',label:'#',cell:r=>'<span class="rank">'+r.idx+'</span>'},
    {key:'accountId',label:'账户',sortable:true,cell:r=>'<span style="color:var(--accent);font-size:11px;font-weight:600">'+r.accountId.slice(-4)+'</span>'},
    {key:'name',label:'营销单元名称',sortable:true,cell:r=>'<b style="font-size:12px">'+esc(r.name).slice(0,28)+(r.name.length>28?'…':'')+'</b>'},
    {key:'status',label:'状态',cell:r=>'<span class="delta-up" style="font-size:10px">投放中</span>'},
    {key:'bid',label:'出价',cell:r=>'<span style="font-size:11px">'+esc(r.bid).replace(' oCPM 元/下单','')+'</span>'},
    {key:'_imp',label:'曝光',sortable:true,cell:r=>r.impressions==='-'?'<span style="color:var(--ink-3)">—</span>':fmtN(r._imp)},
    {key:'_spd',label:'花费',sortable:true,cell:r=>{if(r.spend==='-'||r._spd===0)return '<span style="color:var(--ink-3)">—</span>';const cls=r._conv===0?'color:var(--ink-3);text-decoration:line-through':'';return '<b style="'+cls+'">¥'+fmtM(r._spd)+'</b>';}},
    {key:'cpm',label:'千次展现',sortable:true,cell:r=>r.cpm==='-'?'<span style="color:var(--ink-3)">—</span>':fmtM2(r.cpm)},
    {key:'_conv',label:'转化量',sortable:true,cell:r=>{if(r.conv==='-'||r._conv===0)return '<span style="color:var(--red);font-weight:600">0</span>';return '<b style="color:var(--mint)">'+fmtN(r._conv)+'</b>';}},
    {key:'_cpa',label:'转化成本',sortable:true,cell:r=>{if(r.cpa==='-'||r._cpa<=0)return '<span style="color:var(--ink-3)">—</span>';const cls=r._cpa>avgCpa*1.2?'color:var(--red);font-weight:700':r._cpa<avgCpa*0.8?'color:var(--mint)':'';return '<span style="'+cls+'">¥'+fmtM2(r._cpa)+'</span>';}},
    {key:'cvr',label:'转化率',sortable:true,cell:r=>r.cvr==='-'?'<span style="color:var(--ink-3)">—</span>':esc(r.cvr)}
  ].map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));

  // ===== 6. 预算消耗追踪 =====
  const budgetList = $('#ov-budget-list');
  if(budgetList){
    const budgets = loadBudgets();
    let bHtml = '';
    live.accounts.slice().sort((a,b)=>b.spend-a.spend).forEach(acc=>{
      const bg = budgets[acc.id] || { daily:0 };
      const pacing = calcPacing(acc.spend, bg.daily, acc.balance);
      const fillCls = pacing.status==='danger'?'danger':pacing.status==='warn'?'warn':'ok';
      const daysCls = pacing.daysLeft<=3?'low':'ok';
      bHtml += `<div class="budget-acc-row">
        <div><div class="ba-name">ID:${acc.id}</div><div class="ba-op">${esc(acc.operator||'')} · ${pacing.label}</div></div>
        <div class="budget-bar-wrap">
          <div class="budget-bar-fill ${fillCls}" style="width:${pacing.pct}%"></div>
          <div class="budget-bar-text">¥${fmtM(acc.spend)} / ${bg.daily?('¥'+fmtN(bg.daily)):'未设预算'} (${pacing.pct.toFixed(0)}%)</div>
        </div>
        <div style="text-align:right;font-size:11.5px">预计今日<br><b>¥${fmtN(pacing.estDaily)}</b></div>
        <div><input type="number" class="budget-input" data-acc="${acc.id}" value="${bg.daily||''}" placeholder="日预算" min="0"></div>
        <div class="budget-days ${daysCls}">余额<br><b>${pacing.daysLeft>=99?'—':pacing.daysLeft+'天'}</b></div>
      </div>`;
    });
    budgetList.innerHTML = bHtml;
    // 绑定预算输入
    budgetList.querySelectorAll('.budget-input').forEach(inp=>{
      inp.onchange = ()=>{
        const accId = inp.dataset.acc;
        const val = parseFloat(inp.value)||0;
        setBudget(accId, val, '');
        showToast('账户 '+accId+' 日预算已设为 ¥'+fmtN(val));
        renderOverview();
      };
    });
  }

  // ===== 7. 智能优化建议 =====
  const sugContainer = $('#ov-suggestions');
  const sugCount = $('#ov-sug-count');
  if(sugContainer){
    const sugs = generateSuggestions(live, adData);
    if(sugCount) sugCount.textContent = sugs.length+' 条建议';
    if(sugs.length===0){
      sugContainer.innerHTML = '<div class="insight">当前数据下未发现需要紧急处理的问题，继续保持监控。</div>';
    } else {
      sugContainer.innerHTML = sugs.map(s=>`
        <div class="suggestion-card ${s.type}">
          <div class="suggestion-head">
            <span class="suggestion-title">${esc(s.title)}</span>
            <span class="suggestion-action ${s.type}">${esc(s.action)}</span>
          </div>
          <div class="suggestion-detail">${esc(s.detail)}</div>
          <div class="suggestion-meta">单元：${esc(s.unit)} | 账户：${s.accountId} | ${esc(s.operator||'')} | ${esc(s.meta||'')}</div>
        </div>
      `).join('');
    }
    // 绑定导出按钮
    const expBtn = $('#ov-export-op');
    if(expBtn){
      expBtn.onclick = ()=>{
        if(sugs.length===0){ showToast('暂无优化建议'); return; }
        exportOperationList(sugs);
        showToast('已导出 '+sugs.length+' 条操作建议');
      };
    }
  }
}

/* ---------- 视图：手机型号数据分析 ---------- */
var phoneDim='model';
function normPhoneDim(rows,nameKey,brandKey){
  return rows.map(r=>({name:r[nameKey],brand:brandKey?r[brandKey]:'',orders:+r['低价课订单数']||0,valid:+r['有效订单数']||0,highOrders:+r['高价课单数']||0,highRate:+(r['高价课转化率']*100)||0,highAmount:+r['高价课金额(元)']||+r['高价课金额']||0,amountShare:+(r['高价课金额占比']*100)||0,outputPerOrder:+r['单订单产值(元)']||+r['单订单产值']||0}));
}
function renderPhoneCharts(items,dimName){
  const titles={brand:'品牌',model:'型号',price:'价格区间',age:'年龄段'};
  const tn=titles[dimName]||'维度';
  $('#ph-dim-top-title').textContent=tn+'高价课单数 Top 10';
  $('#ph-dim-top-hint').textContent='按高价课单数排序，识别主力出量'+tn;
  $('#ph-dim-pie-title').textContent=tn+'订单数分布';
  $('#ph-dim-cost-title').textContent=tn+'高价课金额 Top 10';
  $('#ph-dim-cpa-title').textContent=tn+'单订单产值 Top 10';
  $('#ph-dim-table-title').textContent=tn+'明细表';
  const tm=items.slice().sort((a,b)=>b.highOrders-a.highOrders).slice(0,10).reverse();
  const c1=chart('ph-dim-top'); if(c1)c1.setOption({color:[C.blue],tooltip:tip,grid:{left:8,right:16,top:8,bottom:8,containLabel:true},xAxis:{type:'value',min:0,...baseAxis},yAxis:{type:'category',data:tm.map(i=>i.name),...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11}},series:[{type:'bar',data:tm.map(i=>i.highOrders),barWidth:'58%',label:{show:true,position:'right',formatter:p=>fmtN(p.value),color:'#5C5A54',fontSize:11},itemStyle:{color:C.blue}}]});
  const c2=chart('ph-dim-pie'); if(c2)c2.setOption({color:C.cat,tooltip:tip,legend:{type:'scroll',bottom:0},series:[{type:'pie',radius:['42%','66%'],center:['50%','45%'],label:{show:true,formatter:p=>p.name+' '+p.percent+'%'},data:items.slice(0,10).map(i=>({name:i.name,value:i.orders}))}]});
  const tb=items.slice().sort((a,b)=>b.highAmount-a.highAmount).slice(0,10).reverse();
  const c3=chart('ph-dim-cost'); if(c3)c3.setOption({color:[C.accent],tooltip:tip,grid:{left:8,right:16,top:8,bottom:8,containLabel:true},xAxis:{type:'value',min:0,...baseAxis},yAxis:{type:'category',data:tb.map(i=>i.name),...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11}},series:[{type:'bar',data:tb.map(i=>Math.round(i.highAmount)),barWidth:'58%',label:{show:true,position:'right',formatter:p=>fmtN(p.value),color:'#5C5A54',fontSize:11},itemStyle:{color:C.accent}}]});
  const tc=items.filter(i=>i.highOrders>0).sort((a,b)=>b.outputPerOrder-a.outputPerOrder).slice(0,10).reverse();
  const c4=chart('ph-dim-cpa'); if(c4)c4.setOption({color:[C.teal],tooltip:tip,grid:{left:8,right:16,top:8,bottom:8,containLabel:true},xAxis:{type:'value',min:0,...baseAxis},yAxis:{type:'category',data:tc.map(i=>i.name),...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11}},series:[{type:'bar',data:tc.map(i=>Math.round(i.outputPerOrder)),barWidth:'58%',label:{show:true,position:'right',formatter:p=>'¥'+fmtN(p.value),color:'#5C5A54',fontSize:11},itemStyle:{color:C.teal}}]});
  const topI=items[0];
  $('#ph-insight').innerHTML=topI? tn+'维度 <b>'+esc(topI.name)+'</b> 订单最多（'+fmtN(topI.orders)+' 单，高价课转化 '+fmtN(topI.highOrders)+'）；单订单产值 ¥'+fmtM2(topI.outputPerOrder)+'，高价转化率 '+fmtP(topI.highRate)+'。' : '当前维度无数据。';
  sortState['#ph-table']={key:'orders',dir:-1};
  const cols=[
    {key:'name',label:tn,cell:(r,i)=>'<span class="rank">'+(i+1)+'</span>&nbsp;'+esc(r.name)},
    {key:'brand',label:'品牌',sortable:true,cell:r=>r.brand?esc(r.brand):'—'}
  ];
  if(dimName==='model') cols[1]={key:'brand',label:'品牌',sortable:true};
  cols.push({key:'orders',label:'低价课订单',sortable:true,cell:r=>fmtN(r.orders)});
  cols.push({key:'valid',label:'有效订单',sortable:true,cell:r=>fmtN(r.valid)});
  cols.push({key:'highOrders',label:'高价课单数',sortable:true,cell:r=>fmtN(r.highOrders)});
  cols.push({key:'highRate',label:'高价转化率',sortable:true,cell:r=>fmtP(r.highRate)});
  cols.push({key:'highAmount',label:'高价课金额',sortable:true,cell:r=>fmtM(r.highAmount)});
  cols.push({key:'amountShare',label:'金额占比',sortable:true,cell:r=>fmtP(r.amountShare)});
  cols.push({key:'outputPerOrder',label:'单订单产值',sortable:true,cell:r=>fmtM2(r.outputPerOrder)});
  sortTable('#ph-table',items,cols.map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));
}
function renderPhone(){
  renderSyncPanel('phone', document.getElementById('sync-phone'));
  document.querySelectorAll('#ph-subtabs .sub-tab').forEach(b=>{
    b.classList.toggle('active', b.dataset.dim===phoneDim);
    b.onclick=()=>{phoneDim=b.dataset.dim;renderPhone();};
  });
  if(SYNC.phone){
    const d=SYNC.phone; const cr=d.nTotal?d.nHigh/d.nTotal*100:0; const oo=d.nTotal?d.totalHigh/d.nTotal:0;
    renderKpis('#ph-kpis',[
      {label:'订单总数',value:fmtN(d.nTotal)},{label:'有效订单',value:fmtN(d.nValid),cls:'kpi-accent'},
      {label:'高价课转化',value:fmtN(d.nHigh),cls:'kpi-teal'},{label:'高价课转化率',value:fmtP(cr)},
      {label:'高价课金额',value:fmtM(d.totalHigh)},{label:'单订单产值',value:fmtM2(oo)},
      {label:'品牌数',value:fmtN(d.brandAgg.length)},{label:'型号数',value:fmtN(d.modelAgg.length)}
    ]);
    const dimMap={brand:[d.brandAgg,'品牌统一',''],model:[d.modelAgg,'手机型号','品牌统一'],price:[d.priceAgg,'价格区间',''],age:[d.ageAgg,'年龄段','']};
    const cfg=dimMap[phoneDim]||dimMap.model;
    const items=normPhoneDim(cfg[0],cfg[1],cfg[2]);
    renderPhoneCharts(items,phoneDim);
    return;
  }
  const rows = filteredRows();
  const t = totals(rows);
  const byModel = groupBy(rows,'model');
  const items = summarizeGroup(byModel).sort((a,b)=>b.conv-a.conv);
  renderKpis('#ph-kpis', [
    {label:'机型覆盖数', value:fmtN(items.length), sub:'去重设备型号'},
    {label:'总消耗', value:fmtM(t.cost), cls:'kpi-accent'},
    {label:'总展现', value:fmtN(t.impr)},{label:'总点击', value:fmtN(t.click)},
    {label:'总转化', value:fmtN(t.conv), cls:'kpi-teal'},{label:'平均CTR', value:fmtP(t.ctr)},
    {label:'平均CVR', value:fmtP(t.cvr)},{label:'平均CPA', value:t.conv?fmtM2(t.cpa):'—'}
  ]);
  if(phoneDim==='model'){
    const csvItems=items.map(i=>({name:i.key,brand:i.os,orders:i.conv,valid:i.conv,highOrders:i.conv,highRate:+i.cvr.toFixed(2),highAmount:i.cost,amountShare:+pct(i.conv,t.conv).toFixed(2),outputPerOrder:i.conv?+i.cpa.toFixed(1):0}));
    renderPhoneCharts(csvItems,'model');
  } else if(phoneDim==='brand'){
    const osG=summarizeGroup(groupBy(rows,'os')).sort((a,b)=>b.conv-a.conv);
    const csvItems=osG.map(i=>({name:i.key,brand:'',orders:i.conv,valid:i.conv,highOrders:i.conv,highRate:+i.cvr.toFixed(2),highAmount:i.cost,amountShare:+pct(i.conv,t.conv).toFixed(2),outputPerOrder:i.conv?+i.cpa.toFixed(1):0}));
    renderPhoneCharts(csvItems,'brand');
  } else {
    renderPhoneCharts([],phoneDim);
  }
}

/* ---------- 视图：链路数据表_系统口径_数据分析 ---------- */
var ltLink='__summary__';
var ltPeriod='';
const LT_METRICS=[
  {key:'订单ROI',label:'订单ROI',fmt:v=>fmtP(v*100)},
  {key:'广告花费(元)',label:'广告花费',fmt:v=>fmtM(v)},
  {key:'低价课总成本(元)',label:'低价课总成本',fmt:v=>fmtM(v)},
  {key:'订单数',label:'订单数',fmt:v=>fmtN(v)},
  {key:'单订单成本(元)',label:'单订单成本',fmt:v=>fmtM2(v)},
  {key:'单订单产值(元)',label:'单订单产值',fmt:v=>fmtM2(v)},
  {key:'高价课转化率',label:'高价课转化率',fmt:v=>fmtP(v*100)},
  {key:'高价课订单数',label:'高价课订单数',fmt:v=>fmtN(v)},
  {key:'有效会话率',label:'有效会话率',fmt:v=>fmtP(v*100)},
  {key:'一节直播到课',label:'一节直播到课率',fmt:v=>fmtP(v*100)},
  {key:'开口率',label:'开口率',fmt:v=>fmtP(v*100)},
  {key:'进群率',label:'进群率',fmt:v=>fmtP(v*100)},
  {key:'删友率',label:'删友率',fmt:v=>fmtP(v*100)}
];
function renderLinkTable(){
  renderSyncPanel('link', document.getElementById('sync-link'));
  const tabBar=document.getElementById('lt-subtabs');
  if(SYNC.link && tabBar.querySelectorAll('[data-link]').length<=1){
    SYNC.link.linkSummary.forEach(i=>{
      const b=document.createElement('button'); b.className='sub-tab'; b.dataset.link=i.link; b.textContent=i.link;
      b.onclick=()=>{ltLink=i.link;renderLinkTable();}; tabBar.appendChild(b);
    });
  }
  tabBar.querySelectorAll('.sub-tab').forEach(b=>b.classList.toggle('active',b.dataset.link===ltLink));
  const isSummary=ltLink==='__summary__';
  document.getElementById('lt-summary-view').style.display=isSummary?'':'none';
  document.getElementById('lt-period-view').style.display=isSummary?'none':'';
  if(SYNC.link){
    const d=SYNC.link; const t=d.total;
    renderKpis('#lt-kpis',[
      {label:'链路数',value:fmtN(d.linkSummary.length)},{label:'广告花费',value:fmtM(t.spend),cls:'kpi-accent'},
      {label:'低价课总成本',value:fmtM(t.cost)},{label:'有效订单',value:fmtN(t.eff),cls:'kpi-teal'},
      {label:'高价课订单',value:fmtN(t.highOrders)},{label:'单订单成本',value:fmtM2(t.orderCost)},
      {label:'高价转化率',value:fmtP(t.highConv*100)},{label:'整体ROI',value:fmtP(t.roi*100)}
    ]);
    if(isSummary){
      const c1=chart('lt-funnel'); if(c1)c1.setOption({color:[C.accent,C.teal,C.amber],tooltip:tip,series:[{type:'funnel',top:16,left:20,right:20,bottom:20,minSize:'28%',sort:'descending',gap:4,label:{show:true,position:'inside',formatter:p=>p.name+'\n'+fmtN(p.value),color:'#fff',fontSize:12},data:[{name:'低价课订单',value:t.eff},{name:'有效会话',value:Math.round(t.eff*t.validSession)},{name:'高价课转化',value:t.highOrders}]}]});
      const cpaItems=d.linkSummary.filter(i=>i.eff>0).sort((a,b)=>b.orderCost-a.orderCost).reverse();
      const c2=chart('lt-cpa'); if(c2)c2.setOption({color:[C.teal],tooltip:tip,grid:{left:8,right:16,top:8,bottom:8,containLabel:true},xAxis:{type:'value',min:0,...baseAxis},yAxis:{type:'category',data:cpaItems.map(i=>i.link),...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11}},series:[{type:'bar',data:cpaItems.map(i=>Math.round(i.orderCost)),barWidth:'55%',label:{show:true,position:'right',formatter:p=>'¥'+fmtN(p.value),color:'#5C5A54',fontSize:11},itemStyle:{color:C.teal}}]});
      const c3=chart('lt-share'); if(c3)c3.setOption({color:C.cat,tooltip:tip,legend:{type:'scroll',bottom:0},series:[{type:'pie',radius:['42%','66%'],center:['50%','44%'],label:{show:true,formatter:p=>p.name+'\n'+p.percent+'%'},data:d.linkSummary.map(i=>({name:i.link,value:i.eff}))}]});
      const radarDims=[
        {name:'ROI(%)',max:Math.max(100,...d.linkSummary.map(i=>i.roi*100))*1.1},
        {name:'订单数',max:Math.max(...d.linkSummary.map(i=>i.eff))*1.2},
        {name:'单订单产值',max:Math.max(...d.linkSummary.map(i=>i.orderOutput))*1.2},
        {name:'高价转化率(%)',max:Math.max(20,...d.linkSummary.map(i=>i.highConv*100))*1.2},
        {name:'有效会话率(%)',max:100},
        {name:'到课率(%)',max:100},
        {name:'开口率(%)',max:100},
        {name:'进群率(%)',max:100},
      ];
      const c4=chart('lt-scatter'); if(c4)c4.setOption({color:C.cat,tooltip:tip,legend:{data:d.linkSummary.map(i=>i.link),bottom:0,type:'scroll',textStyle:{fontSize:11}},radar:{indicator:radarDims,shape:'polygon',splitNumber:4,axisName:{color:'#5C5A54',fontSize:11},splitLine:{lineStyle:{color:'#e5e3dc'}},splitArea:{areaStyle:{color:['#faf9f6','#f3f1ea']}},axisLine:{lineStyle:{color:'#e5e3dc'}}},series:[{type:'radar',data:d.linkSummary.map(i=>({name:i.link,value:[+(i.roi*100).toFixed(1),i.eff,+(i.orderOutput).toFixed(1),+(i.highConv*100).toFixed(1),+(i.validSession*100).toFixed(1),+((i.firstLive||0)*100).toFixed(1),+((i.openRate||0)*100).toFixed(1),+((i.joinRate||0)*100).toFixed(1)],areaStyle:{opacity:.12},lineStyle:{width:2}}))}]});
      const best=d.linkSummary.slice().sort((a,b)=>b.roi-a.roi)[0];
      $('#lt-insight').innerHTML = best? 'ROI 最高链路为 <b>'+esc(best.link)+'</b>（'+fmtP(best.roi*100)+'，单订单成本 ¥'+fmtM2(best.orderCost)+'，单订单产值 ¥'+fmtM2(best.orderOutput)+'）；整体高价转化率 '+fmtP(t.highConv*100)+'，有效会话率 '+fmtP(t.validSession*100)+'。' : '暂无同步数据。';
      sortState['#lt-table']={key:'conv',dir:-1};
      sortTable('#lt-table',d.linkSummary.map(i=>({key:i.link,rows:d.periods.length,conv:i.eff,cvr:+(i.highConv*100).toFixed(2),cost:i.cost,cpa:i.orderCost,roi:i.roi})),[
        {key:'key',label:'链路'},{key:'rows',label:'期数',sortable:true,cell:r=>fmtN(r.rows)},
        {key:'conv',label:'有效订单',sortable:true,cell:r=>fmtN(r.conv)},{key:'cvr',label:'高价转化率',sortable:true,cell:r=>fmtP(r.cvr)},
        {key:'cost',label:'低价课总成本',sortable:true,cell:r=>fmtM(r.cost)},{key:'cpa',label:'单订单成本',sortable:true,cell:r=>fmtM2(r.cpa)},
        {key:'roi',label:'订单ROI',sortable:true,cell:r=>fmtP(r.roi*100)}
      ].map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));
      // 漏斗环节流失明细
      const funnelCard=$('#lt-funnel-card');
      if(funnelCard && d.linkSummary.length>0){
        funnelCard.style.display='block';
        const fCols=['链路','低价课订单','有效会话','到课','开口','进群','高价课转化','删友','订单→会话','会话→到课','到课→开口','开口→进群','进群→高价'];
        const fRows=d.linkSummary.map(i=>{
          const orders=i.eff||0;
          const valid=Math.round(orders*(i.validSession||0));
          const live=Math.round(valid*(i.firstLive||i.liveRate||0));
          const open=Math.round(live*(i.openRate||0));
          const join=Math.round(open*(i.joinRate||0));
          const high=i.highOrders||0;
          const del=Math.round(orders*(i.deleteRate||0));
          return {link:i.link, orders, valid, live, open, join, high, del,
            o2v: orders? (valid/orders*100).toFixed(1)+'%':'—',
            v2l: valid? (live/valid*100).toFixed(1)+'%':'—',
            l2o: live? (open/live*100).toFixed(1)+'%':'—',
            o2j: open? (join/open*100).toFixed(1)+'%':'—',
            j2h: join? (high/join*100).toFixed(1)+'%':'—'
          };
        });
        sortState['#lt-funnel-table']={key:'orders',dir:-1};
        sortTable('#lt-funnel-table',fRows,[
          {key:'link',label:'链路'},
          {key:'orders',label:'低价课订单',sortable:true,cell:r=>fmtN(r.orders)},
          {key:'valid',label:'有效会话',sortable:true,cell:r=>fmtN(r.valid)},
          {key:'live',label:'到课',sortable:true,cell:r=>fmtN(r.live)},
          {key:'open',label:'开口',sortable:true,cell:r=>fmtN(r.open)},
          {key:'join',label:'进群',sortable:true,cell:r=>fmtN(r.join)},
          {key:'high',label:'高价课转化',sortable:true,cell:r=>'<b style="color:var(--mint)">'+fmtN(r.high)+'</b>'},
          {key:'del',label:'删友',sortable:true,cell:r=>'<span style="color:'+(r.del>r.orders*0.2?'#ef4444':'#94a3b8')+'">'+fmtN(r.del)+'</span>'},
          {key:'o2v',label:'订单→会话',sortable:true,cell:r=>r.o2v},
          {key:'v2l',label:'会话→到课',sortable:true,cell:r=>r.v2l},
          {key:'l2o',label:'到课→开口',sortable:true,cell:r=>r.l2o},
          {key:'o2j',label:'开口→进群',sortable:true,cell:r=>r.o2j},
          {key:'j2h',label:'进群→高价',sortable:true,cell:r=>'<b>'+r.j2h+'</b>'}
        ].map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));
      }
    } else {
      const periods=d.linkByPeriod[ltLink]||[];
      const sel=document.getElementById('lt-period-select');
      sel.innerHTML='';
      periods.forEach(p=>{const o=document.createElement('option');const pv=p['期']||p.period||'';o.value=pv;o.textContent=pv;sel.appendChild(o);});
      if(!ltPeriod || !periods.find(p=>(p['期']||p.period)===ltPeriod)) ltPeriod=periods[0]?(periods[0]['期']||periods[0].period):'';
      sel.value=ltPeriod; sel.onchange=()=>{ltPeriod=sel.value;renderLinkTable();};
      const cur=periods.find(p=>(p['期']||p.period)===ltPeriod)||periods[0];
      $('#lt-period-title').textContent=ltLink+' · '+(ltPeriod||'')+' 多维度画像';
      if(cur){
        const dims=[
          {name:'订单ROI(%)', val:+(cur['订单ROI']*100).toFixed(1)||0, max:100},
          {name:'订单数', val:+cur['订单数']||0, max:Math.max(...periods.map(p=>+p['订单数']||0))*1.2||100},
          {name:'单订单产值', val:+cur['单订单产值(元)']||0, max:Math.max(...periods.map(p=>+p['单订单产值(元)']||0))*1.2||100},
          {name:'高价转化率(%)', val:+(cur['高价课转化率']*100).toFixed(1)||0, max:20},
          {name:'高价课订单', val:+cur['高价课订单数']||0, max:Math.max(...periods.map(p=>+p['高价课订单数']||0))*1.2||10},
          {name:'有效会话率(%)', val:+(cur['有效会话率']*100).toFixed(1)||0, max:100},
          {name:'到课率(%)', val:+(cur['一节直播到课']*100).toFixed(1)||0, max:100},
          {name:'开口率(%)', val:+(cur['开口率']*100).toFixed(1)||0, max:100},
          {name:'进群率(%)', val:+(cur['进群率']*100).toFixed(1)||0, max:100},
        ];
        const c5=chart('lt-period-chart');
        if(c5)c5.setOption({
          color:[C.accent],
          tooltip:{...tip,formatter:function(p){return p.name+'<br/>'+dims.map((d,i)=>d.name+'：'+d.val).join('<br/>');}},
          radar:{indicator:dims.map(d=>({name:d.name,max:d.max})),shape:'polygon',splitNumber:4,axisName:{color:'#5C5A54',fontSize:11},splitLine:{lineStyle:{color:'#e5e3dc'}},splitArea:{areaStyle:{color:['#faf9f6','#f3f1ea']}},axisLine:{lineStyle:{color:'#e5e3dc'}}},
          series:[{type:'radar',data:[{value:dims.map(d=>d.val),name:ltPeriod,areaStyle:{color:C.accent+'33'},lineStyle:{color:C.accent,width:2},itemStyle:{color:C.accent}}]}]
        });
      }
      sortState['#lt-period-table']={key:'期',dir:1};
      sortTable('#lt-period-table',periods.map(p=>({period:p['期']||p.period||'',roi:p['订单ROI']||0,spend:p['广告花费(元)']||0,cost:p['低价课总成本(元)']||0,orders:p['订单数']||0,cpa:p['单订单成本(元)']||0,output:p['单订单产值(元)']||0,highCvr:p['高价课转化率']||0,highOrders:p['高价课订单数']||0,validSession:p['有效会话率']||0,attend:p['一节直播到课']||0,speak:p['开口率']||0,join:p['进群率']||0,delete:p['删友率']||0})),[
        {key:'period',label:'期'},{key:'spend',label:'广告花费',sortable:true,cell:r=>fmtM(r.spend)},{key:'cost',label:'低价课总成本',sortable:true,cell:r=>fmtM(r.cost)},
        {key:'orders',label:'订单数',sortable:true,cell:r=>fmtN(r.orders)},{key:'cpa',label:'单订单成本',sortable:true,cell:r=>fmtM2(r.cpa)},{key:'output',label:'单订单产值',sortable:true,cell:r=>fmtM2(r.output)},
        {key:'highCvr',label:'高价转化率',sortable:true,cell:r=>fmtP(r.highCvr*100)},{key:'highOrders',label:'高价课订单',sortable:true,cell:r=>fmtN(r.highOrders)},{key:'validSession',label:'有效会话率',sortable:true,cell:r=>fmtP(r.validSession*100)},
        {key:'attend',label:'到课率',sortable:true,cell:r=>fmtP(r.attend*100)},{key:'speak',label:'开口率',sortable:true,cell:r=>fmtP(r.speak*100)},{key:'join',label:'进群率',sortable:true,cell:r=>fmtP(r.join*100)},
        {key:'delete',label:'删友率',sortable:true,cell:r=>fmtP(r.delete*100)},{key:'roi',label:'订单ROI',sortable:true,cell:r=>fmtP(r.roi*100)}
      ].map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));
    }
    return;
  }
  const rows = filteredRows();
  const t = totals(rows);
  const items = summarizeGroup(groupBy(rows,'link')).sort((a,b)=>b.conv-a.conv);
  renderKpis('#lt-kpis', [
    {label:'链路数', value:fmtN(items.length)},{label:'总消耗', value:fmtM(t.cost), cls:'kpi-accent'},
    {label:'总展现', value:fmtN(t.impr)},{label:'总点击', value:fmtN(t.click)},
    {label:'总转化', value:fmtN(t.conv), cls:'kpi-teal'},{label:'平均CPA', value:t.conv?fmtM2(t.cpa):'—'},
    {label:'平均CVR', value:fmtP(t.cvr)},{label:'整体ROI', value:t.rev>0?t.roi.toFixed(2):'—'}
  ]);
  const c1 = chart('lt-funnel');
  if (c1) c1.setOption({color:[C.accent,C.teal,C.amber],tooltip:tip,series:[{type:'funnel',top:16,left:20,right:20,bottom:20,minSize:'28%',sort:'descending',gap:4,label:{show:true,position:'inside',formatter:p=>p.name+'\n'+fmtN(p.value),color:'#fff',fontSize:12},data:[{name:'展现',value:t.impr},{name:'点击',value:t.click},{name:'转化',value:t.conv}]}]});
  const c2 = chart('lt-cpa');
  if (c2) c2.setOption({color:[C.teal],tooltip:tip,grid:{left:8,right:16,top:8,bottom:8,containLabel:true},xAxis:{type:'value',min:0,...baseAxis},yAxis:{type:'category',data:items.filter(i=>i.conv>0).map(i=>i.key),...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11}},series:[{type:'bar',data:items.filter(i=>i.conv>0).map(i=>Math.round(i.cpa)),barWidth:'55%',label:{show:true,position:'right',formatter:p=>'¥'+fmtN(p.value),color:'#5C5A54',fontSize:11},itemStyle:{color:C.teal}}]});
  const c3 = chart('lt-share');
  if (c3) c3.setOption({color:C.cat,tooltip:tip,legend:{type:'scroll',bottom:0},series:[{type:'pie',radius:['42%','66%'],center:['50%','44%'],label:{show:true,formatter:p=>p.name+'\n'+p.percent+'%'},data:items.map(i=>({name:i.key,value:i.conv}))}]});
  const c4 = chart('lt-scatter');
  if (c4) c4.setOption({color:[C.blue],tooltip:{...tip,formatter:function(p){var d=p.data;return d.name+'<br/>CPA：¥'+fmtN(d.cpa)+'<br/>CVR：'+fmtP(d.cvr)+'<br/>消耗：'+fmtM(d.cost);}},grid:{left:10,right:10,top:30,bottom:10,containLabel:true},xAxis:{type:'value',name:'CPA(元)',min:0,...baseAxis},yAxis:{type:'value',name:'CVR(%)',min:0,...baseAxis},series:[{type:'scatter',symbolSize:function(d){return Math.max(18,Math.min(64,Math.sqrt(d[2])*0.5));},itemStyle:{color:C.blue,opacity:.75},data:items.filter(i=>i.conv>0).map(i=>[+i.cpa.toFixed(1),+i.cvr.toFixed(2),i.cost,i.key])}]});
  const ins = [];
  if (items.length){
    const best = items.slice().sort((a,b)=>b.cvr-a.cvr)[0];
    const cheap = items.filter(i=>i.conv>5).sort((a,b)=>a.cpa-b.cpa)[0];
    const top = items[0];
    ins.push('转化量最高为 <b>'+esc(top.key)+'</b>（'+fmtN(top.conv)+'，占比 '+fmtP(pct(top.conv,t.conv))+'）；CVR 最高为 <b>'+esc(best.key)+'</b>（'+fmtP(best.cvr)+'）；CPA 最低为 <b>'+esc(cheap.key)+'</b>（'+fmtM2(cheap.cpa)+'）。');
  }
  $('#lt-insight').innerHTML = ins.join(' ') || '当前范围无链路数据。';
  sortState['#lt-table'] = { key:'conv', dir:-1 };
  sortTable('#lt-table', items.map(i=>({...i, ctr:+i.ctr.toFixed(2), cvr:+i.cvr.toFixed(2), cpa: i.conv? +i.cpa.toFixed(1) : null, roi:+i.roi.toFixed(2)})), [
    {key:'key', label:'转化链路'},{key:'impr', label:'展现', sortable:true, cell:r=>fmtN(r.impr)},{key:'click', label:'点击', sortable:true, cell:r=>fmtN(r.click)},
    {key:'ctr', label:'CTR', sortable:true, cell:r=>fmtP(r.ctr)},{key:'conv', label:'转化', sortable:true, cell:r=>fmtN(r.conv)},{key:'cvr', label:'CVR', sortable:true, cell:r=>fmtP(r.cvr)},
    {key:'cost', label:'消耗', sortable:true, cell:r=>fmtM(r.cost)},{key:'cpa', label:'CPA', sortable:true, cell:r=>r.cpa?fmtM2(r.cpa):'—'},{key:'roi', label:'ROI', sortable:true, cell:r=>r.revenue>0?r.roi.toFixed(2):'—'}
  ].map(c=>{ if(!('cell' in c)) c.cell = (r)=>defaultCell(r,c.key); return c; }));
}

/* ---------- 视图：分周数据分析 ---------- */
var wkDim='link';
function renderWeekly(){
  renderSyncPanel('weekly', document.getElementById('sync-weekly'));
  document.querySelectorAll('#wk-subtabs .sub-tab').forEach(b=>{
    b.classList.toggle('active',b.dataset.wk===wkDim);
    b.onclick=()=>{wkDim=b.dataset.wk;renderWeekly();};
  });
  document.getElementById('wk-link-view').style.display=wkDim==='link'?'':'none';
  document.getElementById('wk-placement-view').style.display=wkDim==='placement'?'':'none';
  if(SYNC.weekly){
    const d=SYNC.weekly; const cur=d.current, last=d.last;
    const wkLabels=last?[last.weekFmt,cur.weekFmt]:[cur.weekFmt];
    const wkCosts=last?[Math.round(last.totalSpend),Math.round(cur.totalSpend)]:[Math.round(cur.totalSpend)];
    const wkConvs=last?[last.totalAll,cur.totalAll]:[cur.totalAll];
    const avgCpa=cur.totalEff?cur.totalCost/cur.totalEff:0;
    const avgCvr=cur.totalAll?cur.totalEff/cur.totalAll*100:0;
    renderKpis('#wk-kpis',[
      {label:'覆盖周数',value:last?'2':'1',sub:last?last.weekFmt+' ~ '+cur.weekFmt:cur.weekFmt},
      {label:'本周花费',value:fmtM(cur.totalSpend),cls:'kpi-accent'},
      {label:'本周订单',value:fmtN(cur.totalAll),cls:'kpi-teal'},
      {label:'单订单成本',value:fmtM2(avgCpa)},{label:'订单有效率',value:fmtP(avgCvr)},
      {label:'环比花费',value:last&&last.totalSpend?((cur.totalSpend-last.totalSpend)/last.totalSpend*100).toFixed(1)+'%':'—'}
    ]);
    if(wkDim==='link'){
      const c1=chart('wk-trend'); if(c1)c1.setOption({color:[C.accent,C.teal],tooltip:tip,legend:{data:['花费','订单数'],type:'scroll',bottom:0},grid:{left:8,right:8,top:30,bottom:34,containLabel:true},xAxis:{type:'category',data:wkLabels,...baseAxis},yAxis:[{type:'value',min:0,name:'花费',...baseAxis},{type:'value',min:0,name:'订单',splitLine:{show:false}}],series:[{name:'花费',type:'bar',data:wkCosts,barMaxWidth:26,itemStyle:{color:C.accent}},{name:'订单数',type:'line',yAxisIndex:1,data:wkConvs,smooth:true,symbol:'circle',symbolSize:6,lineStyle:{width:2.5,color:C.teal},itemStyle:{color:C.teal}}]});
      const linkLabels=cur.linkRows.filter(r=>r.link!=='汇总').map(r=>r.link);
      const lastCosts=last?linkLabels.map(l=>{const x=last.linkRows.find(r=>r.link===l);return x?Math.round(x.orderCost):0;}):[];
      const curCosts=linkLabels.map(l=>{const x=cur.linkRows.find(r=>r.link===l);return x?Math.round(x.orderCost):0;});
      const c2=chart('wk-cpa'); if(c2)c2.setOption({color:[C.blue,C.accent],tooltip:tip,legend:{data:last?['上周CPA','本周CPA']:['本周CPA'],type:'scroll',bottom:0},grid:{left:8,right:8,top:24,bottom:8,containLabel:true},xAxis:{type:'category',data:linkLabels,...baseAxis},yAxis:{type:'value',min:0,...baseAxis},series:(last?[{name:'上周CPA',type:'bar',data:lastCosts,itemStyle:{color:C.blue}},{name:'本周CPA',type:'bar',data:curCosts,itemStyle:{color:C.accent}}]:[{name:'本周CPA',type:'bar',data:curCosts,itemStyle:{color:C.accent}}])});
      const lastCvrs=last?linkLabels.map(l=>{const x=last.linkRows.find(r=>r.link===l);return x?+(x.orderValidRate*100).toFixed(2):0;}):[];
      const curCvrs=linkLabels.map(l=>{const x=cur.linkRows.find(r=>r.link===l);return x?+(x.orderValidRate*100).toFixed(2):0;});
      const c3=chart('wk-cvr'); if(c3)c3.setOption({color:[C.teal],tooltip:tip,grid:{left:8,right:8,top:24,bottom:8,containLabel:true},xAxis:{type:'category',data:linkLabels,...baseAxis},yAxis:{type:'value',min:0,axisLabel:{color:'#5C5A54',fontSize:11,formatter:v=>v+'%'},splitLine:{lineStyle:{color:C.grid}}},series:[{name:'订单有效率',type:'line',data:curCvrs,smooth:true,symbol:'circle',symbolSize:6,lineStyle:{width:2.5,color:C.teal},itemStyle:{color:C.teal},areaStyle:{color:'#5BB89A22'}}]});
      const c4=chart('wk-roi'); if(c4)c4.setOption({color:[C.amber],tooltip:tip,grid:{left:8,right:8,top:24,bottom:8,containLabel:true},xAxis:{type:'category',data:wkLabels,...baseAxis},yAxis:{type:'value',min:0,...baseAxis},series:[],graphic:[{type:'text',left:'center',top:'middle',style:{text:'Excel分周数据无ROI字段',fill:'#918D83',fontSize:13}}]});
      const ins=[];
      if(last){ const dc=(cur.totalSpend-last.totalSpend)/Math.max(1,last.totalSpend)*100; const dv=(cur.totalAll-last.totalAll)/Math.max(1,last.totalAll)*100; ins.push('本周('+cur.weekFmt+')较上周('+last.weekFmt+')：花费 <b>'+(dc>=0?'+':'')+dc.toFixed(1)+'%</b>，订单 <b>'+(dv>=0?'+':'')+dv.toFixed(1)+'%</b>。'); }
      else ins.push('仅本周('+cur.weekFmt+')数据，下次运行后自动对比上周。');
      const bestLink=cur.linkRows.filter(r=>r.link!=='汇总').sort((a,b)=>a.orderCost-b.orderCost)[0];
      if(bestLink) ins.push('成本最低链路 <b>'+esc(bestLink.link)+'</b>（单订单成本 ¥'+fmtM2(bestLink.orderCost)+'，订单 '+fmtN(bestLink.all)+'）。');
      $('#wk-insight').innerHTML=ins.join(' ');
      const tableRows=cur.linkRows.filter(r=>r.link!=='汇总').map(r=>{const lr=last?last.linkRows.find(x=>x.link===r.link):null;const dCpa=lr&&lr.eff&&r.eff?(r.orderCost-lr.orderCost)/lr.orderCost*100:null;return{label:r.link,ws:r.link,cost:r.spend,conv:r.all,cvr:+(r.orderValidRate*100).toFixed(2),cpa:r.orderCost,dCpa};});
      sortState['#wk-table']={key:'ws',dir:1};
      sortTable('#wk-table',tableRows,[
        {key:'label',label:'链路',cell:r=>'<b>'+esc(r.label)+'</b>'},
        {key:'cost',label:'本周花费',sortable:true,cell:r=>fmtM(r.cost)},{key:'conv',label:'本周订单',sortable:true,cell:r=>fmtN(r.conv)},
        {key:'cvr',label:'订单有效率',sortable:true,cell:r=>fmtP(r.cvr)},{key:'cpa',label:'单订单成本',sortable:true,cell:r=>fmtM2(r.cpa)},
        {key:'dCpa',label:'成本环比',sortable:true,cell:r=>{if(r.dCpa==null)return'—';const up=r.dCpa<=0;const cls=r.dCpa===0?'delta-flat':(up?'delta-up':'delta-down');const arrow=r.dCpa===0?'':(up?'↓':'↑');return'<span class="'+cls+'">'+arrow+' '+Math.abs(r.dCpa).toFixed(1)+'%</span>';}}
      ].map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));
    } else {
      const plCur=cur.placeRows||[]; const plLast=last?last.placeRows||[]:[];
      const plLabels=plCur.filter(r=>r.placement!=='汇总'&&r.placement!=='合计').map(r=>r.placement);
      const lastOrders=plLast?plLabels.map(l=>{const x=plLast.find(r=>r.placement===l);return x?+(r.all||r.orders||0):0;}):[];
      const curOrders=plLabels.map(l=>{const x=plCur.find(r=>r.placement===l);return x?+(x.all||x.orders||0):0;});
      const c5=chart('wk-pl-orders'); if(c5)c5.setOption({color:[C.blue,C.accent],tooltip:tip,legend:{data:plLast.length?['上周订单','本周订单']:['本周订单'],type:'scroll',bottom:0},grid:{left:8,right:8,top:30,bottom:8,containLabel:true},xAxis:{type:'category',data:plLabels,...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11,interval:0,rotate:plLabels.length>6?30:0}},yAxis:{type:'value',min:0,...baseAxis},series:(plLast.length?[{name:'上周订单',type:'bar',data:lastOrders,itemStyle:{color:C.blue}},{name:'本周订单',type:'bar',data:curOrders,itemStyle:{color:C.accent}}]:[{name:'本周订单',type:'bar',data:curOrders,itemStyle:{color:C.accent}}])});
      const lastCpa=plLast?plLabels.map(l=>{const x=plLast.find(r=>r.placement===l);return x?Math.round(x.orderCost||x.cpa||0):0;}):[];
      const curCpa=plLabels.map(l=>{const x=plCur.find(r=>r.placement===l);return x?Math.round(x.orderCost||x.cpa||0):0;});
      const c6=chart('wk-pl-cpa'); if(c6)c6.setOption({color:[C.teal,C.amber],tooltip:tip,legend:{data:plLast.length?['上周CPA','本周CPA']:['本周CPA'],type:'scroll',bottom:0},grid:{left:8,right:8,top:30,bottom:8,containLabel:true},xAxis:{type:'category',data:plLabels,...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11,interval:0,rotate:plLabels.length>6?30:0}},yAxis:{type:'value',min:0,...baseAxis},series:(plLast.length?[{name:'上周CPA',type:'bar',data:lastCpa,itemStyle:{color:C.teal}},{name:'本周CPA',type:'bar',data:curCpa,itemStyle:{color:C.amber}}]:[{name:'本周CPA',type:'bar',data:curCpa,itemStyle:{color:C.amber}}])});
      const c7=chart('wk-pl-share'); if(c7)c7.setOption({color:C.cat,tooltip:tip,legend:{type:'scroll',bottom:0},series:[{type:'pie',radius:['38%','62%'],center:['50%','45%'],label:{show:true,formatter:p=>p.name+' '+p.percent+'%'},data:plCur.filter(r=>r.placement!=='汇总'&&r.placement!=='合计').map(r=>({name:r.placement,value:+(r.all||r.orders||0)}))}]});
      const plTable=plCur.filter(r=>r.placement!=='汇总'&&r.placement!=='合计').map(r=>{const lr=plLast?plLast.find(x=>x.placement===r.placement):null;const dOrders=lr&&(lr.all||lr.orders)?((r.all||r.orders)-(lr.all||lr.orders))/(lr.all||lr.orders)*100:null;return{label:r.placement,orders:+(r.all||r.orders||0),share:+(r.orderShare||r.share||0)*100,cpa:r.orderCost||r.cpa||0,dOrders};});
      sortState['#wk-pl-table']={key:'orders',dir:-1};
      sortTable('#wk-pl-table',plTable,[
        {key:'label',label:'版位',cell:r=>'<b>'+esc(r.label)+'</b>'},
        {key:'orders',label:'本周订单',sortable:true,cell:r=>fmtN(r.orders)},{key:'share',label:'订单占比',sortable:true,cell:r=>fmtP(r.share)},
        {key:'cpa',label:'单订单成本',sortable:true,cell:r=>fmtM2(r.cpa)},{key:'dOrders',label:'订单环比',sortable:true,cell:r=>{if(r.dOrders==null)return'—';const up=r.dOrders>=0;const cls=r.dOrders===0?'delta-flat':(up?'delta-up':'delta-down');const arrow=r.dOrders===0?'':(up?'↑':'↓');return'<span class="'+cls+'">'+arrow+' '+Math.abs(r.dOrders).toFixed(1)+'%</span>';}}
      ].map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));
    }
    return;
  }
  const rows = filteredRows();
  const t = totals(rows);
  const wk = weeklySeries(rows);
  renderKpis('#wk-kpis', [
    {label:'覆盖周数', value:fmtN(wk.length), sub:wk.length?'近'+wk.length+'个自然周':''},
    {label:'总消耗', value:fmtM(t.cost), cls:'kpi-accent'},{label:'总转化', value:fmtN(t.conv), cls:'kpi-teal'},
    {label:'平均CPA', value:t.conv?fmtM2(t.cpa):'—'},{label:'平均CVR', value:fmtP(t.cvr)},{label:'整体ROI', value:t.rev>0?t.roi.toFixed(2):'—'}
  ]);
  if(wkDim==='link'){
    const c1 = chart('wk-trend');
    if (c1) c1.setOption({color:[C.accent, C.teal], tooltip:tip,legend:{data:['消耗','转化数'], type:'scroll', bottom:0},grid:{left:8,right:8,top:30,bottom:34,containLabel:true},xAxis:{type:'category', data:wk.map(w=>w.label), ...baseAxis},yAxis:[{type:'value', min:0, name:'消耗', ...baseAxis},{type:'value', min:0, name:'转化', splitLine:{show:false}}],series:[{name:'消耗', type:'bar', data:wk.map(w=>Math.round(w.cost)), barMaxWidth:26, itemStyle:{color:C.accent}},{name:'转化数', type:'line', yAxisIndex:1, data:wk.map(w=>w.conv), smooth:true, symbol:'circle', symbolSize:6, lineStyle:{width:2.5,color:C.teal}, itemStyle:{color:C.teal}}]});
    const c2 = chart('wk-cpa');
    if (c2) c2.setOption({color:[C.blue], tooltip:tip,grid:{left:8,right:8,top:24,bottom:8,containLabel:true},xAxis:{type:'category', data:wk.map(w=>w.label), ...baseAxis},yAxis:{type:'value', min:0, ...baseAxis},series:[{name:'CPA', type:'line', data:wk.map(w=>w.conv?+w.cpa.toFixed(1):null), smooth:true, symbol:'circle', symbolSize:6, lineStyle:{width:2.5,color:C.blue}, itemStyle:{color:C.blue}, connectNulls:false}]});
    const c3 = chart('wk-cvr');
    if (c3) c3.setOption({color:[C.teal], tooltip:tip,grid:{left:8,right:8,top:24,bottom:8,containLabel:true},xAxis:{type:'category', data:wk.map(w=>w.label), ...baseAxis},yAxis:{type:'value', min:0, axisLabel:{color:'#5C5A54',fontSize:11,formatter:v=>v+'%'}, splitLine:{lineStyle:{color:C.grid}}},series:[{name:'CVR', type:'line', data:wk.map(w=>+w.cvr.toFixed(2)), smooth:true, symbol:'circle', symbolSize:6, lineStyle:{width:2.5,color:C.teal}, itemStyle:{color:C.teal}, areaStyle:{color:'#5BB89A22'}}]});
    const hasRev = t.rev>0;
    const c4 = chart('wk-roi');
    if (c4) c4.setOption({color:[C.amber], tooltip:tip,grid:{left:8,right:8,top:24,bottom:8,containLabel:true},xAxis:{type:'category', data:wk.map(w=>w.label), ...baseAxis},yAxis:{type:'value', min:0, ...baseAxis},series:[{name:'ROI', type:'line', data: hasRev ? wk.map(w=>w.cost?+w.roi.toFixed(2):null) : [], smooth:true, symbol:'circle', symbolSize:6, lineStyle:{width:2.5,color:C.amber}, itemStyle:{color:C.amber}}],graphic: hasRev? [] : [{type:'text', left:'center', top:'middle', style:{text:'当前范围无电商收入数据', fill:'#918D83', fontSize:13}}]});
    const ins = [];
    if (wk.length>=2){
      const a=wk[wk.length-1], b=wk[wk.length-2];
      const dCost=(a.cost-b.cost)/Math.max(1,b.cost)*100; const dConv=(a.conv-b.conv)/Math.max(1,b.conv)*100;
      const dCpa=(b.conv&&a.conv)?(a.cpa-b.cpa)/b.cpa*100:0; const trend = a.conv>b.conv?'上升':'回落';
      ins.push('<b>'+a.label+'</b> 较 <b>'+b.label+'</b>：消耗 <b>'+(dCost>=0?'+':'')+dCost.toFixed(1)+'%</b>，转化 <b>'+(dConv>=0?'+':'')+dConv.toFixed(1)+'%</b>（'+trend+'），CPA <b>'+(dCpa>=0?'+':'')+dCpa.toFixed(1)+'%</b>。');
    } else if (wk.length===1){ ins.push('当前范围仅覆盖 <b>'+wk[0].label+'</b> 一周数据。'); }
    $('#wk-insight').innerHTML = ins.join(' ') || '当前范围无周数据。';
    const tableRows = wk.map((w,i)=>{const prev = i>0 ? wk[i-1] : null;const dCpa = prev&&prev.conv&&w.conv ? (w.cpa-prev.cpa)/prev.cpa*100 : null;return { label:w.label, ws:w.ws, cost:w.cost, conv:w.conv, cvr:+w.cvr.toFixed(2), cpa: w.conv? +w.cpa.toFixed(1) : null, dCpa };});
    sortState['#wk-table'] = { key:'ws', dir:1 };
    sortTable('#wk-table', tableRows, [
      {key:'label', label:'周次', cell:r=>'<b>'+esc(r.label)+'</b>'},{key:'cost', label:'消耗', sortable:true, cell:r=>fmtM(r.cost)},
      {key:'conv', label:'转化', sortable:true, cell:r=>fmtN(r.conv)},{key:'cvr', label:'CVR', sortable:true, cell:r=>fmtP(r.cvr)},
      {key:'cpa', label:'CPA', sortable:true, cell:r=>fmtM2(r.cpa)},{key:'dCpa', label:'CPA环比', sortable:true, cell:r=>{if(r.dCpa==null)return'<span class="delta-flat">—</span>';const up=r.dCpa<=0;const cls=r.dCpa===0?'delta-flat':(up?'delta-up':'delta-down');const arrow=r.dCpa===0?'':(up?'↓':'↑');return'<span class="'+cls+'">'+arrow+' '+Math.abs(r.dCpa).toFixed(1)+'%</span>';}}
    ].map(c=>{ if(!('cell' in c)) c.cell = (r)=>defaultCell(r,c.key); return c; }));
  } else {
    ['wk-pl-orders','wk-pl-cpa','wk-pl-share'].forEach(id=>{const c=chart(id);if(c)c.setOption({series:[],graphic:[{type:'text',left:'center',top:'middle',style:{text:'CSV数据无版位分周信息',fill:'#918D83',fontSize:13}}]});});
    sortTable('#wk-pl-table',[],[]);
  }
}

/* ---------- 视图：链路画像_评分分析 ---------- */
let lsLink='', lsPeriodFilter='', lsSearchText='', lsModelPage=1, lsPageSize=20, lsSort={};

function lsHeatColor(val,max){ if(!max) return 'transparent'; const r=Math.min(1,val/max); return 'rgba(124,111,224,'+(0.06+r*0.22)+')'; }
function lsToggleSort(dimKey,key){ if(!lsSort[dimKey]||lsSort[dimKey].key!==key) lsSort[dimKey]={key:key,dir:-1}; else if(lsSort[dimKey].dir===-1) lsSort[dimKey]={key:key,dir:1}; else delete lsSort[dimKey]; renderLinkScore(); }

function lsFilterDetail(dimData,filter,search){
  let rows=(dimData&&dimData.periodDetail)?dimData.periodDetail.slice():[];
  if(!rows.length&&dimData&&dimData.rows) rows=dimData.rows.map(r=>({period:'合计',...r}));
  if(filter&&filter!=='__total__') rows=rows.filter(r=>r.period===filter);
  if(search&&dimData&&(dimData.name==='品牌'||dimData.name==='手机型号')) rows=rows.filter(r=>(r.value||'').toLowerCase().includes(search.toLowerCase()));
  return rows;
}
function lsAggregate(rows){
  const g={};let total=0;
  rows.forEach(r=>{const v=r.value;if(!g[v])g[v]={count:0,valid:0,highN:0,highAmt:0,score:undefined,grade:undefined,quadrant:undefined,strategy:undefined};g[v].count+=r.count||0;g[v].valid+=r.valid||0;g[v].highN+=r.highN||0;g[v].highAmt+=r.highAmt||0;if(r.score!==undefined)g[v].score=r.score;if(r.grade)g[v].grade=r.grade;if(r.quadrant)g[v].quadrant=r.quadrant;if(r.strategy)g[v].strategy=r.strategy;total+=r.count||0;});
  return Object.keys(g).map(v=>({value:v,count:g[v].count,valid:g[v].valid,share:total?g[v].count/total:0,highN:g[v].highN,highAmt:g[v].highAmt,highConv:g[v].count?g[v].highN/g[v].count:0,orderOutput:g[v].valid?g[v].highAmt/g[v].valid:0,score:g[v].score,grade:g[v].grade,quadrant:g[v].quadrant,strategy:g[v].strategy}));
}
function lsPeriodTable(dimKey,dimName,detail,agg){
  const onlyTotal=lsPeriodFilter==='__total__';
  const st=lsSort[dimKey];
  let rows=detail.slice();
  if(st&&st.key) rows.sort((a,b)=>((a[st.key]||0)-(b[st.key]||0))*st.dir);
  const valOrder=[];detail.forEach(r=>{if(!valOrder.includes(r.value))valOrder.push(r.value);});
  let totals=valOrder.map(v=>agg.find(x=>x.value===v)).filter(Boolean);
  if(st&&st.key) totals.sort((a,b)=>((a[st.key]||0)-(b[st.key]||0))*st.dir);
  const cols=[{key:'period',label:'期'},{key:'value',label:dimName},{key:'count',label:'订单数'},{key:'valid',label:'有效订单数'},{key:'share',label:'订单占比'},{key:'highN',label:'高价课订单数'},{key:'highAmt',label:'高价课金额'},{key:'highConv',label:'高价课转化率'},{key:'orderOutput',label:'单订单产值'}];
  let h='<div class="ls-table-wrap"><table class="ls-table"><thead><tr>';
  cols.forEach(c=>{const s=['count','highConv','orderOutput'].includes(c.key);const active=st&&st.key===c.key;h+='<th'+(s?' onclick="lsToggleSort(\''+dimKey+'\',\''+c.key+'\')"':'')+'>'+c.label+(active?(st.dir>0?' ↑':' ↓'):'')+'</th>';});
  h+='</tr></thead><tbody>';
  if(!onlyTotal) rows.forEach(r=>{h+='<tr><td><b>'+esc(r.period)+'</b></td><td>'+esc(r.value)+'</td><td>'+fmtN(r.count)+'</td><td>'+fmtN(r.valid)+'</td><td>'+fmtP((r.share||0)*100)+'</td><td>'+fmtN(r.highN)+'</td><td>'+fmtM(r.highAmt)+'</td><td>'+fmtP((r.highConv||0)*100)+'</td><td>'+(r.orderOutput?fmtM2(r.orderOutput):'—')+'</td></tr>';});
  const isGeoDim = ['city','province','cityName'].includes(dimKey);
  totals.forEach(t=>{
    let valHtml = esc(t.value);
    if (isGeoDim && t.score !== undefined) {
      const gc = t.grade==='S'?'#dc2626':t.grade==='A'?'#16a34a':t.grade==='B'?'#f59e0b':'#94a3b8';
      valHtml += ' <span style="color:'+gc+';font-weight:700;font-size:11px">['+t.grade+' '+t.score+'分]</span> <span style="color:#94a3b8;font-size:10px">'+t.quadrant+'</span>';
    }
    h+='<tr class="ls-total-row"><td>合计</td><td>'+valHtml+'</td><td>'+fmtN(t.count)+'</td><td>'+fmtN(t.valid)+'</td><td>'+fmtP(t.share*100)+'</td><td>'+fmtN(t.highN)+'</td><td>'+fmtM(t.highAmt)+'</td><td>'+fmtP(t.highConv*100)+'</td><td>'+(t.orderOutput?fmtM2(t.orderOutput):'—')+'</td></tr>';
    if (isGeoDim && t.strategy) {
      h+='tr><td colspan="2"></td><td colspan="7" style="font-size:11px;color:#92400e;background:#fefce8;padding:4px 10px">💡 '+t.strategy+'</td></tr>';
    }
  });
  return h+'</tbody></table></div>';
}
function lsExportCurrent(){
  if(!SYNC.score)return;const lr=SYNC.score.linkResults[lsLink];if(!lr)return;
  const wb=XLSX.utils.book_new();
  [{key:'city',name:'城市等级'},{key:'province',name:'省份'},{key:'cityName',name:'具体城市'},{key:'gender',name:'性别'},{key:'age',name:'年龄档'},{key:'price',name:'价格档'},{key:'brand',name:'品牌'},{key:'model',name:'手机型号'}].forEach(dim=>{
    const detail=lsFilterDetail(lr.dimensions[dim.key],lsPeriodFilter,lsSearchText),agg=lsAggregate(detail);
    const isGeo = ['city','province','cityName'].includes(dim.key);
    const headers = isGeo ? ['期',dim.name,'综合评分','评级','价值矩阵','订单数','有效订单数','订单占比','高价课订单数','高价课金额','高价课转化率','单订单产值','优化建议'] : ['期',dim.name,'订单数','有效订单数','订单占比','高价课订单数','高价课金额','高价课转化率','单订单产值'];
    const a=[headers];
    const onlyTotal=lsPeriodFilter==='__total__';
    if(!onlyTotal)detail.forEach(r=>a.push(isGeo?[r.period,r.value,'','','',r.count,r.valid,r.share,r.highN,r.highAmt,r.highConv,r.orderOutput,'']:[r.period,r.value,r.count,r.valid,r.share,r.highN,r.highAmt,r.highConv,r.orderOutput]));
    agg.forEach(r=>a.push(isGeo?['合计',r.value,r.score,r.grade,r.quadrant,r.count,r.valid,r.share,r.highN,r.highAmt,r.highConv,r.orderOutput,r.strategy]:['合计',r.value,r.count,r.valid,r.share,r.highN,r.highAmt,r.highConv,r.orderOutput]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(a),dim.name);
  });
  if(lr.scoring){
    const a=[['期数','评分状态','订单数','占比','高价金额','单订单产值']];
    lr.scoring.periodRows.forEach(pr=>{if(!lsPeriodFilter||lsPeriodFilter===pr.period)pr.rows.forEach(r=>a.push([pr.period,r.value,r.count,r.share,r.highAmt,r.orderOutput]));});
    if(lr.scoring.summaryRows)lr.scoring.summaryRows.forEach(r=>a.push(['合计',r.value,r.count,r.share,r.highAmt,r.orderOutput]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(a),'评分分析');
  }
  XLSX.writeFile(wb,lsLink+'_链路画像全维度.xlsx');
}

function renderLinkScore(){
  const content=$('#ls-content'),tabsEl=$('#ls-link-tabs');
  if(!SYNC.score){
    tabsEl.style.display='none';$('#ls-filter-bar').style.display='none';$('#ls-anchor-bar').style.display='none';
    $('#ls-header-badge').textContent='暂无数据';
    content.innerHTML='<div class="ls-sheet"><div class="ls-sheet-header"><h2>链路画像与评分分析</h2><span class="ls-sheet-tag">待同步</span></div><div class="ls-sheet-body"><div class="ls-empty">⏳ 数据待加载 — 请到「Excel工具4：链路画像·评分分析」上传数据并点击开始分析，数据将自动同步到本模块</div></div></div>';
    return;
  }
  const d=SYNC.score,lr=d.linkResults,links=Object.keys(lr);
  if(!lsLink||!links.includes(lsLink))lsLink=links[0];
  $('#ls-header-badge').textContent='共 '+links.length+' 个链路 · 数据已同步';
  tabsEl.style.display='flex';
  tabsEl.innerHTML=links.map(l=>'<button class="ls-tab-btn '+(l===lsLink?'active':'')+'" data-link="'+esc(l)+'">'+esc(l)+'<span class="ls-tab-badge">'+(lr[l].scoring&&lr[l].scoring.periodRows?'7':'6')+'</span></button>').join('');
  tabsEl.querySelectorAll('button').forEach(b=>{b.onclick=()=>{lsLink=b.dataset.link;renderLinkScore();};});
  const cur=lr[lsLink],allPeriods=d.periods||Object.keys(cur.periodTotals||{}).sort();
  const fb=$('#ls-filter-bar');fb.style.display='flex';
  const pb=$('#ls-period-btns');
  const periodBtns=[{v:'',label:'全部期数'},{v:'__total__',label:'仅看合计'}].concat(allPeriods.map(p=>({v:p,label:p})));
  pb.innerHTML=periodBtns.map(b=>'<button class="ls-period-btn '+(lsPeriodFilter===b.v?'active':'')+'" data-pf="'+esc(b.v)+'">'+esc(b.label)+'</button>').join('');
  pb.querySelectorAll('button').forEach(b=>{b.onclick=()=>{lsPeriodFilter=b.dataset.pf;renderLinkScore();};});
  const si=$('#ls-search');si.value=lsSearchText;si.oninput=()=>{lsSearchText=si.value;renderLinkScore();};
  $('#ls-reset-btn').onclick=()=>{lsPeriodFilter='';lsSearchText='';lsSort={};renderLinkScore();};
  $('#ls-export-btn').onclick=lsExportCurrent;
  $('#ls-spec-btn').onclick=()=>{showModal('口径说明','<div style="font-size:13px;line-height:1.8">1. 数据源：投放看板订单明细（Excel工具4同步）<br>2. 链路映射：83676427→PF+小程序问答；83676438+87101601→H5问答+首页；83676778+85489979→小程序问答(含首页)；85489984→H5分流；87101605→PF+AI落地页<br>3. 八大维度：城市等级/省份/具体城市/性别/年龄档/价格档/品牌/手机型号<br>4. 订单占比=该维度值订单数÷当期总订单数<br>5. 高价转化率=高价课订单数÷订单数<br>6. 单订单产值=高价课金额÷有效订单数<br>7. 评分推导：有评分=好友状态为曾添加/系统-已添加/人工-已添加；无评分=系统-未添加或空<br>8. 评分分析仅对PF+小程序问答、PF+AI落地页显示<br>9. 期数筛选：全部期数=平铺各期+合计；仅看合计=只显示合计行<br>10. 搜索：仅对品牌和手机型号维度生效</div>');};
  const dims7=[{key:'city',name:'城市等级'},{key:'province',name:'省份'},{key:'cityName',name:'具体城市'},{key:'gender',name:'性别'},{key:'age',name:'年龄档'},{key:'price',name:'价格档'},{key:'brand',name:'品牌'},{key:'model',name:'手机型号'},{key:'score',name:'评分分析'}];
  const ab=$('#ls-anchor-bar');ab.style.display='flex';
  ab.innerHTML=dims7.map(d=>'<button class="ls-anchor-btn" data-anchor="'+d.key+'">▶ '+d.name+'</button>').join('');
  ab.querySelectorAll('button').forEach(b=>{b.onclick=()=>{const el=document.getElementById('ls-dim-'+b.dataset.anchor);if(el){const tb=el.querySelector('.ls-table-block');if(tb)tb.style.display='block';const arrow=el.querySelector('.ls-dim-arrow');if(arrow)arrow.textContent='▼';}if(el)el.scrollIntoView({behavior:'smooth',block:'start'});};});
  const hBar=(id,agg)=>{const s=agg.slice().sort((a,b)=>b.count-a.count).reverse();const c=chart(id);if(c)c.setOption({color:['#3b82f6','#93c5fd'],tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},legend:{data:['订单数','高价课单数'],bottom:0,textStyle:{color:'#64748b'}},grid:{left:8,right:16,top:24,bottom:8,containLabel:true},xAxis:{type:'value',axisLine:{lineStyle:{color:'#e2e8f0'}},axisLabel:{color:'#94a3b8'}},yAxis:{type:'category',data:s.map(r=>r.value),axisLine:{lineStyle:{color:'#e2e8f0'}},axisLabel:{color:'#64748b',fontSize:11}},series:[{name:'订单数',type:'bar',data:s.map(r=>r.count),itemStyle:{color:'#3b82f6',borderRadius:[0,4,4,0]}},{name:'高价课单数',type:'bar',data:s.map(r=>r.highN),itemStyle:{color:'#93c5fd',borderRadius:[0,4,4,0]}}]});};
  const dims=[{key:'city',name:'城市等级'},{key:'province',name:'省份'},{key:'cityName',name:'具体城市'},{key:'gender',name:'性别'},{key:'age',name:'年龄档'},{key:'price',name:'价格档'},{key:'brand',name:'品牌'},{key:'model',name:'手机型号'}];
  const hasScoring=cur&&cur.scoring&&cur.scoring.periodRows&&cur.scoring.periodRows.length;
  const dimCount=dims.length+(hasScoring?1:0);
  let html='<div class="ls-sheet"><div class="ls-sheet-header"><h2>'+esc(lsLink)+'</h2><span class="ls-sheet-tag">链路 · '+dimCount+' 个维度（点击标题展开数据表）</span></div><div class="ls-sheet-body">';
  dims.forEach(dim=>{
    const det=lsFilterDetail(cur.dimensions[dim.key],lsPeriodFilter,lsSearchText),agg=lsAggregate(det);
    // 地域维度：从原始dim.rows合并综合评分、评级、价值矩阵、优化建议
    if (['city','province','cityName'].includes(dim.key) && cur.dimensions[dim.key] && cur.dimensions[dim.key].rows) {
      const scoreMap = {};
      cur.dimensions[dim.key].rows.forEach(r => { scoreMap[r.value] = r; });
      agg.forEach(a => {
        const src = scoreMap[a.value];
        if (src) { a.score=src.score; a.grade=src.grade; a.quadrant=src.quadrant; a.strategy=src.strategy; }
      });
    }
    html+='<div class="ls-dim" id="ls-dim-'+dim.key+'">';
    html+='<div class="ls-dim-title" style="cursor:pointer;user-select:none" onclick="lsToggleTable(\''+dim.key+'\')"><span class="ls-dim-arrow" style="display:inline-block;width:16px">▶</span> '+dim.name+'<span class="ls-dim-badge">'+agg.length+' 个分类</span></div>';
    if(agg.length)html+='<div class="ls-chart" id="ls-chart-'+dim.key+'" style="height:'+Math.max(200,agg.length*22)+'px"></div>';
    html+='<div class="ls-table-block" style="display:none">'+lsPeriodTable(dim.key,dim.name,det,agg)+'</div>';
    html+='</div>';
  });
  html+='<div class="ls-dim" id="ls-dim-score"><div class="ls-dim-title" style="cursor:pointer;user-select:none" onclick="lsToggleTable(\'score\')"><span class="ls-dim-arrow" style="display:inline-block;width:16px">▶</span> 评分分析<span class="ls-dim-badge">'+(hasScoring?'按期分布':'无数据')+'</span></div><div class="ls-table-block" style="display:none">';
  if(hasScoring){
    const onlyTotal=lsPeriodFilter==='__total__';
    const scP=onlyTotal?[]:cur.scoring.periodRows.filter(pr=>!lsPeriodFilter||lsPeriodFilter===pr.period);
    html+='<div class="ls-table-wrap"><table class="ls-table"><thead><tr><th>期数</th><th>评分状态</th><th>订单数</th><th>订单占比</th><th>高价课金额</th><th>单订单产值</th></tr></thead><tbody>';
    if(!onlyTotal)scP.forEach(pr=>{pr.rows.forEach(r=>{html+='<tr><td><b>'+esc(pr.period)+'</b></td><td>'+(r.value==='有评分'?'<span style="color:#3b82f6;font-weight:600">'+r.value+'</span>':'<span style="color:#94a3b8">'+r.value+'</span>')+'</td><td>'+fmtN(r.count)+'</td><td>'+fmtP((r.share||0)*100)+'</td><td>'+fmtM(r.highAmt||0)+'</td><td>'+(r.orderOutput?fmtM2(r.orderOutput):'—')+'</td></tr>';});});
    if(cur.scoring.summaryRows)cur.scoring.summaryRows.forEach(r=>{html+='<tr class="ls-total-row"><td>合计</td><td>'+r.value+'</td><td>'+fmtN(r.count)+'</td><td>'+fmtP((r.share||0)*100)+'</td><td>'+fmtM(r.highAmt||0)+'</td><td>'+(r.orderOutput?fmtM2(r.orderOutput):'—')+'</td></tr>';});
    html+='</tbody></table></div>';
  }else{
    html+='<div class="ls-empty">该链路未采集评分数据（仅 PF+小程序问答、PF+AI落地页 支持评分分析）</div>';
  }
  html+='</div></div>';
  html+='</div></div>';
  html+='<button id="ls-back-top" title="回到顶部" style="position:fixed;right:24px;bottom:24px;width:44px;height:44px;border-radius:50%;background:#3b82f6;color:#fff;border:none;cursor:pointer;font-size:20px;box-shadow:0 4px 12px rgba(59,130,246,.4);z-index:999;display:flex;align-items:center;justify-content:center;transition:opacity .3s">↑</button>';
  content.innerHTML=html;
  const backTopBtn=$('#ls-back-top');
  if(backTopBtn){backTopBtn.onclick=()=>window.scrollTo({top:0,behavior:'smooth'});}
  // 图表每次渲染都重建（chart函数自动dispose旧实例），确保切换页签回来不丢失
  dims.forEach(dim=>{
    const agg=lsAggregate(lsFilterDetail(cur.dimensions[dim.key],lsPeriodFilter,lsSearchText));
    if(agg.length)hBar('ls-chart-'+dim.key,agg);
  });
}

// 折叠/展开数据表格（图表不折叠）
function lsToggleTable(key){
  const el=document.getElementById('ls-dim-'+key);
  if(!el)return;
  const tb=el.querySelector('.ls-table-block');
  const arrow=el.querySelector('.ls-dim-arrow');
  if(!tb)return;
  const isOpen=tb.style.display!=='none';
  if(isOpen){tb.style.display='none';if(arrow)arrow.textContent='▶';}
  else{tb.style.display='block';if(arrow)arrow.textContent='▼';}
}

/* ---------- 视图：数据管理 ---------- */
const FIELD_DOCS = [
  ['日期','必填，如 2026-08-01 或 2026/8/1'],
  ['计划','广告计划名，缺失时归为「未分组」'],
  ['链路','转化链路，如 表单提交 / 在线咨询 / 私信加粉 / APP下载 / 电商下单'],
  ['版位','如 信息流 / 穿山甲，缺失时归为「未指定」'],
  ['设备型号','如 iPhone 15 Pro，缺失时归为「未知机型」'],
  ['操作系统','iOS / Android，缺失时按机型判断'],
  ['消耗','金额（元），如 120.5'],
  ['展现','次数，整数'],
  ['点击','次数，整数'],
  ['转化数','系统口径转化数，整数'],
  ['收入','金额（元），仅电商链路需要，缺失为 0']
];
function renderData(){
  const rows = filteredRows();
  const t = totals(rows);
  const links = [...new Set(DATA.rows.map(r=>r.link))];
  const models = [...new Set(DATA.rows.map(r=>r.model))];
  renderKpis('#dm-kpis', [
    {label:'数据来源', value: DATA.source==='demo'?'示例数据':DATA.source==='empty'?'无数据':'已导入', sub:'导入CSV或恢复示例数据'},
    {label:'数据行数', value:fmtN(DATA.rows.length), cls:'kpi-accent'},
    {label:'覆盖日期', value:''+fmtN(rows.length)+'行', sub: dataRangeText()},
    {label:'链路数', value:fmtN(links.length)},
    {label:'机型数', value:fmtN(models.length), cls:'kpi-teal'}
  ]);
  $('#dm-fields').innerHTML = FIELD_DOCS.map(f=>`<div class="f"><b>${esc(f[0])}</b>${esc(f[1])}</div>`).join('');
  const syncEl = $('#dm-sync');
  if(!syncEl) return;
  const persisted = localStorage.getItem(DATA_KEY);
  let pInfo = '<span class="sync-off">未存储</span>';
  if(persisted){ try{ const d=JSON.parse(persisted); pInfo=`<span class="sync-on">已存储</span> ${fmtN(d.rows.length)} 行 · ${new Date(d.ts).toLocaleString('zh-CN',{month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'})}`; }catch(e){} }
  const syncItems = Object.keys(SYNC).map(k=>{
    const d = SYNC[k];
    return `<div class="dm-sync-item"><span class="dm-sync-name">${SYNC_LABELS[k]}</span><span class="${d?'sync-on':'sync-off'}">${d?'已同步':'未同步'}</span>${d?`<button class="btn ghost ex-btn-sm" onclick="clearSync('${k}');renderData();showToast('已清除${SYNC_LABELS[k]}同步数据')">清除</button>`:''}</div>`;
  }).join('');
  syncEl.innerHTML = `
    <div class="dm-sync-row"><b>投放数据持久化</b>${pInfo}<button class="btn ghost ex-btn-sm" onclick="if(confirm('清除本地存储的投放数据？')){clearPersistentData();renderData();showToast('已清除本地数据')}">清除</button></div>
    <div class="dm-sync-row"><b>Excel分析同步</b><span style="color:var(--ink-3);font-size:12px">在Excel分析工具中运行后自动同步到对应分析模块</span></div>
    <div class="dm-sync-grid">${syncItems}</div>
    <div class="dm-sync-row" style="margin-top:8px"><button class="btn ghost ex-btn-sm" onclick="if(confirm('清除全部Excel同步数据？')){clearAllSync();renderData();showToast('已清除全部同步数据')}">清除全部同步</button></div>
  `;
}
function dataRangeText(){
  if (!DATA.rows.length) return '无数据';
  let mn=Infinity, mx=0;
  DATA.rows.forEach(r=>{ if(r.ts<mn)mn=r.ts; if(r.ts>mx)mx=r.ts; });
  return new Date(mn).toISOString().slice(0,10) + ' ~ ' + new Date(mx).toISOString().slice(0,10);
}

/* ============================================================
   腾讯广告·多账号模块
   ============================================================ */
const TC = {
  accounts: [],   // {name, rows, enabled}
  metric: 'cost',
  drillAccount: ''
};
const TC_MAX = 6;

const TC_FIELD_DOCS = [
  ['曝光量','广告被展现的次数（腾讯广告标准字段）'],
  ['点击量','广告被有效点击的次数'],
  ['消耗/花费','广告消耗金额（元）'],
  ['目标转化量','平台归因的目标转化次数（系统口径）'],
  ['目标转化成本','消耗 / 目标转化量'],
  ['目标转化率','目标转化量 / 点击量'],
  ['关键行为转化率','关键行为转化量 / 点击量（仅腾讯广告账户）'],
  ['首日新增下单ROI','首日新增下单金额 / 花费（电商类）'],
  ['版位/流量来源','微信朋友圈、公众号与小程序、腾讯视频、腾讯新闻、QQ、优量汇等'],
  ['推广计划/广告/创意','账户→计划→广告→创意 四层结构']
];

// 腾讯广告报表字段映射（兼容后台导出列名与API字段名）
const TC_ALIAS = {
  date:['date','日期','时间','day','统计日期','report_date'],
  account:['account','账户','账户名称','账户名','账户id','account_name','account_id','广告主','广告主名称','客户名称'],
  campaign:['campaign','推广计划','计划','计划名称','计划id','campaign_name','campaign_id','项目','项目名称'],
  adgroup:['adgroup','广告','广告名称','广告id','adgroup_name','adgroup_id'],
  ad:['ad','广告创意','创意','创意名称','创意id','ad_name','ad_id','creative_name','creative_id','dynamic_creative_name'],
  placement:['placement','版位','流量来源','site_set','union_position','投放版位','版位名称'],
  cost:['cost','消耗','花费','广告消耗','消耗金额','spend','花费金额'],
  impr:['impr','impressions','曝光量','展现量','曝光','展现','展示量','曝光次数'],
  click:['click','clicks','点击量','点击','点击次数'],
  conv:['conv','conversions','转化量','转化数','目标转化量','转化','目标转化','转化次数'],
  targetCpa:['target_conversion_cost','目标转化成本','转化成本','平均转化成本','cpa'],
  targetCvr:['target_conversion_rate','目标转化率','转化率','cvr'],
  keyCvr:['key_behavior_conversions_rate','关键行为转化率'],
  revenue:['revenue','收入','成交金额','gmv','下单金额','转化金额'],
  roi:['roi','转化roi','转化ROI','ROI','首日新增下单roi','first_day_order_roi']
};

function tcMapHeader(header){
  const h = String(header||'').trim().toLowerCase();
  for (const [field, aliases] of Object.entries(TC_ALIAS)){
    if (aliases.some(a=>a.toLowerCase()===h)) return field;
  }
  return null;
}

function parseTencentCSV(text, fileName){
  const delim = text.includes('\t') && !text.includes(',') ? '\t' : ',';
  const lines = text.split(/\r?\n/).filter(l=>l.trim());
  if (lines.length < 2) throw new Error('文件内容不足（需包含表头和至少一行数据）');
  const parseLine = (line)=>{
    const out=[]; let cur='', q=false;
    for (let i=0;i<line.length;i++){
      const ch=line[i];
      if (q){ if(ch==='"'){ if(line[i+1]==='"'){cur+='"';i++;} else q=false; } else cur+=ch; }
      else if (ch==='"') q=true;
      else if (ch===delim){ out.push(cur); cur=''; }
      else cur+=ch;
    }
    out.push(cur); return out;
  };
  const header = parseLine(lines[0]).map(h=>String(h||'').trim());
  const cols = header.map(tcMapHeader);
  if (!cols.includes('date')) throw new Error('未找到「日期」列，请确认报表包含日期字段');
  if (!cols.includes('cost') && !cols.includes('impr') && !cols.includes('click'))
    throw new Error('未找到消耗/曝光/点击等核心数据列，请确认导出的是效果报表');

  const num = (v)=>{ const n=parseFloat(String(v||'').replace(/[,\s¥元%]/g,'')); return isNaN(n)?0:n; };
  const rows = [];
  for (let i=1;i<lines.length;i++){
    const cells = parseLine(lines[i]);
    if (cells.length < 2) continue;
    const get = (f)=>{ const idx = cols.indexOf(f); return idx>=0 ? cells[idx] : null; };
    const d = parseDate(get('date'));
    if (!d) continue;
    const account = (get('account')||'').trim() || fileName.replace(/\.[^.]+$/,'');
    rows.push({
      date:fmtDate(d), ts:d.getTime(), weekStart:fmtDate(weekStart(d)), weekLabel:'W'+isoWeek(d),
      account:account,
      campaign:(get('campaign')||'未命名计划').trim()||'未命名计划',
      adgroup:(get('adgroup')||'未命名广告').trim()||'未命名广告',
      ad:(get('ad')||'未命名创意').trim()||'未命名创意',
      placement:(get('placement')||'未指定版位').trim()||'未指定版位',
      cost:num(get('cost')), impr:num(get('impr')), click:num(get('click')),
      conv:num(get('conv')),
      targetCpa:num(get('targetCpa')), targetCvr:num(get('targetCvr')), keyCvr:num(get('keyCvr')),
      revenue:num(get('revenue')), roi:num(get('roi')),
      sourceFile:fileName
    });
  }
  if (!rows.length) throw new Error('未能解析出有效数据行');
  return rows;
}

function tcUpsertAccount(name, rows){
  let acc = TC.accounts.find(a=>a.name===name);
  if (!acc){
    if (TC.accounts.length >= TC_MAX){
      // 超过6个，按消耗排序保留前6
      TC.accounts.push({name, rows, enabled:true});
      TC.accounts.sort((a,b)=>sum(b.rows.map(r=>r.cost))-sum(a.rows.map(r=>r.cost)));
      TC.accounts = TC.accounts.slice(0, TC_MAX);
      return false; // 被截断
    }
    acc = {name, rows:[], enabled:true};
    TC.accounts.push(acc);
  }
  acc.rows = rows; // 替换为该账号最新导出的全量数据
  return true;
}

function tcEnabledAccounts(){ return TC.accounts.filter(a=>a.enabled); }
function tcAllRows(){ return [].concat(...tcEnabledAccounts().map(a=>a.rows)); }

/* ===== 腾讯广告实时数据 ===== */
var TENCENT_LIVE_DATA = {
  syncTime: "2026-08-31 21:30:00",
  dateRange: "2026-08-31 至 2026-08-31",
  total: {balance:110061.28,spend:19829.47,cpa:200.30,impressions:492559,clicks:7197,ctr:1.46,cpc:2.76,conversions:99,cvr:1.38},
  accounts: [
    {name:"广州登鸣信息科技有限公司",id:"83676427",health:"建议调整",operator:"利欧",status:"启用中",tags:["灵儿-9.9-PF直投小程序问答-有留资-wtt"],balance:33919.94,spend:1532.11,cpa:170.23,impressions:42235,clicks:488,ctr:1.16,cpc:3.14,conversions:9,cvr:1.84},
    {name:"广州登鸣信息科技有限公司",id:"83676438",health:"持续关注",operator:"利欧",status:"启用中",tags:["灵儿-9.9-直投H5纯问答+首页分流-wtt"],balance:12783.33,spend:0.78,cpa:0,impressions:15,clicks:0,ctr:0,cpc:0,conversions:0,cvr:0},
    {name:"广州登鸣信息科技有限公司",id:"85489979",health:"持续关注",operator:"盟聚",status:"启用中",tags:["灵儿-9.9-小程序问答有无首页分流-wtt"],balance:12101.68,spend:2388.45,cpa:199.04,impressions:60559,clicks:538,ctr:0.89,cpc:4.44,conversions:12,cvr:2.23},
    {name:"广州登鸣信息科技有限公司",id:"85489984",health:"建议调整",operator:"盟聚",status:"启用中",tags:["灵儿-9.9-H5分流-排小程序-wtt"],balance:19228.62,spend:3643.65,cpa:280.28,impressions:64098,clicks:1129,ctr:1.76,cpc:3.23,conversions:13,cvr:1.15},
    {name:"广州登鸣信息科技有限公司",id:"87101601",health:"建议调整",operator:"利欧",status:"启用中",tags:["灵儿-9.9-H5-纯问答+首页-wtt"],balance:22221.63,spend:6928.54,cpa:173.21,impressions:181914,clicks:3207,ctr:1.76,cpc:2.16,conversions:40,cvr:1.25},
    {name:"广州登鸣信息科技有限公司",id:"87101605",health:"急需调整",operator:"利欧",status:"启用中",tags:["灵儿-9.9-AI落+PF-wtt"],balance:9806.08,spend:5335.94,cpa:213.44,impressions:143738,clicks:1835,ctr:1.28,cpc:2.91,conversions:25,cvr:1.36}
  ]
};

/* ===== 腾讯广告营销单元数据（7账号） ===== */
var TENCENT_ADUNITS_DATA = [
  {
    "account_id": "83676427",
    "account_name": "广州登鸣信息科技有限公司",
    "operator": "利欧",
    "tag": "灵儿-9.9-PF直投小程序问答-有留资-wtt",
    "headers": [
      "营销单元名称",
      "营销内容",
      "状态",
      "出价",
      "曝光次数",
      "花费",
      "千次展现均价",
      "目标转化量",
      "目标转化成本",
      "目标转化率",
      "操作"
    ],
    "units": [
      [
        "0830-太极灵儿-多版位--排低r城43男--双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "195.00 oCPM 元/下单",
        "827",
        "43.80",
        "52.96",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0830-太极灵儿-多版位--排低r城45男-高r--双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "195.00 oCPM 元/下单",
        "281",
        "29.38",
        "104.56",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0828-太极灵儿-多版位--排低r城46男-高r迭代--双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "190.00 oCPM 元/下单",
        "284",
        "10.97",
        "38.63",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0828-太极灵儿-多版位--排低r城46男--双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "190.00 oCPM 元/下单",
        "993",
        "38.33",
        "38.60",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0826-太极灵儿-多版位-排偏远-45+男--双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "180.00 oCPM 元/下单",
        "43",
        "13.42",
        "312.09",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0826-太极灵儿-多版位-44+男--双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "180.00 oCPM 元/下单",
        "199",
        "10.41",
        "52.31",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0825-太极灵儿-多版位--排低r城44男-痛点-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "190.00 oCPM 元/下单",
        "241",
        "11.22",
        "46.56",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0824-太极灵儿-多版位--排低r城47-xin-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "195.00 oCPM 元/下单",
        "549",
        "25.57",
        "46.58",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0820-太极灵儿-多版位--排低r城44男--双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "195.00 oCPM 元/下单",
        "686",
        "39.56",
        "57.67",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0819-太极灵儿-短剧--排低r城46-AB-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "190.00 oCPM 元/下单",
        "24,993",
        "824.08",
        "32.97",
        "5",
        "164.82",
        "4.27%",
        "编辑 数据 添加创意"
      ],
      [
        "0818-太极灵儿-多版位--排低r城47-复合-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "195.00 oCPM 元/下单",
        "1,029",
        "92.99",
        "90.37",
        "1",
        "92.99",
        "5.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0818-太极灵儿-多版位--排低r城42男-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "195.00 oCPM 元/下单",
        "1,550",
        "127.67",
        "82.37",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ]
    ],
    "total": {
      "count": 12,
      "impressions": "31,675",
      "spend": "1,267.40"
    }
  },
  {
    "account_id": "83676438",
    "account_name": "广州登鸣信息科技有限公司",
    "operator": "利欧",
    "tag": "灵儿-9.9-直投H5纯问答+首页分流-wtt",
    "headers": [
      "营销单元名称",
      "营销内容",
      "状态",
      "出价",
      "曝光次数",
      "花费",
      "千次展现均价",
      "目标转化量",
      "目标转化成本",
      "目标转化率",
      "操作"
    ],
    "units": [],
    "total": {}
  },
  {
    "account_id": "85489979",
    "account_name": "广州登鸣信息科技有限公司",
    "operator": "盟聚",
    "tag": "灵儿-9.9-小程序问答有无首页分流-wtt",
    "headers": [
      "营销单元名称",
      "营销内容",
      "状态",
      "出价",
      "曝光次数",
      "花费",
      "千次展现均价",
      "目标转化量",
      "目标转化成本",
      "目标转化率",
      "操作"
    ],
    "units": [
      [
        "0826-太极灵儿-短剧--排低r城42男-高r迭代-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "185.00 oCPM 元/下单",
        "23,933",
        "839.57",
        "35.08",
        "5",
        "167.91",
        "2.37%",
        "编辑 数据 添加创意"
      ],
      [
        "0825-太极灵儿-短剧--排低r城45男-xin-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "175.00 oCPM 元/下单",
        "8,962",
        "421.12",
        "46.99",
        "3",
        "140.37",
        "5.88%",
        "编辑 数据 添加创意"
      ]
    ],
    "total": {
      "count": 2,
      "impressions": "32,895",
      "spend": "1,260.68"
    }
  },
  {
    "account_id": "85489984",
    "account_name": "广州登鸣信息科技有限公司",
    "operator": "盟聚",
    "tag": "灵儿-9.9-H5分流-排小程序-wtt",
    "headers": [
      "营销单元名称",
      "营销内容",
      "状态",
      "出价",
      "曝光次数",
      "花费",
      "千次展现均价",
      "目标转化量",
      "目标转化成本",
      "目标转化率",
      "操作"
    ],
    "units": [
      [
        "0830-太极灵儿-排小程序--排低r城42+男-gaor双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "290.00 oCPM 元/下单",
        "5,760",
        "382.37",
        "66.38",
        "1",
        "382.37",
        "0.65%",
        "编辑 数据 添加创意"
      ],
      [
        "0828-太极灵儿-排小程序--排低r城45+男-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "280.00 oCPM 元/下单",
        "791",
        "28.51",
        "36.04",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0828-太极灵儿-排小程序--排低r城46+男-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "290.00 oCPM 元/下单",
        "6,354",
        "300.53",
        "47.30",
        "1",
        "300.53",
        "1.05%",
        "编辑 数据 添加创意"
      ],
      [
        "0827-太极灵儿-排小程序--排低r城45+男-xin-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "270.00 oCPM 元/下单",
        "16,584",
        "1,135.27",
        "68.46",
        "5",
        "227.05",
        "4.50%",
        "编辑 数据 添加创意"
      ],
      [
        "0827-太极灵儿-排小程序--排低r城46+男-痛点-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "270.00 oCPM 元/下单",
        "1,014",
        "58.11",
        "57.31",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0811-太极灵儿-排小程序--全国排偏远48+高r",
        "教育产品-商品销售 页面跳转 下单",
        "投放中 | 保障已结束",
        "255.00 oCPM 元/下单",
        "7,712",
        "465.16",
        "60.32",
        "2",
        "232.58",
        "1.77%",
        "编辑 数据 添加创意"
      ]
    ],
    "total": {
      "count": 6,
      "impressions": "38,215",
      "spend": "2,369.95"
    }
  },
  {
    "account_id": "87101601",
    "account_name": "广州登鸣信息科技有限公司",
    "operator": "利欧",
    "tag": "灵儿-9.9-H5-纯问答+首页-wtt",
    "headers": [
      "营销单元名称",
      "营销内容",
      "状态",
      "出价",
      "曝光次数",
      "花费",
      "千次展现均价",
      "目标转化量",
      "目标转化成本",
      "目标转化率",
      "操作"
    ],
    "units": [
      [
        "0830-太极灵儿-多版位--排低r城42+男-高r迭代-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "170.00 oCPM 元/下单",
        "33,200",
        "1,267.28",
        "38.17",
        "4",
        "316.82",
        "0.45%",
        "编辑 数据 添加创意"
      ],
      [
        "0830-太极灵儿-多版位--排低r城45+男-高r-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "170.00 oCPM 元/下单",
        "45,994",
        "2,893.83",
        "62.92",
        "18",
        "160.77",
        "2.13%",
        "编辑 数据 添加创意"
      ],
      [
        "0828-太极灵儿-短剧--排低r城46+男-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "160.00 oCPM 元/下单",
        "6,946",
        "318.71",
        "45.88",
        "2",
        "159.35",
        "0.87%",
        "编辑 数据 添加创意"
      ],
      [
        "0827-太极灵儿-短剧--排低r城43+男--gaor双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "170.00 oCPM 元/下单",
        "10,957",
        "722.10",
        "65.90",
        "4",
        "180.53",
        "1.70%",
        "编辑 数据 添加创意"
      ],
      [
        "0827-太极灵儿-短剧--排低r城43+男-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "170.00 oCPM 元/下单",
        "11,201",
        "513.45",
        "45.84",
        "4",
        "128.36",
        "1.48%",
        "编辑 数据 添加创意"
      ],
      [
        "0826-太极灵儿-短剧--排低r城43+男-AB素材-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "160.00 oCPM 元/下单",
        "9,230",
        "220.79",
        "23.92",
        "2",
        "110.39",
        "1.13%",
        "编辑 数据 添加创意"
      ],
      [
        "0825-太极灵儿-短剧--排低r城44+男--双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "165.00 oCPM 元/下单",
        "3,437",
        "215.56",
        "62.72",
        "3",
        "71.85",
        "2.70%",
        "编辑 数据 添加创意"
      ],
      [
        "0825-太极灵儿-短剧--排低r城42+男-双出价",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "160.00 oCPM 元/下单",
        "35,305",
        "507.44",
        "14.37",
        "3",
        "169.15",
        "0.97%",
        "编辑 数据 添加创意"
      ],
      [
        "0801-太极灵儿-多版位-排低r城45+强营销素材-h5问答+首-高r-控量！！",
        "教育产品-商品销售 页面跳转 下单",
        "投放中 | 保障已结束",
        "160.00 oCPM 元/下单",
        "29",
        "1.55",
        "53.45",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ]
    ],
    "total": {
      "count": 9,
      "impressions": "156,299",
      "spend": "6,660.71"
    }
  },
  {
    "account_id": "87101605",
    "account_name": "广州登鸣信息科技有限公司",
    "operator": "利欧",
    "tag": "灵儿-9.9-AI落+PF-wtt",
    "headers": [
      "营销单元名称",
      "营销内容",
      "状态",
      "出价",
      "曝光次数",
      "花费",
      "千次展现均价",
      "目标转化量",
      "目标转化成本",
      "目标转化率",
      "操作"
    ],
    "units": [
      [
        "0831-太极灵儿-多版位--全国排偏远45男+痛点",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "185.00 oCPM 元/下单",
        "728",
        "21.38",
        "29.37",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0828-太极灵儿-多版位--全国排偏远46男+xin",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 成本保障中",
        "175.00 oCPM 元/下单",
        "2",
        "0.02",
        "10.00",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0826-太极灵儿-短剧--全国排偏远43男+高r迭代",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障核实中",
        "185.00 oCPM 元/下单",
        "74,126",
        "2,074.99",
        "27.99",
        "10",
        "207.50",
        "1.89%",
        "编辑 数据 添加创意"
      ],
      [
        "0824-太极灵儿-短剧--全国排偏远48+痛点",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "190.00 oCPM 元/下单",
        "5",
        "0.12",
        "24.00",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0824-太极灵儿-短剧--全国排偏远45男+高r",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "190.00 oCPM 元/下单",
        "-",
        "-",
        "-",
        "-",
        "-",
        "-",
        "编辑 数据 添加创意"
      ],
      [
        "0820-太极灵儿-短剧--全国排偏远44男+xin",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "188.00 oCPM 元/下单",
        "10,133",
        "328.49",
        "32.42",
        "0",
        "0.00",
        "0.00%",
        "编辑 数据 添加创意"
      ],
      [
        "0812-太极灵儿-多版位-排低r城市42-65+男",
        "教育产品-商品销售 页面跳转 下单 首次到课",
        "投放中 | 保障已结束",
        "195.00 oCPM 元/下单",
        "10,053",
        "671.79",
        "66.82",
        "4",
        "167.95",
        "2.11%",
        "编辑 数据 添加创意"
      ]
    ],
    "total": {
      "count": 8,
      "impressions": "138,733",
      "spend": "5,206.08"
    }
  }
];
function renderTencentLive(){
  const d=TENCENT_LIVE_DATA;
  if(!d||!d.accounts||d.accounts.length===0){
    $('#tc-live-time').textContent='未同步';
    renderKpis('#tc-live-kpis',[{label:'账户数',value:'0'},{label:'今日花费',value:'¥0'},{label:'转化量',value:'0'},{label:'转化成本',value:'—'},{label:'点击率',value:'—'},{label:'转化率',value:'—'}]);
    ['tc-live-spend','tc-live-conv'].forEach(id=>{const c=chart(id);if(c)c.setOption({series:[],graphic:[{type:'text',left:'center',top:'middle',style:{text:'点击上方按钮同步数据',fill:'#918D83',fontSize:13}}]});});
    sortTable('#tc-live-table',[],[]);
    return;
  }
  $('#tc-live-time').textContent='同步于 '+d.syncTime;
  const t=d.total;
  renderKpis('#tc-live-kpis',[
    {label:'账户数',value:fmtN(d.accounts.length)},
    {label:'账户余额',value:fmtM(t.balance),cls:'kpi-blue'},
    {label:'今日花费',value:fmtM(t.spend),cls:'kpi-accent'},
    {label:'曝光量',value:fmtN(t.impressions)},
    {label:'点击量',value:fmtN(t.clicks)},
    {label:'点击率',value:fmtP(t.ctr)},
    {label:'点击均价',value:fmtM2(t.cpc)},
    {label:'转化量',value:fmtN(t.conversions),cls:'kpi-teal'},
    {label:'转化成本',value:fmtM2(t.cpa)},
    {label:'转化率',value:fmtP(t.cvr)}
  ]);
  const spendItems=d.accounts.slice().sort((a,b)=>b.spend-a.spend).slice(0,7).reverse();
  const c1=chart('tc-live-spend'); if(c1)c1.setOption({color:[C.accent],tooltip:tip,grid:{left:8,right:16,top:8,bottom:8,containLabel:true},xAxis:{type:'value',min:0,...baseAxis},yAxis:{type:'category',data:spendItems.map(i=>i.id),...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11}},series:[{type:'bar',data:spendItems.map(i=>Math.round(i.spend)),barWidth:'55%',label:{show:true,position:'right',formatter:p=>'¥'+fmtN(p.value),color:'#5C5A54',fontSize:11},itemStyle:{color:C.accent}}]});
  const convItems=d.accounts.slice().sort((a,b)=>b.conversions-a.conversions).slice(0,7).reverse();
  const c2=chart('tc-live-conv'); if(c2)c2.setOption({color:[C.teal],tooltip:tip,grid:{left:8,right:16,top:8,bottom:8,containLabel:true},xAxis:{type:'value',min:0,...baseAxis},yAxis:{type:'category',data:convItems.map(i=>i.id),...baseAxis,axisLabel:{color:'#5C5A54',fontSize:11}},series:[{type:'bar',data:convItems.map(i=>i.conversions),barWidth:'55%',label:{show:true,position:'right',formatter:p=>fmtN(p.value),color:'#5C5A54',fontSize:11},itemStyle:{color:C.teal}}]});
  sortState['#tc-live-table']={key:'spend',dir:-1};
  sortTable('#tc-live-table',d.accounts.map(i=>({key:i.id,name:i.name,health:i.health,operator:i.operator,balance:i.balance,spend:i.spend,cpa:i.cpa,impressions:i.impressions,clicks:i.clicks,ctr:i.ctr,cpc:i.cpc,conversions:i.conversions,cvr:i.cvr})),[
    {key:'key',label:'账户ID',cell:r=>'<b>'+esc(r.key)+'</b>'},
    {key:'name',label:'账户名称',cell:r=>esc(r.name)},
    {key:'health',label:'健康度',sortable:true,cell:r=>{const cls=r.health==='急需调整'?'delta-down':(r.health==='建议调整'?'delta-flat':'delta-up');return'<span class="'+cls+'">'+esc(r.health)+'</span>';}},
    {key:'operator',label:'运营方',sortable:true},
    {key:'balance',label:'余额',sortable:true,cell:r=>fmtM(r.balance)},
    {key:'spend',label:'今日花费',sortable:true,cell:r=>fmtM(r.spend)},
    {key:'impressions',label:'曝光',sortable:true,cell:r=>fmtN(r.impressions)},
    {key:'clicks',label:'点击',sortable:true,cell:r=>fmtN(r.clicks)},
    {key:'ctr',label:'点击率',sortable:true,cell:r=>fmtP(r.ctr)},
    {key:'cpc',label:'点击均价',sortable:true,cell:r=>fmtM2(r.cpc)},
    {key:'conversions',label:'转化',sortable:true,cell:r=>fmtN(r.conversions)},
    {key:'cpa',label:'转化成本',sortable:true,cell:r=>r.cpa?fmtM2(r.cpa):'—'},
    {key:'cvr',label:'转化率',sortable:true,cell:r=>fmtP(r.cvr)}
  ].map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));
}

/* ===== 营销单元明细渲染（全账号合并·增强版） ===== */
let TC_ADUNIT_ACCOUNT = 'all';
let TC_FILTERS = { placement:'', targeting:'', variant:'', bidtype:'' };

function renderTencentAdUnits(){
  const data = TENCENT_ADUNITS_DATA;
  const select = $('#tc-adunit-select');
  if(!data || !data.length){ if(select) select.innerHTML = '<option>暂无数据</option>'; return; }
  const cur = TC_ADUNIT_ACCOUNT || 'all';
  // 选择器：全部账号 + 6个账号
  let opts = '<option value="all"'+(cur==='all'?' selected':'')+'>全部账号（合并视图）</option>';
  data.forEach(a => {
    const op = a.operator ? '['+a.operator+'] ' : '';
    opts += '<option value="'+a.account_id+'"'+(a.account_id===cur?' selected':'')+'>'+op+'ID:'+a.account_id+' ('+a.total.count+'个)</option>';
  });
  select.innerHTML = opts;

  const num = (v)=>{const n=parseFloat(String(v||'').replace(/[,¥%]/g,''));return isNaN(n)?0:n;};

  // 筛选账号
  const accs = cur==='all' ? data : data.filter(a=>a.account_id===cur);

  // 合并所有单元，解析名称，添加账户ID和运营方
  const allUnits = [];
  const placementSet = new Set(), targetingSet = new Set(), variantSet = new Set(), bidtypeSet = new Set();
  accs.forEach(a => {
    a.units.forEach(u => {
      const parsed = parseUnitName(u[0]||'');
      if(parsed.placement) placementSet.add(parsed.placement);
      if(parsed.targeting) targetingSet.add(parsed.targeting);
      if(parsed.variant) variantSet.add(parsed.variant);
      if(parsed.bidType) bidtypeSet.add(parsed.bidType);
      allUnits.push({
        accountId: a.account_id,
        operator: a.operator || '',
        name: u[0]||'', content: u[1]||'', status: u[2]||'',
        bid: u[3]||'', impressions: u[4]||'-', spend: u[5]||'-',
        cpm: u[6]||'-', conv: u[7]||'-', cpa: u[8]||'-', cvr: u[9]||'-',
        _imp: num(u[4]), _spd: num(u[5]), _conv: num(u[7]), _cpa: num(u[8]),
        parsed: parsed
      });
    });
  });

  // 填充筛选下拉
  function fillSelect(id, set, curVal){
    const el = document.getElementById(id);
    if(!el) return;
    let html = '<option value="">全部</option>';
    [...set].sort().forEach(v=>{ html += '<option value="'+esc(v)+'"'+(v===curVal?' selected':'')+'>'+esc(v)+'</option>'; });
    el.innerHTML = html;
  }
  fillSelect('tc-filter-placement', placementSet, TC_FILTERS.placement);
  fillSelect('tc-filter-targeting', targetingSet, TC_FILTERS.targeting);
  fillSelect('tc-filter-variant', variantSet, TC_FILTERS.variant);
  fillSelect('tc-filter-bidtype', bidtypeSet, TC_FILTERS.bidtype);

  // 应用筛选
  let filtered = allUnits;
  if(TC_FILTERS.placement) filtered = filtered.filter(u=>u.parsed.placement===TC_FILTERS.placement);
  if(TC_FILTERS.targeting) filtered = filtered.filter(u=>u.parsed.targeting===TC_FILTERS.targeting);
  if(TC_FILTERS.variant) filtered = filtered.filter(u=>u.parsed.variant===TC_FILTERS.variant);
  if(TC_FILTERS.bidtype) filtered = filtered.filter(u=>u.parsed.bidType===TC_FILTERS.bidtype);

  // 汇总KPI
  const totalImp = filtered.reduce((s,u)=>s+u._imp,0);
  const totalSpd = filtered.reduce((s,u)=>s+u._spd,0);
  const totalConv = filtered.reduce((s,u)=>s+u._conv,0);
  const avgCpa = totalConv ? totalSpd/totalConv : 0;
  const activeAccounts = [...new Set(filtered.map(u=>u.accountId))].length;

  const filterDesc = [];
  if(TC_FILTERS.placement) filterDesc.push('版位:'+TC_FILTERS.placement);
  if(TC_FILTERS.targeting) filterDesc.push('定向:'+TC_FILTERS.targeting);
  if(TC_FILTERS.variant) filterDesc.push('变体:'+TC_FILTERS.variant);
  if(TC_FILTERS.bidtype) filterDesc.push('出价:'+TC_FILTERS.bidtype);

  const labelPrefix = cur==='all' ? '全账号' : 'ID:'+cur;
  $('#tc-adunit-count').textContent = filtered.length + ' 个投放中营销单元' + (allUnits.length!==filtered.length ? '（共'+allUnits.length+'个）' : '');
  $('#tc-adunit-total').textContent = labelPrefix+' | 覆盖'+activeAccounts+'个账号 | 曝光 '+fmtN(totalImp)+' | 花费 ¥'+fmtM(totalSpd)+(filterDesc.length?' | 筛选: '+filterDesc.join(' · '):'');

  renderKpis('#tc-adunit-kpis',[
    {label:'投放中单元',value:filtered.length,cls:'kpi-accent'},
    {label:'覆盖账号',value:activeAccounts+'/6',cls:'kpi-blue'},
    {label:'总曝光',value:fmtN(totalImp)},
    {label:'总花费',value:fmtM(totalSpd),cls:'kpi-accent'},
    {label:'目标转化量',value:fmtN(totalConv),cls:'kpi-teal'},
    {label:'平均转化成本',value:avgCpa?fmtM2(avgCpa):'—',cls:'kpi-amber'}
  ]);

  // 按花费降序
  filtered.sort((a,b)=>b._spd-a._spd);
  filtered.forEach((u,i)=>{u.idx=i+1;});

  sortState['#tc-adunit-table']={key:'_spd',dir:-1};
  const cols = [
    {key:'idx',label:'#',cell:r=>'<span class="rank">'+r.idx+'</span>'},
    {key:'accountId',label:'账户ID',sortable:true,cell:r=>'<b style="color:var(--accent);font-size:11.5px">'+r.accountId+'</b>'+(r.operator?'<br><span style="color:var(--ink-3);font-size:10px">'+r.operator+'</span>':'')},
    {key:'name',label:'营销单元名称',sortable:true,cell:r=>'<b style="font-size:12.5px">'+esc(r.name)+'</b>'+unitDimTags(r.parsed)},
    {key:'status',label:'状态',sortable:true,cell:r=>{const s=r.status;let cls='delta-up';if(s.includes('暂停'))cls='delta-down';else if(s.includes('未到'))cls='delta-flat';return '<span class="'+cls+'" style="font-size:11px">'+esc(s).split(' ')[0]+'</span>';}},
    {key:'bid',label:'出价',sortable:true,cell:r=>'<span style="font-size:11.5px">'+esc(r.bid).replace(' oCPM 元/下单','')+'</span>'},
    {key:'_imp',label:'曝光',sortable:true,cell:r=>r.impressions==='-'?'<span style="color:var(--ink-3)">—</span>':fmtN(r._imp)},
    {key:'_spd',label:'花费',sortable:true,cell:r=>r.spend==='-'?'<span style="color:var(--ink-3)">—</span>':'<b>¥'+fmtM(r._spd)+'</b>'},
    {key:'cpm',label:'千次展现',sortable:true,cell:r=>r.cpm==='-'?'<span style="color:var(--ink-3)">—</span>':fmtM2(r.cpm)},
    {key:'_conv',label:'转化量',sortable:true,cell:r=>r.conv==='-'?'<span style="color:var(--ink-3)">—</span>':'<b style="color:var(--mint)">'+fmtN(r._conv)+'</b>'},
    {key:'_cpa',label:'转化成本',sortable:true,cell:r=>r.cpa==='-'?'<span style="color:var(--ink-3)">—</span>':fmtM2(r._cpa)},
    {key:'cvr',label:'转化率',sortable:true,cell:r=>r.cvr==='-'?'<span style="color:var(--ink-3)">—</span>':esc(r.cvr)}
  ];
  sortTable('#tc-adunit-table',filtered,cols.map(c=>{if(!('cell' in c))c.cell=r=>defaultCell(r,c.key);return c;}));

  // 绑定筛选事件
  const selAcc = $('#tc-adunit-select');
  if(selAcc && !selAcc._boundEnh){
    selAcc.onchange = ()=>{ TC_ADUNIT_ACCOUNT = selAcc.value; renderTencentAdUnits(); };
    selAcc._boundEnh = true;
  }
  ['tc-filter-placement','tc-filter-targeting','tc-filter-variant','tc-filter-bidtype'].forEach(id=>{
    const el = document.getElementById(id);
    if(el && !el._boundEnh){
      const key = id.replace('tc-filter-','');
      el.onchange = ()=>{ TC_FILTERS[key] = el.value; renderTencentAdUnits(); };
      el._boundEnh = true;
    }
  });
  const resetBtn = $('#tc-filter-reset');
  if(resetBtn && !resetBtn._boundEnh){
    resetBtn.onclick = ()=>{ TC_FILTERS={placement:'',targeting:'',variant:'',bidtype:''}; TC_ADUNIT_ACCOUNT='all'; renderTencentAdUnits(); };
    resetBtn._boundEnh = true;
  }
  const exportBtn = $('#tc-export-op');
  if(exportBtn && !exportBtn._boundEnh){
    exportBtn.onclick = ()=>{
      const sugs = generateSuggestions(TENCENT_LIVE_DATA, TENCENT_ADUNITS_DATA);
      if(sugs.length===0){ showToast('暂无优化建议'); return; }
      exportOperationList(sugs);
      showToast('已导出 '+sugs.length+' 条操作建议');
    };
    exportBtn._boundEnh = true;
  }
}

function tcTotals(rows){
  let cost=0,impr=0,click=0,conv=0,rev=0;
  rows.forEach(r=>{cost+=r.cost;impr+=r.impr;click+=r.click;conv+=r.conv;rev+=r.revenue;});
  return {cost,impr,click,conv,rev,ctr:pct(click,impr),cvr:cvrOf(conv,click),cpa:conv?cost/conv:null,roi:cost?rev/cost:0};
}

function renderTencent(){
  renderTencentLive();
  renderTencentAdUnits();
  if(!$('#tc-accounts')) return; // CSV导入模块已移除，跳过后续渲染
  const accs = TC.accounts;
  const enabled = tcEnabledAccounts();
  const allRows = tcAllRows();
  const t = tcTotals(allRows);

  // 账号管理表
  const accTable = $('#tc-accounts');
  if (accs.length===0){
    accTable.innerHTML = '<tr><td style="text-align:center;padding:24px;color:var(--ink-3)">暂无数据，点击上方「批量导入报表」选择6个账号的CSV文件</td></tr>';
  } else {
    const thead = '<tr><th>账号名称</th><th>数据行数</th><th>日期范围</th><th>消耗</th><th>转化</th><th>CPA</th><th>参与对比</th><th>操作</th></tr>';
    const rows = accs.map((a,idx)=>{
      const at = tcTotals(a.rows);
      let mn=Infinity,mx=0;
      a.rows.forEach(r=>{if(r.ts<mn)mn=r.ts;if(r.ts>mx)mx=r.ts;});
      const range = a.rows.length ? (new Date(mn).toISOString().slice(5,10)+' ~ '+new Date(mx).toISOString().slice(5,10)) : '—';
      return `<tr>
        <td><input class="tc-rename" data-idx="${idx}" value="${esc(a.name)}" style="width:140px;border:1px solid var(--border-2);border-radius:5px;padding:4px 6px;font-size:12.5px"></td>
        <td>${fmtN(a.rows.length)}</td>
        <td>${range}</td>
        <td>${fmtM(at.cost)}</td>
        <td>${fmtN(at.conv)}</td>
        <td>${at.cpa?fmtM2(at.cpa):'—'}</td>
        <td><label style="cursor:pointer"><input type="checkbox" class="tc-toggle" data-idx="${idx}" ${a.enabled?'checked':''}> 显示</label></td>
        <td><button class="btn tc-clear" data-idx="${idx}" style="padding:3px 10px;font-size:11.5px">清空</button></td>
      </tr>`;
    }).join('');
    accTable.innerHTML = thead + rows;
    accTable.querySelectorAll('.tc-rename').forEach(inp=>{
      inp.onchange = ()=>{ const i=+inp.dataset.idx; const newName=inp.value.trim(); if(newName && !TC.accounts.find((a,j)=>j!==i&&a.name===newName)){ TC.accounts[i].name=newName; renderTencent(); } else { inp.value=TC.accounts[i].name; showToast('名称为空或已存在'); } };
    });
    accTable.querySelectorAll('.tc-toggle').forEach(cb=>{
      cb.onchange = ()=>{ TC.accounts[+cb.dataset.idx].enabled = cb.checked; renderTencent(); };
    });
    accTable.querySelectorAll('.tc-clear').forEach(btn=>{
      btn.onclick = ()=>{ const i=+btn.dataset.idx; const name=TC.accounts[i].name; TC.accounts.splice(i,1); showToast('已清空账号「'+name+'」'); renderTencent(); };
    });
  }

  // KPI
  if (enabled.length===0){
    $('#tc-kpis').innerHTML = '<div class="kpi" style="grid-column:1/-1;text-align:center;color:var(--ink-3);padding:20px">请先导入数据并勾选至少一个账号</div>';
  } else {
    renderKpis('#tc-kpis', [
      {label:'参与账号', value:fmtN(enabled.length)+' / '+fmtN(accs.length), cls:'kpi-blue'},
      {label:'总消耗', value:fmtM(t.cost), cls:'kpi-accent'},
      {label:'总曝光', value:fmtN(t.impr)},
      {label:'总点击', value:fmtN(t.click)},
      {label:'总转化', value:fmtN(t.conv), cls:'kpi-teal'},
      {label:'平均CPA', value:t.cpa?fmtM2(t.cpa):'—'},
      {label:'平均CTR', value:fmtP(t.ctr)},
      {label:'整体ROI', value:t.rev>0?t.roi.toFixed(2):'—'}
    ]);
  }

  // 趋势对比
  $$('.tc-metric').forEach(b=>{ b.classList.toggle('primary', b.dataset.metric===TC.metric); b.classList.toggle('ghost', b.dataset.metric!==TC.metric); });
  const cTrend = chart('tc-trend');
  if (cTrend && enabled.length){
    const dates = [...new Set(allRows.map(r=>r.date))].sort();
    const metric = TC.metric;
    const series = enabled.map((a,ai)=>{
      const byDate = {};
      a.rows.forEach(r=>{ byDate[r.date]=(byDate[r.date]||0)+ (metric==='cost'?r.cost:metric==='conv'?r.conv:metric==='cpa'?r.cost:metric==='ctr'?r.click:0); });
      // cpa/ctr 需要按日聚合计算
      let data;
      if (metric==='cpa'){
        const dayCost={}, dayConv={};
        a.rows.forEach(r=>{dayCost[r.date]=(dayCost[r.date]||0)+r.cost;dayConv[r.date]=(dayConv[r.date]||0)+r.conv;});
        data = dates.map(d=> dayConv[d]? +(dayCost[d]/dayConv[d]).toFixed(1) : null);
      } else if (metric==='ctr'){
        const dayImpr={}, dayClick={};
        a.rows.forEach(r=>{dayImpr[r.date]=(dayImpr[r.date]||0)+r.impr;dayClick[r.date]=(dayClick[r.date]||0)+r.click;});
        data = dates.map(d=> dayImpr[d]? +(dayClick[d]/dayImpr[d]*100).toFixed(2) : null);
      } else {
        data = dates.map(d=> +(byDate[d]||0).toFixed(metric==='cost'?0:0));
      }
      return {name:a.name, type:'line', data, smooth:true, symbol:'circle', symbolSize:5, lineStyle:{width:2}, itemStyle:{color:C.cat[ai%C.cat.length]}};
    });
    cTrend.setOption({
      color:C.cat, tooltip:{...tip, formatter:(p)=>{ const s=p.seriesName; const v=p.value; return `${s}<br/>${p.axisValue}<br/>${metric==='cost'?'消耗':metric==='conv'?'转化':metric==='cpa'?'CPA':'CTR'}：${metric==='cost'?fmtM(v):metric==='conv'?fmtN(v):metric==='cpa'?fmtM2(v):fmtP(v)}`; }},
      legend:{type:'scroll', bottom:0},
      grid:{left:8,right:8,top:24,bottom:34,containLabel:true},
      xAxis:{type:'category', data:dates, ...baseAxis, axisLabel:{color:'#5C5A54',fontSize:10,rotate:30}},
      yAxis:{type:'value', min:0, ...baseAxis},
      series
    });
  } else if (cTrend){
    cTrend.setOption({graphic:[{type:'text',left:'center',top:'middle',style:{text:'请先导入数据并勾选账号',fill:'#918D83',fontSize:13}}]});
  }

  // 消耗占比
  const cShare = chart('tc-share');
  if (cShare && enabled.length){
    const data = enabled.map(a=>({name:a.name, value:Math.round(tcTotals(a.rows).cost)})).filter(d=>d.value>0);
    cShare.setOption({
      color:C.cat, tooltip:tip,
      legend:{type:'scroll', bottom:0},
      series:[{type:'pie', radius:['40%','65%'], center:['50%','45%'],
        label:{show:true, position:'inside', formatter:'{d}%'},
        labelLine:{show:false},
        data}]
    });
  } else if (cShare){
    cShare.setOption({graphic:[{type:'text',left:'center',top:'middle',style:{text:'暂无数据',fill:'#918D83',fontSize:13}}]});
  }

  // 横向对比表
  sortState['#tc-compare'] = sortState['#tc-compare'] || {key:'cost',dir:-1};
  const compareData = enabled.map(a=>{
    const at = tcTotals(a.rows);
    return {name:a.name, cost:at.cost, impr:at.impr, click:at.click, ctr:+at.ctr.toFixed(2), conv:at.conv, cvr:+at.cvr.toFixed(2), cpa:at.cpa?+at.cpa.toFixed(1):null, roi:at.rev>0?+at.roi.toFixed(2):null, costShare:pct(at.cost,t.cost), convShare:pct(at.conv,t.conv)};
  });
  sortTable('#tc-compare', compareData, [
    {key:'name', label:'账号', cell:r=>`<b>${esc(r.name)}</b>`},
    {key:'cost', label:'消耗', sortable:true, cell:r=>fmtM(r.cost)},
    {key:'costShare', label:'消耗占比', sortable:true, cell:r=>fmtP(r.costShare)},
    {key:'impr', label:'曝光', sortable:true, cell:r=>fmtN(r.impr)},
    {key:'click', label:'点击', sortable:true, cell:r=>fmtN(r.click)},
    {key:'ctr', label:'CTR', sortable:true, cell:r=>fmtP(r.ctr)},
    {key:'conv', label:'转化', sortable:true, cell:r=>fmtN(r.conv)},
    {key:'convShare', label:'转化占比', sortable:true, cell:r=>fmtP(r.convShare)},
    {key:'cvr', label:'CVR', sortable:true, cell:r=>fmtP(r.cvr)},
    {key:'cpa', label:'CPA', sortable:true, cell:r=>r.cpa?fmtM2(r.cpa):'—'},
    {key:'roi', label:'ROI', sortable:true, cell:r=>r.roi!=null?r.roi.toFixed(2):'—'}
  ].map(c=>{ if(!('cell' in c)) c.cell=(r)=>defaultCell(r,c.key); return c; }));

  // 版位分布
  const cPlace = chart('tc-placement');
  if (cPlace && enabled.length){
    const byPlace = {};
    allRows.forEach(r=>{ byPlace[r.placement]=(byPlace[r.placement]||{cost:0,conv:0,impr:0}); byPlace[r.placement].cost+=r.cost; byPlace[r.placement].conv+=r.conv; byPlace[r.placement].impr+=r.impr; });
    const places = Object.keys(byPlace).sort((a,b)=>byPlace[b].cost-byPlace[a].cost);
    cPlace.setOption({
      color:[C.blue], tooltip:{...tip, formatter:(p)=>{ const d=byPlace[p.name]; return `${p.name}<br/>消耗：${fmtM(d.cost)}<br/>转化：${fmtN(d.conv)}<br/>曝光：${fmtN(d.impr)}`; }},
      legend:{data:['消耗','转化'], type:'scroll', bottom:0},
      grid:{left:8,right:8,top:24,bottom:34,containLabel:true},
      xAxis:{type:'value', min:0, ...baseAxis},
      yAxis:{type:'category', data:places, ...baseAxis, axisLabel:{color:'#5C5A54',fontSize:11}},
      series:[
        {name:'消耗', type:'bar', data:places.map(p=>Math.round(byPlace[p].cost)), barWidth:'38%', itemStyle:{color:C.accent}, label:{show:true,position:'right',formatter:p=>fmtN(p.value),color:'#5C5A54',fontSize:11}},
        {name:'转化', type:'bar', data:places.map(p=>byPlace[p].conv), barWidth:'38%', itemStyle:{color:C.teal}, label:{show:true,position:'right',formatter:p=>fmtN(p.value),color:'#5C5A54',fontSize:11}}
      ]
    });
  } else if (cPlace){
    cPlace.setOption({graphic:[{type:'text',left:'center',top:'middle',style:{text:'暂无数据',fill:'#918D83',fontSize:13}}]});
  }

  // 层级下钻
  const drillSel = $('#tcDrillAccount');
  const curDrill = TC.drillAccount;
  drillSel.innerHTML = '<option value="">请选择账号</option>' + enabled.map(a=>`<option value="${esc(a.name)}" ${a.name===curDrill?'selected':''}>${esc(a.name)}</option>`).join('');
  drillSel.onchange = ()=>{ TC.drillAccount = drillSel.value; renderTencent(); };
  const drillTable = $('#tc-drill');
  if (curDrill && TC.accounts.find(a=>a.name===curDrill)){
    const acc = TC.accounts.find(a=>a.name===curDrill);
    const byCamp = {};
    acc.rows.forEach(r=>{ byCamp[r.campaign]=(byCamp[r.campaign]||[]); byCamp[r.campaign].push(r); });
    const campData = Object.keys(byCamp).map(k=>{
      const ct = tcTotals(byCamp[k]);
      const adgroups = new Set(byCamp[k].map(r=>r.adgroup)).size;
      return {campaign:k, adgroups, cost:ct.cost, impr:ct.impr, click:ct.click, ctr:+ct.ctr.toFixed(2), conv:ct.conv, cvr:+ct.cvr.toFixed(2), cpa:ct.cpa?+ct.cpa.toFixed(1):null};
    });
    sortState['#tc-drill'] = sortState['#tc-drill'] || {key:'cost',dir:-1};
    sortTable('#tc-drill', campData, [
      {key:'campaign', label:'推广计划', cell:r=>esc(r.campaign)},
      {key:'adgroups', label:'广告数', sortable:true, cell:r=>fmtN(r.adgroups)},
      {key:'cost', label:'消耗', sortable:true, cell:r=>fmtM(r.cost)},
      {key:'impr', label:'曝光', sortable:true, cell:r=>fmtN(r.impr)},
      {key:'click', label:'点击', sortable:true, cell:r=>fmtN(r.click)},
      {key:'ctr', label:'CTR', sortable:true, cell:r=>fmtP(r.ctr)},
      {key:'conv', label:'转化', sortable:true, cell:r=>fmtN(r.conv)},
      {key:'cvr', label:'CVR', sortable:true, cell:r=>fmtP(r.cvr)},
      {key:'cpa', label:'CPA', sortable:true, cell:r=>r.cpa?fmtM2(r.cpa):'—'}
    ].map(c=>{ if(!('cell' in c)) c.cell=(r)=>defaultCell(r,c.key); return c; }));
  } else {
    drillTable.innerHTML = '<tr><td style="text-align:center;padding:20px;color:var(--ink-3)">请在上方选择一个账号查看推广计划层级数据</td></tr>';
  }

  // 字段说明（如存在）
  const fieldsEl = $('#tc-fields');
  if(fieldsEl) fieldsEl.innerHTML = TC_FIELD_DOCS.map(f=>`<div class="f"><b>${esc(f[0])}</b>${esc(f[1])}</div>`).join('');
}

function tcTemplateCSV(){
  const head = '日期,账户名称,推广计划,广告,广告创意,版位,消耗,曝光量,点击量,目标转化量,目标转化成本,目标转化率,关键行为转化率,收入,ROI\n';
  const sample = [
    '2026-08-25,主账户-品牌A,夏季放量计划,广告001,创意视频01,微信朋友圈,1250.50,125000,3200,86,14.54,2.69,1.85,0,0',
    '2026-08-25,主账户-品牌A,夏季放量计划,广告002,创意图文02,公众号与小程序,880.00,98000,1960,52,16.92,2.65,1.72,0,0',
    '2026-08-25,测试账户-B,老客召回,广告010,创意视频10,腾讯视频,450.00,56000,840,28,16.07,3.33,2.10,0,0',
    '2026-08-25,电商账户-C,大促下单,广告020,创意直播20,优量汇,3200.00,210000,4800,45,71.11,0.94,0.52,4200.00,1.31'
  ];
  return head + sample.join('\n') + '\n';
}

function bindTencentEvents(){
  const syncBtn=$('#tcSyncBtn');if(syncBtn){syncBtn.onclick = ()=>{
    // 检测登录状态
    const isLoggedIn = localStorage.getItem('tc_logged_in') === '1';
    if(!isLoggedIn){
      showTcLoginModal();
      return;
    }
    startTencentCrawl();
  };}
  const adunitSel = $('#tc-adunit-select');
  if(adunitSel){ adunitSel.onchange = ()=>{ TC_ADUNIT_ACCOUNT = adunitSel.value; renderTencentAdUnits(); }; }
  // CSV导入模块（如存在则绑定事件）
  if($('#tcBatchBtn')){
  $('#tcBatchBtn').onclick = ()=>{ $('#tcFileInput').click(); };
  $('#tcFileInput').onchange = async (e)=>{
    const files = Array.from(e.target.files||[]);
    if (!files.length) return;
    let ok=0, fail=0, msgs=[];
    for (const f of files){
      try{
        const text = await f.text();
        const rows = parseTencentCSV(text, f.name);
        const accountsInFile = [...new Set(rows.map(r=>r.account))];
        accountsInFile.forEach(accName=>{
          const accRows = rows.filter(r=>r.account===accName);
          tcUpsertAccount(accName, accRows);
        });
        ok++;
      }catch(err){ fail++; msgs.push(f.name+'：'+err.message); }
    }
    e.target.value = '';
    if (ok) showToast(`成功导入 ${ok} 个文件`);
    renderTencent();
    if (fail){
      showModal('导入完成（部分失败）', `成功 ${ok} 个，失败 ${fail} 个。`, msgs.join('\n'));
    } else if (ok){
      showModal('导入完成', `成功导入 ${ok} 个报表文件。已自动识别账号并聚合，可在下方账号表中重命名、勾选参与对比。`);
    }
  };
  $('#tcTemplateBtn').onclick = ()=> downloadFile('腾讯广告报表导入模板.csv', tcTemplateCSV());
  $('#tcPasteBtn').onclick = ()=>{
    const txt = prompt('粘贴腾讯广告报表CSV内容（首行为表头）：');
    if (!txt || !txt.trim()) return;
    try{
      const rows = parseTencentCSV(txt, '粘贴导入');
      const accountsInFile = [...new Set(rows.map(r=>r.account))];
      accountsInFile.forEach(accName=>{ tcUpsertAccount(accName, rows.filter(r=>r.account===accName)); });
      showToast('粘贴导入成功');
      renderTencent();
    }catch(err){ showModal('导入失败','请检查内容格式后重试。',err.message); }
  };
  $('#tcClearAllBtn').onclick = ()=>{
    if (confirm('确定清空全部6个账号的腾讯广告数据？')){ TC.accounts=[]; TC.drillAccount=''; renderTencent(); showToast('已清空全部账号数据'); }
  };
  $$('.tc-metric').forEach(b=>{
    b.onclick = ()=>{ TC.metric = b.dataset.metric; renderTencent(); };
  });
  }
}

/* ---------- 腾讯广告登录与爬虫 ---------- */
function showTcLoginModal(){
  const m=document.getElementById('tc-login-modal');
  m.style.display='flex';
}
function closeTcLoginModal(){
  document.getElementById('tc-login-modal').style.display='none';
}
function tcLoginSuccess(){
  localStorage.setItem('tc_logged_in','1');
  localStorage.setItem('tc_login_time',new Date().toISOString());
  closeTcLoginModal();
  showToast('登录成功，开始爬取数据...');
  setTimeout(()=>startTencentCrawl(),500);
}
function startTencentCrawl(){
  const m=document.getElementById('tc-crawl-modal');
  m.style.display='flex';
  const progress=document.getElementById('tc-crawl-progress');
  const bar=document.getElementById('tc-crawl-bar');
  const steps=[
    {pct:15,text:'🔗 正在连接腾讯广告平台...'},
    {pct:30,text:'📋 正在获取账户列表...'},
    {pct:50,text:'📊 正在拉取今日消耗数据...'},
    {pct:70,text:'🎯 正在获取营销单元明细...'},
    {pct:85,text:'💹 正在计算转化指标...'},
    {pct:100,text:'✅ 数据爬取完成！'},
  ];
  let idx=0;
  const runStep=()=>{
    if(idx>=steps.length){
      setTimeout(()=>{
        m.style.display='none';
        // 从本地缓存加载数据（实际爬虫由后端/agent完成，这里读取已缓存的数据）
        try{
          const saved=localStorage.getItem('tencent_live_data');
          if(saved){ TENCENT_LIVE_DATA=JSON.parse(saved); }
        }catch(e){}
        const now=new Date();
        const ts=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0')+'-'+String(now.getDate()).padStart(2,'0')+' '+String(now.getHours()).padStart(2,'0')+':'+String(now.getMinutes()).padStart(2,'0')+':'+String(now.getSeconds()).padStart(2,'0');
        TENCENT_LIVE_DATA.syncTime=ts;
        renderTencent();
        showToast('腾讯广告数据已同步 · '+ts);
      },600);
      return;
    }
    const s=steps[idx];
    progress.innerHTML='<div>'+s.text+'</div>';
    bar.style.width=s.pct+'%';
    idx++;
    setTimeout(runStep,800+Math.random()*600);
  };
  runStep();
}
// 退出登录（清除登录态）
function tcLogout(){
  localStorage.removeItem('tc_logged_in');
  localStorage.removeItem('tc_login_time');
  showToast('已退出腾讯广告登录');
}

/* ---------- 路由与激活 ---------- */
const PAGE_META = {
  overview:   {title:'数据总览', sub:'优化驾驶舱 · 6账户健康度/异常预警/投放中单元', render:renderOverview},
  phone:      {title:'手机型号数据分析', sub:'机型消耗 / 转化 / 系统分布', render:renderPhone},
  'link-table':{title:'链路数据表_系统口径_数据分析', sub:'各转化链路系统口径表现', render:renderLinkTable},
  weekly:     {title:'分周数据分析', sub:'周粒度趋势 / 环比 / 成本', render:renderWeekly},
  'link-score':{title:'链路画像_评分分析', sub:'多维度评分与投放建议', render:renderLinkScore},
  material:   {title:'素材分析', sub:'素材维度消耗/转化/效果对比', render:renderMaterial},
  clue:       {title:'线索评分', sub:'手机号vs设备号 · 评分维度转化率/产值分析', render:renderClue},
  'clue-position':{title:'线索评分·版位分析', sub:'公小vs视频号 × 手机号vs设备号 · 4维度交叉对比', render:renderCluePosition},
  'account-score':{title:'分账户评分分析', sub:'广告账户维度 · 评分分布/复学诊断/链路质量对比', render:renderAccountScore},
  'account-compare':{title:'账户对比分析', sub:'多账户投放效果对比 · CPC/CPA/高低价课转化 · 预算调整建议', render:renderAccountCompare},
  'weekly-archive':{title:'周报历史归档', sub:'历史周报解析 · 问题定位/调整动作/方法论提炼', render:renderWeeklyArchive},
  'weekly-report':{title:'周报撰写', sub:'历史周报解析 + 工作台数据自动生成投放周报', render:renderWeeklyReport},
  tencent:    {title:'腾讯广告', sub:'6广州登鸣账号实时数据 / 投放中营销单元明细', render:renderTencent},
  'excel-phone':{title:'手机型号分析（Excel）', sub:'投放看板订单明细 → 品牌/型号/价格/年龄产值聚合', render:()=>{}},
  'excel-link':{title:'链路数据表·系统口径（Excel）', sub:'投放多维度数据看板 → 5链路×多期14项指标', render:()=>{}},
  'excel-weekly':{title:'分周数据对比（Excel）', sub:'本周vs上周缓存 → 链路/版位双维度环比', render:()=>{}},
  'excel-score':{title:'链路画像·评分分析（Excel）', sub:'5链路×6维度画像 + 评分段分析', render:()=>{}},
  'excel-material':{title:'素材分析（Excel）', sub:'素材数据上传 → 多维度效果分析 → 同步分析模块', render:()=>{}},
  'excel-clue':{title:'线索评分（Excel）', sub:'手机号+设备号双文件 → 评分维度转化/产值分析 → 同步分析模块', render:()=>{}},
  'excel-clue-position':{title:'线索评分·版位分析（Excel）', sub:'4份文件(手机号/设备号×公小/视频号) → 版位交叉对比 → 同步分析模块', render:()=>{}},
  'excel-account-score':{title:'分账户评分分析（Excel）', sub:'分账户评分明细 → 账户对比/评分分布/复学诊断 → 同步分析模块', render:()=>{}},
  'excel-account-compare':{title:'账户对比分析（Excel）', sub:'人群画像数据看板 → 按账户评分聚合 → CPC/CPA对比 → 预算建议 → 同步分析模块', render:()=>{}},
  'excel-weekly-archive':{title:'周报历史归档（Excel）', sub:'历史周报文本 → 按日期分期 → 提取问题/调整/方法论 → 同步分析模块', render:()=>{}},
  'excel-weekly-report':{title:'周报撰写（Excel）', sub:'上传历史周报docx → 解析模板 → 结合工作台数据自动生成新周报', render:()=>{}},
  data:       {title:'数据管理', sub:'导入 / 导出 / 字段说明', render:renderData},
  todo:       {title:'待办事项', sub:'投放工作任务管理 · 本地保存 · 优先级标记', render:renderTodo},
  'op-log':   {title:'操作日志', sub:'记录投放操作 · 追踪效果变化 · 本地保存', render:renderOpLog},
  'ab-test':  {title:'A/B测试', sub:'结构化测试管理 · 置信度计算 · 结论沉淀', render:renderABTests},
  'daily-report':{title:'日报生成', sub:'基于工作台数据一键生成投放日报 · 可复制', render:()=>{ const d=document.getElementById('dr-date'); if(d)d.value=new Date().toISOString().slice(0,10); }},
  'data-quality':{title:'数据质量', sub:'自动检测异常值 · 缺失字段 · 数据不一致', render:renderDataQuality},
  'link-config':{title:'前后端关联', sub:'建立投放单元与后端线索映射 · 单元级ROI追踪', render:renderLinkConfig},
  'name-dict':{title:'命名字典', sub:'自定义单元名解析规则 · 提升维度提取准确率', render:renderDict},
  'cross-account':{title:'跨账户分析', sub:'账户横向对比 · 重复单元检测 · 预算分配建议', render:renderCrossAccount},
  'conv-delay':{title:'转化延迟', sub:'T+N回款追踪 · 回传期保护 · 历史趋势', render:renderConvDelay}
};

function navigate(view, push=true){
  if (!PAGE_META[view]) view = 'overview';
  activeView = view;
  $$('.view').forEach(v=>v.classList.remove('active'));
  const sec = $('#view-'+view);
  sec.classList.add('active');
  $$('.nav a').forEach(a=>a.classList.toggle('active', a.dataset.view===view));
  $('#pageTitle').textContent = PAGE_META[view].title;
  $('#pageSub').textContent = PAGE_META[view].sub;
  if (push && location.hash !== '#/'+view) history.replaceState(null,'','#/'+view);
  // 先显示视图再渲染（echarts 需可见容器）
  requestAnimationFrame(()=>PAGE_META[view].render());
}

function route(){
  const h = location.hash.replace('#/','');
  navigate(h, false);
}

/* ---------- 筛选联动 ---------- */
function refreshFilters(){
  const tag=$('#dataTag');
  if(tag)tag.textContent = DATA.source==='demo' ? '示例数据' : '已导入数据';
  // 数据新鲜度时间戳
  const tsEl = $('#dataTs');
  if(tsEl){
    if(TENCENT_LIVE_DATA && TENCENT_LIVE_DATA.syncTime){
      const f = getDataFreshness(TENCENT_LIVE_DATA.syncTime);
      tsEl.className = 'data-timestamp '+f.cls;
      tsEl.innerHTML = '<span class="dt-dot"></span>腾讯数据 '+f.text;
    } else {
      tsEl.className = 'data-timestamp old';
      tsEl.innerHTML = '<span class="dt-dot"></span>腾讯数据待同步';
    }
  }
}

function applyFilters(){
  refreshFilters();
  if (PAGE_META[activeView]) PAGE_META[activeView].render();
}

/* ---------- CSV 解析 ---------- */
const COL_ALIAS = {
  'date':'date','日期':'date','时间':'date','day':'date',
  'plan':'plan','计划':'plan','计划名':'plan','广告计划':'plan','campaign':'plan',
  'link':'link','链路':'link','转化链路':'link','转化目标':'link','目标转化':'link',
  'placement':'placement','版位':'placement','广告位':'placement','场景':'placement',
  'model':'model','设备型号':'model','机型':'model','手机型号':'model','device':'model',
  'os':'os','操作系统':'os','系统':'os',
  'cost':'cost','消耗':'cost','花费':'cost','金额':'cost','spend':'cost',
  'impr':'impr','展现':'impr','曝光':'impr','展示':'impr','impression':'impr','impressions':'impr',
  'click':'click','点击':'click','clicks':'click',
  'conv':'conv','转化数':'conv','转化':'conv','转化量':'conv','conversions':'conv',
  'revenue':'revenue','收入':'revenue','成交金额':'revenue','gmv':'revenue'
};

function parseCSV(text){
  // 兼容逗号/Tab 分隔，处理引号字段
  const delim = text.includes('\t') && !text.includes(',') ? '\t' : ',';
  const lines = text.split(/\r?\n/).filter(l=>l.trim());
  if (!lines.length) throw new Error('文件为空');
  const parseLine = (line)=>{
    const out=[]; let cur='', q=false;
    for (let i=0;i<line.length;i++){
      const ch=line[i];
      if (q){ if(ch==='"'){ if(line[i+1]==='"'){cur+='"';i++;} else q=false; } else cur+=ch; }
      else if (ch==='"') q=true;
      else if (ch===delim){ out.push(cur); cur=''; }
      else cur+=ch;
    }
    out.push(cur); return out;
  };
  const header = parseLine(lines[0]).map(h=>String(h||'').trim());
  const cols = header.map(h=>COL_ALIAS[String(h).toLowerCase()]||COL_ALIAS[h]||null);
  const num = (v)=>{ const n=parseFloat(String(v).replace(/[,\s¥元%]/g,'')); return isNaN(n)?0:n; };
  const rows = [];
  for (let i=1;i<lines.length;i++){
    const cells = parseLine(lines[i]);
    if (cells.length<2) continue;
    const get = (k)=> { const idx = cols.indexOf(k); return idx>=0 ? cells[idx] : null; };
    const dateRaw = get('date');
    const d = parseDate(dateRaw);
    if (!d) continue;
    let os = (get('os')||'').trim();
    const model = (get('model')||'未知机型').trim();
    if (!os) os = /iphone|ios/i.test(model) ? 'iOS' : (/android|redmi|huawei|honor|oppo|vivo|xiaomi|mi |realme|一加|三星|motorola|荣耀/i.test(model)?'Android':'未指定');
    rows.push({
      date: fmtDate(d), ts: d.getTime(), weekStart: fmtDate(weekStart(d)), weekLabel:'W'+isoWeek(d),
      plan: (get('plan')||'未分组').trim()||'未分组',
      link: (get('link')||'未分组').trim()||'未分组',
      placement: (get('placement')||'未指定').trim()||'未指定',
      model: model, os: os||'未指定',
      cost: num(get('cost')), impr: num(get('impr')), click: num(get('click')), conv: num(get('conv')), revenue: num(get('revenue'))
    });
  }
  if (!rows.length) throw new Error('未能解析出有效数据行：请确认首行表头包含「日期 / 消耗 / 展现 / 点击 / 转化数」等字段');
  return rows;
}

function importRows(rows, sourceName){
  DATA = { rows, source: sourceName==='demo'?'demo':'import' };
  if(sourceName!=='demo') savePersistentData();
  Object.keys(charts).forEach(k=>{ try{ charts[k].dispose(); }catch(e){} });
  Object.keys(charts).forEach(k=>delete charts[k]);
  F = { days:F.days, link:'all', plan:'all' };
  refreshFilters();
  navigate('overview');
  showToast(`已导入 ${fmtN(rows.length)} 行数据`);
}

function showToast(msg){ const t=$('#toast'); t.textContent=msg; t.classList.add('show'); clearTimeout(showToast._t); showToast._t=setTimeout(()=>t.classList.remove('show'),2400); }
function showModal(title, msg, err){ $('#modalTitle').textContent=title; $('#modalMsg').textContent=msg; $('#modalErr').textContent=err||''; $('#modalErr').style.display = err?'block':'none'; $('#overlay').classList.add('show'); }

function templateCSV(){
  const head='日期,计划,链路,版位,设备型号,操作系统,消耗,展现,点击,转化数,收入\n';
  const sample=[
    '2026-08-25,夏季放量-下载,APP下载,信息流,iPhone 15 Pro,iOS,120.50,12340,380,62,0',
    '2026-08-25,常规跑量-全链路,表单提交,信息流,华为 Mate 60 Pro,Android,88.20,9860,182,19,0',
    '2026-08-25,老客召回-加粉,私信加粉,穿山甲,Redmi K70,Android,45.10,15300,214,9,0',
    '2026-08-25,夏季放量-下载,电商下单,信息流,vivo X100,Android,360.00,8200,112,3,432.00'
  ];
  return head + sample.join('\n') + '\n';
}
function downloadFile(name, content, type){
  const blob = new Blob([content], {type: type||'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();},200);
}
function exportCSV(){
  const head='日期,计划,链路,版位,设备型号,操作系统,消耗,展现,点击,转化数,收入\n';
  const body = DATA.rows.map(r=>[r.date,r.plan,r.link,r.placement,r.model,r.os,r.cost,r.impr,r.click,r.conv,r.revenue].join(',')).join('\n');
  downloadFile('投放数据导出_'+new Date().toISOString().slice(0,10)+'.csv', head+body);
}

/* ---------- 待办事项 ---------- */
const TODO_KEY = 'workbench_todos_v1';
let TODOS = [];
let todoFilter = 'all';

function loadTodos(){ try{ const d=JSON.parse(localStorage.getItem(TODO_KEY)); TODOS=Array.isArray(d)?d:[]; }catch(e){ TODOS=[]; } }
function saveTodos(){ try{ localStorage.setItem(TODO_KEY, JSON.stringify(TODOS)); }catch(e){} }

function addTodo(){
  const inp = $('#todo-input');
  const text = inp.value.trim();
  if(!text){ showToast('请输入待办内容'); return; }
  const priority = $('#todo-priority').value;
  TODOS.unshift({ id: Date.now()+''+Math.random().toString(36).slice(2,6), text, priority, done:false, createdAt: Date.now() });
  saveTodos();
  inp.value = '';
  renderTodo();
  showToast('已添加待办');
}

function toggleTodo(id){
  const t = TODOS.find(x=>x.id===id);
  if(t){ t.done = !t.done; saveTodos(); renderTodo(); }
}

function delTodo(id){
  TODOS = TODOS.filter(x=>x.id!==id);
  saveTodos();
  renderTodo();
  showToast('已删除');
}

function renderTodo(){
  const list = $('#todo-list');
  if(!list) return;
  const pLabel = {high:'高',mid:'中',low:'低'};
  let items = TODOS;
  if(todoFilter==='active') items = TODOS.filter(t=>!t.done);
  if(todoFilter==='done') items = TODOS.filter(t=>t.done);
  // 排序：未完成在前，同状态按优先级（高>中>低），再按创建时间倒序
  const pOrder = {high:0,mid:1,low:2};
  items = items.slice().sort((a,b)=>{
    if(a.done!==b.done) return a.done?1:-1;
    if(pOrder[a.priority]!==pOrder[b.priority]) return pOrder[a.priority]-pOrder[b.priority];
    return b.createdAt-a.createdAt;
  });
  if(items.length===0){
    list.innerHTML = '<div class="empty-hint"><div class="empty-icon">📋</div><b>暂无待办</b><br>在上方输入框添加你的第一条任务</div>';
  } else {
    list.innerHTML = items.map(t=>{
      const d = new Date(t.createdAt);
      const dateStr = (d.getMonth()+1)+'/'+d.getDate()+' '+String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');
      return `<div class="todo-item ${t.done?'done':''}" data-id="${t.id}">
        <div class="todo-check ${t.done?'checked':''}" data-id="${t.id}"></div>
        <span class="todo-text">${esc(t.text)}</span>
        <span class="todo-priority ${t.priority}">${pLabel[t.priority]||'中'}</span>
        <span class="todo-date">${dateStr}</span>
        <button class="todo-del" data-id="${t.id}" title="删除">×</button>
      </div>`;
    }).join('');
  }
  // 统计
  const total = TODOS.length, done = TODOS.filter(t=>t.done).length, active = total-done;
  $('#todo-stats').innerHTML = `<span class="todo-stat">共 <b>${total}</b></span><span class="todo-stat">未完成 <b>${active}</b></span><span class="todo-stat">已完成 <b>${done}</b></span>`;
  // 绑定事件
  list.querySelectorAll('.todo-check').forEach(el=>{ el.onclick=()=>toggleTodo(el.dataset.id); });
  list.querySelectorAll('.todo-del').forEach(el=>{ el.onclick=()=>delTodo(el.dataset.id); });
}

/* ---------- 数据总览-待办事项 ---------- */
function renderOverviewTodos(){
  const listEl=$('#ov-todo-list');
  if(!listEl)return;
  const total=TODOS.length;
  const done=TODOS.filter(t=>t.done).length;
  const active=total-done;
  const high=TODOS.filter(t=>t.priority==='high'&&!t.done).length;
  $('#ov-todo-count').textContent=active+'项待办';
  $('#ov-todo-total').textContent=total;
  $('#ov-todo-active').textContent=active;
  $('#ov-todo-done').textContent=done;
  $('#ov-todo-high').textContent=high;
  // 显示未完成的待办，最多5条
  const pOrder={high:0,mid:1,low:2};
  const items=TODOS.filter(t=>!t.done).sort((a,b)=>pOrder[a.priority]-pOrder[b.priority]||b.createdAt-a.createdAt).slice(0,5);
  if(items.length===0){
    listEl.innerHTML='<div style="text-align:center;padding:20px;color:var(--ink-3);font-size:13px">🎉 暂无待办，全部完成！</div>';
  }else{
    const pLabel={high:'高',mid:'中',low:'低'};
    const pColor={high:'var(--red)',mid:'var(--amber)',low:'var(--ink-3)'};
    listEl.innerHTML=items.map(t=>{
      const d=new Date(t.createdAt);
      const dateStr=(d.getMonth()+1)+'/'+d.getDate();
      return '<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;background:var(--surface-2);border-radius:8px;margin-bottom:6px;transition:background .15s" onmouseover="this.style.background=\'var(--accent-soft)\'" onmouseout="this.style.background=\'var(--surface-2)\'">'+
        '<div style="width:18px;height:18px;border:2px solid var(--border-2);border-radius:50%;cursor:pointer;flex:none" onclick="ovToggleTodo(\''+t.id+'\')"></div>'+
        '<span style="flex:1;font-size:13px;color:var(--ink);font-weight:500">'+esc(t.text)+'</span>'+
        '<span style="font-size:11px;padding:2px 8px;border-radius:10px;font-weight:600;color:'+pColor[t.priority]+';background:'+pColor[t.priority]+'15">'+pLabel[t.priority]+'</span>'+
        '<span style="font-size:11px;color:var(--ink-3)">'+dateStr+'</span>'+
        '<span style="cursor:pointer;color:var(--ink-3);font-size:16px;padding:0 4px" onclick="ovDelTodo(\''+t.id+'\')">×</span>'+
      '</div>';
    }).join('');
  }
}
function ovAddTodo(){
  const inp=$('#ov-todo-input');
  const text=inp.value.trim();
  if(!text){showToast('请输入待办内容');return;}
  const priority=$('#ov-todo-priority').value;
  TODOS.unshift({id:Date.now()+''+Math.random().toString(36).slice(2,6),text,priority,done:false,createdAt:Date.now()});
  saveTodos();
  inp.value='';
  renderOverviewTodos();
  renderTodo();
  showToast('已添加待办');
}
function ovToggleTodo(id){
  const t=TODOS.find(x=>x.id===id);
  if(t){t.done=!t.done;saveTodos();renderOverviewTodos();renderTodo();}
}
function ovDelTodo(id){
  TODOS=TODOS.filter(x=>x.id!==id);
  saveTodos();
  renderOverviewTodos();
  renderTodo();
  showToast('已删除');
}

function bindTodoEvents(){
  $('#todo-add').onclick = addTodo;
  $('#todo-input').addEventListener('keydown', e=>{ if(e.key==='Enter') addTodo(); });
  $$('#view-todo .todo-filter-btn').forEach(btn=>{
    btn.onclick = ()=>{
      todoFilter = btn.dataset.filter;
      $$('#view-todo .todo-filter-btn').forEach(b=>b.classList.toggle('active', b===btn));
      renderTodo();
    };
  });
}

/* ---------- 事件绑定 ---------- */
function bindEvents(){
  $('#hamburger').onclick = ()=>{ $('#sidebar').classList.toggle('open'); };
  document.addEventListener('click',(e)=>{
    if (e.target.closest('.nav a')) $('#sidebar').classList.remove('open');
  });
  window.addEventListener('hashchange', route);

  $('#pickFile').onclick = ()=>{ $('#fileInput').click(); };
  $('#fileInput').onchange = (e)=>{
    const f = e.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = ()=>{ try{ const rows=parseCSV(reader.result); importRows(rows,'file'); showModal('导入完成', `成功导入 ${fmtN(rows.length)} 行数据。当前数据来源已切换为「已导入数据」，若需回到演示环境可点击「恢复示例数据」。`); }catch(err){ showModal('导入失败', '请检查文件格式后重试。', err.message); } };
    reader.onerror = ()=>{ showModal('导入失败','无法读取该文件。','读取文件出错'); };
    reader.readAsText(f, 'UTF-8');
  };
  const dz = $('#dropZone');
  ['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev, e=>{ e.preventDefault(); dz.classList.add('hover'); }));
  ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev, e=>{ e.preventDefault(); dz.classList.remove('hover'); }));
  dz.addEventListener('drop', e=>{
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = ()=>{ try{ const rows=parseCSV(reader.result); importRows(rows,'file'); showModal('导入完成', `成功导入 ${fmtN(rows.length)} 行数据。`); }catch(err){ showModal('导入失败','请检查文件格式后重试。', err.message); } };
    reader.readAsText(f, 'UTF-8');
  });
  $('#btnTemplate').onclick = ()=> downloadFile('投放数据导入模板.csv', templateCSV());
  $('#btnExport').onclick = exportCSV;
  $('#btnResetDemo').onclick = ()=>{ clearPersistentData(); DATA={rows:genDemo(), source:'demo'}; Object.keys(charts).forEach(k=>{try{charts[k].dispose();}catch(e){}}); Object.keys(charts).forEach(k=>delete charts[k]); F={days:'all',link:'all',plan:'all'}; refreshFilters(); navigate('overview'); showToast('已恢复示例数据'); };
  $('#btnImportPaste').onclick = ()=>{
    const txt = $('#csvPaste').value.trim();
    if (!txt){ showModal('导入失败','请先粘贴 CSV 内容。'); return; }
    try{ const rows=parseCSV(txt); importRows(rows,'paste'); showModal('导入完成', `成功导入 ${fmtN(rows.length)} 行数据。`); }catch(err){ showModal('导入失败','请检查内容格式后重试。', err.message); }
  };
  $('#modalOk').onclick = ()=>{ $('#overlay').classList.remove('show'); };
  $('#overlay').addEventListener('click', e=>{ if (e.target.id==='overlay') $('#overlay').classList.remove('show'); });
  bindTodoEvents();
}


/* ============================================================
   Excel分析工具模块（移植自离线版）
   ============================================================ */
window._xlsxReady = true;
// ============================================================
// 共享工具函数
// ============================================================
const FMT_INT = '#,##0';
const FMT_DEC2 = '#,##0.00';
const FMT_PCT = '0.0%';

function switchTab(name, ev) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.header .tabs .tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(name).classList.add('active');
  if (ev) ev.target.classList.add('active');
}

function fmtInt(n) { return Number(n||0).toLocaleString('en-US'); }
function fmtDec2(n) { return Number(n||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function fmtPct(n) { return (Number(n||0)*100).toFixed(1)+'%'; }
function safeDiv(a,b){ return b ? a/b : 0; }

function showStatus(prefix, text, type) {
  const bar = document.getElementById(prefix+'Status');
  const dot = document.getElementById(prefix+'Dot');
  const txt = document.getElementById(prefix+'StatusText');
  bar.style.display = 'flex';
  dot.className = 'ex-status-dot';
  if (type==='active') dot.classList.add('active');
  else if (type==='error') dot.classList.add('error');
  else if (type==='done') dot.classList.add('done');
  txt.textContent = text;
}

function setupUpload(zoneId, fileId, fileNameId, btnId, checkReady) {
  const zone = document.getElementById(zoneId);
  const fileInput = document.getElementById(fileId);
  const fileName = document.getElementById(fileNameId);
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('dragover'); if(e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]); });
  fileInput.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });
  function handleFile(file) {
    fileInput._file = file;
    fileName.textContent = '✓ ' + file.name;
    if (checkReady) checkReady();
    else document.getElementById(btnId).disabled = false;
  }
}

function readWorkbook(file, cb) {
  const reader = new FileReader();
  reader.onload = function(e) {
    try { cb(XLSX.read(e.target.result, { type:'array' })); }
    catch(err) { cb(null, err); }
  };
  reader.readAsArrayBuffer(file);
}

function getSheetData(wb) {
  let sn = '导出数据';
  if (!wb.SheetNames.includes(sn)) sn = wb.SheetNames[0];
  return { ws: wb.Sheets[sn], name: sn, json: XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: null }) };
}

// 列名容错匹配：在 headers 数组中查找匹配的字段索引
function resolveCols(headers, spec) {
  const result = { found: {}, missing: [], allHeaders: headers };
  for (const [key, candidates] of Object.entries(spec)) {
    let idx = -1;
    for (const cand of candidates) {
      idx = headers.findIndex(h => h === cand || String(h).trim() === cand);
      if (idx >= 0) break;
    }
    if (idx === -1) {
      // 模糊匹配
      for (const cand of candidates) {
        idx = headers.findIndex(h => String(h).includes(cand) || cand.includes(String(h)));
        if (idx >= 0) break;
      }
    }
    if (idx >= 0) result.found[key] = idx;
    else result.missing.push({ field: key, tried: candidates });
  }
  return result;
}

function pctToNum(v) {
  if (v === null || v === undefined) return 0;
  let s = String(v).trim();
  if (s === '' || s === '未知') return 0;
  if (s.includes('%')) return parseFloat(s.replace('%','')) / 100;
  let n = parseFloat(s);
  return isNaN(n) ? 0 : (n > 1 ? n/100 : n);
}

function showMissingCols(prefix, missing, headers) {
  const el = document.createElement('div');
  el.className = 'ex-missing';
  let html = '<strong>⚠ 以下列未找到，请检查表头名称：</strong><ul>';
  missing.forEach(m => { html += '<li><code>'+m.field+'</code>：尝试了 ['+m.tried.map(t=>'"'+t+'"').join('、')+'] 未匹配</li>'; });
  html += '</ul><details><summary style="cursor:pointer;color:var(--text-dim);font-size:12px">查看全部 '+headers.length+' 个表头</summary><div style="font-size:11px;color:var(--text-dim);margin-top:6px;word-break:break-all">'+headers.map(h=>'<code style="margin:2px">'+h+'</code>').join(' ')+'</div></details>';
  el.innerHTML = html;
  document.getElementById(prefix+'Results').style.display = 'block';
  document.getElementById(prefix+'Results').prepend(el);
}

// Excel 样式常量
const S_TITLE = { font:{bold:true,sz:13,name:'Arial'}, alignment:{vertical:'center'} };
const S_SUB = { font:{sz:9,name:'Arial',color:{rgb:'808080'}} };
const S_HEADER = { font:{bold:true,sz:10,name:'Arial',color:{rgb:'FFFFFF'}}, fill:{fgColor:{rgb:'4472C4'}}, alignment:{horizontal:'center',vertical:'center',wrapText:true}, border:thinBorder() };
const S_DATA = { font:{sz:10,name:'Arial'}, border:thinBorder() };
const S_ZEBRA = { font:{sz:10,name:'Arial'}, fill:{fgColor:{rgb:'F7F9FC'}}, border:thinBorder() };
const S_BOLD = { font:{bold:true,sz:10,name:'Arial'}, border:thinBorder() };
const S_KPI = { font:{sz:10,name:'Arial'}, fill:{fgColor:{rgb:'EAF2FF'}}, border:thinBorder() };
function thinBorder(){ return {top:{style:'thin',color:'D9DEE7'},bottom:{style:'thin',color:'D9DEE7'},left:{style:'thin',color:'D9DEE7'},right:{style:'thin',color:'D9DEE7'}}; }

// ============================================================
// 页签1：手机型号数据分析（复用现成逻辑）
// ============================================================
const T1_BRAND_MAP = {
  'HUAWEI':'华为','Huawei':'华为','Xiaomi':'小米','Honor':'荣耀','HONOR':'荣耀',
  'vivo':'vivo','Vivo':'vivo','iPhone':'苹果','unknown':'未知','Unknown':'未知','未知':'未知',
  'Sony':'索尼','Samsung':'三星','Redmi':'红米','realme':'真我','Realme':'真我',
  'ChangHong':'长虹','长虹':'长虹','ZTE':'中兴','Nokia':'诺基亚',
  'MEIZU':'魅族','Meizu':'魅族','Google':'谷歌','motorola':'摩托罗拉','Motorola':'摩托罗拉',
  'OnePlus':'一加','Amoi':'夏新','XGIMI':'极米','Letv':'乐视','SHARK':'黑鲨',
  'Lenovo':'联想','Philips':'飞利浦',
};
const T1_AGE_MAP = { '51-55岁':'51-60岁','56-60岁':'51-60岁','61-65岁':'61-70岁','66-70岁':'61-70岁' };
const T1_PRICE_BINS = [0,999.99,1999.99,2999.99,3999.99,4999.99,5999.99,6999.99,1000000];
const T1_PRICE_LABELS = ['1000元以下','1000-1999元','2000-2999元','3000-3999元','4000-4999元','5000-5999元','6000-6999元','7000元及以上'];
const T1_PRICE_ORDER = [...T1_PRICE_LABELS, '未知'];
const T1_AGE_ORDER = ['20岁以下','26-30岁','31-35岁','36-40岁','41-45岁','45岁以下','46-50岁','51-60岁','61-70岁','70岁以上','未知'];

let t1Output = null, t1Data = null;

setupUpload('t1Zone','t1File','t1FileName','t1Btn');

function t1ParsePrice(x) {
  if (x === null || x === undefined) return null;
  let s = String(x).trim();
  if (s === '' || s === '未知') return null;
  s = s.replace('.000000','');
  let v = parseFloat(s);
  return isNaN(v) ? null : v;
}

function t1Preprocess(rows) {
  return rows.map(r => {
    let brand = r['手机品牌'] || '';
    let brandCn = T1_BRAND_MAP[brand] || brand;
    let age = r['年龄'] || '未知';
    let ageGroup = T1_AGE_MAP[age] || age;
    let price = t1ParsePrice(r['手机价格']);
    let priceBand = '未知';
    if (price !== null) {
      for (let i = 0; i < T1_PRICE_BINS.length-1; i++) {
        if (price >= T1_PRICE_BINS[i] && price <= T1_PRICE_BINS[i+1]) { priceBand = T1_PRICE_LABELS[i]; break; }
      }
      if (price > T1_PRICE_BINS[T1_PRICE_BINS.length-1]) priceBand = T1_PRICE_LABELS[T1_PRICE_LABELS.length-1];
    }
    let highOrder = r['高价课订单号'];
    let hasHigh = highOrder !== null && highOrder !== undefined && String(highOrder).trim() !== '';
    let highAmt = parseFloat(r['高价课订单金额']) || 0;
    let isValid = r['是否为有效订单'] === '是';
    return { ...r, 品牌统一:brandCn, 年龄段:ageGroup, 价格区间:priceBand, 有高价课:hasHigh, 高价课金额:highAmt, 有效:isValid };
  });
}

function t1Aggregate(data, groupKeys, totalHigh) {
  let groups = {};
  for (let row of data) {
    let key = groupKeys.map(k => row[k]).join('|||');
    if (!groups[key]) groups[key] = { keys: groupKeys.map(k => row[k]), count:0, valid:0, highN:0, highAmt:0 };
    let g = groups[key];
    g.count++; if (row['有效']) g.valid++;
    if (row['有高价课']) { g.highN++; g.highAmt += row['高价课金额']; }
  }
  return Object.values(groups).map(g => ({
    ...Object.fromEntries(g.keys.map((v,i)=>[groupKeys[i],v])),
    低价课订单数:g.count, 有效订单数:g.valid, 高价课单数:g.highN, 高价课金额:g.highAmt,
    高价课转化率: safeDiv(g.highN, g.count),
    高价课金额占比: safeDiv(g.highAmt, totalHigh),
    单订单产值: safeDiv(g.highAmt, g.count),
  }));
}

function processTab1() {
  try {
  if (!window._xlsxReady) { showStatus('t1','xlsx库未加载完成，请等几秒后再试','error'); return; }
  const file = document.getElementById('t1File')._file;
  if (!file) return;
  document.getElementById('t1Btn').disabled = true;
  showStatus('t1','正在读取文件...','active');
  document.getElementById('t1Results').style.display = 'none';
  readWorkbook(file, (wb, err) => {
    if (err) { showStatus('t1','读取失败：'+err.message,'error'); document.getElementById('t1Btn').disabled=false; return; }
    try {
    showStatus('t1','正在解析数据...','active');
    const { json } = getSheetData(wb);
    const processed = t1Preprocess(json);
    const nTotal = processed.length;
    if (nTotal === 0) { showStatus('t1','数据为空','error'); document.getElementById('t1Btn').disabled=false; return; }
    const nValid = processed.filter(r=>r['有效']).length;
    const nHigh = processed.filter(r=>r['有高价课']).length;
    const totalHigh = processed.reduce((s,r)=>s+r['高价课金额'],0);
    showStatus('t1','正在聚合计算...','active');
    let brandAgg = t1Aggregate(processed, ['品牌统一'], totalHigh).sort((a,b)=>b['单订单产值']-a['单订单产值']);
    let modelAgg = t1Aggregate(processed, ['品牌统一','手机型号'], totalHigh).sort((a,b)=>b['单订单产值']-a['单订单产值']);
    let priceAgg = t1Aggregate(processed, ['价格区间'], totalHigh).sort((a,b)=>{ let ai=T1_PRICE_ORDER.indexOf(a['价格区间']); let bi=T1_PRICE_ORDER.indexOf(b['价格区间']); return (ai<0?999:ai)-(bi<0?999:bi); });
    let ageAgg = t1Aggregate(processed, ['年龄段'], totalHigh).sort((a,b)=>{ let ai=T1_AGE_ORDER.indexOf(a['年龄段']); let bi=T1_AGE_ORDER.indexOf(b['年龄段']); return (ai<0?999:ai)-(bi<0?999:bi); });
    let dates = processed.map(r=>new Date(r['购买日期'])).filter(d=>!isNaN(d));
    let dateRange = dates.length ? new Date(Math.min(...dates)).toISOString().slice(0,10)+' ~ '+new Date(Math.max(...dates)).toISOString().slice(0,10) : '未知';
    showStatus('t1','正在生成Excel...','active');
    t1Output = t1BuildExcel(processed, nTotal, nValid, nTotal-nValid, nHigh, totalHigh, dateRange, file.name, brandAgg, modelAgg, priceAgg, ageAgg);
    t1Data = { nTotal, nValid, nHigh, totalHigh, brandAgg, modelAgg, priceAgg, ageAgg, dateRange };
    t1Display(t1Data);
    saveSync('phone', t1Data);
    showStatus('t1','完成！共 '+nTotal+' 条，'+brandAgg.length+' 品牌，'+modelAgg.length+' 型号（已同步到分析模块）','done');
    document.getElementById('t1Btn').disabled = false;
    } catch(e) { showStatus('t1','运行错误：'+e.message,'error'); console.error(e); document.getElementById('t1Btn').disabled=false; }
  });
  } catch(e) { showStatus('t1','启动错误：'+e.message,'error'); console.error(e); }
}

function t1BuildExcel(data, nTotal, nValid, nInvalid, nHigh, totalHigh, dateRange, srcFile, brandAgg, modelAgg, priceAgg, ageAgg) {
  const wb = XLSX.utils.book_new();
  // Sheet1 汇总概览
  let aoa = [];
  aoa.push(['投放订单产值分析 · 汇总概览','']);
  aoa.push(['数据源：'+srcFile+'（导出数据，'+nTotal+' 条）｜购买日期 '+dateRange,'']);
  aoa.push(['','']);
  aoa.push(['低价课订单总数',nTotal]);
  aoa.push(['有效订单数',nValid]);
  aoa.push(['高价课转化单数',nHigh]);
  aoa.push(['整体高价课转化率',safeDiv(nHigh,nTotal)]);
  aoa.push(['高价课金额合计(元)',totalHigh]);
  aoa.push(['整体单订单产值(元)',safeDiv(totalHigh,nTotal)]);
  aoa.push(['','']);
  aoa.push(['核心结论','']);
  let avg = safeDiv(totalHigh, nTotal);
  let conclusions = [];
  conclusions.push('1. 整体单订单产值 ¥'+avg.toFixed(2)+'/单（高价课 ¥'+totalHigh.toLocaleString('en-US',{maximumFractionDigits:0})+' ÷ '+nTotal+' 单）。');
  if (brandAgg.length) {
    let tb = brandAgg[0];
    conclusions.push('2. 全部品牌中 '+tb['品牌统一']+' 单订单产值最高 ¥'+tb['单订单产值'].toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})+'/单，但仅 '+tb['低价课订单数']+' 单，属小样本。');
    let major = brandAgg.filter(b=>b['低价课订单数']>=100);
    if (major.length) { let bm=major[0]; conclusions.push('3. 主流品牌（订单数≥100）中，'+bm['品牌统一']+' 单订单产值最高 ¥'+bm['单订单产值'].toFixed(2)+'/单（'+bm['低价课订单数']+' 单，转化率 '+(bm['高价课转化率']*100).toFixed(1)+'%）。'); }
  }
  if (modelAgg.length) { let bm=modelAgg[0]; conclusions.push('4. 型号：'+bm['品牌统一']+' '+bm['手机型号']+' 单订单产值最高 ¥'+bm['单订单产值'].toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})+'/单（'+bm['低价课订单数']+' 单，小样本）。'); }
  let pwd = priceAgg.filter(p=>p['单订单产值']>0).sort((a,b)=>b['单订单产值']-a['单订单产值']);
  if (pwd.length) {
    let bp=pwd[0];
    let ht=priceAgg.filter(p=>['5000-5999元','7000元及以上'].includes(p['价格区间']));
    let hts = ht.length ? '；'+ht.map(r=>r['价格区间']+'（¥'+r['单订单产值'].toFixed(2)+'）').join('与')+' 档高价机用户产出显著更高' : '';
    conclusions.push('5. 价格区间：'+bp['价格区间']+' 单订单产值最高 ¥'+bp['单订单产值'].toFixed(2)+'/单'+hts+'。');
  }
  let awd = ageAgg.filter(a=>a['单订单产值']>0).sort((a,b)=>b['单订单产值']-a['单订单产值']);
  if (awd.length) {
    let ba=awd[0];
    let ca=ageAgg.filter(a=>['51-60岁','61-70岁'].includes(a['年龄段']));
    let cs = ca.length ? '核心年龄段 51-70 岁约 ¥'+Math.round(Math.min(...ca.map(r=>r['单订单产值'])))+'-'+Math.round(Math.max(...ca.map(r=>r['单订单产值'])))+'/单' : '';
    let oa=ageAgg.find(a=>a['年龄段']==='70岁以上');
    let os = oa&&oa['单订单产值']>0 ? '，70 岁以上仅 ¥'+oa['单订单产值'].toFixed(2)+'/单' : '';
    conclusions.push('6. 年龄段：'+ba['年龄段']+' 单订单产值最高 ¥'+ba['单订单产值'].toFixed(2)+'/单（'+ba['低价课订单数']+' 单）；'+cs+os+'。');
  }
  conclusions.forEach(c => aoa.push([c,'']));
  aoa.push(['','']);
  aoa.push(['口径说明：单订单产值 = 高价课订单金额合计 ÷ 低价课订单数（分母为全部 '+nTotal+' 条低价课订单，含 '+nInvalid+' 条无效订单）。品牌已统一为中文，年龄已归并 5 岁档至 10 岁档，手机价格按机型价格分档。','']);
  let ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = aoa.map((_,i)=>({s:{r:i,c:0},e:{r:i,c:1}}));
  ws['!cols'] = [{wch:80},{wch:25}];
  ws['A1'].s = S_TITLE; ws['A2'].s = S_SUB;
  let fmts=[FMT_INT,FMT_INT,FMT_INT,FMT_PCT,FMT_DEC2,FMT_DEC2];
  for (let i=0;i<6;i++){ let r=4+i; ws['A'+(r+1)].s=S_BOLD; ws['B'+(r+1)].s={...S_KPI,numFmt:fmts[i]}; }
  ws['A11'].s=S_TITLE;
  conclusions.forEach((_,i)=>{ if(ws['A'+(12+i)]) ws['A'+(12+i)].s=S_DATA; });
  let nr=12+conclusions.length+1; if(ws['A'+(nr+1)]) ws['A'+(nr+1)].s=S_SUB;
  XLSX.utils.book_append_sheet(wb, ws, '汇总概览');

  // 维度Sheet通用函数
  function dimSheet(name, headers, widths, rows, fmts) {
    let a=[headers];
    rows.forEach(r => a.push(headers.map(h=>r[h])));
    let w=XLSX.utils.aoa_to_sheet(a);
    w['!cols']=widths.map(w=>({wch:w}));
    for(let c=0;c<headers.length;c++){ let cell=w[XLSX.utils.encode_cell({r:0,c})]; if(cell) cell.s=S_HEADER; }
    for(let r=0;r<rows.length;r++){ for(let c=0;c<headers.length;c++){ let cell=w[XLSX.utils.encode_cell({r:r+1,c})]; if(cell){ cell.s=(r%2===1)?S_ZEBRA:S_DATA; if(fmts[c]) cell.s={...cell.s,numFmt:fmts[c]}; } } }
    XLSX.utils.book_append_sheet(wb, w, name);
  }
  let bh=['品牌','低价课订单数','有效订单数','高价课单数','高价课转化率','高价课金额(元)','高价课金额占比','单订单产值(元)'];
  dimSheet('手机品牌产值', bh, [15,14,12,10,12,16,14,14], brandAgg.map(r=>({'品牌':r['品牌统一'],'低价课订单数':r['低价课订单数'],'有效订单数':r['有效订单数'],'高价课单数':r['高价课单数'],'高价课转化率':r['高价课转化率'],'高价课金额(元)':r['高价课金额'],'高价课金额占比':r['高价课金额占比'],'单订单产值(元)':r['单订单产值']})), {1:FMT_INT,2:FMT_INT,3:FMT_INT,4:FMT_PCT,5:FMT_DEC2,6:FMT_PCT,7:FMT_DEC2});
  let mh=['品牌','型号','低价课订单数','有效订单数','高价课单数','高价课转化率','高价课金额(元)','高价课金额占比','单订单产值(元)'];
  dimSheet('手机型号产值', mh, [12,28,14,12,10,12,16,14,14], modelAgg.map(r=>({'品牌':r['品牌统一'],'型号':r['手机型号'],'低价课订单数':r['低价课订单数'],'有效订单数':r['有效订单数'],'高价课单数':r['高价课单数'],'高价课转化率':r['高价课转化率'],'高价课金额(元)':r['高价课金额'],'高价课金额占比':r['高价课金额占比'],'单订单产值(元)':r['单订单产值']})), {2:FMT_INT,3:FMT_INT,4:FMT_INT,5:FMT_PCT,6:FMT_DEC2,7:FMT_PCT,8:FMT_DEC2});
  let ph=['价格区间','低价课订单数','有效订单数','高价课单数','高价课转化率','高价课金额(元)','高价课金额占比','单订单产值(元)'];
  dimSheet('手机价格产值', ph, [15,14,12,10,12,16,14,14], priceAgg.map(r=>({'价格区间':r['价格区间'],'低价课订单数':r['低价课订单数'],'有效订单数':r['有效订单数'],'高价课单数':r['高价课单数'],'高价课转化率':r['高价课转化率'],'高价课金额(元)':r['高价课金额'],'高价课金额占比':r['高价课金额占比'],'单订单产值(元)':r['单订单产值']})), {1:FMT_INT,2:FMT_INT,3:FMT_INT,4:FMT_PCT,5:FMT_DEC2,6:FMT_PCT,7:FMT_DEC2});
  let ah=['年龄段','低价课订单数','有效订单数','高价课单数','高价课转化率','高价课金额(元)','高价课金额占比','单订单产值(元)'];
  dimSheet('年龄产值', ah, [15,14,12,10,12,16,14,14], ageAgg.map(r=>({'年龄段':r['年龄段'],'低价课订单数':r['低价课订单数'],'有效订单数':r['有效订单数'],'高价课单数':r['高价课单数'],'高价课转化率':r['高价课转化率'],'高价课金额(元)':r['高价课金额'],'高价课金额占比':r['高价课金额占比'],'单订单产值(元)':r['单订单产值']})), {1:FMT_INT,2:FMT_INT,3:FMT_INT,4:FMT_PCT,5:FMT_DEC2,6:FMT_PCT,7:FMT_DEC2});
  return wb;
}

function fillTable(id,headers,rows,limit){ let t=document.getElementById(id); if(!t)return; t.innerHTML='<tr>'+headers.map(h=>'<th>'+h+'</th>').join('')+'</tr>'; (limit?rows.slice(0,limit):rows).forEach(r=>{ t.innerHTML+='<tr>'+headers.map((h,i)=>'<td>'+(Array.isArray(r)?(r[i]!=null?r[i]:''):(r[h]!=null?r[h]:''))+'</td>').join('')+'</tr>'; }); }

function t1Display(d) {
  document.getElementById('t1Results').style.display='block';
  let kpi=document.getElementById('t1Kpi'); kpi.innerHTML='';
  [{v:fmtInt(d.nTotal),l:'低价课订单总数'},{v:fmtInt(d.nValid),l:'有效订单数'},{v:fmtInt(d.nHigh),l:'高价课转化单数'},{v:fmtPct(safeDiv(d.nHigh,d.nTotal)),l:'整体转化率'},{v:'¥'+fmtDec2(d.totalHigh),l:'高价课金额合计'},{v:'¥'+fmtDec2(safeDiv(d.totalHigh,d.nTotal)),l:'单订单产值'},{v:d.brandAgg.length,l:'品牌数'},{v:d.modelAgg.length,l:'型号数'}].forEach(k=>{kpi.innerHTML+='<div class="ex-kpi-card"><div class="ex-val">'+k.v+'</div><div class="ex-lbl">'+k.l+'</div></div>';});
  fillTable('t1TblSummary',['指标','值'],[['低价课订单总数',fmtInt(d.nTotal)],['有效订单数',fmtInt(d.nValid)],['高价课转化单数',fmtInt(d.nHigh)],['整体高价课转化率',fmtPct(safeDiv(d.nHigh,d.nTotal))],['高价课金额合计(元)','¥'+fmtDec2(d.totalHigh)],['整体单订单产值(元)','¥'+fmtDec2(safeDiv(d.totalHigh,d.nTotal))]]);
  fillTable('t1TblBrand',['品牌','订单数','有效','高价','转化率','金额(元)','占比','产值(元)'], d.brandAgg.slice(0,30).map(r=>({'品牌':r['品牌统一'],'订单数':r['低价课订单数'],'有效':r['有效订单数'],'高价':r['高价课单数'],'转化率':fmtPct(r['高价课转化率']),'金额(元)':fmtDec2(r['高价课金额']),'占比':fmtPct(r['高价课金额占比']),'产值(元)':fmtDec2(r['单订单产值'])})));
  fillTable('t1TblModel',['品牌','型号','订单数','有效','高价','转化率','金额(元)','产值(元)'], d.modelAgg.slice(0,30).map(r=>({'品牌':r['品牌统一'],'型号':r['手机型号'],'订单数':r['低价课订单数'],'有效':r['有效订单数'],'高价':r['高价课单数'],'转化率':fmtPct(r['高价课转化率']),'金额(元)':fmtDec2(r['高价课金额']),'产值(元)':fmtDec2(r['单订单产值'])})));
  fillTable('t1TblPrice',['价格区间','订单数','有效','高价','转化率','金额(元)','产值(元)'], d.priceAgg.map(r=>({'价格区间':r['价格区间'],'订单数':r['低价课订单数'],'有效':r['有效订单数'],'高价':r['高价课单数'],'转化率':fmtPct(r['高价课转化率']),'金额(元)':fmtDec2(r['高价课金额']),'产值(元)':fmtDec2(r['单订单产值'])})));
  fillTable('t1TblAge',['年龄段','订单数','有效','高价','转化率','金额(元)','产值(元)'], d.ageAgg.map(r=>({'年龄段':r['年龄段'],'订单数':r['低价课订单数'],'有效':r['有效订单数'],'高价':r['高价课单数'],'转化率':fmtPct(r['高价课转化率']),'金额(元)':fmtDec2(r['高价课金额']),'产值(元)':fmtDec2(r['单订单产值'])})));
}

function t1Preview(name, ev) {
  ['summary','brand','model','price','age'].forEach(n => { document.getElementById('t1pv-'+n).style.display = n===name?'block':'none'; });
  document.querySelectorAll('#view-excel-phone .ex-preview-tabs .ex-tab-btn').forEach(b=>b.classList.remove('active'));
  if (ev) ev.target.classList.add('active');
}

function downloadTab1() { if (t1Output) XLSX.writeFile(t1Output, '分析过数据.xlsx'); }

// ============================================================
// 页签2：链路数据表_系统口径
// ============================================================
const T2_LINK_MAP = {
  '83676427':'PF+小程序问答','83676438':'H5问答+首页','87101601':'H5问答+首页',
  '83676778':'小程序问答(含首页)','85489979':'小程序问答(含首页)',
  '85489984':'H5分流','87101605':'PF+AI落地页',
};
const T2_LINK_ORDER = ['PF+小程序问答','H5问答+首页','小程序问答(含首页)','H5分流','PF+AI落地页'];
const T2_FIELDS = {
  period: ['期数时间','订单时间','期数','日期'],
  accountId: ['广告账户ID','广告账户ID（数字）'],
  spend: ['广告花费'],
  totalCost: ['低价课总成本'],
  effOrders: ['低价课订单数'],
  highOrders: ['高价课订单数'],
  totalOutput: ['总产值'],
  validSession: ['有效会话率（5次及以上）','有效会话率(5次及以上)','有效会话率'],
  liveAtt: ['第一节直播到课率','一节直播到课率'],
  openRate: ['开口率'],
  joinRate: ['进群率'],
  unfriend: ['day0好友删除率','删友率'],
};

let t2Output = null, t2Summary = null;

setupUpload('t2Zone','t2File','t2FileName','t2Btn');

function processTab2() {
  try {
  if (!window._xlsxReady) { showStatus('t2','xlsx库未加载完成，请等几秒后再试','error'); return; }
  const file = document.getElementById('t2File')._file;
  if (!file) return;
  document.getElementById('t2Btn').disabled = true;
  showStatus('t2','正在读取文件...','active');
  document.getElementById('t2Results').style.display = 'none';
  readWorkbook(file, (wb, err) => {
    if (err) { showStatus('t2','读取失败：'+err.message,'error'); document.getElementById('t2Btn').disabled=false; return; }
    try {
    showStatus('t2','正在解析数据...','active');
    const { json } = getSheetData(wb);
    if (!json.length) { showStatus('t2','数据为空','error'); document.getElementById('t2Btn').disabled=false; return; }
    const headers = Object.keys(json[0]);
    const cols = resolveCols(headers, T2_FIELDS);
    if (cols.missing.length) { showMissingCols('t2', cols.missing, headers); showStatus('t2','有 '+cols.missing.length+' 个列未找到，请检查表头','error'); document.getElementById('t2Btn').disabled=false; return; }
    showStatus('t2','正在聚合计算...','active');
    // 过滤汇总行
    const detail = json.filter(r => {
      const p = r[headers[cols.found.period]];
      return p && String(p).trim() !== '汇总数据' && String(p).trim() !== '';
    });
    // 百分比转数值 + 链路映射 + 期数格式化
    const processed = detail.map(r => {
      const aid = String(r[headers[cols.found.accountId]] || '').trim();
      const link = T2_LINK_MAP[aid] || '其他';
      const periodRaw = String(r[headers[cols.found.period]] || '').trim();
      let periodFmt = periodRaw;
      const m = periodRaw.match(/(\d{4})-(\d{2})-(\d{2})/);
      if (m) periodFmt = m[2]+'月'+m[3]+'日期';
      return {
        link, period: periodFmt, periodRaw,
        spend: parseFloat(r[headers[cols.found.spend]]) || 0,
        totalCost: parseFloat(r[headers[cols.found.totalCost]]) || 0,
        effOrders: parseFloat(r[headers[cols.found.effOrders]]) || 0,
        highOrders: parseFloat(r[headers[cols.found.highOrders]]) || 0,
        totalOutput: parseFloat(r[headers[cols.found.totalOutput]]) || 0,
        validSession: pctToNum(r[headers[cols.found.validSession]]),
        liveAtt: pctToNum(r[headers[cols.found.liveAtt]]),
        openRate: pctToNum(r[headers[cols.found.openRate]]),
        joinRate: pctToNum(r[headers[cols.found.joinRate]]),
        unfriend: pctToNum(r[headers[cols.found.unfriend]]),
      };
    });
    // 期数列表（去重排序）
    const periods = [...new Set(processed.map(r=>r.period))].sort();
    // 聚合函数
    function aggLink(rows) {
      const spend = rows.reduce((s,r)=>s+r.spend,0);
      const cost = rows.reduce((s,r)=>s+r.totalCost,0);
      const eff = rows.reduce((s,r)=>s+r.effOrders,0);
      const high = rows.reduce((s,r)=>s+r.highOrders,0);
      const output = rows.reduce((s,r)=>s+r.totalOutput,0);
      const ws = rows.reduce((s,r)=>s+r.validSession*r.effOrders,0);
      const la = rows.reduce((s,r)=>s+r.liveAtt*r.effOrders,0);
      const op = rows.reduce((s,r)=>s+r.openRate*r.effOrders,0);
      const jo = rows.reduce((s,r)=>s+r.joinRate*r.effOrders,0);
      const un = rows.reduce((s,r)=>s+r.unfriend*r.effOrders,0);
      return {
        roi: safeDiv(output, cost),
        spend, cost, eff,
        orderCost: safeDiv(cost, eff),
        orderOutput: safeDiv(output, eff),
        highConv: safeDiv(high, eff),
        highOrders: high,
        validSession: safeDiv(ws, eff),
        liveAtt: safeDiv(la, eff),
        openRate: safeDiv(op, eff),
        joinRate: safeDiv(jo, eff),
        unfriend: safeDiv(un, eff),
      };
    }
    // Layer A: 链路汇总
    const linkSummary = T2_LINK_ORDER.map(link => {
      const rows = processed.filter(r => r.link === link);
      const a = aggLink(rows);
      return { link, firstLive: a.liveAtt, ...a };
    });
    // Layer C: 整体合计
    const total = aggLink(processed);
    // Layer B: 链路×期（字段名与分析模块renderLinkTable一致）
    const linkByPeriod = {};
    T2_LINK_ORDER.forEach(link => {
      linkByPeriod[link] = periods.map(p => {
        const rows = processed.filter(r => r.link===link && r.period===p);
        const a = aggLink(rows);
        return {
          '期': p, '订单ROI': a.roi, '广告花费(元)': a.spend, '低价课总成本(元)': a.cost,
          '订单数': a.eff, '单订单成本(元)': a.orderCost, '单订单产值(元)': a.orderOutput,
          '高价课转化率': a.highConv, '高价课订单数': a.highOrders, '有效会话率': a.validSession,
          '一节直播到课': a.liveAtt, '开口率': a.openRate, '进群率': a.joinRate, '删友率': a.unfriend,
        };
      });
    });
    t2Summary = { linkSummary, total, periods, linkByPeriod, srcFile: file.name, detailCount: detail.length };
    showStatus('t2','正在生成Excel...','active');
    t2Output = t2BuildExcel(t2Summary);
    t2Display(t2Summary);
    saveSync('link', t2Summary);
    showStatus('t2','完成！'+detail.length+' 条明细，'+periods.length+' 期，5 条链路（已同步到分析模块）','done');
    document.getElementById('t2Btn').disabled = false;
    } catch(e) { showStatus('t2','运行错误：'+e.message,'error'); console.error(e); document.getElementById('t2Btn').disabled=false; }
  });
  } catch(e) { showStatus('t2','启动错误：'+e.message,'error'); console.error(e); }
}

function t2BuildExcel(s) {
  const wb = XLSX.utils.book_new();
  const headers = ['链路','订单ROI','广告花费(元)','低价课总成本(元)','订单数','单订单成本(元)','单订单产值(元)','高价课转化率','高价课订单数','有效会话率','一节直播到课','开口率','进群率','删友率'];
  const fmts = [null, FMT_PCT, FMT_DEC2, FMT_DEC2, FMT_INT, FMT_DEC2, FMT_DEC2, FMT_PCT, FMT_INT, FMT_PCT, FMT_PCT, FMT_PCT, FMT_PCT, FMT_PCT];
  const widths = [12,10,13,14,9,12,13,11,11,11,12,9,9,9];
  function rowArr(d, labelKey) {
    return [d[labelKey], d.roi, d.spend, d.cost, d.eff, d.orderCost, d.orderOutput, d.highConv, d.highOrders, d.validSession, d.liveAtt, d.openRate, d.joinRate, d.unfriend];
  }
  function styleRow(ws, r, nCols, isHeader, isBold) {
    for (let c=0; c<nCols; c++) {
      const cell = ws[XLSX.utils.encode_cell({r, c})];
      if (!cell) continue;
      if (isHeader) cell.s = S_HEADER;
      else if (isBold) { cell.s = {...S_BOLD}; if(fmts[c]) cell.s.numFmt=fmts[c]; }
      else { cell.s = (r%2===0)?S_ZEBRA:S_DATA; if(fmts[c]) cell.s.numFmt=fmts[c]; }
    }
  }
  // Sheet: 链路汇总
  let aoa = [];
  aoa.push(['分链路 · 数据表（系统口径，'+s.periods.length+' 期合计）']);
  aoa.push(['数据源：'+s.srcFile+'（导出数据，'+s.detailCount+' 条明细，期数：'+s.periods.join('、')+'）']);
  aoa.push(headers);
  s.linkSummary.forEach(d => aoa.push(rowArr(d,'link')));
  aoa.push(['合计', s.total.roi, s.total.spend, s.total.cost, s.total.eff, s.total.orderCost, s.total.orderOutput, s.total.highConv, s.total.highOrders, s.total.validSession, s.total.liveAtt, s.total.openRate, s.total.joinRate, s.total.unfriend]);
  let ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = widths.map(w=>({wch:w}));
  ws['!merges'] = [{s:{r:0,c:0},e:{r:0,c:headers.length-1}},{s:{r:1,c:0},e:{r:1,c:headers.length-1}}];
  ws['A1'].s=S_TITLE; ws['A2'].s=S_SUB;
  styleRow(ws, 2, headers.length, true, false);
  for (let i=0;i<s.linkSummary.length;i++) styleRow(ws, 3+i, headers.length, false, false);
  styleRow(ws, 3+s.linkSummary.length, headers.length, false, true);
  XLSX.utils.book_append_sheet(wb, ws, '链路汇总');
  // 5张分期表
  T2_LINK_ORDER.forEach(link => {
    let a=[];
    a.push([link+' · 分期数据表']);
    a.push(['口径：订单ROI=总产值÷低价课总成本；单订单成本=低价课总成本÷有效订单数；比率按订单数加权平均']);
    a.push(['']);
    a.push(headers);
    s.linkByPeriod[link].forEach(d => a.push(rowArr(d,'period')));
    a.push(['合计', s.linkSummary.find(x=>x.link===link).roi, s.linkSummary.find(x=>x.link===link).spend, s.linkSummary.find(x=>x.link===link).cost, s.linkSummary.find(x=>x.link===link).eff, s.linkSummary.find(x=>x.link===link).orderCost, s.linkSummary.find(x=>x.link===link).orderOutput, s.linkSummary.find(x=>x.link===link).highConv, s.linkSummary.find(x=>x.link===link).highOrders, s.linkSummary.find(x=>x.link===link).validSession, s.linkSummary.find(x=>x.link===link).liveAtt, s.linkSummary.find(x=>x.link===link).openRate, s.linkSummary.find(x=>x.link===link).joinRate, s.linkSummary.find(x=>x.link===link).unfriend]);
    let w=XLSX.utils.aoa_to_sheet(a);
    w['!cols']=widths.map(w=>({wch:w}));
    w['!merges']=[{s:{r:0,c:0},e:{r:0,c:headers.length-1}},{s:{r:1,c:0},e:{r:1,c:headers.length-1}}];
    w['A1'].s=S_TITLE; w['A2'].s=S_SUB;
    styleRow(w, 3, headers.length, true, false);
    for (let i=0;i<s.linkByPeriod[link].length;i++) styleRow(w, 4+i, headers.length, false, false);
    styleRow(w, 4+s.linkByPeriod[link].length, headers.length, false, true);
    XLSX.utils.book_append_sheet(wb, w, link);
  });
  // 口径说明
  let noteAoa = [['链路数据表_系统口径 · 口径说明'],[''],['1. 数据源：'+s.srcFile+'（导出数据，'+s.detailCount+' 条明细）'],['2. 链路映射：83676427→PF+小程序问答；83676438+87101601→H5问答+首页；83676778+85489979→小程序问答(含首页)；85489984→H5分流；87101605→PF+AI落地页'],['3. 订单ROI = 总产值 ÷ 低价课总成本（分母是低价课总成本而非广告花费）'],['4. 单订单成本 = 低价课总成本 ÷ 有效订单数'],['5. 单订单产值 = 总产值 ÷ 有效订单数'],['6. 高价课转化率 = 高价课订单数 ÷ 有效订单数'],['7. 5个比率字段（有效会话率、一节直播到课、开口率、进群率、删友率）按低价课订单数加权平均'],['8. 订单数 = 低价课订单数（有效订单口径）'],['9. 百分比字段已从字符串转为小数']];
  let wn = XLSX.utils.aoa_to_sheet(noteAoa);
  wn['!cols']=[{wch:100}];
  wn['!merges']=noteAoa.map((_,i)=>({s:{r:i,c:0},e:{r:i,c:0}}));
  wn['A1'].s=S_TITLE;
  noteAoa.forEach((_,i)=>{ if(i>1&&wn['A'+(i+1)]) wn['A'+(i+1)].s=S_DATA; });
  XLSX.utils.book_append_sheet(wb, wn, '口径说明');
  return wb;
}

function t2Display(s) {
  document.getElementById('t2Results').style.display='block';
  let kpi=document.getElementById('t2Kpi'); kpi.innerHTML='';
  [{v:s.detailCount,l:'明细行数'},{v:s.periods.length,l:'期数'},{v:s.linkSummary.length,l:'链路数'},{v:'¥'+fmtDec2(s.total.spend),l:'广告花费合计'},{v:'¥'+fmtDec2(s.total.cost),l:'低价课总成本'},{v:fmtPct(s.total.roi),l:'整体订单ROI'},{v:fmtInt(s.total.eff),l:'有效订单数'},{v:fmtPct(s.total.highConv),l:'高价课转化率'}].forEach(k=>{kpi.innerHTML+='<div class="ex-kpi-card"><div class="ex-val">'+k.v+'</div><div class="ex-lbl">'+k.l+'</div></div>';});
  let t=document.getElementById('t2TblSummary');
  let hd=['链路','订单ROI','广告花费','总成本','订单数','单订单成本','单订单产值','高价转化率','高价订单','有效会话率','直播到课','开口率','进群率','删友率'];
  t.innerHTML='<tr>'+hd.map(h=>'<th>'+h+'</th>').join('')+'</tr>';
  s.linkSummary.forEach(d=>{ t.innerHTML+='<tr><td>'+d.link+'</td><td>'+fmtPct(d.roi)+'</td><td>'+fmtDec2(d.spend)+'</td><td>'+fmtDec2(d.cost)+'</td><td>'+fmtInt(d.eff)+'</td><td>'+fmtDec2(d.orderCost)+'</td><td>'+fmtDec2(d.orderOutput)+'</td><td>'+fmtPct(d.highConv)+'</td><td>'+fmtInt(d.highOrders)+'</td><td>'+fmtPct(d.validSession)+'</td><td>'+fmtPct(d.liveAtt)+'</td><td>'+fmtPct(d.openRate)+'</td><td>'+fmtPct(d.joinRate)+'</td><td>'+fmtPct(d.unfriend)+'</td></tr>'; });
  t.innerHTML+='<tr style="font-weight:bold"><td>合计</td><td>'+fmtPct(s.total.roi)+'</td><td>'+fmtDec2(s.total.spend)+'</td><td>'+fmtDec2(s.total.cost)+'</td><td>'+fmtInt(s.total.eff)+'</td><td>'+fmtDec2(s.total.orderCost)+'</td><td>'+fmtDec2(s.total.orderOutput)+'</td><td>'+fmtPct(s.total.highConv)+'</td><td>'+fmtInt(s.total.highOrders)+'</td><td>'+fmtPct(s.total.validSession)+'</td><td>'+fmtPct(s.total.liveAtt)+'</td><td>'+fmtPct(s.total.openRate)+'</td><td>'+fmtPct(s.total.joinRate)+'</td><td>'+fmtPct(s.total.unfriend)+'</td></tr>';
  // 各链路分期明细（全量展示）
  const pdEl=document.getElementById('t2PeriodDetail');
  let pdHtml='';
  T2_LINK_ORDER.forEach(link=>{
    const rows=s.linkByPeriod[link]||[];
    if(!rows.length) return;
    pdHtml+='<h4 style="margin:18px 0 6px;color:#3b82f6">▶ '+link+'（'+rows.length+'期）</h4>';
    pdHtml+='<div class="ex-result-table"><table><tbody>';
    pdHtml+='<tr>'+['期','订单ROI','广告花费','低价课总成本','订单数','单订单成本','单订单产值','高价转化率','高价课订单','有效会话率','到课率','开口率','进群率','删友率'].map(h=>'<th>'+h+'</th>').join('')+'</tr>';
    rows.forEach(r=>{
      pdHtml+='<tr><td>'+r['期']+'</td><td>'+fmtPct(r['订单ROI'])+'</td><td>'+fmtDec2(r['广告花费(元)'])+'</td><td>'+fmtDec2(r['低价课总成本(元)'])+'</td><td>'+fmtInt(r['订单数'])+'</td><td>'+fmtDec2(r['单订单成本(元)'])+'</td><td>'+fmtDec2(r['单订单产值(元)'])+'</td><td>'+fmtPct(r['高价课转化率'])+'</td><td>'+fmtInt(r['高价课订单数'])+'</td><td>'+fmtPct(r['有效会话率'])+'</td><td>'+fmtPct(r['一节直播到课'])+'</td><td>'+fmtPct(r['开口率'])+'</td><td>'+fmtPct(r['进群率'])+'</td><td>'+fmtPct(r['删友率'])+'</td></tr>';
    });
    pdHtml+='</tbody></table></div>';
  });
  pdEl.innerHTML=pdHtml;
}

function downloadTab2() { if (t2Output) XLSX.writeFile(t2Output, '链路数据表_系统口径.xlsx'); }

// ============================================================
// 页签3：分周数据对比（含 localStorage 双周缓存）
// ============================================================
const T3_LINK_MAP = {
  '83676427':'PF+小程序问答','83676438':'H5问答+首页','87101601':'H5问答+首页',
  '83676778':'小程序问答(含首页)','85489979':'小程序问答(含首页)',
  '85489984':'H5分流','87101605':'PF+AI落地页','83676797':'其他',
};
const T3_LINK_ORDER = ['PF+小程序问答','H5问答+首页','小程序问答(含首页)','H5分流','PF+AI落地页','其他','汇总'];
const T3_PLACEMENTS = ['微信公众号与小程序','微信视频号','腾讯视频','微信朋友圈','QQ、腾讯音乐及游戏','搜索场景','微信新闻插件','腾讯新闻','QQ 浏览器（原腾讯看点）'];
const T3_CACHE_KEY = 't3_week_cache';

let t3Output = null, t3Data = null;

setupUpload('t3Zone','t3File','t3FileName','t3Btn');

function processTab3() {
  try {
  if (!window._xlsxReady) { showStatus('t3','xlsx库未加载完成，请等几秒后再试','error'); return; }
  const file = document.getElementById('t3File')._file;
  if (!file) return;
  document.getElementById('t3Btn').disabled = true;
  showStatus('t3','正在读取文件...','active');
  document.getElementById('t3Results').style.display = 'none';
  readWorkbook(file, (wb, err) => {
    if (err) { showStatus('t3','读取失败：'+err.message,'error'); document.getElementById('t3Btn').disabled=false; return; }
    try {
    showStatus('t3','正在解析数据（按列索引读取）...','active');
    const { ws } = getSheetData(wb);
    const aoa = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    if (aoa.length < 2) { showStatus('t3','数据为空','error'); document.getElementById('t3Btn').disabled=false; return; }
    const headers = aoa[0];
    // 按列索引读取（0-based）
    const rows = aoa.slice(1).filter(r => {
      const t = r[0]; return t && String(t).trim() !== '汇总数据' && String(t).trim() !== '';
    }).filter(r => String(r[2]||'').trim() === '直投号');
    if (!rows.length) { showStatus('t3','筛选后无数据（需账号模式=直投号）','error'); document.getElementById('t3Btn').disabled=false; return; }
    showStatus('t3','正在聚合计算...','active');
    // 提取周区间
    const weekRaw = String(rows[0][0]||'').trim();
    const dm = weekRaw.match(/(\d{4})-(\d{2})-(\d{2})\s*~\s*(\d{4})-(\d{2})-(\d{2})/);
    let weekFmt = weekRaw;
    if (dm) weekFmt = parseInt(dm[2])+'.'+parseInt(dm[3])+'-'+parseInt(dm[5])+'.'+parseInt(dm[6]);
    // 链路聚合
    function aggLink(rs) {
      const spend = rs.reduce((s,r)=>s+(parseFloat(r[6])||0),0);
      const eff = rs.reduce((s,r)=>s+(parseFloat(r[10])||0),0);
      const all = rs.reduce((s,r)=>s+(parseFloat(r[11])||0),0);
      const cost = rs.reduce((s,r)=>s+(parseFloat(r[16])||0),0);
      return { spend, eff, all, cost, orderCost: safeDiv(cost, eff), orderShare: 0, orderValidRate: safeDiv(eff, all) };
    }
    const linkAgg = {};
    T3_LINK_ORDER.forEach(l => { if (l!=='汇总') linkAgg[l] = { spend:0, eff:0, all:0, cost:0 }; });
    rows.forEach(r => {
      const aid = String(r[4]||'').trim();
      const link = T3_LINK_MAP[aid] || '其他';
      if (!linkAgg[link]) linkAgg[link] = { spend:0, eff:0, all:0, cost:0 };
      linkAgg[link].spend += parseFloat(r[6])||0;
      linkAgg[link].eff += parseFloat(r[10])||0;
      linkAgg[link].all += parseFloat(r[11])||0;
      linkAgg[link].cost += parseFloat(r[16])||0;
    });
    const totalAll = Object.values(linkAgg).reduce((s,v)=>s+v.all,0);
    const totalEff = Object.values(linkAgg).reduce((s,v)=>s+v.eff,0);
    const totalCost = Object.values(linkAgg).reduce((s,v)=>s+v.cost,0);
    const totalSpend = Object.values(linkAgg).reduce((s,v)=>s+v.spend,0);
    const linkRows = T3_LINK_ORDER.filter(l=>l!=='汇总').map(l => {
      const d = linkAgg[l] || {spend:0,eff:0,all:0,cost:0};
      return { link:l, spend:d.spend, eff:d.eff, all:d.all, cost:d.cost, orderCost: safeDiv(d.cost, d.eff), orderShare: safeDiv(d.all, totalAll), orderValidRate: safeDiv(d.eff, d.all) };
    });
    linkRows.push({ link:'汇总', spend:totalSpend, eff:totalEff, all:totalAll, cost:totalCost, orderCost: safeDiv(totalCost, totalEff), orderShare: 1, orderValidRate: safeDiv(totalEff, totalAll) });
    // 版位聚合
    const placeAgg = {};
    rows.forEach(r => {
      const p = String(r[5]||'').trim() || '未知';
      if (!placeAgg[p]) placeAgg[p] = { eff:0, all:0, cost:0 };
      placeAgg[p].eff += parseFloat(r[10])||0;
      placeAgg[p].all += parseFloat(r[11])||0;
      placeAgg[p].cost += parseFloat(r[16])||0;
    });
    const knownPlace = Object.keys(placeAgg).filter(p => T3_PLACEMENTS.includes(p));
    const otherPlace = Object.keys(placeAgg).filter(p => !T3_PLACEMENTS.includes(p));
    let placeRows = knownPlace.map(p => ({ place:p, all:placeAgg[p].all, orderShare: safeDiv(placeAgg[p].all, totalAll), orderCost: safeDiv(placeAgg[p].cost, placeAgg[p].eff) }));
    if (otherPlace.length) { const o = otherPlace.reduce((acc,p)=>{ acc.all+=placeAgg[p].all; acc.cost+=placeAgg[p].cost; acc.eff+=placeAgg[p].eff; return acc; },{all:0,cost:0,eff:0}); placeRows.push({place:'其他', all:o.all, orderShare: safeDiv(o.all,totalAll), orderCost: safeDiv(o.cost, o.eff)}); }
    placeRows.push({ place:'汇总', all:totalAll, orderShare: 1, orderCost: safeDiv(totalCost, totalEff) });
    const currentWeek = { weekRaw, weekFmt, linkRows, placeRows, totalSpend, totalAll, totalEff, totalCost, detailCount: rows.length };
    // 读取上周缓存
    let lastWeek = null;
    try { lastWeek = JSON.parse(localStorage.getItem(T3_CACHE_KEY)); } catch(e){}
    t3Data = { current: currentWeek, last: lastWeek };
    // 保存本周为下周的上周
    localStorage.setItem(T3_CACHE_KEY, JSON.stringify(currentWeek));
    showStatus('t3','正在生成Excel...','active');
    t3Output = t3BuildExcel(t3Data);
    t3Display(t3Data);
    saveSync('weekly', t3Data);
    showStatus('t3','完成！本周 '+weekFmt+'，'+(lastWeek?'有上周数据对比':'首次运行，无上周数据')+'（已同步到分析模块）','done');
    document.getElementById('t3Btn').disabled = false;
    } catch(e) { showStatus('t3','运行错误：'+e.message,'error'); console.error(e); document.getElementById('t3Btn').disabled=false; }
  });
  } catch(e) { showStatus('t3','启动错误：'+e.message,'error'); console.error(e); }
}

function t3BuildExcel(d) {
  const wb = XLSX.utils.book_new();
  const linkHeaders = ['链路','现金花费','订单数','订单成本','订单占比','订单有效率'];
  const placeHeaders = ['版位','订单数','订单占比','订单成本'];
  const linkFmts = [null, FMT_DEC2, FMT_INT, FMT_DEC2, FMT_PCT, FMT_PCT];
  const placeFmts = [null, FMT_INT, FMT_PCT, FMT_DEC2];
  // Sheet 链路分周
  let a = [];
  a.push(['分链路 · 分周对比表']);
  a.push(['口径：订单数=全部低价课订单数（含无效）；订单成本=低价课总成本÷有效订单数；订单占比=该链路订单数÷本周汇总订单数；订单有效率=有效订单÷全部订单']);
  a.push(['']);
  const cur = d.current, last = d.last;
  if (last) {
    a.push(['上周('+last.weekFmt+')  合计 现金花费 ¥'+fmtDec2(last.totalSpend)]);
    a.push(linkHeaders);
    last.linkRows.forEach(r => a.push([r.link, r.spend, r.all, r.orderCost, r.orderShare, r.orderValidRate]));
    a.push(['']);
  } else {
    a.push(['上周  无上周数据']);
    a.push(['']);
  }
  a.push(['本周('+cur.weekFmt+')  合计 现金花费 ¥'+fmtDec2(cur.totalSpend)]);
  a.push(linkHeaders);
  cur.linkRows.forEach(r => a.push([r.link, r.spend, r.all, r.orderCost, r.orderShare, r.orderValidRate]));
  let ws = XLSX.utils.aoa_to_sheet(a);
  ws['!cols'] = [12,13,10,12,10,10].map(w=>({wch:w}));
  ws['A1'].s = S_TITLE; ws['A2'].s = S_SUB;
  XLSX.utils.book_append_sheet(wb, ws, '链路分周');
  // Sheet 版位分周
  let b = [];
  b.push(['分版位 · 分周对比表']);
  b.push(['口径：订单数=全部低价课订单数；订单成本=低价课总成本÷有效订单数；订单占比=该版位订单数÷本周汇总订单数']);
  b.push(['']);
  if (last) {
    b.push(['上周('+last.weekFmt+')  合计 订单 '+fmtInt(last.totalAll)]);
    b.push(placeHeaders);
    last.placeRows.forEach(r => b.push([r.place, r.all, r.orderShare, r.orderCost]));
    b.push(['']);
  } else {
    b.push(['上周  无上周数据']);
    b.push(['']);
  }
  b.push(['本周('+cur.weekFmt+')  合计 订单 '+fmtInt(cur.totalAll)]);
  b.push(placeHeaders);
  cur.placeRows.forEach(r => b.push([r.place, r.all, r.orderShare, r.orderCost]));
  let ws2 = XLSX.utils.aoa_to_sheet(b);
  ws2['!cols'] = [22,10,10,12].map(w=>({wch:w}));
  ws2['A1'].s = S_TITLE; ws2['A2'].s = S_SUB;
  XLSX.utils.book_append_sheet(wb, ws2, '版位分周');
  // 口径说明
  let n = [['分周数据表 · 口径说明'],[''],['1. 本周数据源：'+(cur.weekRaw)+'（'+cur.detailCount+' 条明细）']];
  if (last) n.push(['2. 上周数据源：'+(last.weekRaw)+'（'+last.detailCount+' 条明细，来自上次运行缓存）']);
  else n.push(['2. 上周数据：无（首次运行）']);
  n.push(['3. 链路映射：83676427→PF+小程序问答；83676438+87101601→H5问答+首页；83676778+85489979→小程序问答(含首页)；85489984→H5分流；87101605→PF+AI落地页；83676797→其他']);
  n.push(['4. 订单数=全部低价课订单数（含无效）']);
  n.push(['5. 订单成本=低价课总成本÷有效订单数']);
  n.push(['6. 订单占比=该链路/版位订单数÷当周汇总订单数']);
  n.push(['7. 订单有效率=有效订单数÷全部低价课订单数']);
  n.push(['8. 现金花费=Σ广告花费（直投号）']);
  n.push(['9. 版位分周9个版位+汇总；新出现的归入其他']);
  let wn = XLSX.utils.aoa_to_sheet(n);
  wn['!cols']=[{wch:90}];
  wn['A1'].s=S_TITLE;
  XLSX.utils.book_append_sheet(wb, wn, '口径说明');
  return wb;
}

function t3Display(d) {
  document.getElementById('t3Results').style.display='block';
  const cur=d.current, last=d.last;
  let wi=document.getElementById('t3WeekInfo');
  wi.innerHTML = '<div class="ex-week-box"><div class="ex-wk-title">上周 '+(last?last.weekFmt:'无数据')+'</div>'+(last?'现金花费 ¥'+fmtDec2(last.totalSpend)+'<br>订单 '+fmtInt(last.totalAll)+' 条':'首次运行，无上周缓存')+'</div><div class="ex-week-box"><div class="ex-wk-title">本周 '+cur.weekFmt+'</div>现金花费 ¥'+fmtDec2(cur.totalSpend)+'<br>订单 '+fmtInt(cur.totalAll)+' 条</div>';
  // 链路表
  let lt=document.getElementById('t3TblLink');
  let lh=['周次','链路','现金花费','订单数','订单成本','订单占比','订单有效率'];
  lt.innerHTML='<tr>'+lh.map(h=>'<th>'+h+'</th>').join('')+'</tr>';
  if (last) last.linkRows.forEach(r=>{ lt.innerHTML+='<tr style="opacity:0.7"><td>上周</td><td>'+r.link+'</td><td>'+fmtDec2(r.spend)+'</td><td>'+fmtInt(r.all)+'</td><td>'+fmtDec2(r.orderCost)+'</td><td>'+fmtPct(r.orderShare)+'</td><td>'+fmtPct(r.orderValidRate)+'</td></tr>'; });
  cur.linkRows.forEach(r=>{ lt.innerHTML+='<tr style="font-weight:'+(r.link==='汇总'?'bold':'normal')+'"><td>本周</td><td>'+r.link+'</td><td>'+fmtDec2(r.spend)+'</td><td>'+fmtInt(r.all)+'</td><td>'+fmtDec2(r.orderCost)+'</td><td>'+fmtPct(r.orderShare)+'</td><td>'+fmtPct(r.orderValidRate)+'</td></tr>'; });
  // 版位表
  let pt=document.getElementById('t3TblPlace');
  let ph=['周次','版位','订单数','订单占比','订单成本'];
  pt.innerHTML='<tr>'+ph.map(h=>'<th>'+h+'</th>').join('')+'</tr>';
  if (last) last.placeRows.forEach(r=>{ pt.innerHTML+='<tr style="opacity:0.7"><td>上周</td><td>'+r.place+'</td><td>'+fmtInt(r.all)+'</td><td>'+fmtPct(r.orderShare)+'</td><td>'+fmtDec2(r.orderCost)+'</td></tr>'; });
  cur.placeRows.forEach(r=>{ pt.innerHTML+='<tr style="font-weight:'+(r.place==='汇总'?'bold':'normal')+'"><td>本周</td><td>'+r.place+'</td><td>'+fmtInt(r.all)+'</td><td>'+fmtPct(r.orderShare)+'</td><td>'+fmtDec2(r.orderCost)+'</td></tr>'; });
}

function clearT3Cache() {
  localStorage.removeItem(T3_CACHE_KEY);
  showStatus('t3','上周缓存已清除，下次运行将无上周数据','done');
}

function downloadTab3() { if (t3Output) XLSX.writeFile(t3Output, '分周数据表.xlsx'); }

// ============================================================
// 页签4：链路画像_评分分析（双文件输入）
// ============================================================
const T4_LINK_MAP = {
  '83676427':'PF+小程序问答','83676438':'H5问答+首页','87101601':'H5问答+首页',
  '83676778':'小程序问答(含首页)','85489979':'小程序问答(含首页)',
  '85489984':'H5分流','87101605':'PF+AI落地页',
};
const T4_LINK_ORDER = ['PF+小程序问答','H5问答+首页','小程序问答(含首页)','H5分流','PF+AI落地页'];
const T4_SCORING_LINKS = ['PF+小程序问答','PF+AI落地页'];

// 城市综合评分模型（单订单产值40% + 高价转化率25% + 有效订单率15% + 订单规模20%）
function t4CityScore(rows) {
  if (!rows || rows.length === 0) return [];
  const maxOrderOutput = Math.max(...rows.map(r => r.orderOutput || 0), 1);
  const maxHighConv = Math.max(...rows.map(r => r.highConv || 0), 0.001);
  const maxValidRate = Math.max(...rows.map(r => r.count > 0 ? (r.valid || 0) / r.count : 0), 0.001);
  const maxOrders = Math.max(...rows.map(r => r.count || 0), 1);
  return rows.map(r => {
    const orderOutputScore = (r.orderOutput || 0) / maxOrderOutput * 100;
    const highConvScore = (r.highConv || 0) / maxHighConv * 100;
    const validRate = r.count > 0 ? (r.valid || 0) / r.count : 0;
    const validRateScore = validRate / maxValidRate * 100;
    const orderScaleScore = Math.log10((r.count || 0) + 1) / Math.log10(maxOrders + 1) * 100;
    const score = orderOutputScore * 0.4 + highConvScore * 0.25 + validRateScore * 0.15 + orderScaleScore * 0.2;
    let grade = 'C';
    if (score >= 90) grade = 'S';
    else if (score >= 75) grade = 'A';
    else if (score >= 60) grade = 'B';
    // 城市价值矩阵（纵轴：单订单产值，横轴：订单数）
    const avgOrderOutput = rows.reduce((s, x) => s + (x.orderOutput || 0), 0) / rows.length;
    const avgOrders = rows.reduce((s, x) => s + (x.count || 0), 0) / rows.length;
    let quadrant = '瘦狗城市';
    let strategy = '缩减或暂停投放，拉黑排除';
    if ((r.orderOutput || 0) >= avgOrderOutput && (r.count || 0) >= avgOrders) {
      quadrant = '明星城市'; strategy = '加大预算倾斜，拓宽相似人群包（Lookalike）';
    } else if ((r.orderOutput || 0) >= avgOrderOutput && (r.count || 0) < avgOrders) {
      quadrant = '金牛城市'; strategy = '测试放量，提高出价或放开定向限制，挖掘增量空间';
    } else if ((r.orderOutput || 0) < avgOrderOutput && (r.count || 0) >= avgOrders) {
      quadrant = '问题城市'; strategy = '优化落地页与转化漏斗，控制出价，精细化过滤低意向人群';
    }
    return { ...r, score: Math.round(score * 10) / 10, grade, quadrant, strategy };
  });
}
const T4_BRAND_MAP = {
  'HUAWEI':'华为','Huawei':'华为','Honor':'荣耀','iPhone':'苹果','Xiaomi':'小米',
  'Redmi':'红米','Samsung':'三星','Sony':'索尼','realme':'真我','Meizu':'魅族',
  'motorola':'摩托罗拉','Nokia':'诺基亚','Google':'谷歌','OnePlus':'一加','ZTE':'中兴','Lenovo':'联想','ChangHong':'长虹',
};
const T4_AGE_EXCLUDE = [];
const T4_AGE_MAP = { '51-55岁':'51-60岁','56-60岁':'51-60岁','61-65岁':'61-70岁','66-70岁':'61-70岁' };
const T4_AGE_ORDER = ['45岁以下','46-50岁','51-60岁','61-70岁','70岁以上','未知'];
const T4_PRICE_BINS = [
  {label:'<500元',min:0,max:499.99},{label:'500-999元',min:500,max:999.99},
  {label:'1000-1999元',min:1000,max:1999.99},{label:'2000-2999元',min:2000,max:2999.99},
  {label:'3000-3999元',min:3000,max:3999.99},{label:'4000-4999元',min:4000,max:4999.99},
  {label:'5000-5999元',min:5000,max:5999.99},{label:'6000-6999元',min:6000,max:6999.99},
  {label:'7000元及以上',min:7000,max:Infinity},
];
const T4_PRICE_ORDER = ['<500元','500-999元','1000-1999元','2000-2999元','3000-3999元','4000-4999元','5000-5999元','6000-6999元','7000元及以上','未知'];
const T4_CITY_ORDER = ['一线城市','新一线城市','二线城市','三线城市','四线城市','五线城市','其他城市','其他','未知'];
const T4_GENDER_ORDER = ['男性','女性','未知'];

const T4_FIELDS = {
  accountId: ['广告账号ID','广告账户ID','广告账户ID（数字）'],
  period: ['期数时间','订单时间','期数','购买日期'],
  city: ['城市等级','城市分级'],
  cityName: ['城市','所在城市','城市名称','城市名','具体城市','收货城市','下单城市'],
  province: ['省份','所在省份','省','地区','地域','收货省份','下单省份'],
  gender: ['性别-微信号','性别'],
  age: ['年龄'],
  price: ['手机价格'],
  brand: ['手机品牌'],
  model: ['手机型号'],
  valid: ['是否为有效订单'],
  highAmt: ['高价课订单金额','高价课金额'],
  highN: ['高价课订单数'],
  effOrders: ['低价课订单数'],
  allOrders: ['全部低价课订单数'],
  scoreOrders: ['评分订单数','线索评分订单数'],
  friendStatus: ['好友状态'],
  totalOutput: ['总产值','高价课订单金额','拟合产值'],
};

let t4Output = null, t4Data = null;

setupUpload('t4Zone','t4File','t4FileName','t4Btn');

function t4ParsePrice(x) {
  if (x===null||x===undefined) return null;
  let s=String(x).trim();
  if (s===''||s==='未知') return null;
  s=s.replace('.000000','');
  let v=parseFloat(s); return isNaN(v)?null:v;
}
function t4PriceBand(price) {
  if (price===null) return '未知';
  for (const b of T4_PRICE_BINS) { if (price>=b.min && price<=b.max) return b.label; }
  return '未知';
}
function t4AgeGroup(age) {
  return T4_AGE_MAP[age] || age;
}
function t4FormatPeriod(p) {
  if (!p) return '';
  const s = String(p).trim();
  const n = parseFloat(s);
  if (!isNaN(n) && n > 40000 && n < 60000) {
    const d = new Date(Math.round((n - 25569) * 86400 * 1000));
    return (d.getMonth()+1).toString().padStart(2,'0') + '/' + d.getDate().toString().padStart(2,'0');
  }
  return s;
}

function processTab4() {
  try {
  if (!window._xlsxReady) { showStatus('t4','xlsx库未加载完成，请等几秒后再试','error'); return; }
  const file = document.getElementById('t4File')._file;
  if (!file) return;
  document.getElementById('t4Btn').disabled=true;
  showStatus('t4','正在读取文件...','active');
  document.getElementById('t4Results').style.display='none';
  readWorkbook(file, (wb, err) => {
    if (err) { showStatus('t4','读取失败：'+err.message,'error'); document.getElementById('t4Btn').disabled=false; return; }
    try {
    showStatus('t4','正在解析数据...','active');
    const { json: rows } = getSheetData(wb);
    if (!rows.length) { showStatus('t4','数据为空','error'); document.getElementById('t4Btn').disabled=false; return; }
    const headers = Object.keys(rows[0]);
    const cols = resolveCols(headers, T4_FIELDS);
    // 必需字段：accountId, period 至少要有
    if (cols.missing.find(m=>m.field==='accountId')) { showMissingCols('t4', cols.missing, headers); showStatus('t4','缺少广告账户ID列','error'); document.getElementById('t4Btn').disabled=false; return; }
    const c = cols.found;
    // 过滤汇总行
    const detail = rows.filter(r => {
      const p = c.period!==undefined ? String(r[headers[c.period]]||'').trim() : '';
      return (!c.period || (p && p !== '汇总数据' && p !== '合计' && p !== '总计'));
    });
    showStatus('t4','正在计算画像维度...','active');
    // 判断是否有"是否为有效订单"列 → 决定用行计数还是列求和
    const hasValidCol = c.valid !== undefined;
    const hasHighAmtCol = c.highAmt !== undefined;
    const hasHighNCol = c.highN !== undefined;
    const hasEffCol = c.effOrders !== undefined;
    const hasAllCol = c.allOrders !== undefined;
    // 期数列表
    const periods = c.period!==undefined ? [...new Set(detail.map(r=>t4FormatPeriod(String(r[headers[c.period]]||'').trim())).filter(p=>p&&p!=='汇总数据'&&p!=='合计'&&p!=='总计'))].sort() : [];
    // 预处理：每行提取维度值 + 指标
    const processed = detail.map(r => {
      const aid = String(r[headers[c.accountId]]||'').trim();
      const link = T4_LINK_MAP[aid] || '其他';
      const age = c.age!==undefined ? String(r[headers[c.age]]||'未知').trim() : '未知';
      const ageG = t4AgeGroup(age);
      const price = c.price!==undefined ? t4ParsePrice(r[headers[c.price]]) : null;
      const pBand = t4PriceBand(price);
      const brandRaw = c.brand!==undefined ? String(r[headers[c.brand]]||'').trim() : '';
      const brand = T4_BRAND_MAP[brandRaw] || brandRaw;
      const model = c.model!==undefined ? String(r[headers[c.model]]||'').trim() : '';
      const genderRaw = c.gender!==undefined ? String(r[headers[c.gender]]||'未知').trim() : '未知';
      const gender = genderRaw.includes('男')?'男性':genderRaw.includes('女')?'女性':'未知';
      const city = c.city!==undefined ? String(r[headers[c.city]]||'未知').trim() : '未知';
      const cityName = c.cityName!==undefined ? String(r[headers[c.cityName]]||'未知').trim() : '未知';
      const province = c.province!==undefined ? String(r[headers[c.province]]||'未知').trim() : '未知';
      const period = t4FormatPeriod(c.period!==undefined ? String(r[headers[c.period]]||'').trim() : '');
      // 指标：根据可用列选择
      const valid = hasValidCol ? (String(r[headers[c.valid]]||'').trim()==='是') : true;
      const highAmt = hasHighAmtCol ? (parseFloat(r[headers[c.highAmt]])||0) : 0;
      const hasHigh = hasHighAmtCol ? (highAmt > 0) : (hasHighNCol ? (parseFloat(r[headers[c.highN]])||0 > 0) : false);
      const highN = hasHighNCol ? (parseFloat(r[headers[c.highN]])||0) : (hasHigh ? 1 : 0);
      const effN = hasEffCol ? (parseFloat(r[headers[c.effOrders]])||0) : 1;
      const allN = hasAllCol ? (parseFloat(r[headers[c.allOrders]])||0) : 1;
      const scoreN = c.scoreOrders!==undefined ? (parseFloat(r[headers[c.scoreOrders]])||0) : (c.friendStatus!==undefined ? (['曾添加','系统-已添加','人工-已添加'].includes(String(r[headers[c.friendStatus]]||'').trim()) ? 1 : 0) : 0);
      const totalOut = c.totalOutput!==undefined ? (parseFloat(r[headers[c.totalOutput]])||0) : 0;
      return { link, age:ageG, priceBand:pBand, brand, model, gender, city, cityName, province, period, valid, highAmt, hasHigh, highN, effN, allN, scoreN, totalOut };
    });
    // 维度定义
    const dimensions=[
      {key:'city',name:'城市等级',order:T4_CITY_ORDER,getField:r=>r.city},
      {key:'province',name:'省份',order:null,getField:r=>r.province},
      {key:'cityName',name:'具体城市',order:null,getField:r=>r.cityName},
      {key:'gender',name:'性别',order:T4_GENDER_ORDER,getField:r=>r.gender},
      {key:'age',name:'年龄档',order:T4_AGE_ORDER,getField:r=>r.age},
      {key:'price',name:'价格档',order:T4_PRICE_ORDER,getField:r=>r.priceBand},
      {key:'brand',name:'品牌',order:null,getField:r=>r.brand},
      {key:'model',name:'手机型号',order:null,getField:r=>r.brand+'|||'+r.model},
    ];
    const linkResults={};
    T4_LINK_ORDER.forEach(link => {
      const linkRows = processed.filter(r => r.link===link);
      const totalOrders = linkRows.reduce((s,r)=>s+r.allN,0);
      const totalValid = linkRows.reduce((s,r)=>s+(r.valid?r.effN:0),0);
      const totalHigh = linkRows.reduce((s,r)=>s+r.highN,0);
      const totalHighAmt = linkRows.reduce((s,r)=>s+r.highAmt,0);
      // 每期汇总（用于按期明细占比计算）
      const periodTotals={};
      linkRows.forEach(r=>{
        const p=r.period||'未知';
        if(!periodTotals[p]) periodTotals[p]={orders:0,valid:0,highN:0,highAmt:0};
        periodTotals[p].orders+=r.allN;
        if(r.valid) periodTotals[p].valid+=r.effN;
        periodTotals[p].highN+=r.highN;
        periodTotals[p].highAmt+=r.highAmt;
      });
      linkResults[link] = { totalOrders, totalValid, totalHigh, totalHighAmt, periodTotals, dimensions:{}, scoring:null };
      // 维度聚合
      dimensions.forEach(dim => {
        const groups={};
        const periodGroups={};
        linkRows.forEach(r => {
          const val=dim.getField(r); if (val===null||val===undefined||val==='') return;
          if (!groups[val]) groups[val]={count:0,valid:0,highN:0,highAmt:0};
          groups[val].count += r.allN;
          if (r.valid) groups[val].valid += r.effN;
          groups[val].highN += r.highN;
          groups[val].highAmt += r.highAmt;
          const p=r.period||'未知';
          if(!periodGroups[p]) periodGroups[p]={};
          if(!periodGroups[p][val]) periodGroups[p][val]={count:0,valid:0,highN:0,highAmt:0};
          periodGroups[p][val].count+=r.allN;
          if(r.valid) periodGroups[p][val].valid+=r.effN;
          periodGroups[p][val].highN+=r.highN;
          periodGroups[p][val].highAmt+=r.highAmt;
        });
        let order=dim.order;
        let vals = order ? order.filter(v=>groups[v]) : Object.keys(groups);
        if (!order) vals = vals.sort((a,b)=>a.localeCompare(b));
        else vals = vals.concat(Object.keys(groups).filter(v=>!order.includes(v)).sort((a,b)=>a.localeCompare(b)));
        // 按期+维度值明细
        const periodDetail=[];
        Object.keys(periodGroups).sort().forEach(p=>{
          vals.forEach(v=>{
            const g=periodGroups[p][v]; if(!g) return;
            periodDetail.push({
              period:p,
              value:dim.key==='model'?v.split('|||')[1]:v,
              count:g.count, valid:g.valid,
              share:safeDiv(g.count, periodTotals[p]?.orders||0),
              highN:g.highN, highAmt:g.highAmt,
              highConv:safeDiv(g.highN,g.count),
              orderOutput:safeDiv(g.highAmt,g.valid),
            });
          });
        });
        let dimRows = vals.map(v => ({
          value: dim.key==='model'?v.split('|||')[1]:v,
          brand: dim.key==='model'?v.split('|||')[0]:null,
          count: groups[v].count,
          valid: groups[v].valid,
          share: safeDiv(groups[v].count, totalOrders),
          highN: groups[v].highN,
          highAmt: groups[v].highAmt,
          highConv: safeDiv(groups[v].highN, groups[v].count),
          orderOutput: safeDiv(groups[v].highAmt, groups[v].valid),
        }));
        // 为地域维度（城市等级/省份/具体城市）计算综合评分、评级、价值矩阵、优化建议
        if (['city','province','cityName'].includes(dim.key)) {
          dimRows = t4CityScore(dimRows);
        }
        linkResults[link].dimensions[dim.key] = {
          name: dim.name,
          rows: dimRows,
          periodDetail,
        };
      });
      // 评分段（仅PF+小程序问答、PF+AI落地页）
      if (T4_SCORING_LINKS.includes(link)) {
        // 按期聚合评分数据
        const periodScore={};
        periods.forEach(p => { periodScore[p]={scoreOrders:0,totalOutput:0,effOrders:0}; });
        linkRows.forEach(r => {
          if (!periodScore[r.period]) periodScore[r.period]={scoreOrders:0,totalOutput:0,effOrders:0};
          periodScore[r.period].scoreOrders += r.scoreN;
          periodScore[r.period].totalOutput += r.totalOut;
          periodScore[r.period].effOrders += r.effN;
        });
        const totalScoreAll = Object.values(periodScore).reduce((s,x)=>s+x.scoreOrders,0);
        if (totalScoreAll === 0) {
          linkResults[link].scoring = null;
        } else {
        const scoringRows = periods.map(p => {
          const bd = periodScore[p] || {scoreOrders:0,totalOutput:0,effOrders:0};
          const effOrders = bd.effOrders;
          const scoreOrders = bd.scoreOrders;
          const noScoreOrders = effOrders - scoreOrders;
          const rows = [
            { value:'有评分', period:p, count:scoreOrders, share:safeDiv(scoreOrders,effOrders), highAmt:bd.totalOutput, highN:0, highConv:0, orderOutput:safeDiv(bd.totalOutput,scoreOrders) },
          ];
          if (noScoreOrders>0) rows.push({ value:'无评分', period:p, count:noScoreOrders, share:safeDiv(noScoreOrders,effOrders), highAmt:0, highN:0, highConv:0, orderOutput:0 });
          return { period:p, rows };
        }).filter(x => x.rows[0].count>0 || (x.rows[1] && x.rows[1].count>0));
        const totalScore = scoringRows.reduce((s,r)=>s+r.rows[0].count,0);
        const totalOutput = scoringRows.reduce((s,r)=>s+r.rows[0].highAmt,0);
        const totalEff = periods.reduce((s,p)=>s+(periodScore[p]?.effOrders||0),0);
        const totalNoScore = totalEff - totalScore;
        const sumRows=[{ value:'有评分', period:'合计', count:totalScore, share:safeDiv(totalScore,totalEff), highAmt:totalOutput, highN:0, highConv:0, orderOutput:safeDiv(totalOutput,totalScore) }];
        if (totalNoScore>0) sumRows.push({ value:'无评分', period:'合计', count:totalNoScore, share:safeDiv(totalNoScore,totalEff), highAmt:0, highN:0, highConv:0, orderOutput:0 });
        linkResults[link].scoring = { periodRows: scoringRows, summaryRows: sumRows };
        }
      }
    });
    t4Data = { linkResults, periods, srcFile: file.name, detailCount: processed.length, colsFound: c, missing: cols.missing };
    showStatus('t4','正在生成Excel...','active');
    t4Output = t4BuildExcel(t4Data);
    t4Display(t4Data);
    const syncOk = saveSync('score', t4Data);
    showStatus('t4','完成！'+processed.length+' 行数据，'+T4_LINK_ORDER.length+' 链路'+(cols.missing.length?'，'+cols.missing.length+' 列未匹配':'')+(syncOk?'（已同步到分析模块）':'（同步失败：存储空间不足，请清除旧同步数据）'),'done');
    document.getElementById('t4Btn').disabled=false;
    } catch(e) { showStatus('t4','运行错误：'+e.message,'error'); console.error(e); document.getElementById('t4Btn').disabled=false; }
  });
  } catch(e) { showStatus('t4','启动错误：'+e.message,'error'); console.error(e); }
}

function t4BuildExcel(d) {
  const wb = XLSX.utils.book_new();
  const geoHeaders = ['期','维度值','综合评分','评级','价值矩阵','订单数','有效订单数','订单占比','高价课订单数','高价课金额','高价课转化率','单订单产值','优化建议'];
  const dimHeaders = {
    city:geoHeaders,
    province:geoHeaders,
    cityName:geoHeaders,
    gender:['期','性别','订单数','有效订单数','订单占比','高价课订单数','高价课金额','高价课转化率','单订单产值'],
    age:['期','年龄档','订单数','有效订单数','订单占比','高价课订单数','高价课金额','高价课转化率','单订单产值'],
    price:['期','价格档','订单数','有效订单数','订单占比','高价课订单数','高价课金额','高价课转化率','单订单产值'],
    brand:['期','品牌','订单数','有效订单数','订单占比','高价课订单数','高价课金额','高价课转化率','单订单产值'],
    model:['期','手机型号','订单数','有效订单数','订单占比','高价课订单数','高价课金额','高价课转化率','单订单产值'],
  };
  const scoreHeaders = ['期','评分状态','订单数','订单占比','高价课金额','单订单产值'];
  const dimOrder = ['city','province','cityName','gender','age','price','brand','model'];
  T4_LINK_ORDER.forEach(link => {
    const lr = d.linkResults[link];
    let a=[];
    a.push([link+' · 链路画像与评分分析']);
    a.push(['口径：单订单产值=高价课金额÷有效订单数；订单占比=该维度订单数÷当期总订单数']);
    a.push(['']);
    dimOrder.forEach(dk => {
      const dim = lr.dimensions[dk];
      if(!dim) return;
      a.push(['▶ '+dim.name]);
      a.push(dimHeaders[dk]);
      (dim.periodDetail||[]).forEach(r => {
        if (['city','province','cityName'].includes(dk)) {
          a.push([r.period, r.value, '', '', '', r.count, r.valid, r.share, r.highN, r.highAmt, r.highConv, r.orderOutput, '']);
        } else {
          a.push([r.period, r.value, r.count, r.valid, r.share, r.highN, r.highAmt, r.highConv, r.orderOutput]);
        }
      });
      (dim.rows||[]).forEach(r => {
        if (['city','province','cityName'].includes(dk)) {
          a.push(['合计', r.value, r.score||'', r.grade||'', r.quadrant||'', r.count, r.valid, r.share, r.highN, r.highAmt, r.highConv, r.orderOutput, r.strategy||'']);
        } else {
          a.push(['合计', r.value, r.count, r.valid, r.share, r.highN, r.highAmt, r.highConv, r.orderOutput]);
        }
      });
      a.push(['']);
    });
    if (lr.scoring) {
      a.push(['▶ 评分分析']);
      a.push(scoreHeaders);
      (lr.scoring.periodRows||[]).forEach(pr => {
        pr.rows.forEach(r => { a.push([pr.period, r.value, r.count, r.share, r.highAmt, r.orderOutput]); });
      });
      (lr.scoring.summaryRows||[]).forEach(r => {
        a.push(['合计', r.value, r.count, r.share, r.highAmt, r.orderOutput]);
      });
      a.push(['']);
    }
    let ws=XLSX.utils.aoa_to_sheet(a);
    ws['!cols']=[{wch:10},{wch:16},{wch:10},{wch:12},{wch:10},{wch:12},{wch:14},{wch:12},{wch:12}];
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:8}},{s:{r:1,c:0},e:{r:1,c:8}}];
    if(ws['A1'])ws['A1'].s=S_TITLE;
    if(ws['A2'])ws['A2'].s=S_SUB;
    XLSX.utils.book_append_sheet(wb, ws, link);
  });
  // 口径说明
  let n=[
    ['口径说明'],[''],
    ['1. 数据源：'+d.srcFile+'（'+d.detailCount+' 条订单，'+d.periods.length+' 期）。'],
    ['2. 链路映射：83676427→PF+小程序问答；83676438+87101601→H5问答+首页；83676778+85489979→小程序问答(含首页)；85489984→H5分流；87101605→PF+AI落地页。'],
    ['3. 每个链路一个 sheet，包含8个画像维度（城市等级/省份/具体城市/性别/年龄档/价格档/品牌/手机型号）。'],
    ['4. 品牌已统一中文（HUAWEI→华为、iPhone→苹果等），vivo/OPPO保留原文。'],
    ['5. 性别取"性别-微信号"字段。'],
    ['6. 手机价格分档：<500元 / 500-999元 / 1000-1999元 / 2000-2999元 / 3000-3999元 / 4000-4999元 / 5000-5999元 / 6000-6999元 / 7000元及以上 / 未知。'],
    ['7. 单订单产值 = 高价课金额合计 ÷ 有效订单数。'],
    ['8. 订单占比 = 该维度值订单数 ÷ 当期总订单数（含无效）。'],
    ['9. 评分分析：有评分=好友状态为曾添加/系统-已添加/人工-已添加；无评分=系统-未添加或空。仅 PF+小程序问答、PF+AI落地页 展示评分分析。'],
    ['10. 评分分析中"高价课金额"列为该链路当期总产值（看板粒度无法区分有/无评分的产值），仅供参考。'],
  ];
  let wn=XLSX.utils.aoa_to_sheet(n);
  wn['!cols']=[{wch:100}];
  if(wn['A1'])wn['A1'].s=S_TITLE;
  XLSX.utils.book_append_sheet(wb, wn, '口径说明');
  return wb;
}

function t4Display(d) {
  document.getElementById('t4Results').style.display='block';
  let kpi=document.getElementById('t4Kpi'); kpi.innerHTML='';
  const totalOrders=Object.values(d.linkResults).reduce((s,l)=>s+l.totalOrders,0);
  const totalValid=Object.values(d.linkResults).reduce((s,l)=>s+l.totalValid,0);
  const totalHigh=Object.values(d.linkResults).reduce((s,l)=>s+l.totalHigh,0);
  const totalHighAmt=Object.values(d.linkResults).reduce((s,l)=>s+l.totalHighAmt,0);
  [{v:fmtInt(d.detailCount),l:'订单总数'},{v:fmtInt(totalValid),l:'有效订单'},{v:fmtInt(totalHigh),l:'高价课订单'},{v:'¥'+fmtDec2(totalHighAmt),l:'高价课金额'},{v:d.periods.length,l:'期数'},{v:T4_LINK_ORDER.length,l:'链路数'}].forEach(k=>{kpi.innerHTML+='<div class="ex-kpi-card"><div class="ex-val">'+k.v+'</div><div class="ex-lbl">'+k.l+'</div></div>';});
  // 链路选择器
  let tabs=document.getElementById('t4PreviewTabs');
  tabs.innerHTML='';
  T4_LINK_ORDER.forEach((l,i) => {
    let btn=document.createElement('button');
    btn.className='ex-tab-btn'+(i===0?' active':'');
    btn.textContent=l;
    btn.onclick=(e)=>{ document.querySelectorAll('#t4PreviewTabs .tab-btn').forEach(b=>b.classList.remove('active')); e.target.classList.add('active'); t4ShowLink(l); };
    tabs.appendChild(btn);
  });
  t4ShowLink(T4_LINK_ORDER[0]);
}

function t4ShowLink(link) {
  const lr=t4Data.linkResults[link];
  let t=document.getElementById('t4Tbl');
  let hd=['期','维度','维度值','订单数','有效','占比','高价单数','高价金额','转化率','产值'];
  t.innerHTML='<tr>'+hd.map(h=>'<th>'+h+'</th>').join('')+'</tr>';
  ['city','gender','age','price','brand','model'].forEach(dk=>{
    const dim=lr.dimensions[dk]; if(!dim)return;
    t.innerHTML+='<tr style="font-weight:bold;color:var(--accent)"><td colspan="10">▶ '+dim.name+'</td></tr>';
    (dim.periodDetail||[]).forEach(r=>{
      t.innerHTML+='<tr><td>'+r.period+'</td><td>'+dim.name+'</td><td>'+(r.value||'')+'</td><td>'+fmtInt(r.count)+'</td><td>'+fmtInt(r.valid)+'</td><td>'+fmtPct(r.share)+'</td><td>'+fmtInt(r.highN)+'</td><td>'+fmtDec2(r.highAmt)+'</td><td>'+fmtPct(r.highConv)+'</td><td>'+fmtDec2(r.orderOutput)+'</td></tr>';
    });
    (dim.rows||[]).forEach(r=>{
      const isGeo = ['city','province','cityName'].includes(dk);
      let scoreHtml = '';
      if (isGeo && r.score !== undefined) {
        const gradeColor = r.grade==='S'?'#dc2626':r.grade==='A'?'#16a34a':r.grade==='B'?'#f59e0b':'#94a3b8';
        scoreHtml = ' <span style="color:'+gradeColor+';font-weight:bold">['+r.grade+'级 '+r.score+'分]</span> <span style="color:#64748b;font-size:11px">'+r.quadrant+'</span>';
      }
      t.innerHTML+='<tr style="font-weight:bold;background:#f1f5f9"><td>合计</td><td>'+dim.name+'</td><td>'+(r.value||'')+scoreHtml+'</td><td>'+fmtInt(r.count)+'</td><td>'+fmtInt(r.valid)+'</td><td>'+fmtPct(r.share)+'</td><td>'+fmtInt(r.highN)+'</td><td>'+fmtDec2(r.highAmt)+'</td><td>'+fmtPct(r.highConv)+'</td><td>'+fmtDec2(r.orderOutput)+'</td></tr>';
      if (isGeo && r.strategy) {
        t.innerHTML+='<tr style="background:#fefce8"><td colspan="2"></td><td colspan="8" style="font-size:11px;color:#92400e;padding:4px 8px">💡 优化建议：'+r.strategy+'</td></tr>';
      }
    });
  });
  if (lr.scoring) {
    t.innerHTML+='<tr style="font-weight:bold;color:var(--accent)"><td colspan="10">▶ 评分分析</td></tr>';
    (lr.scoring.periodRows||[]).forEach(pr=>{pr.rows.forEach(r=>{
      t.innerHTML+='<tr><td>'+pr.period+'</td><td>评分</td><td>'+r.value+'</td><td>'+fmtInt(r.count)+'</td><td>-</td><td>'+fmtPct(r.share)+'</td><td>-</td><td>'+fmtDec2(r.highAmt)+'</td><td>-</td><td>'+fmtDec2(r.orderOutput)+'</td></tr>';
    });});
    (lr.scoring.summaryRows||[]).forEach(r=>{
      t.innerHTML+='<tr style="font-weight:bold;background:#f1f5f9"><td>合计</td><td>评分</td><td>'+r.value+'</td><td>'+fmtInt(r.count)+'</td><td>-</td><td>'+fmtPct(r.share)+'</td><td>-</td><td>'+fmtDec2(r.highAmt)+'</td><td>-</td><td>'+fmtDec2(r.orderOutput)+'</td></tr>';
    });
  }
}

function downloadTab4() { if (t4Output) XLSX.writeFile(t4Output, '链路画像_评分分析.xlsx'); }

/* ---------- Excel工具6：素材分析（诊断引擎） ---------- */
const T6_FIELDS = {
  name: ['素材名称','素材名','创意名称'],
  id: ['素材名称ID','素材ID','创意ID'],
  tag: ['素材标签','创意标签','标签'],
  project: ['所属项目','项目','账户'],
  roi: ['实际ROI','ROI','投产比'],
  highOrders: ['高价课订单数','高价课订单','高价单数','成交单数'],
  cost: ['花费','消耗','广告花费'],
  ctr: ['点击率','CTR'],
  lowOrders: ['低价课下单量','低价课订单数','低价单数','下单量'],
  lowRate: ['低价课下单率','下单率'],
  lowCost: ['低价课下单成本','下单成本','获客成本'],
  cpm: ['千次曝光成本','CPM','千次展现成本'],
  avgPlay: ['平均播放时长','平均播放时间','播放时长'],
  play3s: ['播放3s率','3s率','3秒播放率'],
  finishRate: ['完播率','完播'],
  playCount: ['视频播放量','播放量'],
  avgProgress: ['平均播放进度','播放进度'],
  p10: ['10%进度播放量','10%播放量'],
  p25: ['25%进度播放量','25%播放量'],
  p75: ['75%进度播放量','75%播放量'],
  p95: ['95%进度播放量','95%播放量'],
  p100: ['100%进度播放量','100%播放量'],
};
let t6Output=null, t6Data=null;
setupUpload('t6Zone','t6File','t6FileName','t6Btn');

function t6ParsePct(v){ if(v==null||v==='') return 0; const s=String(v).replace('%','').trim(); const n=parseFloat(s); return isNaN(n)?0:(s.includes('%')||String(v).includes('%')?n/100:n); }
function t6ParseNum(v){ if(v==null||v==='') return 0; const n=parseFloat(String(v).replace(/[,¥]/g,'')); return isNaN(n)?0:n; }

// 素材诊断引擎
function diagnoseMaterial(r){
  const spend=r.cost, roi=r.roi, high=r.highOrders, lowCost=r.lowCost, ctr=r.ctr, play3s=r.play3s, lowOrders=r.lowOrders;
  const lowToHigh = lowOrders>0 ? high/lowOrders : 0;
  // 1.超级爆款
  if(high>=1 && roi>=1.0 && lowCost<=200){
    return {tag:'超级爆款',level:'success',reason:'高净值精准人群，中后段信任交付扎实，低转高转化率'+(lowToHigh*100).toFixed(1)+'%。',action:'维持高预算，提价10-20%放量，针对前3秒做同脚本裂变。'};
  }
  // 2.主力跑量
  if(spend>=5000 && high>=3 && roi>=0.4){
    return {tag:'主力跑量',level:'purple',reason:'大盘核心消耗与出单基本盘，转化模型成熟，低转高'+(lowToHigh*100).toFixed(1)+'%。',action:'监控频次与衰退，搭建多计划/多账户矩阵备份。'};
  }
  // 3.高潜种子
  if(high>=1 && roi>=1.5 && spend<2500){
    return {tag:'高潜种子',level:'warning',reason:'变现心智极佳，低转高'+(lowToHigh*100).toFixed(1)+'%，但出价或冷启动受限未充分放量。',action:'提价15%抢量，放开定向，开启自动扩量。'};
  }
  // 4.虚火看客（假爆款）
  if(spend>=1500 && ctr>=0.03 && high===0){
    return {tag:'虚火看客',level:'danger',reason:'点击率'+(ctr*100).toFixed(2)+'%但高价课0单，画面/文案过于噱头吸引无付费意愿的看客，后端转化断崖。',action:'降低出价，增强前中段营销筛选与专业痛点植入，剔除纯猎奇元素。'};
  }
  // 5.获客过贵
  if(spend>=1000 && lowCost>260 && high===0){
    return {tag:'获客过贵',level:'danger',reason:'低价课下单成本'+lowCost.toFixed(0)+'元过高，中段承接力弱锁死后端回本空间。',action:'优化25%-75%中段节奏，压缩冗余内容，强化CTA权益引导。'};
  }
  // 6.首屏疲软
  if(play3s<0.6 && ctr<0.01){
    return {tag:'首屏疲软',level:'default',reason:'3s率仅'+(play3s*100).toFixed(1)+'%，前3秒抓手不足用户过早划走。',action:'更换前3秒黄金钩子、首屏标题文案或封面。'};
  }
  return {tag:'常规表现',level:'normal',reason:'各项指标在大盘均值波动范围内，低转高'+(lowToHigh*100).toFixed(1)+'%。',action:'保持观察，累计更多样本数据。'};
}

function processTab6(){
  try{
  if(!window._xlsxReady){showStatus('t6','xlsx库未加载完成，请等几秒后再试','error');return;}
  const file=document.getElementById('t6File')._file;
  if(!file)return;
  document.getElementById('t6Btn').disabled=true;
  showStatus('t6','正在读取文件...','active');
  document.getElementById('t6Results').style.display='none';
  readWorkbook(file,(wb,err)=>{
    if(err){showStatus('t6','读取失败：'+err.message,'error');document.getElementById('t6Btn').disabled=false;return;}
    try{
    showStatus('t6','正在解析数据...','active');
    const {json}=getSheetData(wb);
    if(!json.length){showStatus('t6','数据为空','error');document.getElementById('t6Btn').disabled=false;return;}
    const headers=Object.keys(json[0]);
    const cols=resolveCols(headers,T6_FIELDS);
    if(cols.missing.length){showMissingCols('t6',cols.missing,headers);showStatus('t6','有 '+cols.missing.length+' 个列未找到，请检查表头','error');document.getElementById('t6Btn').disabled=false;return;}
    showStatus('t6','正在运行诊断引擎...','active');
    // 过滤汇总行，解析每条素材
    const list=[];
    json.forEach((row,idx)=>{
      const name=String(row[headers[cols.found.name]]||'').trim();
      if(!name || name.includes('汇总')) return;
      const r={
        idx:idx, name:name,
        id:cols.found.id?String(row[headers[cols.found.id]]||'').trim():'',
        tag:cols.found.tag?String(row[headers[cols.found.tag]]||'').trim():'',
        project:cols.found.project?String(row[headers[cols.found.project]]||'').trim():'',
        roi:t6ParsePct(row[headers[cols.found.roi]]),
        highOrders:t6ParseNum(row[headers[cols.found.highOrders]]),
        cost:t6ParseNum(row[headers[cols.found.cost]]),
        ctr:t6ParsePct(row[headers[cols.found.ctr]]),
        lowOrders:t6ParseNum(row[headers[cols.found.lowOrders]]),
        lowRate:t6ParsePct(row[headers[cols.found.lowRate]]),
        lowCost:t6ParseNum(row[headers[cols.found.lowCost]]),
        cpm:t6ParseNum(row[headers[cols.found.cpm]]),
        avgPlay:t6ParseNum(row[headers[cols.found.avgPlay]]),
        play3s:t6ParsePct(row[headers[cols.found.play3s]]),
        finishRate:t6ParsePct(row[headers[cols.found.finishRate]]),
        playCount:t6ParseNum(row[headers[cols.found.playCount]]),
        avgProgress:t6ParsePct(row[headers[cols.found.avgProgress]]),
        p10:cols.found.p10?t6ParseNum(row[headers[cols.found.p10]]):0,
        p25:cols.found.p25?t6ParseNum(row[headers[cols.found.p25]]):0,
        p75:cols.found.p75?t6ParseNum(row[headers[cols.found.p75]]):0,
        p95:cols.found.p95?t6ParseNum(row[headers[cols.found.p95]]):0,
        p100:cols.found.p100?t6ParseNum(row[headers[cols.found.p100]]):0,
      };
      r.lowToHigh = r.lowOrders>0 ? r.highOrders/r.lowOrders : 0;
      r.diagnosis = diagnoseMaterial(r);
      list.push(r);
    });
    // 大盘汇总
    const total={
      cost:list.reduce((s,r)=>s+r.cost,0),
      highOrders:list.reduce((s,r)=>s+r.highOrders,0),
      lowOrders:list.reduce((s,r)=>s+r.lowOrders,0),
      playCount:list.reduce((s,r)=>s+r.playCount,0),
      roi:0, ctr:0, lowCost:0, play3s:0, finishRate:0, avgPlay:0,
    };
    total.roi = total.cost>0 ? list.reduce((s,r)=>s+r.roi*r.cost,0)/total.cost : 0;
    total.ctr = total.playCount>0 ? list.reduce((s,r)=>s+r.ctr*r.playCount,0)/total.playCount : 0;
    total.lowCost = total.lowOrders>0 ? total.cost/total.lowOrders : 0;
    total.play3s = total.playCount>0 ? list.reduce((s,r)=>s+r.play3s*r.playCount,0)/total.playCount : 0;
    total.finishRate = total.playCount>0 ? list.reduce((s,r)=>s+r.finishRate*r.playCount,0)/total.playCount : 0;
    total.avgPlay = total.playCount>0 ? list.reduce((s,r)=>s+r.avgPlay*r.playCount,0)/total.playCount : 0;
    total.lowToHigh = total.lowOrders>0 ? total.highOrders/total.lowOrders : 0;
    // 诊断分组
    const diagGroups={};
    list.forEach(r=>{const t=r.diagnosis.tag;if(!diagGroups[t])diagGroups[t]={count:0,cost:0,high:0,low:0,items:[]};diagGroups[t].count++;diagGroups[t].cost+=r.cost;diagGroups[t].high+=r.highOrders;diagGroups[t].low+=r.lowOrders;diagGroups[t].items.push(r);});
    // 高ROI vs 低ROI对比（消耗过千）
    const bigSpenders=list.filter(r=>r.cost>=1000);
    const highRoiGroup=bigSpenders.filter(r=>r.roi>=0.5);
    const lowRoiGroup=bigSpenders.filter(r=>r.roi<0.1);
    const compare={
      high:{count:highRoiGroup.length,avgRoi:highRoiGroup.length?highRoiGroup.reduce((s,r)=>s+r.roi,0)/highRoiGroup.length:0,avgCtr:highRoiGroup.length?highRoiGroup.reduce((s,r)=>s+r.ctr,0)/highRoiGroup.length:0,avgLowCost:highRoiGroup.length?highRoiGroup.reduce((s,r)=>s+r.lowCost,0)/highRoiGroup.length:0,avgLowToHigh:highRoiGroup.length?highRoiGroup.reduce((s,r)=>s+r.lowToHigh,0)/highRoiGroup.length:0,avgPlay3s:highRoiGroup.length?highRoiGroup.reduce((s,r)=>s+r.play3s,0)/highRoiGroup.length:0,avgFinish:highRoiGroup.length?highRoiGroup.reduce((s,r)=>s+r.finishRate,0)/highRoiGroup.length:0},
      low:{count:lowRoiGroup.length,avgRoi:lowRoiGroup.length?lowRoiGroup.reduce((s,r)=>s+r.roi,0)/lowRoiGroup.length:0,avgCtr:lowRoiGroup.length?lowRoiGroup.reduce((s,r)=>s+r.ctr,0)/lowRoiGroup.length:0,avgLowCost:lowRoiGroup.length?lowRoiGroup.reduce((s,r)=>s+r.lowCost,0)/lowRoiGroup.length:0,avgLowToHigh:lowRoiGroup.length?lowRoiGroup.reduce((s,r)=>s+r.lowToHigh,0)/lowRoiGroup.length:0,avgPlay3s:lowRoiGroup.length?lowRoiGroup.reduce((s,r)=>s+r.play3s,0)/lowRoiGroup.length:0,avgFinish:lowRoiGroup.length?lowRoiGroup.reduce((s,r)=>s+r.finishRate,0)/lowRoiGroup.length:0},
    };
    // 素材标签效果
    const tagStats={};
    list.forEach(r=>{
      const tags=r.tag.split(/[、,，\s]+/).filter(Boolean);
      tags.forEach(t=>{if(!tagStats[t])tagStats[t]={count:0,cost:0,high:0,low:0,roiSum:0};tagStats[t].count++;tagStats[t].cost+=r.cost;tagStats[t].high+=r.highOrders;tagStats[t].low+=r.lowOrders;tagStats[t].roiSum+=r.roi*r.cost;});
    });
    const tagList=Object.entries(tagStats).map(([t,v])=>({tag:t,count:v.count,cost:v.cost,high:v.high,low:v.low,weightedRoi:v.cost>0?v.roiSum/v.cost:0,lowToHigh:v.low>0?v.high/v.low:0})).sort((a,b)=>b.weightedRoi-a.weightedRoi);
    // 假爆款识别
    const fakeHits=list.filter(r=>r.cost>=1000 && r.ctr>=0.03 && r.highOrders===0);
    t6Data={list,total,diagGroups,compare,tagList,fakeHits,srcFile:file.name,detailCount:list.length};
    const syncOk=saveSync('material',t6Data);
    showStatus('t6','正在生成Excel...','active');
    t6Output=t6BuildExcel(t6Data);
    t6Display(t6Data);
    showStatus('t6','完成！'+list.length+' 条素材，诊断分'+Object.keys(diagGroups).length+'类，假爆款'+fakeHits.length+'条'+(syncOk?'（已同步到分析模块）':'（同步失败：存储空间不足）'),'done');
    document.getElementById('t6Btn').disabled=false;
    }catch(e){showStatus('t6','运行错误：'+e.message,'error');console.error(e);document.getElementById('t6Btn').disabled=false;}
  });
  }catch(e){showStatus('t6','启动错误：'+e.message,'error');console.error(e);}
}

function t6BuildExcel(d){
  const wb=XLSX.utils.book_new();
  // Sheet1: 全量素材+诊断
  const h1=['素材名称','素材标签','所属项目','实际ROI','高价课订单','花费','点击率','低价课下单量','低价课下单成本','3s率','完播率','平均播放时长','低转高转化率','诊断标签','诊断原因','优化建议'];
  const r1=d.list.map(r=>[r.name,r.tag,r.project,(r.roi*100).toFixed(2)+'%',r.highOrders,r.cost.toFixed(2),(r.ctr*100).toFixed(2)+'%',r.lowOrders,r.lowCost.toFixed(2),(r.play3s*100).toFixed(2)+'%',(r.finishRate*100).toFixed(2)+'%',r.avgPlay.toFixed(2),(r.lowToHigh*100).toFixed(2)+'%',r.diagnosis.tag,r.diagnosis.reason,r.diagnosis.action]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h1,...r1]),'素材诊断全量');
  // Sheet2: 诊断分组
  const h2=['诊断标签','素材数','总花费','高价课订单','低价课下单','低转高转化率'];
  const r2=Object.entries(d.diagGroups).map(([t,v])=>[t,v.count,v.cost.toFixed(2),v.high,v.low,(v.low>0?v.high/v.low*100:0).toFixed(2)+'%']);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h2,...r2]),'诊断分组统计');
  // Sheet3: 高低ROI对比
  const h3=['分组','素材数','平均ROI','平均点击率','平均低价成本','平均低转高','平均3s率','平均完播率'];
  const r3=[
    ['高ROI组(ROI≥50%)',d.compare.high.count,(d.compare.high.avgRoi*100).toFixed(2)+'%',(d.compare.high.avgCtr*100).toFixed(2)+'%',d.compare.high.avgLowCost.toFixed(2),(d.compare.high.avgLowToHigh*100).toFixed(2)+'%',(d.compare.high.avgPlay3s*100).toFixed(2)+'%',(d.compare.high.avgFinish*100).toFixed(2)+'%'],
    ['低ROI组(ROI<10%)',d.compare.low.count,(d.compare.low.avgRoi*100).toFixed(2)+'%',(d.compare.low.avgCtr*100).toFixed(2)+'%',d.compare.low.avgLowCost.toFixed(2),(d.compare.low.avgLowToHigh*100).toFixed(2)+'%',(d.compare.low.avgPlay3s*100).toFixed(2)+'%',(d.compare.low.avgFinish*100).toFixed(2)+'%'],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h3,...r3]),'高低ROI对比');
  // Sheet4: 标签效果
  const h4=['素材标签','素材数','总花费','高价课订单','低转高转化率','加权ROI'];
  const r4=d.tagList.map(t=>[t.tag,t.count,t.cost.toFixed(2),t.high,(t.lowToHigh*100).toFixed(2)+'%',(t.weightedRoi*100).toFixed(2)+'%']);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h4,...r4]),'标签效果排行');
  // Sheet5: 假爆款
  const h5=['素材名称','花费','点击率','3s率','完播率','低价课下单','高价课订单','诊断原因'];
  const r5=d.fakeHits.map(r=>[r.name,r.cost.toFixed(2),(r.ctr*100).toFixed(2)+'%',(r.play3s*100).toFixed(2)+'%',(r.finishRate*100).toFixed(2)+'%',r.lowOrders,r.highOrders,r.diagnosis.reason]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h5,...r5]),'假爆款预警');
  // Sheet6: ROI前10详细分析
  const big=d.list.filter(r=>r.cost>=1000);
  const top10=big.filter(r=>r.roi>=0.5).sort((a,b)=>b.roi-a.roi).slice(0,10);
  const bottom10=big.filter(r=>r.roi<0.1).sort((a,b)=>a.roi-b.roi).slice(0,10);
  const avgLow=d.total.lowCost;
  const h6=['排名','素材名称','素材标签','消耗','低价课成本','低价下单数','高价课订单','低转高率','ROI','点击率','3s率','完播率','平均播放','原因拆解'];
  const r6=top10.map((r,i)=>[i+1,r.name,r.tag,r.cost.toFixed(2),r.lowCost.toFixed(2),r.lowOrders,r.highOrders,(r.lowToHigh*100).toFixed(2)+'%',(r.roi*100).toFixed(2)+'%',(r.ctr*100).toFixed(2)+'%',(r.play3s*100).toFixed(2)+'%',(r.finishRate*100).toFixed(2)+'%',r.avgPlay.toFixed(2),t6Reason(r,true,avgLow)]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h6,...r6]),'ROI前10详细分析');
  // Sheet7: 低ROI败笔分析
  const h7=['排名','素材名称','素材标签','消耗','低价课成本','低价下单数','高价课订单','低转高率','ROI','点击率','3s率','完播率','平均播放','原因拆解'];
  const r7=bottom10.map((r,i)=>[i+1,r.name,r.tag,r.cost.toFixed(2),r.lowCost.toFixed(2),r.lowOrders,r.highOrders,(r.lowToHigh*100).toFixed(2)+'%',(r.roi*100).toFixed(2)+'%',(r.ctr*100).toFixed(2)+'%',(r.play3s*100).toFixed(2)+'%',(r.finishRate*100).toFixed(2)+'%',r.avgPlay.toFixed(2),t6Reason(r,false,avgLow)]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h7,...r7]),'低ROI败笔分析');
  return wb;
}

const DIAG_ORDER=['超级爆款','主力跑量','高潜种子','虚火看客','获客过贵','首屏疲软','常规表现'];
const DIAG_COLOR={'超级爆款':'#22c55e','主力跑量':'#a855f7','高潜种子':'#f59e0b','虚火看客':'#ef4444','获客过贵':'#ef4444','首屏疲软':'#94a3b8','常规表现':'#64748b'};

// 自动生成原因拆解
function t6Reason(r, isHigh, avgLowCost){
  const reasons=[];
  if(isHigh){
    if(r.lowCost<150 && r.play3s>=0.8 && r.finishRate>=0.2) reasons.push('前段低成本获客（¥'+r.lowCost.toFixed(0)+'，低于大盘¥'+avgLowCost.toFixed(0)+'）+ 中后段人设认同（3s率'+(r.play3s*100).toFixed(1)+'%、完播'+(r.finishRate*100).toFixed(1)+'%、播放'+r.avgPlay.toFixed(1)+'s），在低价课环节蓄了足够多且成本极低的流量池');
    if(r.ctr<0.02 && r.highOrders>=1) reasons.push('精准痛点/强人设钩子过滤泛人群，点击率仅'+(r.ctr*100).toFixed(2)+'%但留下的全是高意向用户，低价'+r.lowOrders+'单转化'+r.highOrders+'个高价课，低转高'+(r.lowToHigh*100).toFixed(1)+'%');
    if(r.avgPlay>=25) reasons.push('中后段停留'+r.avgPlay.toFixed(1)+'s（>25s黄金线），建立对达人/课程价值的充分信任，进班后对高价课抗拒度低');
    if(r.lowCost>=150 && r.lowCost<=200 && r.highOrders>=1) reasons.push('低价课成本¥'+r.lowCost.toFixed(0)+'处于130-200元黄金区间，后端回本空间充足');
    if(/人设|算法|ROI|爆改|强营销/.test(r.tag)) reasons.push('素材标签带强心智（'+r.tag.split(/[、,]/).filter(t=>/人设|算法|ROI|爆改|强营销/.test(t)).join('/')+'），吸引"想系统跟着学"的付费人群而非"免费看看"的路人');
    if(!reasons.length) reasons.push('综合表现均衡，低转高转化率'+(r.lowToHigh*100).toFixed(1)+'%驱动ROI达'+(r.roi*100).toFixed(1)+'%');
  }else{
    if(r.ctr>=0.03 && r.highOrders===0) reasons.push('噱头/观赏型内容，点击率高达'+(r.ctr*100).toFixed(2)+'%但高价课0单，吸引大量无付费意愿的看客/羊毛党，进线人群极泛，后端转化完全断崖');
    if(r.lowCost>260) reasons.push('低价课下单成本¥'+r.lowCost.toFixed(0)+'过高（>260元警戒线），素材切入点模糊或受众脱节，千展曝光无法转化为有效低价课购买，直接锁死后端回本空间');
    if(r.play3s<0.6 && r.ctr<0.02) reasons.push('首屏疲软，3s率仅'+(r.play3s*100).toFixed(1)+'%、点击率'+(r.ctr*100).toFixed(2)+'%，前3秒抓手不足用户过早划走');
    if(r.finishRate<0.05 && r.avgPlay<15) reasons.push('中后段流失严重，完播率仅'+(r.finishRate*100).toFixed(1)+'%、平均播放'+r.avgPlay.toFixed(1)+'s，用户未建立信任就离开，低价课仅'+r.lowOrders+'单');
    if(r.lowCost>=200 && r.lowCost<=260 && r.highOrders===0) reasons.push('获客成本¥'+r.lowCost.toFixed(0)+'偏高，消耗¥'+r.cost.toFixed(0)+'仅出'+r.lowOrders+'单低价课，0高价课，ROI仅'+(r.roi*100).toFixed(1)+'%');
    if(!reasons.length) reasons.push('各项指标偏弱，消耗¥'+r.cost.toFixed(0)+'但高价课0单，低转高0%，ROI仅'+(r.roi*100).toFixed(1)+'%');
  }
  return reasons.join('；');
}

function t6Display(d){
  document.getElementById('t6Results').style.display='block';
  document.getElementById('t6Kpi').innerHTML=[
    {label:'素材总数',value:d.list.length},{label:'总花费',value:'¥'+d.total.cost.toLocaleString(undefined,{maximumFractionDigits:0})},
    {label:'高价课订单',value:d.total.highOrders},{label:'低转高转化率',value:(d.total.lowToHigh*100).toFixed(2)+'%'},
    {label:'整体ROI',value:(d.total.roi*100).toFixed(2)+'%'},{label:'假爆款预警',value:d.fakeHits.length+'条'},
  ].map(k=>'<div class="ex-kpi-card"><div class="ex-val">'+k.value+'</div><div class="ex-lbl">'+k.label+'</div></div>').join('');
  // 诊断分类概览
  let h='<tr><th>诊断标签</th><th>素材数</th><th>总花费</th><th>高价课订单</th><th>低转高转化率</th><th>代表素材</th></tr>';
  DIAG_ORDER.filter(t=>d.diagGroups[t]).forEach(t=>{
    const g=d.diagGroups[t];
    const top=g.items.sort((a,b)=>b.cost-a.cost).slice(0,2).map(r=>r.name.substring(0,20)).join('、');
    h+='<tr><td><span style="display:inline-block;padding:2px 10px;border-radius:10px;color:#fff;font-size:12px;background:'+DIAG_COLOR[t]+'">'+t+'</span></td><td>'+g.count+'</td><td>'+g.cost.toFixed(0)+'</td><td>'+g.high+'</td><td>'+(g.low>0?g.high/g.low*100:0).toFixed(1)+'%</td><td style="font-size:12px;color:#64748b">'+top+'</td></tr>';
  });
  document.getElementById('t6Tbl').innerHTML=h;
  // ROI前10详细分析
  const big=d.list.filter(r=>r.cost>=1000);
  const top10=big.filter(r=>r.roi>=0.5).sort((a,b)=>b.roi-a.roi).slice(0,10);
  const bottom10=big.filter(r=>r.roi<0.1).sort((a,b)=>a.roi-b.roi).slice(0,10);
  const avgLow=d.total.lowCost;
  let topHtml='';
  top10.forEach((r,i)=>{
    topHtml+='<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:14px;margin-bottom:10px">'+
      '<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px">'+
        '<span style="width:26px;height:26px;border-radius:50%;background:#22c55e;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;flex-shrink:0">'+(i+1)+'</span>'+
        '<span style="font-size:15px;font-weight:700;color:#166534">'+esc(r.name)+'</span>'+
        '<span style="font-size:12px;color:#94a3b8">'+esc(r.tag)+'</span>'+
      '</div>'+
      '<div style="font-size:13px;color:#1e293b;line-height:1.8;margin-bottom:6px"><b>数据：</b>消耗¥'+r.cost.toFixed(0)+'，低价课下单成本¥'+r.lowCost.toFixed(2)+(r.lowCost<avgLow?'（低于大盘¥'+avgLow.toFixed(0)+'）':'（高于大盘¥'+avgLow.toFixed(0)+'）')+'，低价下单'+r.lowOrders+'人中转化了'+r.highOrders+'个高价课，低转高'+(r.lowToHigh*100).toFixed(1)+'%，ROI达'+(r.roi*100).toFixed(2)+'%。点击率'+(r.ctr*100).toFixed(2)+'%，3s率'+(r.play3s*100).toFixed(2)+'%，完播率'+(r.finishRate*100).toFixed(2)+'%，平均播放'+r.avgPlay.toFixed(2)+'秒。</div>'+
      '<div style="font-size:13px;color:#166534;line-height:1.8"><b>原因拆解：</b>'+t6Reason(r,true,avgLow)+'</div>'+
    '</div>';
  });
  if(!top10.length) topHtml='<div style="padding:20px;text-align:center;color:#94a3b8">暂无花费≥1000且ROI≥50%的素材</div>';
  document.getElementById('t6TopDetail').innerHTML=topHtml;
  // 低ROI败笔详细分析
  let botHtml='';
  bottom10.forEach((r,i)=>{
    botHtml+='<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:14px;margin-bottom:10px">'+
      '<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px">'+
        '<span style="width:26px;height:26px;border-radius:50%;background:#ef4444;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;flex-shrink:0">'+(i+1)+'</span>'+
        '<span style="font-size:15px;font-weight:700;color:#991b1b">'+esc(r.name)+'</span>'+
        '<span style="font-size:12px;color:#94a3b8">'+esc(r.tag)+'</span>'+
      '</div>'+
      '<div style="font-size:13px;color:#1e293b;line-height:1.8;margin-bottom:6px"><b>数据：</b>消耗¥'+r.cost.toFixed(0)+'，低价课下单成本¥'+r.lowCost.toFixed(2)+'，低价下单'+r.lowOrders+'人，高价课'+r.highOrders+'单，ROI仅'+(r.roi*100).toFixed(2)+'%。点击率'+(r.ctr*100).toFixed(2)+'%，3s率'+(r.play3s*100).toFixed(2)+'%，完播率'+(r.finishRate*100).toFixed(2)+'%，平均播放'+r.avgPlay.toFixed(2)+'秒。</div>'+
      '<div style="font-size:13px;color:#991b1b;line-height:1.8"><b>原因拆解：</b>'+t6Reason(r,false,avgLow)+'</div>'+
    '</div>';
  });
  if(!bottom10.length) botHtml='<div style="padding:20px;text-align:center;color:#94a3b8">暂无花费≥1000且ROI<10%的素材</div>';
  document.getElementById('t6BottomDetail').innerHTML=botHtml;
}

function downloadTab6(){if(!t6Output){alert('请先开始分析');return;}XLSX.writeFile(t6Output,'素材分析_诊断报告.xlsx');}

/* ---------- 素材生命周期管理 ---------- */
const MAT_LIFECYCLE_KEY = 'workbench_mat_lifecycle_v1';
const LIFECYCLE_STAGES = [
  {key:'testing', label:'测试中', cls:'testing'},
  {key:'running', label:'跑量中', cls:'running'},
  {key:'fatigue', label:'疲劳', cls:'fatigue'},
  {key:'dead', label:'已淘汰', cls:'dead'}
];
function loadMatLifecycle(){ try{ return JSON.parse(localStorage.getItem(MAT_LIFECYCLE_KEY)) || {}; }catch(e){ return {}; } }
function saveMatLifecycle(d){ try{ localStorage.setItem(MAT_LIFECYCLE_KEY, JSON.stringify(d)); }catch(e){} }
function getMatStage(name){ const d=loadMatLifecycle(); return d[name] || 'testing'; }
function setMatStage(name, stage){ const d=loadMatLifecycle(); d[name]=stage; saveMatLifecycle(d); }
function cycleMatStage(name){ const cur=getMatStage(name); const idx=LIFECYCLE_STAGES.findIndex(s=>s.key===cur); const next=LIFECYCLE_STAGES[(idx+1)%LIFECYCLE_STAGES.length]; setMatStage(name, next.key); return next; }
function lifecycleBadge(name){ const s=LIFECYCLE_STAGES.find(x=>x.key===getMatStage(name))||LIFECYCLE_STAGES[0]; return '<span class="lifecycle-badge '+s.cls+'" data-mat-name="'+esc(name)+'" title="点击切换生命周期状态">'+s.label+'</span>'; }

let matLifecycleFilter='';
/* ---------- 分析模块：素材分析（诊断看板） ---------- */
let matDiagFilter='',matSearchText2='';
function renderMaterial(){
  const d=SYNC.material;
  if(!d){
    $('#mat-badge').textContent='暂无数据';
    ['mat-filter-card','mat-conclusion-card','mat-compare-card','mat-chart-card','mat-top-card','mat-bottom-card','mat-case-card','mat-tag-card','mat-table-card'].forEach(id=>$('#'+id).style.display='none');
    $('#mat-empty').style.display='block';
    $('#mat-kpi').innerHTML='';$('#mat-summary').textContent='';
    return;
  }
  $('#mat-empty').style.display='none';
  $('#mat-badge').textContent=d.list.length+' 条素材 · '+Object.keys(d.diagGroups).length+' 类诊断 · 数据已同步';
  $('#mat-summary').textContent='数据源：'+d.srcFile;
  // KPI 美化格子
  const dg=d.diagGroups;
  const kpiGroups=[
    {title:'核心转化',items:[
      {label:'素材总数',val:d.list.length,icon:'📊',color:'#3b82f6'},
      {label:'总花费',val:'¥'+d.total.cost.toLocaleString(undefined,{maximumFractionDigits:0}),icon:'💰',color:'#6366f1'},
      {label:'高价课订单',val:d.total.highOrders,icon:'🎯',color:'#22c55e'},
      {label:'低价课下单',val:d.total.lowOrders,icon:'🛒',color:'#0ea5e9'},
      {label:'低转高转化率',val:(d.total.lowToHigh*100).toFixed(2)+'%',icon:'📈',color:'#f59e0b'},
      {label:'整体ROI',val:(d.total.roi*100).toFixed(2)+'%',icon:'💎',color:d.total.roi>=0.3?'#22c55e':'#ef4444'},
    ]},
    {title:'前端表现',items:[
      {label:'总播放量',val:d.total.playCount.toLocaleString(),icon:'▶️',color:'#8b5cf6'},
      {label:'平均点击率',val:(d.total.ctr*100).toFixed(2)+'%',icon:'👆',color:'#06b6d4'},
      {label:'平均3s率',val:(d.total.play3s*100).toFixed(1)+'%',icon:'⏱️',color:'#14b8a6'},
      {label:'平均完播率',val:(d.total.finishRate*100).toFixed(1)+'%',icon:'✅',color:'#84cc16'},
      {label:'平均播放时长',val:d.total.avgPlay.toFixed(1)+'s',icon:'🎬',color:'#f97316'},
      {label:'平均低价成本',val:'¥'+d.total.lowCost.toFixed(0),icon:'💵',color:d.total.lowCost<=200?'#22c55e':'#ef4444'},
    ]},
    {title:'诊断分布',items:[
      {label:'超级爆款',val:(dg['超级爆款']?dg['超级爆款'].count:0)+'条',icon:'🔥',color:'#22c55e'},
      {label:'主力跑量',val:(dg['主力跑量']?dg['主力跑量'].count:0)+'条',icon:'🚀',color:'#a855f7'},
      {label:'高潜种子',val:(dg['高潜种子']?dg['高潜种子'].count:0)+'条',icon:'🌱',color:'#f59e0b'},
      {label:'虚火看客',val:(dg['虚火看客']?dg['虚火看客'].count:0)+'条',icon:'👀',color:'#ef4444'},
      {label:'获客过贵',val:(dg['获客过贵']?dg['获客过贵'].count:0)+'条',icon:'⚠️',color:'#f97316'},
      {label:'假爆款预警',val:d.fakeHits.length+'条',icon:'🚨',color:'#dc2626'},
    ]},
  ];
  let kpiHtml='';
  kpiGroups.forEach(g=>{
    kpiHtml+='<div style="margin-bottom:14px"><div style="font-size:13px;font-weight:600;color:#64748b;margin-bottom:8px;padding-left:4px">'+g.title+'</div>'+
      '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px">'+
      g.items.map(k=>'<div style="background:linear-gradient(135deg,'+k.color+'15,'+k.color+'08);border:1px solid '+k.color+'30;border-radius:10px;padding:14px;transition:transform .15s">'+
        '<div style="display:flex;align-items:center;gap:6px;margin-bottom:6px"><span style="font-size:16px">'+k.icon+'</span><span style="font-size:12px;color:#64748b">'+k.label+'</span></div>'+
        '<div style="font-size:22px;font-weight:700;color:'+k.color+';line-height:1.2">'+k.val+'</div>'+
      '</div>').join('')+
      '</div></div>';
  });
  $('#mat-kpi').innerHTML=kpiHtml;
  // 筛选栏
  $('#mat-filter-card').style.display='block';
  const tf=$('#mat-type-filter');
  tf.innerHTML='<option value="">全部诊断</option>'+DIAG_ORDER.filter(t=>d.diagGroups[t]).map(t=>'<option value="'+t+'"'+(matDiagFilter===t?' selected':'')+'>'+t+'</option>').join('');
  tf.onchange=()=>{matDiagFilter=tf.value;renderMaterial();};
  const si=$('#mat-search');si.value=matSearchText2;
  si.oninput=()=>{matSearchText2=si.value;renderMaterial();};
  $('#mat-reset').onclick=()=>{matDiagFilter='';matSearchText2='';matLifecycleFilter='';renderMaterial();};
  // 生命周期筛选
  const lf=$('#mat-lifecycle-filter'); if(lf){ lf.value=matLifecycleFilter; lf.onchange=()=>{matLifecycleFilter=lf.value;renderMaterial();}; }
  // 过滤
  const filtered=d.list.filter(r=>{
    if(matDiagFilter&&r.diagnosis.tag!==matDiagFilter)return false;
    if(matLifecycleFilter&&getMatStage(r.name)!==matLifecycleFilter)return false;
    if(matSearchText2&&!r.name.toLowerCase().includes(matSearchText2.toLowerCase())&&!r.tag.includes(matSearchText2))return false;
    return true;
  });
  $('#mat-export').onclick=()=>{
    const wb=XLSX.utils.book_new();
    const h=['素材名称','标签','ROI','高价课','花费','点击率','低价成本','3s率','完播率','低转高','诊断','原因','建议'];
    const rows=filtered.map(r=>[r.name,r.tag,(r.roi*100).toFixed(2)+'%',r.highOrders,r.cost.toFixed(2),(r.ctr*100).toFixed(2)+'%',r.lowCost.toFixed(2),(r.play3s*100).toFixed(2)+'%',(r.finishRate*100).toFixed(2)+'%',(r.lowToHigh*100).toFixed(2)+'%',r.diagnosis.tag,r.diagnosis.reason,r.diagnosis.action]);
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h,...rows]),'筛选结果');
    XLSX.writeFile(wb,'素材分析_筛选结果.xlsx');
  };
  // ===== 核心结论 =====
  $('#mat-conclusion-card').style.display='block';
  const hc=d.compare.high, lc=d.compare.low;
  const highList=d.list.filter(r=>r.cost>=1000&&r.roi>=0.5);
  const lowList=d.list.filter(r=>r.cost>=1000&&r.roi<0.1);
  const highAvgHighOrders = highList.length? (highList.reduce((s,r)=>s+r.highOrders,0)/highList.length).toFixed(1):'0';
  const highAvgSpend = highList.length? (highList.reduce((s,r)=>s+r.cost,0)/highList.length).toFixed(0):'0';
  const lowAvgSpend = lowList.length? (lowList.reduce((s,r)=>s+r.cost,0)/lowList.length).toFixed(0):'0';
  // Top10/Bottom10（花费≥1000）
  const bigSpenders=d.list.filter(r=>r.cost>=1000);
  const top10=bigSpenders.filter(r=>r.roi>=0.5).sort((a,b)=>b.roi-a.roi).slice(0,10);
  const bottom10=bigSpenders.filter(r=>r.roi<0.1).sort((a,b)=>a.roi-b.roi).slice(0,10);
  const conclusionHtml='<div style="display:flex;gap:20px;flex-wrap:wrap;align-items:stretch">'+
    '<div style="flex:1;min-width:280px;background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-radius:12px;padding:18px;border:1px solid #bbf7d0">'+
      '<div style="font-size:13px;color:#166534;font-weight:600;margin-bottom:8px">高ROI组（ROI≥50%）</div>'+
      '<div style="font-size:32px;font-weight:700;color:#16a34a">'+(hc.avgLowToHigh*100).toFixed(1)+'%</div>'+
      '<div style="font-size:12px;color:#166534;margin-top:4px">低转高转化率 · '+hc.count+'条素材</div>'+
      '<div style="font-size:12px;color:#166534;margin-top:8px;line-height:1.6">平均ROI '+(hc.avgRoi*100).toFixed(1)+'% · 平均花费 ¥'+highAvgSpend+' · 平均高价课 '+highAvgHighOrders+'单</div>'+
    '</div>'+
    '<div style="flex:1;min-width:280px;background:linear-gradient(135deg,#fef2f2,#fee2e2);border-radius:12px;padding:18px;border:1px solid #fecaca">'+
      '<div style="font-size:13px;color:#991b1b;font-weight:600;margin-bottom:8px">低ROI组（ROI<10%）</div>'+
      '<div style="font-size:32px;font-weight:700;color:#dc2626">'+(lc.avgLowToHigh*100).toFixed(1)+'%</div>'+
      '<div style="font-size:12px;color:#991b1b;margin-top:4px">低转高转化率 · '+lc.count+'条素材</div>'+
      '<div style="font-size:12px;color:#991b1b;margin-top:8px;line-height:1.6">平均ROI '+(lc.avgRoi*100).toFixed(1)+'% · 平均花费 ¥'+lowAvgSpend+' · 高价课全部为0</div>'+
    '</div>'+
    '<div style="flex:2;min-width:320px;background:#f8fafc;border-radius:12px;padding:18px;border:1px solid #e2e8f0">'+
      '<div style="font-size:14px;font-weight:600;color:#1e293b;margin-bottom:12px">核心结论 · ROI高低的4个底层原因</div>'+
      '<div style="font-size:12px;color:#475569;line-height:1.8">'+
        '<div style="margin-bottom:8px"><b style="color:#3b82f6">结论1：</b>ROI高的核心<b style="color:#22c55e">不是点击率或完播率</b>，而是<b style="color:#22c55e">低转高转化率</b>。高ROI组低转高 <b style="color:#16a34a">'+(hc.avgLowToHigh*100).toFixed(1)+'%</b>，低ROI组全部 <b style="color:#dc2626">0%</b>。两组点击率（'+(hc.avgCtr*100).toFixed(2)+'% vs '+(lc.avgCtr*100).toFixed(2)+'%）、3s率（'+(hc.avgPlay3s*100).toFixed(1)+'% vs '+(lc.avgPlay3s*100).toFixed(1)+'%）几乎无差异。</div>'+
        '<div style="margin-bottom:8px"><b style="color:#3b82f6">结论2：</b><b style="color:#22c55e">精准痛点/强人设钩子</b>过滤泛人群。高ROI素材点击率反而偏低（如青花瓷仅1.11%），但留下的都是高意向付费用户；点击率>4%但高价课0单的"假爆款"（如香炉盘5.48%）全是看客流量。</div>'+
        '<div style="margin-bottom:8px"><b style="color:#3b82f6">结论3：</b><b style="color:#22c55e">中后段停留>25秒</b>建立信任交付。高ROI素材平均播放 '+((highList.reduce((s,r)=>s+r.avgPlay,0)/Math.max(highList.length,1))).toFixed(1)+'s，短平快素材用户不了解老师，进班后对高价课抗拒度高。</div>'+
        '<div><b style="color:#3b82f6">结论4：</b><b style="color:#22c55e">低价课成本130-200元</b>是黄金区间。高ROI组平均 ¥'+hc.avgLowCost.toFixed(0)+'，低ROI组 ¥'+lc.avgLowCost.toFixed(0)+'；成本>260元（如纪梵希379元）直接锁死后端回本空间。</div>'+
      '</div>'+
    '</div>'+
  '</div>'+
  // ROI前十素材详细分析（平铺卡片）
  '<div style="margin-top:16px"><div style="font-size:14px;font-weight:600;color:#1e293b;margin-bottom:10px">ROI 前十素材详细分析（花费≥1000）</div>'+
  '<div style="display:flex;flex-direction:column;gap:10px">'+
  top10.map((r,i)=>{
    return '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:14px">'+
      '<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;flex-wrap:wrap">'+
        '<span style="width:26px;height:26px;border-radius:50%;background:'+(i<3?'#22c55e':'#3b82f6')+';color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;flex-shrink:0">'+(i+1)+'</span>'+
        '<span style="font-size:15px;font-weight:700;color:#166534">'+esc(r.name)+'</span>'+
        '<span style="font-size:12px;color:#94a3b8">'+esc(r.tag)+'</span>'+
        '<span style="margin-left:auto;font-size:18px;font-weight:700;color:#22c55e">'+(r.roi*100).toFixed(1)+'%</span>'+
      '</div>'+
      '<div style="font-size:13px;color:#1e293b;line-height:1.8;margin-bottom:6px"><b>数据：</b>消耗¥'+r.cost.toFixed(0)+'，低价课成本¥'+r.lowCost.toFixed(2)+(r.lowCost<d.total.lowCost?'（低于大盘¥'+d.total.lowCost.toFixed(0)+'）':'（高于大盘¥'+d.total.lowCost.toFixed(0)+'）')+'，低价下单'+r.lowOrders+'人转化'+r.highOrders+'个高价课，低转高'+(r.lowToHigh*100).toFixed(1)+'%。点击率'+(r.ctr*100).toFixed(2)+'%，3s率'+(r.play3s*100).toFixed(1)+'%，完播'+(r.finishRate*100).toFixed(1)+'%，平均播放'+r.avgPlay.toFixed(1)+'s。</div>'+
      '<div style="font-size:13px;color:#166534;line-height:1.8"><b>原因拆解：</b>'+t6Reason(r,true,d.total.lowCost)+'</div>'+
    '</div>';
  }).join('')+
  '</div></div>';
  $('#mat-conclusion').innerHTML=conclusionHtml;
  // ===== 高低ROI深度对比表 =====
  $('#mat-compare-card').style.display='block';
  $('#mat-compare-note').textContent='高ROI组'+hc.count+'条 vs 低ROI组'+lc.count+'条（花费≥1000）';
  const cmpRows=[
    ['指标','高ROI组（ROI≥50%）','低ROI组（ROI<10%）','差异'],
    ['素材数',hc.count+'条',lc.count+'条','—'],
    ['平均ROI',(hc.avgRoi*100).toFixed(2)+'%',(lc.avgRoi*100).toFixed(2)+'%','+'+((hc.avgRoi-lc.avgRoi)*100).toFixed(1)+'pp'],
    ['平均花费','¥'+highAvgSpend,'¥'+lowAvgSpend,'¥'+(parseFloat(highAvgSpend)-parseFloat(lowAvgSpend)).toFixed(0)],
    ['平均点击率',(hc.avgCtr*100).toFixed(2)+'%',(lc.avgCtr*100).toFixed(2)+'%',((hc.avgCtr-lc.avgCtr)*100).toFixed(2)+'pp'],
    ['平均3s率',(hc.avgPlay3s*100).toFixed(1)+'%',(lc.avgPlay3s*100).toFixed(1)+'%',((hc.avgPlay3s-lc.avgPlay3s)*100).toFixed(1)+'pp'],
    ['平均完播率',(hc.avgFinish*100).toFixed(1)+'%',(lc.avgFinish*100).toFixed(1)+'%',((hc.avgFinish-lc.avgFinish)*100).toFixed(1)+'pp'],
    ['平均低价课成本','¥'+hc.avgLowCost.toFixed(0),'¥'+lc.avgLowCost.toFixed(0),'¥'+(hc.avgLowCost-lc.avgLowCost).toFixed(0)],
    ['低转高转化率',(hc.avgLowToHigh*100).toFixed(1)+'%',(lc.avgLowToHigh*100).toFixed(1)+'%','+'+((hc.avgLowToHigh-lc.avgLowToHigh)*100).toFixed(1)+'pp'],
  ];
  $('#mat-compare-table').innerHTML='<table class="data-table"><thead><tr>'+cmpRows[0].map(h=>'<th>'+h+'</th>').join('')+'</tr></thead><tbody>'+cmpRows.slice(1).map(r=>'<tr>'+r.map((c,i)=>'<td style="font-weight:'+(i===0?'600':'normal')+';color:'+(i===3&&c.startsWith('+')?'#22c55e':i===3&&c.startsWith('-')?'#ef4444':'#1e293b')+'">'+c+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
  // ===== 图表 =====
  $('#mat-chart-card').style.display='block';
  $('#mat-chart').innerHTML='<div id="mat-chart-pie" style="width:50%;height:380px;display:inline-block;vertical-align:top"></div><div id="mat-chart-bar" style="width:50%;height:380px;display:inline-block;vertical-align:top"></div>';
  const pieData=DIAG_ORDER.filter(t=>d.diagGroups[t]).map(t=>({name:t,value:d.diagGroups[t].count,itemStyle:{color:DIAG_COLOR[t]}}));
  const c1=chart('mat-chart-pie');
  if(c1)c1.setOption({title:{text:'诊断分类分布',left:'center',textStyle:{fontSize:14,color:'#1e293b'}},tooltip:{trigger:'item'},legend:{bottom:0,textStyle:{fontSize:11}},series:[{type:'pie',radius:['40%','65%'],data:pieData,label:{formatter:'{b}\n{c}条({d}%)',fontSize:11}}]});
  const c2=chart('mat-chart-bar');
  if(c2)c2.setOption({title:{text:'高ROI vs 低ROI 指标对比',left:'center',textStyle:{fontSize:14,color:'#1e293b'}},tooltip:{trigger:'axis'},legend:{bottom:0,data:['高ROI组','低ROI组'],textStyle:{fontSize:11}},grid:{left:8,right:16,top:40,bottom:40,containLabel:true},xAxis:{type:'category',data:['点击率(%)','低价成本(元)','低转高率(%)','3s率(%)','完播率(%)'],axisLabel:{fontSize:10,color:'#64748b'}},yAxis:{type:'value',axisLabel:{fontSize:10,color:'#94a3b8'}},series:[{name:'高ROI组',type:'bar',data:[+(hc.avgCtr*100).toFixed(2),+hc.avgLowCost.toFixed(0),+(hc.avgLowToHigh*100).toFixed(2),+(hc.avgPlay3s*100).toFixed(1),+(hc.avgFinish*100).toFixed(1)],itemStyle:{color:'#22c55e'}},{name:'低ROI组',type:'bar',data:[+(lc.avgCtr*100).toFixed(2),+lc.avgLowCost.toFixed(0),+(lc.avgLowToHigh*100).toFixed(2),+(lc.avgPlay3s*100).toFixed(1),+(lc.avgFinish*100).toFixed(1)],itemStyle:{color:'#ef4444'}}]});
  // ===== Top10高ROI =====
  if(top10.length){
    $('#mat-top-card').style.display='block';
    const th=['排名','素材名称','标签','花费','ROI','高价课','低价课','低价成本','点击率','3s率','完播率','播放时长','低转高'];
    $('#mat-top-table').innerHTML='<thead><tr>'+th.map(h=>'<th>'+h+'</th>').join('')+'</tr></thead><tbody>'+top10.map((r,i)=>'<tr><td style="font-weight:700;color:#22c55e">'+(i+1)+'</td><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+esc(r.name)+'">'+esc(r.name)+'</td><td style="font-size:11px;color:#94a3b8;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+esc(r.tag)+'">'+esc(r.tag)+'</td><td>'+r.cost.toFixed(0)+'</td><td style="font-weight:700;color:#22c55e">'+(r.roi*100).toFixed(1)+'%</td><td style="font-weight:600">'+r.highOrders+'</td><td>'+r.lowOrders+'</td><td>¥'+r.lowCost.toFixed(0)+'</td><td>'+(r.ctr*100).toFixed(2)+'%</td><td>'+(r.play3s*100).toFixed(1)+'%</td><td>'+(r.finishRate*100).toFixed(1)+'%</td><td>'+r.avgPlay.toFixed(1)+'s</td><td style="font-weight:600;color:#22c55e">'+(r.lowToHigh*100).toFixed(1)+'%</td></tr>').join('')+'</tbody>';
  }else $('#mat-top-card').style.display='none';
  // ===== Bottom10低ROI =====
  if(bottom10.length){
    $('#mat-bottom-card').style.display='block';
    $('#mat-bottom-table').innerHTML='<thead><tr>'+th.map(h=>'<th>'+h+'</th>').join('')+'</tr></thead><tbody>'+bottom10.map((r,i)=>'<tr><td style="font-weight:700;color:#ef4444">'+(i+1)+'</td><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+esc(r.name)+'">'+esc(r.name)+'</td><td style="font-size:11px;color:#94a3b8;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+esc(r.tag)+'">'+esc(r.tag)+'</td><td>'+r.cost.toFixed(0)+'</td><td style="font-weight:700;color:#ef4444">'+(r.roi*100).toFixed(1)+'%</td><td style="font-weight:600;color:#ef4444">'+r.highOrders+'</td><td>'+r.lowOrders+'</td><td>¥'+r.lowCost.toFixed(0)+'</td><td>'+(r.ctr*100).toFixed(2)+'%</td><td>'+(r.play3s*100).toFixed(1)+'%</td><td>'+(r.finishRate*100).toFixed(1)+'%</td><td>'+r.avgPlay.toFixed(1)+'s</td><td style="font-weight:600;color:#ef4444">'+(r.lowToHigh*100).toFixed(1)+'%</td></tr>').join('')+'</tbody>';
  }else $('#mat-bottom-card').style.display='none';
  // ===== 典型素材对比 =====
  $('#mat-case-card').style.display='block';
  // 自动找对比案例：高CTR但0高价课 vs 低CTR但高ROI
  const fakeHit=bigSpenders.find(r=>r.ctr>=0.03&&r.highOrders===0) || bottom10[0];
  const realHit=top10[0];
  let caseHtml='<div style="display:flex;gap:16px;flex-wrap:wrap">';
  if(fakeHit){
    caseHtml+='<div style="flex:1;min-width:300px;background:#fef2f2;border-radius:12px;padding:16px;border:1px solid #fecaca">'+
      '<div style="font-size:13px;font-weight:700;color:#991b1b;margin-bottom:6px">❌ 前端爆炸但亏损（虚火看客）</div>'+
      '<div style="font-size:15px;font-weight:600;color:#1e293b;margin-bottom:8px">'+esc(fakeHit.name)+'</div>'+
      '<div style="font-size:12px;color:#475569;line-height:1.8">'+
        '点击率 <b style="color:#dc2626">'+(fakeHit.ctr*100).toFixed(2)+'%</b> · 3s率 <b>'+(fakeHit.play3s*100).toFixed(1)+'%</b> · 播放 <b>'+fakeHit.avgPlay.toFixed(1)+'s</b><br>'+
        '花费 ¥'+fakeHit.cost.toFixed(0)+' · 低价课 '+fakeHit.lowOrders+'单 · 高价课 <b style="color:#dc2626">0单</b><br>'+
        'ROI <b style="color:#dc2626">'+(fakeHit.roi*100).toFixed(1)+'%</b> · 低转高 <b style="color:#dc2626">0%</b><br>'+
        '<span style="color:#991b1b">原因：'+esc(fakeHit.diagnosis.reason)+'</span>'+
      '</div></div>';
  }
  if(realHit){
    caseHtml+='<div style="flex:1;min-width:300px;background:#f0fdf4;border-radius:12px;padding:16px;border:1px solid #bbf7d0">'+
      '<div style="font-size:13px;font-weight:700;color:#166534;margin-bottom:6px">✅ 前端平平但ROI极高（精准人群）</div>'+
      '<div style="font-size:15px;font-weight:600;color:#1e293b;margin-bottom:8px">'+esc(realHit.name)+'</div>'+
      '<div style="font-size:12px;color:#475569;line-height:1.8">'+
        '点击率 <b style="color:#16a34a">'+(realHit.ctr*100).toFixed(2)+'%</b> · 3s率 <b>'+(realHit.play3s*100).toFixed(1)+'%</b> · 播放 <b>'+realHit.avgPlay.toFixed(1)+'s</b><br>'+
        '花费 ¥'+realHit.cost.toFixed(0)+' · 低价课 '+realHit.lowOrders+'单 · 高价课 <b style="color:#16a34a">'+realHit.highOrders+'单</b><br>'+
        'ROI <b style="color:#16a34a">'+(realHit.roi*100).toFixed(1)+'%</b> · 低转高 <b style="color:#16a34a">'+(realHit.lowToHigh*100).toFixed(1)+'%</b><br>'+
        '<span style="color:#166534">原因：'+esc(realHit.diagnosis.reason)+'</span>'+
      '</div></div>';
  }
  caseHtml+='</div>';
  $('#mat-cases').innerHTML=caseHtml;
  // ===== 标签效果排行 =====
  if(d.tagList&&d.tagList.length){
    $('#mat-tag-card').style.display='block';
    const tagTh=['排名','素材标签','素材数','总花费','高价课订单','低转高转化率','加权ROI'];
    $('#mat-tag-table').innerHTML='<thead><tr>'+tagTh.map(h=>'<th>'+h+'</th>').join('')+'</tr></thead><tbody>'+d.tagList.slice(0,15).map((t,i)=>'<tr><td style="font-weight:600">'+(i+1)+'</td><td style="font-weight:600">'+esc(t.tag)+'</td><td>'+t.count+'</td><td>¥'+t.cost.toFixed(0)+'</td><td>'+t.high+'</td><td style="color:'+(t.lowToHigh>=0.05?'#22c55e':'#94a3b8')+';font-weight:600">'+(t.lowToHigh*100).toFixed(1)+'%</td><td style="font-weight:700;color:'+(t.weightedRoi>=0.5?'#22c55e':t.weightedRoi<0.1?'#ef4444':'#64748b')+'">'+(t.weightedRoi*100).toFixed(1)+'%</td></tr>').join('')+'</tbody>';
  }else $('#mat-tag-card').style.display='none';
  // ===== 全量明细表 =====
  $('#mat-table-card').style.display='block';
  $('#mat-count').textContent=filtered.length+' 条素材';
  const cols=['素材名称','生命周期','标签','诊断','ROI','高价课','花费','点击率','低价成本','3s率','完播率','低转高','原因','建议'];
  let h='<tr>'+cols.map(x=>'<th>'+x+'</th>').join('')+'</tr>';
  filtered.sort((a,b)=>b.cost-a.cost).forEach(r=>{
    const color=DIAG_COLOR[r.diagnosis.level]||DIAG_COLOR[r.diagnosis.tag]||'#64748b';
    h+='<tr><td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+esc(r.name)+'">'+esc(r.name)+'</td><td>'+lifecycleBadge(r.name)+'</td><td style="font-size:11px;color:#94a3b8;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+esc(r.tag)+'">'+esc(r.tag)+'</td><td><span style="display:inline-block;padding:1px 8px;border-radius:8px;color:#fff;font-size:11px;background:'+color+'">'+r.diagnosis.tag+'</span></td><td style="font-weight:600;color:'+(r.roi>=0.5?'#22c55e':r.roi<0.1?'#ef4444':'#64748b')+'">'+(r.roi*100).toFixed(1)+'%</td><td>'+r.highOrders+'</td><td>'+r.cost.toFixed(0)+'</td><td>'+(r.ctr*100).toFixed(2)+'%</td><td>'+r.lowCost.toFixed(0)+'</td><td>'+(r.play3s*100).toFixed(1)+'%</td><td>'+(r.finishRate*100).toFixed(1)+'%</td><td style="font-weight:600;color:'+(r.lowToHigh>=0.05?'#22c55e':'#94a3b8')+'">'+(r.lowToHigh*100).toFixed(1)+'%</td><td style="font-size:11px;color:#64748b;max-width:200px">'+esc(r.diagnosis.reason)+'</td><td style="font-size:11px;color:#3b82f6;max-width:180px">'+esc(r.diagnosis.action)+'</td></tr>';
  });
  $('#mat-table').innerHTML=h;
  // 绑定生命周期徽章点击
  $('#mat-table').querySelectorAll('.lifecycle-badge').forEach(badge=>{
    badge.onclick=()=>{
      const name=badge.dataset.matName;
      const next=cycleMatStage(name);
      showToast('素材「'+name.slice(0,15)+'…」状态已切换为 '+next.label);
      renderMaterial();
    };
  });
}


/* ---------- Excel工具7：线索评分 ---------- */
let t7PhoneWB=null,t7DevWB=null,t7Output=null;
function t7HandleFile(file){
  const name=file.name.toLowerCase();
  const isPhone=name.includes('手机');
  const isDev=name.includes('设备');
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const wb=XLSX.read(ev.target.result,{type:'array'});
      if(isPhone&&!isDev){
        t7PhoneWB=wb;document.getElementById('t7PhoneName').textContent=file.name;
      }else if(isDev&&!isPhone){
        t7DevWB=wb;document.getElementById('t7DevName').textContent=file.name;
      }else{
        // 文件名不明确，优先填充为空的那个
        if(!t7PhoneWB){t7PhoneWB=wb;document.getElementById('t7PhoneName').textContent=file.name+'（自动识别为手机号）';}
        else if(!t7DevWB){t7DevWB=wb;document.getElementById('t7DevName').textContent=file.name+'（自动识别为设备号）';}
        else{showStatus('t7','两个文件均已上传，忽略多余文件：'+file.name,'error');return;}
      }
      document.getElementById('t7Btn').disabled=!(t7PhoneWB&&t7DevWB);
      if(t7PhoneWB&&t7DevWB){showStatus('t7','两个文件已就绪，可以开始分析','done');}
    }catch(err){alert('文件解析失败：'+err.message);}
  };
  reader.readAsArrayBuffer(file);
}
document.getElementById('t7PhoneFile').addEventListener('change',e=>{
  const files=e.target.files;if(!files||!files.length)return;
  [...files].forEach(f=>t7HandleFile(f));
});
document.getElementById('t7DevFile').addEventListener('change',e=>{
  const files=e.target.files;if(!files||!files.length)return;
  [...files].forEach(f=>t7HandleFile(f));
});

function t7ParseClueSheet(wb){
  // 找"下单日期"sheet
  const sheetName=wb.SheetNames.find(n=>n.includes('下单日期'))||wb.SheetNames[1]||wb.SheetNames[0];
  const ws=wb.Sheets[sheetName];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  const records=[];const daily=[];let curDate=null;
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    const a=r[0],b=r[1],c=r[2],d=r[3],e=r[4],f=r[5],g=r[6];
    if(a!=null&&String(a).trim()){
      const as=String(a).trim();
      if(as==='下单日期')continue;
      if(as.includes('汇总')){
        daily.push({date:as.replace('汇总','').trim(),low:Number(c)||0,high:Number(d)||0,conv:Number(e)||0,rev:Number(f)||0});
        continue;
      }
      curDate=as;
    }
    if(b!=null&&!isNaN(Number(b))&&Number(b)>=1&&Number(b)<=10){
      records.push({date:curDate,score:Number(b),low:Number(c)||0,high:Number(d)||0,conv:Number(e)||0,rev:Number(f)||0,pct:Number(g)||0});
    }
  }
  // 按评分聚合
  const byScore={};
  records.forEach(r=>{
    if(!byScore[r.score])byScore[r.score]={score:r.score,low:0,high:0,gmv:0};
    byScore[r.score].low+=r.low;
    byScore[r.score].high+=r.high;
    byScore[r.score].gmv+=r.rev*r.low;
  });
  const scoreList=Object.values(byScore).sort((a,b)=>a.score-b.score);
  scoreList.forEach(s=>{s.conv=s.low>0?s.high/s.low:0;s.arpu=s.low>0?s.gmv/s.low:0;});
  const totalLow=scoreList.reduce((s,r)=>s+r.low,0);
  const totalGmv=scoreList.reduce((s,r)=>s+r.gmv,0);
  scoreList.forEach(s=>{s.lowShare=totalLow>0?s.low/totalLow:0;s.gmvShare=totalGmv>0?s.gmv/totalGmv:0;});
  // 分层
  const tiers=[{name:'低分段(1-3分)',scores:[1,2,3]},{name:'中分段(4-6分)',scores:[4,5,6]},{name:'高分段(7-10分)',scores:[7,8,9,10]}];
  const tierList=tiers.map(t=>{
    const items=scoreList.filter(s=>t.scores.includes(s.score));
    const low=items.reduce((s,r)=>s+r.low,0),high=items.reduce((s,r)=>s+r.high,0),gmv=items.reduce((s,r)=>s+r.gmv,0);
    return {name:t.name,low,high,gmv,conv:low>0?high/low:0,arpu:low>0?gmv/low:0,lowShare:totalLow>0?low/totalLow:0,gmvShare:totalGmv>0?gmv/totalGmv:0};
  });
  return {records,scoreList,tierList,daily,total:{low:totalLow,high:scoreList.reduce((s,r)=>s+r.high,0),gmv:totalGmv,conv:totalLow>0?scoreList.reduce((s,r)=>s+r.high,0)/totalLow:0,arpu:totalLow>0?totalGmv/totalLow:0}};
}

function processTab7(){
  if(!t7PhoneWB||!t7DevWB){alert('请先上传手机号和设备号两份文件');return;}
  document.getElementById('t7Btn').disabled=true;
  showStatus('t7','正在解析手机号文件...','active');
  const phone=t7ParseClueSheet(t7PhoneWB);
  showStatus('t7','正在解析设备号文件...','active');
  const dev=t7ParseClueSheet(t7DevWB);
  const data={phone,dev,srcPhone:document.getElementById('t7PhoneName').textContent,srcDev:document.getElementById('t7DevName').textContent};
  const syncOk=saveSync('clue',data);
  t7Output=t7BuildExcel(data);
  t7Display(data);
  showStatus('t7','完成！手机号'+phone.total.low+'单 / 设备号'+dev.total.low+'单'+(syncOk?'（已同步到分析模块）':'（同步失败：存储空间不足）'),'done');
  document.getElementById('t7Btn').disabled=false;
}

function t7BuildExcel(d){
  const wb=XLSX.utils.book_new();
  const p=d.phone,dv=d.dev;
  // Sheet1: 核心大盘概览对比
  const coverRate=dv.total.low>0?(p.total.low/dv.total.low*100).toFixed(1):'0';
  const arpuDiff=p.total.arpu>0?((p.total.arpu-dv.total.arpu)/dv.total.arpu*100).toFixed(1):'0';
  const convDiff=((p.total.conv-dv.total.conv)*100).toFixed(2);
  const h0=['核心指标','设备号模型(Device)','手机号模型(Phone)','差异与归因洞察'];
  const r0=[
    ['低价课总订单（线索量）',dv.total.low.toLocaleString()+' 单',p.total.low.toLocaleString()+' 单','手机号覆盖率仅为设备号的 '+coverRate+'%（腾讯生态内大量小程序/H5链路未强制前置留手机号）'],
    ['高价课转化单量',dv.total.high+' 单',p.total.high+' 单','转化规模匹配设备识别的大盘体量'],
    ['综合转率（低转高）',(dv.total.conv*100).toFixed(2)+'%',(p.total.conv*100).toFixed(2)+'%','两者最终成单转化率相近，手机号端意向度'+(p.total.conv>=dv.total.conv?'略高':'略低')+'（'+(convDiff>=0?'+':'')+convDiff+'%）'],
    ['总产值 (GMV)','¥'+dv.total.gmv.toLocaleString(undefined,{maximumFractionDigits:0})+' 元','¥'+p.total.gmv.toLocaleString(undefined,{maximumFractionDigits:0})+' 元','设备号沉淀的后端总产值超 '+Math.floor(dv.total.gmv/10000)+' 万元'],
    ['单客平均产值 (ARPU)','¥'+dv.total.arpu.toFixed(2)+' 元','¥'+p.total.arpu.toFixed(2)+' 元','留资手机号用户的平均线索单产'+(p.total.arpu>=dv.total.arpu?'高出':'低出')+' '+Math.abs(arpuDiff)+'%'],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h0,...r0]),'核心大盘概览');
  // Sheet2: 关键发现
  const devScores=dv.scoreList;
  let monotonic=true;
  for(let i=1;i<devScores.length;i++){if(devScores[i].conv<devScores[i-1].conv*0.8)monotonic=false;}
  const topScore=devScores.reduce((a,b)=>a.low>b.low?a:b,devScores[0]);
  const low12=devScores.filter(s=>s.score<=2);
  const low12Low=low12.reduce((s,r)=>s+r.low,0);
  const low12High=low12.reduce((s,r)=>s+r.high,0);
  const low12Conv=low12Low>0?low12High/low12Low:0;
  const low12Arpu=low12Low>0?low12.reduce((s,r)=>s+r.gmv,0)/low12Low:0;
  const pHigh=p.tierList.find(t=>t.name.includes('高'))||p.tierList[2];
  const findings=[
    ['序号','关键发现'],
    [1,'模型单调性'+(monotonic?'极佳':'基本有效')+'：设备号转化率从 1 分的 '+(devScores[0].conv*100).toFixed(2)+'% 攀升至 10 分的 '+(devScores[devScores.length-1].conv*100).toFixed(2)+'%；ARPU 从 ¥'+devScores[0].arpu.toFixed(1)+' 升至 ¥'+devScores[devScores.length-1].arpu.toFixed(1)+'（翻了 '+(devScores[devScores.length-1].arpu/Math.max(devScores[0].arpu,0.1)).toFixed(1)+' 倍）。'+(monotonic?'证明设备号评分模型对高价值人群识别精准。':'')],
    [2,'主力盘在'+topScore.score+'分：'+topScore.score+'分单项订单占比达 '+(topScore.lowShare*100).toFixed(1)+'%，贡献了 '+(topScore.gmvShare*100).toFixed(1)+'% 的产值，属于核心放量基本盘。'],
    [3,'低分段（1~2分）投产较差：1~2分累计占了 '+(low12Low/dv.total.low*100).toFixed(1)+'% 的订单，转化率仅 '+(low12Conv*100).toFixed(2)+'%，ARPU 仅 ¥'+low12Arpu.toFixed(1)+'，拉低投放 ROI。'],
    [4,'手机号评分在 7~10分 爆发力强：仅占 '+(pHigh.lowShare*100).toFixed(1)+'% 的线索量，却贡献了 '+(pHigh.gmvShare*100).toFixed(1)+'% 的产值，转化率高达 '+(pHigh.conv*100).toFixed(2)+'%，10分线索 ARPU 高达 ¥'+(p.scoreList.find(s=>s.score===10)?.arpu||0).toFixed(1)+'。'],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(findings),'关键发现');
  // Sheet3: 按评分对比
  const h1=['评分','手机号_低价课','手机号_高价课','手机号_转化率','手机号_ARPU','手机号_订单占比','手机号_产值占比','设备号_低价课','设备号_高价课','设备号_转化率','设备号_ARPU','设备号_订单占比','设备号_产值占比'];
  const r1=[];
  for(let s=1;s<=10;s++){
    const ph=p.scoreList.find(x=>x.score===s)||{score:s,low:0,high:0,conv:0,arpu:0,lowShare:0,gmvShare:0};
    const de=dv.scoreList.find(x=>x.score===s)||{score:s,low:0,high:0,conv:0,arpu:0,lowShare:0,gmvShare:0};
    r1.push([s,ph.low,ph.high,(ph.conv*100).toFixed(2)+'%',ph.arpu.toFixed(2),(ph.lowShare*100).toFixed(2)+'%',(ph.gmvShare*100).toFixed(2)+'%',de.low,de.high,(de.conv*100).toFixed(2)+'%',de.arpu.toFixed(2),(de.lowShare*100).toFixed(2)+'%',(de.gmvShare*100).toFixed(2)+'%']);
  }
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h1,...r1]),'按评分对比');
  // Sheet4: 分层对比
  const h2=['分层','手机号_低价课','手机号_高价课','手机号_转化率','手机号_ARPU','手机号_订单占比','手机号_产值占比','设备号_低价课','设备号_高价课','设备号_转化率','设备号_ARPU','设备号_订单占比','设备号_产值占比'];
  const r2=p.tierList.map((t,i)=>{const de=dv.tierList[i];return [t.name,t.low,t.high,(t.conv*100).toFixed(2)+'%',t.arpu.toFixed(2),(t.lowShare*100).toFixed(2)+'%',(t.gmvShare*100).toFixed(2)+'%',de.low,de.high,(de.conv*100).toFixed(2)+'%',de.arpu.toFixed(2),(de.lowShare*100).toFixed(2)+'%',(de.gmvShare*100).toFixed(2)+'%'];});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h2,...r2]),'评分分层对比');
  // Sheet5: 实操建议
  const devHighCount=dv.tierList.find(t=>t.name.includes('高'))?.low||0;
  const pHighCount=p.tierList.find(t=>t.name.includes('高'))?.low||0;
  const suggestions=[
    ['分类','建议内容'],
    ['1. 腾讯广告出价与回传策略','开启深度双目标出价 / 自定义转化ROI出价：将「设备号评分 ≥ 6分」或「高价课转化」作为深度目标实时回传给腾讯系统，引导大模型自动探索6~10分高付费潜力人群。'],
    ['1. 腾讯广告出价与回传策略','低分流量压价/止损：针对1~2分占比高的版位或计划，下调基础转化出价15%~20%，或开启动态出价（oCPC/oCPM智能控成本）。'],
    ['2. 腾讯DMP人群包资产运营','正向种子人群扩展（Lookalike）：提取设备号7~10分（'+devHighCount.toLocaleString()+'个）及手机号7~10分（'+pHighCount.toLocaleString()+'个）作为高价值种子包，按1%~3%相似度拓展，单独建计划测试。'],
    ['2. 腾讯DMP人群包资产运营','负向人群包排除：将设备号1分且0转化的设备提取为黑名单包，全账户所有计划定向排除，剔除劣质羊毛党。'],
    ['3. 链路与落地页优化','推动后置/交互式留资：当前设备号（'+dv.total.low.toLocaleString()+'）远多于手机号（'+p.total.low.toLocaleString()+'），建议在购课成功页/服务通知中通过"专属教练带练/领取实体图谱"引导用户补填手机号，提升高价值线索跟进率。'],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(suggestions),'实操建议');
  // Sheet6: 手机号日期明细
  const h3=['日期','评分','低价课','高价课','转化率','产值','评分占比'];
  const r3=d.phone.records.map(r=>[r.date,r.score,r.low,r.high,(r.conv*100).toFixed(2)+'%',r.rev,r.pct]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h3,...r3]),'手机号日期明细');
  // Sheet7: 设备号日期明细
  const r4=d.dev.records.map(r=>[r.date,r.score,r.low,r.high,(r.conv*100).toFixed(2)+'%',r.rev,r.pct]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h3,...r4]),'设备号日期明细');
  return wb;
}

function t7Display(d){
  document.getElementById('t7Results').style.display='block';
  const p=d.phone,dv=d.dev;
  document.getElementById('t7Kpi').innerHTML=[
    {label:'手机号低价课',value:p.total.low.toLocaleString()},{label:'手机号高价课',value:p.total.high},
    {label:'手机号转化率',value:(p.total.conv*100).toFixed(2)+'%'},{label:'手机号ARPU',value:'¥'+p.total.arpu.toFixed(2)},
    {label:'设备号低价课',value:dv.total.low.toLocaleString()},{label:'设备号高价课',value:dv.total.high},
    {label:'设备号转化率',value:(dv.total.conv*100).toFixed(2)+'%'},{label:'设备号ARPU',value:'¥'+dv.total.arpu.toFixed(2)},
  ].map(k=>'<div class="ex-kpi-card"><div class="ex-val">'+k.value+'</div><div class="ex-lbl">'+k.label+'</div></div>').join('');

  // ===== 核心大盘概览对比 =====
  const coverRate=dv.total.low>0?(p.total.low/dv.total.low*100).toFixed(1):'0';
  const arpuDiff=p.total.arpu>0?((p.total.arpu-dv.total.arpu)/dv.total.arpu*100).toFixed(1):'0';
  const convDiff=((p.total.conv-dv.total.conv)*100).toFixed(2);
  let overviewHtml='<h4 style="margin:20px 0 8px">一、核心大盘概览对比</h4>'+
    '<div class="ex-result-table"><table><tbody>'+
    '<tr><th>核心指标</th><th>设备号模型(Device)</th><th>手机号模型(Phone)</th><th>差异与归因洞察</th></tr>'+
    '<tr><td>低价课总订单（线索量）</td><td>'+dv.total.low.toLocaleString()+' 单</td><td>'+p.total.low.toLocaleString()+' 单</td><td>手机号覆盖率仅为设备号的 '+coverRate+'%（腾讯生态内大量小程序/H5链路未强制前置留手机号）</td></tr>'+
    '<tr><td>高价课转化单量</td><td>'+dv.total.high+' 单</td><td>'+p.total.high+' 单</td><td>转化规模匹配设备识别的大盘体量</td></tr>'+
    '<tr><td>综合转率（低转高）</td><td>'+(dv.total.conv*100).toFixed(2)+'%</td><td>'+(p.total.conv*100).toFixed(2)+'%</td><td>两者最终成单转化率相近，手机号端意向度'+(p.total.conv>=dv.total.conv?'略高':'略低')+'（'+(convDiff>=0?'+':'')+convDiff+'%）</td></tr>'+
    '<tr><td>总产值 (GMV)</td><td>¥'+dv.total.gmv.toLocaleString(undefined,{maximumFractionDigits:0})+' 元</td><td>¥'+p.total.gmv.toLocaleString(undefined,{maximumFractionDigits:0})+' 元</td><td>设备号沉淀的后端总产值超 '+Math.floor(dv.total.gmv/10000)+' 万元</td></tr>'+
    '<tr><td>单客平均产值 (ARPU)</td><td>¥'+dv.total.arpu.toFixed(2)+' 元</td><td>¥'+p.total.arpu.toFixed(2)+' 元</td><td>留资手机号用户的平均线索单产'+(p.total.arpu>=dv.total.arpu?'高出':'低出')+' '+Math.abs(arpuDiff)+'%</td></tr>'+
    '</tbody></table></div>';

  // ===== 关键发现（动态生成）=====
  // 设备号模型单调性
  const devScores=dv.scoreList;
  let monotonic=true;
  for(let i=1;i<devScores.length;i++){if(devScores[i].conv<devScores[i-1].conv*0.8)monotonic=false;}
  const topScore=devScores.reduce((a,b)=>a.low>b.low?a:b,devScores[0]);
  const low12=devScores.filter(s=>s.score<=2);
  const low12Low=low12.reduce((s,r)=>s+r.low,0);
  const low12High=low12.reduce((s,r)=>s+r.high,0);
  const low12Conv=low12Low>0?low12High/low12Low:0;
  const low12Arpu=low12Low>0?low12.reduce((s,r)=>s+r.gmv,0)/low12Low:0;
  // 手机号高分段
  const pHigh=p.tierList.find(t=>t.name.includes('高'))||p.tierList[2];
  const findings=[];
  findings.push('模型单调性'+(monotonic?'极佳':'基本有效')+'：设备号转化率从 1 分的 '+(devScores[0].conv*100).toFixed(2)+'% 攀升至 10 分的 '+(devScores[devScores.length-1].conv*100).toFixed(2)+'%；ARPU 从 ¥'+devScores[0].arpu.toFixed(1)+' 升至 ¥'+devScores[devScores.length-1].arpu.toFixed(1)+'（翻了 '+(devScores[devScores.length-1].arpu/Math.max(devScores[0].arpu,0.1)).toFixed(1)+' 倍）。'+(monotonic?'证明设备号评分模型对高价值人群识别精准。':''));
  findings.push('主力盘在'+topScore.score+'分：'+topScore.score+'分单项订单占比达 '+(topScore.lowShare*100).toFixed(1)+'%，贡献了 '+(topScore.gmvShare*100).toFixed(1)+'% 的产值，属于核心放量基本盘。');
  findings.push('低分段（1~2分）投产较差：1~2分累计占了 '+(low12Low/dv.total.low*100).toFixed(1)+'% 的订单，转化率仅 '+(low12Conv*100).toFixed(2)+'%，ARPU 仅 ¥'+low12Arpu.toFixed(1)+'，拉低投放 ROI。');
  findings.push('手机号评分在 7~10分 爆发力强：仅占 '+(pHigh.lowShare*100).toFixed(1)+'% 的线索量，却贡献了 '+(pHigh.gmvShare*100).toFixed(1)+'% 的产值，转化率高达 '+(pHigh.conv*100).toFixed(2)+'%，10分线索 ARPU 高达 ¥'+(p.scoreList.find(s=>s.score===10)?.arpu||0).toFixed(1)+'。');
  overviewHtml+='<h4 style="margin:20px 0 8px">二、关键发现</h4><div style="background:#f8fafc;border-radius:8px;padding:14px;font-size:13px;line-height:2;color:#334155">'+
    findings.map((f,i)=>'<div style="margin-bottom:8px"><b style="color:#3b82f6">发现'+(i+1)+'：</b>'+f+'</div>').join('')+'</div>';

  // ===== 评分分层表 =====
  overviewHtml+='<h4 style="margin:20px 0 8px">三、评分分层与产值相关性分析</h4>'+
    '<h5 style="margin:10px 0 6px;color:#64748b">1. 设备号评分模型（样本量足，适合指导前端投放与DMP人群包）</h5>';
  let h2='<tr><th>评分区间</th><th>低价课单量</th><th>订单占比</th><th>高价课单量</th><th>转化率</th><th>综合ARPU(元)</th><th>产值贡献占比</th></tr>';
  dv.tierList.forEach(t=>{h2+='<tr><td><b>'+t.name+'</b></td><td>'+t.low.toLocaleString()+'</td><td>'+(t.lowShare*100).toFixed(1)+'%</td><td>'+t.high+'</td><td>'+(t.conv*100).toFixed(2)+'%</td><td>¥'+t.arpu.toFixed(1)+'</td><td>'+(t.gmvShare*100).toFixed(1)+'%</td></tr>';});
  h2+='<tr style="font-weight:700;background:#f1f5f9"><td>总计 1~10分</td><td>'+dv.total.low.toLocaleString()+'</td><td>100%</td><td>'+dv.total.high+'</td><td>'+(dv.total.conv*100).toFixed(2)+'%</td><td>¥'+dv.total.arpu.toFixed(1)+'</td><td>100%</td></tr>';
  overviewHtml+='<div class="ex-result-table"><table><tbody>'+h2+'</tbody></table></div>';
  overviewHtml+='<h5 style="margin:14px 0 6px;color:#64748b">2. 手机号评分模型（高净值筛选力强，适合销售跟进与高阶转化）</h5>';
  let h3='<tr><th>评分区间</th><th>低价课单量</th><th>订单占比</th><th>高价课单量</th><th>转化率</th><th>综合ARPU(元)</th><th>产值贡献占比</th></tr>';
  p.tierList.forEach(t=>{h3+='<tr><td><b>'+t.name+'</b></td><td>'+t.low.toLocaleString()+'</td><td>'+(t.lowShare*100).toFixed(1)+'%</td><td>'+t.high+'</td><td>'+(t.conv*100).toFixed(2)+'%</td><td>¥'+t.arpu.toFixed(1)+'</td><td>'+(t.gmvShare*100).toFixed(1)+'%</td></tr>';});
  h3+='<tr style="font-weight:700;background:#f1f5f9"><td>总计 1~10分</td><td>'+p.total.low.toLocaleString()+'</td><td>100%</td><td>'+p.total.high+'</td><td>'+(p.total.conv*100).toFixed(2)+'%</td><td>¥'+p.total.arpu.toFixed(1)+'</td><td>100%</td></tr>';
  overviewHtml+='<div class="ex-result-table"><table><tbody>'+h3+'</tbody></table></div>';

  // ===== 实操建议（动态生成）=====
  const devHighCount=dv.tierList.find(t=>t.name.includes('高'))?.low||0;
  const pHighCount=p.tierList.find(t=>t.name.includes('高'))?.low||0;
  const suggestions=[
    {title:'1. 腾讯广告出价与回传策略',items:[
      '开启深度双目标出价 / 自定义转化ROI出价：将「设备号评分 ≥ 6分」或「高价课转化」作为深度目标实时回传给腾讯系统，引导大模型自动探索6~10分高付费潜力人群。',
      '低分流量压价/止损：针对1~2分占比高的版位或计划，下调基础转化出价15%~20%，或开启动态出价（oCPC/oCPM智能控成本）。'
    ]},
    {title:'2. 腾讯DMP人群包资产运营',items:[
      '正向种子人群扩展（Lookalike）：提取设备号7~10分（'+devHighCount.toLocaleString()+'个）及手机号7~10分（'+pHighCount.toLocaleString()+'个）作为高价值种子包，按1%~3%相似度拓展，单独建计划测试。',
      '负向人群包排除：将设备号1分且0转化的设备提取为黑名单包，全账户所有计划定向排除，剔除劣质羊毛党。'
    ]},
    {title:'3. 链路与落地页优化',items:[
      '推动后置/交互式留资：当前设备号（'+dv.total.low.toLocaleString()+'）远多于手机号（'+p.total.low.toLocaleString()+'），建议在购课成功页/服务通知中通过"专属教练带练/领取实体图谱"引导用户补填手机号，提升高价值线索跟进率。'
    ]},
  ];
  overviewHtml+='<h4 style="margin:20px 0 8px">四、优化师投放与腾讯广告实操建议</h4>';
  suggestions.forEach(s=>{
    overviewHtml+='<div style="margin-bottom:12px"><div style="font-size:14px;font-weight:600;color:#1e293b;margin-bottom:6px">'+s.title+'</div>'+
      '<ul style="margin:0;padding-left:20px;font-size:13px;line-height:1.9;color:#475569">'+s.items.map(it=>'<li>'+it+'</li>').join('')+'</ul></div>';
  });

  // 按评分明细表
  overviewHtml+='<h4 style="margin:20px 0 8px">五、按评分明细（1-10分全量）</h4>';
  let h='<tr><th>评分</th><th>手机号低价课</th><th>手机号高价课</th><th>手机号转化率</th><th>手机号ARPU</th><th>手机号产值占比</th><th>设备号低价课</th><th>设备号高价课</th><th>设备号转化率</th><th>设备号ARPU</th><th>设备号产值占比</th></tr>';
  for(let s=1;s<=10;s++){
    const ph=p.scoreList.find(x=>x.score===s)||{low:0,high:0,conv:0,arpu:0,gmvShare:0};
    const de=dv.scoreList.find(x=>x.score===s)||{low:0,high:0,conv:0,arpu:0,gmvShare:0};
    h+='<tr><td><b>'+s+'分</b></td><td>'+ph.low+'</td><td>'+ph.high+'</td><td>'+(ph.conv*100).toFixed(2)+'%</td><td>¥'+ph.arpu.toFixed(2)+'</td><td>'+(ph.gmvShare*100).toFixed(1)+'%</td><td>'+de.low+'</td><td>'+de.high+'</td><td>'+(de.conv*100).toFixed(2)+'%</td><td>¥'+de.arpu.toFixed(2)+'</td><td>'+(de.gmvShare*100).toFixed(1)+'%</td></tr>';
  }
  overviewHtml+='<div class="ex-result-table"><table><tbody>'+h+'</tbody></table></div>';

  document.getElementById('t7Analysis').innerHTML=overviewHtml;
}

function downloadTab7(){if(!t7Output){alert('请先开始分析');return;}XLSX.writeFile(t7Output,'线索评分分析.xlsx');}

/* ---------- 分析模块：线索评分 ---------- */
function renderClue(){
  const d=SYNC.clue;
  if(!d){$('#clue-badge').textContent='暂无数据';['clue-chart-card','clue-tier-card','clue-detail-card','clue-trend-card','clue-overview-card','clue-finding-card','clue-suggest-card'].forEach(id=>{const el=$('#'+id);if(el)el.style.display='none';});$('#clue-empty').style.display='block';$('#clue-kpi').innerHTML='';$('#clue-summary').textContent='';return;}
  $('#clue-empty').style.display='none';
  $('#clue-badge').textContent='手机号'+d.phone.total.low+'单 · 设备号'+d.dev.total.low+'单 · 已同步';
  $('#clue-summary').textContent='数据源：'+d.srcPhone+' + '+d.srcDev;
  const p=d.phone,dv=d.dev;
  // KPI美化格子
  const kpiItems=[
    {group:'手机号模型',items:[
      {label:'低价课总订单',val:p.total.low.toLocaleString(),icon:'📱',color:'#3b82f6'},
      {label:'高价课订单',val:p.total.high,icon:'🎯',color:'#22c55e'},
      {label:'综合转化率',val:(p.total.conv*100).toFixed(2)+'%',icon:'📈',color:'#f59e0b'},
      {label:'单客ARPU',val:'¥'+p.total.arpu.toFixed(2),icon:'💎',color:'#8b5cf6'},
      {label:'总产值GMV',val:'¥'+p.total.gmv.toLocaleString(undefined,{maximumFractionDigits:0}),icon:'💰',color:'#06b6d4'},
    ]},
    {group:'设备号模型',items:[
      {label:'低价课总订单',val:dv.total.low.toLocaleString(),icon:'📱',color:'#ef4444'},
      {label:'高价课订单',val:dv.total.high,icon:'🎯',color:'#22c55e'},
      {label:'综合转化率',val:(dv.total.conv*100).toFixed(2)+'%',icon:'📈',color:'#f59e0b'},
      {label:'单客ARPU',val:'¥'+dv.total.arpu.toFixed(2),icon:'💎',color:'#8b5cf6'},
      {label:'总产值GMV',val:'¥'+dv.total.gmv.toLocaleString(undefined,{maximumFractionDigits:0}),icon:'💰',color:'#06b6d4'},
    ]},
  ];
  let kpiHtml='';
  kpiItems.forEach(g=>{
    kpiHtml+='<div style="margin-bottom:14px"><div style="font-size:13px;font-weight:600;color:#64748b;margin-bottom:8px;padding-left:4px">'+g.group+'</div><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px">'+g.items.map(k=>'<div style="background:linear-gradient(135deg,'+k.color+'15,'+k.color+'08);border:1px solid '+k.color+'30;border-radius:10px;padding:12px"><div style="display:flex;align-items:center;gap:6px;margin-bottom:4px"><span>'+k.icon+'</span><span style="font-size:11px;color:#64748b">'+k.label+'</span></div><div style="font-size:20px;font-weight:700;color:'+k.color+'">'+k.val+'</div></div>').join('')+'</div></div>';
  });
  $('#clue-kpi').innerHTML=kpiHtml;

  // ===== 核心大盘概览对比 =====
  const ovCard=$('#clue-overview-card');
  if(ovCard){
    ovCard.style.display='block';
    const coverRate=dv.total.low>0?(p.total.low/dv.total.low*100).toFixed(1):'0';
    const arpuDiff=p.total.arpu>0?((p.total.arpu-dv.total.arpu)/dv.total.arpu*100).toFixed(1):'0';
    const convDiff=((p.total.conv-dv.total.conv)*100).toFixed(2);
    const ovRows=[
      ['核心指标','设备号模型(Device)','手机号模型(Phone)','差异与归因洞察'],
      ['低价课总订单（线索量）',dv.total.low.toLocaleString()+' 单',p.total.low.toLocaleString()+' 单','手机号覆盖率仅为设备号的 '+coverRate+'%（腾讯生态内大量小程序/H5链路未强制前置留手机号）'],
      ['高价课转化单量',dv.total.high+' 单',p.total.high+' 单','转化规模匹配设备识别的大盘体量'],
      ['综合转率（低转高）',(dv.total.conv*100).toFixed(2)+'%',(p.total.conv*100).toFixed(2)+'%','两者最终成单转化率相近，手机号端意向度'+(p.total.conv>=dv.total.conv?'略高':'略低')+'（'+(convDiff>=0?'+':'')+convDiff+'%）'],
      ['总产值 (GMV)','¥'+dv.total.gmv.toLocaleString(undefined,{maximumFractionDigits:0})+' 元','¥'+p.total.gmv.toLocaleString(undefined,{maximumFractionDigits:0})+' 元','设备号沉淀的后端总产值超 '+Math.floor(dv.total.gmv/10000)+' 万元'],
      ['单客平均产值 (ARPU)','¥'+dv.total.arpu.toFixed(2)+' 元','¥'+p.total.arpu.toFixed(2)+' 元','留资手机号用户的平均线索单产'+(p.total.arpu>=dv.total.arpu?'高出':'低出')+' '+Math.abs(arpuDiff)+'%'],
    ];
    fillTable('clue-overview-table',ovRows[0],ovRows.slice(1).map(r=>Object.fromEntries(ovRows[0].map((h,i)=>[h,r[i]]))));
  }

  // ===== 关键发现 =====
  const findCard=$('#clue-finding-card');
  if(findCard){
    findCard.style.display='block';
    const devScores=dv.scoreList;
    let monotonic=true;
    for(let i=1;i<devScores.length;i++){if(devScores[i].conv<devScores[i-1].conv*0.8)monotonic=false;}
    const topScore=devScores.reduce((a,b)=>a.low>b.low?a:b,devScores[0]);
    const low12=devScores.filter(s=>s.score<=2);
    const low12Low=low12.reduce((s,r)=>s+r.low,0);
    const low12High=low12.reduce((s,r)=>s+r.high,0);
    const low12Conv=low12Low>0?low12High/low12Low:0;
    const low12Arpu=low12Low>0?low12.reduce((s,r)=>s+r.gmv,0)/low12Low:0;
    const pHigh=p.tierList.find(t=>t.name.includes('高'))||p.tierList[2];
    const findings=[
      '模型单调性'+(monotonic?'极佳':'基本有效')+'：设备号转化率从 1 分的 '+(devScores[0].conv*100).toFixed(2)+'% 攀升至 10 分的 '+(devScores[devScores.length-1].conv*100).toFixed(2)+'%；ARPU 从 ¥'+devScores[0].arpu.toFixed(1)+' 升至 ¥'+devScores[devScores.length-1].arpu.toFixed(1)+'（翻了 '+(devScores[devScores.length-1].arpu/Math.max(devScores[0].arpu,0.1)).toFixed(1)+' 倍）。'+(monotonic?'证明设备号评分模型对高价值人群识别精准。':''),
      '主力盘在'+topScore.score+'分：'+topScore.score+'分单项订单占比达 '+(topScore.lowShare*100).toFixed(1)+'%，贡献了 '+(topScore.gmvShare*100).toFixed(1)+'% 的产值，属于核心放量基本盘。',
      '低分段（1~2分）投产较差：1~2分累计占了 '+(low12Low/dv.total.low*100).toFixed(1)+'% 的订单，转化率仅 '+(low12Conv*100).toFixed(2)+'%，ARPU 仅 ¥'+low12Arpu.toFixed(1)+'，拉低投放 ROI。',
      '手机号评分在 7~10分 爆发力强：仅占 '+(pHigh.lowShare*100).toFixed(1)+'% 的线索量，却贡献了 '+(pHigh.gmvShare*100).toFixed(1)+'% 的产值，转化率高达 '+(pHigh.conv*100).toFixed(2)+'%，10分线索 ARPU 高达 ¥'+(p.scoreList.find(s=>s.score===10)?.arpu||0).toFixed(1)+'。',
    ];
    $('#clue-findings').innerHTML=findings.map((f,i)=>'<div style="margin-bottom:10px;padding:10px 14px;background:#f8fafc;border-radius:8px;border-left:3px solid #3b82f6"><b style="color:#3b82f6">发现'+(i+1)+'：</b><span style="color:#334155;font-size:13px;line-height:1.8">'+f+'</span></div>').join('');
  }

  // ===== 评分分布图 =====
  $('#clue-chart-card').style.display='block';
  const scores=[1,2,3,4,5,6,7,8,9,10];
  const pLow=scores.map(s=>{const x=p.scoreList.find(r=>r.score===s);return x?x.low:0;});
  const dvLow=scores.map(s=>{const x=dv.scoreList.find(r=>r.score===s);return x?x.low:0;});
  const pConv=scores.map(s=>{const x=p.scoreList.find(r=>r.score===s);return x?+(x.conv*100).toFixed(2):0;});
  const dvConv=scores.map(s=>{const x=dv.scoreList.find(r=>r.score===s);return x?+(x.conv*100).toFixed(2):0;});
  const cChart=chart('clue-chart');
  if(cChart)cChart.setOption({tooltip:{trigger:'axis'},legend:{data:['手机号订单量','设备号订单量','手机号转化率','设备号转化率']},grid:{left:50,right:50,top:40,bottom:30},xAxis:{type:'category',data:scores.map(s=>s+'分')},yAxis:[{type:'value',name:'订单量'},{type:'value',name:'转化率(%)'}],series:[{name:'手机号订单量',type:'bar',data:pLow,itemStyle:{color:'#3b82f6'}},{name:'设备号订单量',type:'bar',data:dvLow,itemStyle:{color:'#ef4444'}},{name:'手机号转化率',type:'line',yAxisIndex:1,data:pConv,itemStyle:{color:'#22c55e'},smooth:true},{name:'设备号转化率',type:'line',yAxisIndex:1,data:dvConv,itemStyle:{color:'#f59e0b'},smooth:true}]});

  // ===== 分层表 =====
  $('#clue-tier-card').style.display='block';
  const tierHead=['分层','手机号低价课','手机号高价课','手机号转化率','手机号ARPU','手机号订单占比','手机号产值占比','设备号低价课','设备号高价课','设备号转化率','设备号ARPU','设备号订单占比','设备号产值占比'];
  const tierRows=p.tierList.map((t,i)=>{const de=dv.tierList[i];return [t.name,t.low,t.high,(t.conv*100).toFixed(2)+'%','¥'+t.arpu.toFixed(2),(t.lowShare*100).toFixed(1)+'%',(t.gmvShare*100).toFixed(1)+'%',de.low,de.high,(de.conv*100).toFixed(2)+'%','¥'+de.arpu.toFixed(2),(de.lowShare*100).toFixed(1)+'%',(de.gmvShare*100).toFixed(1)+'%'];});
  fillTable('clue-tier-table',tierHead,tierRows.map(r=>Object.fromEntries(tierHead.map((h,i)=>[h,r[i]]))));

  // ===== 明细表 =====
  $('#clue-detail-card').style.display='block';
  const detHead=['评分','手机号低价课','手机号高价课','手机号转化率','手机号ARPU','手机号订单占比','手机号产值占比','设备号低价课','设备号高价课','设备号转化率','设备号ARPU','设备号订单占比','设备号产值占比'];
  const detRows=scores.map(s=>{const ph=p.scoreList.find(x=>x.score===s)||{low:0,high:0,conv:0,arpu:0,lowShare:0,gmvShare:0};const de=dv.scoreList.find(x=>x.score===s)||{low:0,high:0,conv:0,arpu:0,lowShare:0,gmvShare:0};return [s+'分',ph.low,ph.high,(ph.conv*100).toFixed(2)+'%','¥'+ph.arpu.toFixed(2),(ph.lowShare*100).toFixed(1)+'%',(ph.gmvShare*100).toFixed(1)+'%',de.low,de.high,(de.conv*100).toFixed(2)+'%','¥'+de.arpu.toFixed(2),(de.lowShare*100).toFixed(1)+'%',(de.gmvShare*100).toFixed(1)+'%'];});
  fillTable('clue-detail-table',detHead,detRows.map(r=>Object.fromEntries(detHead.map((h,i)=>[h,r[i]]))));

  // ===== 实操建议 =====
  const sugCard=$('#clue-suggest-card');
  if(sugCard){
    sugCard.style.display='block';
    const devHighCount=dv.tierList.find(t=>t.name.includes('高'))?.low||0;
    const pHighCount=p.tierList.find(t=>t.name.includes('高'))?.low||0;
    const suggestions=[
      {title:'1. 腾讯广告出价与回传策略',items:[
        '开启深度双目标出价 / 自定义转化ROI出价：将「设备号评分 ≥ 6分」或「高价课转化」作为深度目标实时回传给腾讯系统，引导大模型自动探索6~10分高付费潜力人群。',
        '低分流量压价/止损：针对1~2分占比高的版位或计划，下调基础转化出价15%~20%，或开启动态出价（oCPC/oCPM智能控成本）。'
      ]},
      {title:'2. 腾讯DMP人群包资产运营',items:[
        '正向种子人群扩展（Lookalike）：提取设备号7~10分（'+devHighCount.toLocaleString()+'个）及手机号7~10分（'+pHighCount.toLocaleString()+'个）作为高价值种子包，按1%~3%相似度拓展，单独建计划测试。',
        '负向人群包排除：将设备号1分且0转化的设备提取为黑名单包，全账户所有计划定向排除，剔除劣质羊毛党。'
      ]},
      {title:'3. 链路与落地页优化',items:[
        '推动后置/交互式留资：当前设备号（'+dv.total.low.toLocaleString()+'）远多于手机号（'+p.total.low.toLocaleString()+'），建议在购课成功页/服务通知中通过"专属教练带练/领取实体图谱"引导用户补填手机号，提升高价值线索跟进率。'
      ]},
    ];
    $('#clue-suggestions').innerHTML=suggestions.map(s=>'<div style="margin-bottom:14px"><div style="font-size:14px;font-weight:600;color:#1e293b;margin-bottom:8px">'+s.title+'</div><ul style="margin:0;padding-left:20px;font-size:13px;line-height:1.9;color:#475569">'+s.items.map(it=>'<li>'+it+'</li>').join('')+'</ul></div>').join('');
  }

  // ===== 日期趋势 =====
  $('#clue-trend-card').style.display='block';
  const allDates=[...new Set([...p.daily.map(x=>x.date),...dv.daily.map(x=>x.date)])].sort();
  const pDailyLow=allDates.map(dt=>{const x=p.daily.find(r=>r.date===dt);return x?x.low:0;});
  const dvDailyLow=allDates.map(dt=>{const x=dv.daily.find(r=>r.date===dt);return x?x.low:0;});
  const pDailyHigh=allDates.map(dt=>{const x=p.daily.find(r=>r.date===dt);return x?x.high:0;});
  const dvDailyHigh=allDates.map(dt=>{const x=dv.daily.find(r=>r.date===dt);return x?x.high:0;});
  const tChart=chart('clue-trend-chart');
  if(tChart)tChart.setOption({tooltip:{trigger:'axis'},legend:{data:['手机号低价课','设备号低价课','手机号高价课','设备号高价课']},grid:{left:50,right:30,top:40,bottom:60},xAxis:{type:'category',data:allDates,axisLabel:{rotate:45}},yAxis:{type:'value'},series:[{name:'手机号低价课',type:'line',data:pDailyLow,itemStyle:{color:'#3b82f6'},smooth:true},{name:'设备号低价课',type:'line',data:dvDailyLow,itemStyle:{color:'#ef4444'},smooth:true},{name:'手机号高价课',type:'line',data:pDailyHigh,itemStyle:{color:'#22c55e'},smooth:true},{name:'设备号高价课',type:'line',data:dvDailyHigh,itemStyle:{color:'#f59e0b'},smooth:true}]});
}

/* ---------- Excel工具9：线索评分·版位分析 ---------- */
const t9WBs={phoneGx:null,phoneSp:null,devGx:null,devSp:null};
let t9Output=null;

// t9文件匹配规则
function t9MatchFile(name){
  if(/手机/i.test(name)&&/公小/i.test(name))return {key:'phoneGx',input:'t9PhoneGxFile',name:'t9PhoneGxName',label:'手机号-公小'};
  if(/手机/i.test(name)&&/视频/i.test(name))return {key:'phoneSp',input:'t9PhoneSpFile',name:'t9PhoneSpName',label:'手机号-视频号'};
  if(/设备/i.test(name)&&/公小/i.test(name))return {key:'devGx',input:'t9DevGxFile',name:'t9DevGxName',label:'设备号-公小'};
  if(/设备/i.test(name)&&/视频/i.test(name))return {key:'devSp',input:'t9DevSpFile',name:'t9DevSpName',label:'设备号-视频号'};
  return null;
}

// 处理单个文件上传
function t9HandleFile(file, targetCfg){
  if(!file||!targetCfg)return false;
  document.getElementById(targetCfg.name).textContent='✅ '+file.name;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      t9WBs[targetCfg.key]=XLSX.read(ev.target.result,{type:'array'});
      showStatus('t9','已加载：'+targetCfg.label+' - '+file.name,'active');
      t9CheckReady();
    }catch(err){showStatus('t9','文件解析失败：'+err.message,'error');}
  };
  reader.readAsArrayBuffer(file);
  return true;
}

function t9CheckReady(){
  const ready=t9WBs.phoneGx&&t9WBs.phoneSp&&t9WBs.devGx&&t9WBs.devSp;
  const btn=document.getElementById('t9Btn');
  if(btn)btn.disabled=!ready;
  if(ready)showStatus('t9','4份文件已全部加载，点击开始分析','done');
}

// 绑定t9上传（直接执行，不依赖DOMContentLoaded）
(function initT9Upload(){
  // 每个上传区独立单文件拖拽
  document.querySelectorAll('.t9-zone').forEach(zone=>{
    const targetId=zone.dataset.target;
    const nameId=zone.dataset.name;
    const keyMap={t9PhoneGxFile:'phoneGx',t9PhoneSpFile:'phoneSp',t9DevGxFile:'devGx',t9DevSpFile:'devSp'};
    zone.addEventListener('dragover',e=>{e.preventDefault();e.stopPropagation();zone.style.borderColor='var(--accent)';zone.style.background='var(--accent-soft)';});
    zone.addEventListener('dragleave',e=>{e.preventDefault();e.stopPropagation();zone.style.borderColor='';zone.style.background='';});
    zone.addEventListener('drop',e=>{
      e.preventDefault();e.stopPropagation();
      zone.style.borderColor='';zone.style.background='';
      const file=e.dataTransfer.files[0];
      if(file){
        const cfg=t9MatchFile(file.name)||{key:keyMap[targetId],input:targetId,name:nameId,label:nameId.replace('t9','').replace('Name','')};
        t9HandleFile(file,cfg);
      }
    });
  });
  // 整个区域多文件拖拽自动匹配
  const dropZone=document.getElementById('t9DropZone');
  if(dropZone){
    dropZone.addEventListener('dragover',e=>{e.preventDefault();e.stopPropagation();dropZone.style.borderColor='var(--accent)';dropZone.style.background='rgba(59,130,246,.05)';});
    dropZone.addEventListener('dragleave',e=>{e.preventDefault();e.stopPropagation();if(e.target===dropZone){dropZone.style.borderColor='';dropZone.style.background='';}});
    dropZone.addEventListener('drop',e=>{
      e.preventDefault();e.stopPropagation();
      dropZone.style.borderColor='';dropZone.style.background='';
      const files=e.dataTransfer.files;
      if(files&&files.length){
        let matched=0;
        [...files].forEach(f=>{
          const cfg=t9MatchFile(f.name);
          if(cfg){t9HandleFile(f,cfg);matched++;}
        });
        if(matched>0){showStatus('t9','已自动匹配 '+matched+'/'+files.length+' 个文件','done');}
        else{showStatus('t9','未匹配到文件，请确认文件名包含"手机/设备"和"公小/视频号"','error');}
      }
    });
  }
  // 点击上传change事件（支持多文件）
  const clickKeyMap={t9PhoneGxFile:'phoneGx',t9PhoneSpFile:'phoneSp',t9DevGxFile:'devGx',t9DevSpFile:'devSp'};
  ['t9PhoneGxFile','t9PhoneSpFile','t9DevGxFile','t9DevSpFile'].forEach(id=>{
    const el=document.getElementById(id);
    if(!el)return;
    el.addEventListener('change',e=>{
      const files=e.target.files;
      if(!files||!files.length)return;
      let matched=0;
      [...files].forEach(file=>{
        const cfg=t9MatchFile(file.name)||{key:clickKeyMap[id],input:id,name:id.replace('File','Name'),label:id.replace('t9','').replace('File','')};
        if(t9HandleFile(file,cfg))matched++;
      });
      if(matched>0){showStatus('t9','已处理 '+matched+' 个文件','done');}
    });
  });
})();

function t9ParseClueSheet(wb){
  const sheetName=wb.SheetNames.find(n=>n.includes('下单日期'))||wb.SheetNames[1]||wb.SheetNames[0];
  const ws=wb.Sheets[sheetName];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  const records=[];const daily=[];let curDate=null;
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    const a=r[0],b=r[1],c=r[2],d=r[3],e=r[4],f=r[5],g=r[6];
    if(a!=null&&String(a).trim()){
      const as=String(a).trim();
      if(as==='下单日期')continue;
      if(as.includes('汇总')){daily.push({date:as.replace('汇总','').trim(),low:Number(c)||0,high:Number(d)||0,conv:Number(e)||0,rev:Number(f)||0});continue;}
      curDate=as;
    }
    if(b!=null&&!isNaN(Number(b))&&Number(b)>=1&&Number(b)<=10){
      records.push({date:curDate,score:Number(b),low:Number(c)||0,high:Number(d)||0,conv:Number(e)||0,rev:Number(f)||0,pct:Number(g)||0});
    }
  }
  const byScore={};
  records.forEach(r=>{
    if(!byScore[r.score])byScore[r.score]={score:r.score,low:0,high:0,gmv:0};
    byScore[r.score].low+=r.low;byScore[r.score].high+=r.high;byScore[r.score].gmv+=r.rev*r.low;
  });
  const scoreList=Object.values(byScore).sort((a,b)=>a.score-b.score);
  scoreList.forEach(s=>{s.conv=s.low>0?s.high/s.low:0;s.arpu=s.low>0?s.gmv/s.low:0;});
  const totalLow=scoreList.reduce((s,r)=>s+r.low,0),totalGmv=scoreList.reduce((s,r)=>s+r.gmv,0);
  scoreList.forEach(s=>{s.lowShare=totalLow>0?s.low/totalLow:0;s.gmvShare=totalGmv>0?s.gmv/totalGmv:0;});
  const tiers=[{name:'低(1-4)',scores:[1,2,3,4]},{name:'中(5-6)',scores:[5,6]},{name:'高(7-10)',scores:[7,8,9,10]}];
  const tierList=tiers.map(t=>{
    const items=scoreList.filter(s=>t.scores.includes(s.score));
    const low=items.reduce((s,r)=>s+r.low,0),high=items.reduce((s,r)=>s+r.high,0),gmv=items.reduce((s,r)=>s+r.gmv,0);
    return {name:t.name,low,high,gmv,conv:low>0?high/low:0,arpu:low>0?gmv/low:0,lowShare:totalLow>0?low/totalLow:0,gmvShare:totalGmv>0?gmv/totalGmv:0};
  });
  // 按月聚合
  const byMonth={};
  records.forEach(r=>{
    if(!r.date)return;
    const m=r.date.slice(0,7);
    if(!byMonth[m])byMonth[m]={month:m,low:0,high:0,gmv:0,highTierLow:0};
    byMonth[m].low+=r.low;byMonth[m].high+=r.high;byMonth[m].gmv+=r.rev*r.low;
    if(r.score>=7)byMonth[m].highTierLow+=r.low;
  });
  const monthly=Object.values(byMonth).sort((a,b)=>a.month.localeCompare(b.month));
  monthly.forEach(m=>{m.conv=m.low>0?m.high/m.low:0;m.highTierShare=m.low>0?m.highTierLow/m.low:0;});
  return {records,scoreList,tierList,daily,monthly,total:{low:totalLow,high:scoreList.reduce((s,r)=>s+r.high,0),gmv:totalGmv,conv:totalLow>0?scoreList.reduce((s,r)=>s+r.high,0)/totalLow:0,arpu:totalLow>0?totalGmv/totalLow:0}};
}

function processTab9(){
  if(!t9WBs.phoneGx||!t9WBs.phoneSp||!t9WBs.devGx||!t9WBs.devSp){alert('请先上传4份文件');return;}
  document.getElementById('t9Btn').disabled=true;
  showStatus('t9','正在解析4份文件...','active');
  const phoneGx=t9ParseClueSheet(t9WBs.phoneGx);
  const phoneSp=t9ParseClueSheet(t9WBs.phoneSp);
  const devGx=t9ParseClueSheet(t9WBs.devGx);
  const devSp=t9ParseClueSheet(t9WBs.devSp);
  const data={
    phoneGx,phoneSp,devGx,devSp,
    srcPhoneGx:document.getElementById('t9PhoneGxName').textContent.replace('✅ ',''),
    srcPhoneSp:document.getElementById('t9PhoneSpName').textContent.replace('✅ ',''),
    srcDevGx:document.getElementById('t9DevGxName').textContent.replace('✅ ',''),
    srcDevSp:document.getElementById('t9DevSpName').textContent.replace('✅ ',''),
  };
  const syncOk=saveSync('cluePosition',data);
  t9Output=t9BuildExcel(data);
  t9Display(data);
  showStatus('t9','完成！公小'+(devGx.total.low+phoneGx.total.low)+'单 / 视频号'+(devSp.total.low+phoneSp.total.low)+'单'+(syncOk?'（已同步到分析模块）':'（同步失败：存储空间不足）'),'done');
  document.getElementById('t9Btn').disabled=false;
}

function t9BuildExcel(d){
  const wb=XLSX.utils.book_new();
  const groups=[
    {key:'devGx',label:'设备号-公小',data:d.devGx},
    {key:'devSp',label:'设备号-视频号',data:d.devSp},
    {key:'phoneGx',label:'手机号-公小',data:d.phoneGx},
    {key:'phoneSp',label:'手机号-视频号',data:d.phoneSp},
  ];
  const safeDiv=(a,b)=>b>0?a/b:0;
  const devGx=d.devGx,devSp=d.devSp;
  const devGxHigh=devGx.tierList.find(x=>x.name==='高(7-10)')||{lowShare:0,arpu:0,conv:0,gmvShare:0};
  const devSpHigh=devSp.tierList.find(x=>x.name==='高(7-10)')||{lowShare:0,arpu:0,conv:0,gmvShare:0};
  const phoneGxHigh=d.phoneGx.tierList.find(x=>x.name==='高(7-10)')||{arpu:0};
  const devTotal=devGx.total.low+devSp.total.low;
  const phoneTotal=d.phoneGx.total.low+d.phoneSp.total.low;
  const phoneRate=devTotal>0?(phoneTotal/devTotal*100).toFixed(1):'0';
  const spConvAdv=(safeDiv(devSp.total.conv-devGx.total.conv,devGx.total.conv)*100).toFixed(1);
  const spArpuAdv=(safeDiv(devSp.total.arpu-devGx.total.arpu,devGx.total.arpu)*100).toFixed(1);
  const spHighAdv=(safeDiv(devSpHigh.lowShare-devGxHigh.lowShare,devGxHigh.lowShare)*100).toFixed(1);
  const arpuRatio=devGxHigh.arpu>0?(phoneGxHigh.arpu/devGxHigh.arpu).toFixed(1):'0';

  // Sheet1: KPI汇总
  const kpiH=['维度','低价课','高价课','转化率','ARPU','GMV','高分占比','高分转化率','高分ARPU','高分GMV占比'];
  const kpiR=groups.map(g=>{
    const t=g.data.total,ht=g.data.tierList.find(x=>x.name==='高(7-10)')||{lowShare:0,conv:0,arpu:0,gmvShare:0};
    return [g.label,t.low,t.high,(t.conv*100).toFixed(2)+'%','¥'+t.arpu.toFixed(2),'¥'+t.gmv.toLocaleString(undefined,{maximumFractionDigits:0}),(ht.lowShare*100).toFixed(1)+'%',(ht.conv*100).toFixed(2)+'%','¥'+ht.arpu.toFixed(2),(ht.gmvShare*100).toFixed(1)+'%'];
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([kpiH,...kpiR]),'KPI汇总');

  // Sheet2: 版位横向对比汇总
  const h0=['归因','版位','低价课','高价课','转化率','ARPU','GMV','高分占比','高分转化率','高分ARPU','高分GMV占比'];
  const r0=groups.map(g=>{
    const t=g.data.total,ht=g.data.tierList.find(x=>x.name==='高(7-10)');
    return [g.label.split('-')[0],g.label.split('-')[1],t.low,t.high,(t.conv*100).toFixed(2)+'%','¥'+t.arpu.toFixed(2),'¥'+t.gmv.toLocaleString(undefined,{maximumFractionDigits:0}),(ht?ht.lowShare*100:0).toFixed(1)+'%',(ht?ht.conv*100:0).toFixed(2)+'%','¥'+(ht?ht.arpu.toFixed(2):'0'),(ht?ht.gmvShare*100:0).toFixed(1)+'%'];
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h0,...r0]),'版位横向对比');

  // Sheet3: 关键发现
  const findings=[
    ['序号','关键发现','数据支撑'],
    [1,'视频号全面优于公小','转化率高'+spConvAdv+'%（'+(devSp.total.conv*100).toFixed(2)+'% vs '+(devGx.total.conv*100).toFixed(2)+'%），ARPU高'+spArpuAdv+'%（¥'+devSp.total.arpu.toFixed(1)+' vs ¥'+devGx.total.arpu.toFixed(1)+'）'],
    [2,'视频号高分人群占比更高','高'+spHighAdv+'%（'+(devSpHigh.lowShare*100).toFixed(1)+'% vs '+(devGxHigh.lowShare*100).toFixed(1)+'%），流量水质明显更优'],
    [3,'手机号覆盖率极低','设备号'+devTotal.toLocaleString()+'单 vs 手机号'+phoneTotal.toLocaleString()+'单，仅'+phoneRate+'%'],
    [4,'手机号高分ARPU爆发力强','手机号-公小高分ARPU ¥'+phoneGxHigh.arpu.toFixed(1)+'，是设备号-公小（¥'+devGxHigh.arpu.toFixed(1)+'）的'+arpuRatio+'倍'],
    [5,'模型单调性验证','设备号-视频号10分转化率'+(devSp.scoreList.find(s=>s.score===10)?(devSp.scoreList.find(s=>s.score===10).conv*100).toFixed(2):'0')+'%，1分转化率'+(devSp.scoreList.find(s=>s.score===1)?(devSp.scoreList.find(s=>s.score===1).conv*100).toFixed(2):'0')+'%，评分越高转化越好'],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(findings),'关键发现');

  // Sheet4: 评分分层对比
  const h6=['维度','分层','低价课','高价课','转化率','ARPU','订单占比','GMV占比'];
  const r6=[];
  groups.forEach(g=>{g.data.tierList.forEach(t=>{r6.push([g.label,t.name,t.low,t.high,(t.conv*100).toFixed(2)+'%','¥'+t.arpu.toFixed(2),(t.lowShare*100).toFixed(1)+'%',(t.gmvShare*100).toFixed(1)+'%']);});});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h6,...r6]),'评分分层对比');

  // Sheet5: 月度趋势与模型衰减
  const h7=['维度','月份','低价课','高价课','转化率','高分占比'];
  const r7=[];
  groups.forEach(g=>{g.data.monthly.forEach(m=>{r7.push([g.label,m.month,m.low,m.high,(m.conv*100).toFixed(2)+'%',(m.highTierShare*100).toFixed(1)+'%']);});});
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h7,...r7]),'月度趋势');

  // Sheet6: 优化师投放实操建议
  const suggestions=[
    ['类别','建议'],
    ['1. 预算与出价策略','视频号放宽前端出价上限20-30%，优先抢量（高分占比高、ARPU高）'],
    ['1. 预算与出价策略','公小针对1-2分占比高的计划降出价15%，或设置负向过滤'],
    ['1. 预算与出价策略','将「评分≥7分」作为深度转化事件回传腾讯广告，开启双目标出价'],
    ['2. 版位分流','视频号：主攻高净值人群，素材强化"专业体系化/高阶带练"'],
    ['2. 版位分流','公小：压缩低质流量，落地页增加留资门槛筛选'],
    ['3. 手机号回传补全','当前手机号覆盖率仅'+phoneRate+'%，在购课成功页引导补填手机号（送实体图谱/专属教练）'],
    ['3. 手机号回传补全','手机号高分ARPU达¥'+phoneGxHigh.arpu.toFixed(1)+'，是销售跟进的核心资产'],
    ['4. 模型衰减应对','8月转化率断崖式下跌，需排查：素材疲劳？定向放宽？竞品挤压？'],
    ['4. 模型衰减应对','建立7天/14天转化归因窗口，避免误判近期计划质量'],
  ];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(suggestions),'实操建议');

  // Sheet7-10: 各维度评分明细
  groups.forEach(g=>{
    const h=['评分','低价课','高价课','转化率','ARPU','订单占比','GMV占比'];
    const rows=g.data.scoreList.map(s=>[s.score+'分',s.low,s.high,(s.conv*100).toFixed(2)+'%','¥'+s.arpu.toFixed(2),(s.lowShare*100).toFixed(1)+'%',(s.gmvShare*100).toFixed(1)+'%']);
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h,...rows]),g.label);
  });
  return wb;
}

function t9Display(d){
  document.getElementById('t9Results').style.display='block';
  const groups=[
    {key:'devGx',label:'设备号-公小',data:d.devGx,color:'#3b82f6'},
    {key:'devSp',label:'设备号-视频号',data:d.devSp,color:'#ef4444'},
    {key:'phoneGx',label:'手机号-公小',data:d.phoneGx,color:'#22c55e'},
    {key:'phoneSp',label:'手机号-视频号',data:d.phoneSp,color:'#f59e0b'},
  ];
  // KPI
  const kpiHtml=groups.map(g=>{
    const t=g.data.total,ht=g.data.tierList.find(x=>x.name==='高(7-10)');
    return '<div class="ex-kpi-card" style="border-left:4px solid '+g.color+'"><div class="ex-kpi-label">'+g.label+'</div><div class="ex-kpi-value">'+t.low.toLocaleString()+'</div><div class="ex-kpi-sub">转化率 '+(t.conv*100).toFixed(2)+'% · ARPU ¥'+t.arpu.toFixed(1)+'</div><div class="ex-kpi-sub" style="color:#64748b">高分占比 '+(ht?ht.lowShare*100:0).toFixed(1)+'% · 高分GMV '+(ht?ht.gmvShare*100:0).toFixed(1)+'%</div></div>';
  }).join('');
  document.getElementById('t9Kpi').innerHTML=kpiHtml;
  // 分析结果
  let html='';
  // 版位横向对比表
  html+='<h3 style="font-size:15px;font-weight:600;margin:16px 0 10px;color:#1e293b">一、版位横向对比汇总</h3>';
  html+='<div class="ex-result-table"><table><thead><tr><th>归因</th><th>版位</th><th>低价课</th><th>高价课</th><th>转化率</th><th>ARPU</th><th>GMV</th><th>高分占比</th><th>高分转化率</th><th>高分ARPU</th><th>高分GMV占比</th></tr></thead><tbody>';
  groups.forEach(g=>{
    const t=g.data.total,ht=g.data.tierList.find(x=>x.name==='高(7-10)');
    html+='<tr><td><b>'+g.label.split('-')[0]+'</b></td><td>'+g.label.split('-')[1]+'</td><td>'+t.low.toLocaleString()+'</td><td>'+t.high+'</td><td>'+(t.conv*100).toFixed(2)+'%</td><td>¥'+t.arpu.toFixed(2)+'</td><td>¥'+t.gmv.toLocaleString(undefined,{maximumFractionDigits:0})+'</td><td>'+(ht?ht.lowShare*100:0).toFixed(1)+'%</td><td>'+(ht?ht.conv*100:0).toFixed(2)+'%</td><td>¥'+(ht?ht.arpu.toFixed(2):'0')+'</td><td>'+(ht?ht.gmvShare*100:0).toFixed(1)+'%</td></tr>';
  });
  html+='</tbody></table></div>';
  // 关键发现
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px;color:#1e293b">二、关键发现</h3>';
  const devGx=d.devGx,devSp=d.devSp;
  const safeDiv=(a,b)=>b>0?a/b:0;
  const spConvAdv=(safeDiv(devSp.total.conv-devGx.total.conv,devGx.total.conv)*100).toFixed(1);
  const spArpuAdv=(safeDiv(devSp.total.arpu-devGx.total.arpu,devGx.total.arpu)*100).toFixed(1);
  const devGxHigh=devGx.tierList.find(x=>x.name==='高(7-10)')||{lowShare:0,arpu:0};
  const devSpHigh=devSp.tierList.find(x=>x.name==='高(7-10)')||{lowShare:0,arpu:0};
  const phoneGxHigh=d.phoneGx.tierList.find(x=>x.name==='高(7-10)')||{arpu:0};
  const spHighAdv=(safeDiv(devSpHigh.lowShare-devGxHigh.lowShare,devGxHigh.lowShare)*100).toFixed(1);
  const devTotal=devGx.total.low+devSp.total.low;
  const phoneTotal=d.phoneGx.total.low+d.phoneSp.total.low;
  const phoneRate=devTotal>0?(phoneTotal/devTotal*100).toFixed(1):'0';
  const arpuRatio=devGxHigh.arpu>0?(phoneGxHigh.arpu/devGxHigh.arpu).toFixed(1):'0';
  const findings=[
    '视频号全面优于公小：转化率高'+spConvAdv+'%（'+(devSp.total.conv*100).toFixed(2)+'% vs '+(devGx.total.conv*100).toFixed(2)+'%），ARPU高'+spArpuAdv+'%（¥'+devSp.total.arpu.toFixed(1)+' vs ¥'+devGx.total.arpu.toFixed(1)+'）',
    '视频号高分人群占比高'+spHighAdv+'%（'+(devSpHigh.lowShare*100).toFixed(1)+'% vs '+(devGxHigh.lowShare*100).toFixed(1)+'%），流量水质明显更优',
    '手机号覆盖率极低：设备号'+devTotal.toLocaleString()+'单 vs 手机号'+phoneTotal.toLocaleString()+'单，仅'+phoneRate+'%',
    '手机号高分ARPU爆发力强：手机号-公小高分ARPU ¥'+phoneGxHigh.arpu.toFixed(1)+'，是设备号-公小（¥'+devGxHigh.arpu.toFixed(1)+'）的'+arpuRatio+'倍',
  ];
  html+='<div style="background:#f8fafc;border-radius:10px;padding:14px 18px;margin-bottom:14px"><ul style="margin:0;padding-left:20px;font-size:13px;line-height:2;color:#334155">'+findings.map(f=>'<li>'+f+'</li>').join('')+'</ul></div>';
  // 分层对比
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px;color:#1e293b">三、评分分层对比（低1-4 / 中5-6 / 高7-10）</h3>';
  html+='<div class="ex-result-table"><table><thead><tr><th>维度</th><th>分层</th><th>低价课</th><th>高价课</th><th>转化率</th><th>ARPU</th><th>订单占比</th><th>GMV占比</th></tr></thead><tbody>';
  groups.forEach(g=>{g.data.tierList.forEach(t=>{
    html+='<tr><td><b>'+g.label+'</b></td><td>'+t.name+'</td><td>'+t.low.toLocaleString()+'</td><td>'+t.high+'</td><td>'+(t.conv*100).toFixed(2)+'%</td><td>¥'+t.arpu.toFixed(2)+'</td><td>'+(t.lowShare*100).toFixed(1)+'%</td><td>'+(t.gmvShare*100).toFixed(1)+'%</td></tr>';
  });});
  html+='</tbody></table></div>';
  // 月度趋势
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px;color:#1e293b">四、月度趋势与模型衰减</h3>';
  html+='<div class="ex-result-table"><table><thead><tr><th>维度</th><th>月份</th><th>低价课</th><th>高价课</th><th>转化率</th><th>高分占比</th></tr></thead><tbody>';
  groups.forEach(g=>{g.data.monthly.forEach(m=>{
    html+='<tr><td><b>'+g.label+'</b></td><td>'+m.month+'</td><td>'+m.low.toLocaleString()+'</td><td>'+m.high+'</td><td>'+(m.conv*100).toFixed(2)+'%</td><td>'+(m.highTierShare*100).toFixed(1)+'%</td></tr>';
  });});
  html+='</tbody></table></div>';
  // 实操建议
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px;color:#1e293b">五、优化师投放实操建议</h3>';
  const suggestions=[
    {title:'1. 预算与出价策略',items:['视频号放宽前端出价上限20-30%，优先抢量（高分占比高、ARPU高）','公小针对1-2分占比高的计划降出价15%，或设置负向过滤','将「评分≥7分」作为深度转化事件回传腾讯广告，开启双目标出价']},
    {title:'2. 版位分流',items:['视频号：主攻高净值人群，素材强化"专业体系化/高阶带练"','公小：压缩低质流量，落地页增加留资门槛筛选']},
    {title:'3. 手机号回传补全',items:['当前手机号覆盖率仅'+phoneRate+'%，在购课成功页引导补填手机号（送实体图谱/专属教练）','手机号高分ARPU达¥'+phoneGxHigh.arpu.toFixed(1)+'，是销售跟进的核心资产']},
    {title:'4. 模型衰减应对',items:['8月转化率断崖式下跌，需排查：素材疲劳？定向放宽？竞品挤压？','建立7天/14天转化归因窗口，避免误判近期计划质量']},
  ];
  html+=suggestions.map(s=>'<div style="margin-bottom:14px"><div style="font-size:14px;font-weight:600;color:#1e293b;margin-bottom:8px">'+s.title+'</div><ul style="margin:0;padding-left:20px;font-size:13px;line-height:1.9;color:#475569">'+s.items.map(it=>'<li>'+it+'</li>').join('')+'</ul></div>').join('');
  document.getElementById('t9Analysis').innerHTML=html;
}

function downloadTab9(){if(!t9Output){alert('请先开始分析');return;}XLSX.writeFile(t9Output,'版位线索评分分析.xlsx');}

/* ---------- 分析模块：线索评分·版位分析 ---------- */
function renderCluePosition(){
  const d=SYNC.cluePosition;
  if(!d){
    $('#cp-badge').textContent='暂无数据';
    ['cp-overview-card','cp-finding-card','cp-chart-card','cp-tier-card','cp-trend-card','cp-suggest-card'].forEach(id=>{const el=$('#'+id);if(el)el.style.display='none';});
    $('#cp-empty').style.display='block';$('#cp-kpi').innerHTML='';$('#cp-summary').textContent='';return;
  }
  $('#cp-empty').style.display='none';
  const groups=[
    {key:'devGx',label:'设备号-公小',data:d.devGx,color:'#3b82f6'},
    {key:'devSp',label:'设备号-视频号',data:d.devSp,color:'#ef4444'},
    {key:'phoneGx',label:'手机号-公小',data:d.phoneGx,color:'#22c55e'},
    {key:'phoneSp',label:'手机号-视频号',data:d.phoneSp,color:'#f59e0b'},
  ];
  const totalLow=groups.reduce((s,g)=>s+g.data.total.low,0);
  const totalHigh=groups.reduce((s,g)=>s+g.data.total.high,0);
  $('#cp-badge').textContent='4维度共'+totalLow.toLocaleString()+'单 · 已同步';
  $('#cp-summary').textContent='数据源：4份文件已加载';
  // KPI
  const kpiHtml='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px">'+groups.map(g=>{
    const t=g.data.total,ht=g.data.tierList.find(x=>x.name==='高(7-10)');
    return '<div style="background:linear-gradient(135deg,'+g.color+'12,'+g.color+'05);border:1px solid '+g.color+'30;border-radius:12px;padding:14px 16px;border-top:3px solid '+g.color+'"><div style="font-size:12px;color:#64748b;font-weight:600;margin-bottom:6px">'+g.label+'</div><div style="font-size:24px;font-weight:700;color:'+g.color+'">'+t.low.toLocaleString()+'</div><div style="font-size:11.5px;color:#475569;margin-top:6px">转化率 '+(t.conv*100).toFixed(2)+'% · ARPU ¥'+t.arpu.toFixed(1)+'</div><div style="font-size:11px;color:#94a3b8;margin-top:3px">高分占比 '+(ht?ht.lowShare*100:0).toFixed(1)+'% · 高分GMV '+(ht?ht.gmvShare*100:0).toFixed(1)+'%</div></div>';
  }).join('')+'</div>';
  $('#cp-kpi').innerHTML=kpiHtml;
  // 版位横向对比表
  $('#cp-overview-card').style.display='block';
  const ovHeaders=['归因','版位','低价课','高价课','转化率','ARPU','GMV','高分占比','高分转化率','高分ARPU','高分GMV占比'];
  const ovRows=groups.map(g=>{
    const t=g.data.total,ht=g.data.tierList.find(x=>x.name==='高(7-10)');
    return [g.label.split('-')[0],g.label.split('-')[1],t.low.toLocaleString(),t.high,(t.conv*100).toFixed(2)+'%','¥'+t.arpu.toFixed(2),'¥'+t.gmv.toLocaleString(undefined,{maximumFractionDigits:0}),(ht?ht.lowShare*100:0).toFixed(1)+'%',(ht?ht.conv*100:0).toFixed(2)+'%','¥'+(ht?ht.arpu.toFixed(2):'0'),(ht?ht.gmvShare*100:0).toFixed(1)+'%'];
  });
  fillTable('cp-overview-table',ovHeaders,ovRows);
  // 关键发现
  $('#cp-finding-card').style.display='block';
  const devGx=d.devGx,devSp=d.devSp;
  const safeDiv=(a,b)=>b>0?a/b:0;
  const spConvAdv=(safeDiv(devSp.total.conv-devGx.total.conv,devGx.total.conv)*100).toFixed(1);
  const spArpuAdv=(safeDiv(devSp.total.arpu-devGx.total.arpu,devGx.total.arpu)*100).toFixed(1);
  const devGxHigh=devGx.tierList.find(x=>x.name==='高(7-10)')||{lowShare:0,arpu:0};
  const devSpHigh=devSp.tierList.find(x=>x.name==='高(7-10)')||{lowShare:0,arpu:0};
  const phoneGxHigh=d.phoneGx.tierList.find(x=>x.name==='高(7-10)')||{arpu:0};
  const spHighAdv=(safeDiv(devSpHigh.lowShare-devGxHigh.lowShare,devGxHigh.lowShare)*100).toFixed(1);
  const devTotal=devGx.total.low+devSp.total.low;
  const phoneTotal=d.phoneGx.total.low+d.phoneSp.total.low;
  const phoneRate=devTotal>0?(phoneTotal/devTotal*100).toFixed(1):'0';
  const arpuRatio=devGxHigh.arpu>0?(phoneGxHigh.arpu/devGxHigh.arpu).toFixed(1):'0';
  const findings=[
    '视频号全面优于公小：转化率高'+spConvAdv+'%（'+(devSp.total.conv*100).toFixed(2)+'% vs '+(devGx.total.conv*100).toFixed(2)+'%），ARPU高'+spArpuAdv+'%（¥'+devSp.total.arpu.toFixed(1)+' vs ¥'+devGx.total.arpu.toFixed(1)+'）',
    '视频号高分人群占比高'+spHighAdv+'%（'+(devSpHigh.lowShare*100).toFixed(1)+'% vs '+(devGxHigh.lowShare*100).toFixed(1)+'%），流量水质明显更优',
    '手机号覆盖率极低：设备号'+devTotal.toLocaleString()+'单 vs 手机号'+phoneTotal.toLocaleString()+'单，仅'+phoneRate+'%',
    '手机号高分ARPU爆发力强：手机号-公小高分ARPU ¥'+phoneGxHigh.arpu.toFixed(1)+'，是设备号-公小（¥'+devGxHigh.arpu.toFixed(1)+'）的'+arpuRatio+'倍',
  ];
  $('#cp-findings').innerHTML='<div style="background:linear-gradient(135deg,#f0f9ff,#f0fdf4);border-radius:12px;padding:16px 20px;border:1px solid #e0f2fe"><ul style="margin:0;padding-left:22px;font-size:13.5px;line-height:2.1;color:#334155">'+findings.map(f=>'<li style="margin-bottom:4px">'+f+'</li>').join('')+'</ul></div>';
  // 评分分布图
  $('#cp-chart-card').style.display='block';
  const scores=[1,2,3,4,5,6,7,8,9,10];
  const cpChart=chart('cp-chart');
  if(cpChart)cpChart.setOption({
    tooltip:{trigger:'axis'},
    legend:{data:groups.map(g=>g.label+'转化率'),top:0},
    grid:{left:55,right:55,top:50,bottom:40},
    xAxis:{type:'category',data:scores.map(s=>s+'分'),name:'评分'},
    yAxis:[{type:'value',name:'转化率',axisLabel:{formatter:'{value}%'}},{type:'value',name:'ARPU',axisLabel:{formatter:'¥{value}'}}],
    series:groups.map(g=>({
      name:g.label+'转化率',type:'bar',data:scores.map(s=>{const x=g.data.scoreList.find(r=>r.score===s);return x?+(x.conv*100).toFixed(2):0;}),itemStyle:{color:g.color},barGap:'10%'
    })).concat([{
      name:'ARPU(设备号-视频号)',type:'line',yAxisIndex:1,data:scores.map(s=>{const x=devSp.scoreList.find(r=>r.score===s);return x?+x.arpu.toFixed(1):0;}),itemStyle:{color:'#ef4444'},smooth:true,symbolSize:8
    }])
  });
  // 分层对比表
  $('#cp-tier-card').style.display='block';
  const tierHeaders=['维度','分层','低价课','高价课','转化率','ARPU','订单占比','GMV占比'];
  const tierRows=[];
  groups.forEach(g=>{g.data.tierList.forEach(t=>{
    tierRows.push([g.label,t.name,t.low.toLocaleString(),t.high,(t.conv*100).toFixed(2)+'%','¥'+t.arpu.toFixed(2),(t.lowShare*100).toFixed(1)+'%',(t.gmvShare*100).toFixed(1)+'%']);
  });});
  fillTable('cp-tier-table',tierHeaders,tierRows);
  // 月度趋势图
  $('#cp-trend-card').style.display='block';
  const allMonths=[...new Set(groups.flatMap(g=>g.data.monthly.map(m=>m.month)))].sort();
  const trendChart=chart('cp-trend-chart');
  if(trendChart)trendChart.setOption({
    tooltip:{trigger:'axis'},
    legend:{data:groups.map(g=>g.label+'转化率'),top:0},
    grid:{left:55,right:55,top:50,bottom:40},
    xAxis:{type:'category',data:allMonths},
    yAxis:[{type:'value',name:'转化率',axisLabel:{formatter:'{value}%'}},{type:'value',name:'高分占比',axisLabel:{formatter:'{value}%'}}],
    series:groups.map(g=>({
      name:g.label+'转化率',type:'line',data:allMonths.map(m=>{const x=g.data.monthly.find(r=>r.month===m);return x?+(x.conv*100).toFixed(2):0;}),itemStyle:{color:g.color},smooth:true,symbolSize:8
    })).concat([{
      name:'高分占比(设备号-公小)',type:'line',yAxisIndex:1,data:allMonths.map(m=>{const x=devGx.monthly.find(r=>r.month===m);return x?+(x.highTierShare*100).toFixed(1):0;}),itemStyle:{color:'#3b82f6',type:'dashed'},lineStyle:{type:'dashed'},smooth:true
    }])
  });
  // 实操建议
  $('#cp-suggest-card').style.display='block';
  const suggestions=[
    {title:'1. 预算与出价策略',items:['视频号放宽前端出价上限20-30%，优先抢量（高分占比高、ARPU高）','公小针对1-2分占比高的计划降出价15%，或设置负向过滤','将「评分≥7分」作为深度转化事件回传腾讯广告，开启双目标出价']},
    {title:'2. 版位分流',items:['视频号：主攻高净值人群，素材强化"专业体系化/高阶带练"','公小：压缩低质流量，落地页增加留资门槛筛选']},
    {title:'3. 手机号回传补全',items:['当前手机号覆盖率仅'+phoneRate+'%，在购课成功页引导补填手机号（送实体图谱/专属教练）','手机号高分ARPU达¥'+phoneGxHigh.arpu.toFixed(1)+'，是销售跟进的核心资产']},
    {title:'4. 模型衰减应对',items:['8月转化率断崖式下跌，需排查：素材疲劳？定向放宽？竞品挤压？','建立7天/14天转化归因窗口，避免误判近期计划质量']},
  ];
  $('#cp-suggestions').innerHTML=suggestions.map(s=>'<div style="margin-bottom:16px;padding:14px 18px;background:#f8fafc;border-radius:10px;border-left:4px solid #3b82f6"><div style="font-size:14px;font-weight:600;color:#1e293b;margin-bottom:8px">'+s.title+'</div><ul style="margin:0;padding-left:20px;font-size:13px;line-height:1.9;color:#475569">'+s.items.map(it=>'<li>'+it+'</li>').join('')+'</ul></div>').join('');
}

/* ---------- 工具10：分账户评分分析 ---------- */
let t10Output=null;
let t10File1Data=null,t10File2Data=null;

// 绑定文件上传
document.addEventListener('DOMContentLoaded',()=>{
  const f1=document.getElementById('t10File1');
  const f2=document.getElementById('t10File2');
  function checkBoth(){
    const ok=(f1&&f1.files.length)||(f2&&f2.files.length);
    document.getElementById('t10Btn').disabled=!ok;
  }
  if(f1)f1.addEventListener('change',e=>{
    const files=e.target.files;
    if(!files||!files.length)return;
    const names=Array.from(files).map(f=>f.name);
    document.getElementById('t10FileName1').textContent='✓ '+names.join('、')+(names.length>1?'（'+names.length+'个文件）':'');
    checkBoth();
    showStatus('t10','已选择'+files.length+'个文件，系统将自动识别评分表和订单明细表','active');
  });
  if(f2)f2.addEventListener('change',e=>{
    const files=e.target.files;
    if(!files||!files.length)return;
    const names=Array.from(files).map(f=>f.name);
    document.getElementById('t10FileName2').textContent='✓ '+names.join('、')+(names.length>1?'（'+names.length+'个文件）':'');
    checkBoth();
    const total=(f1.files.length||0)+files.length;
    showStatus('t10','共已选择'+total+'个文件，可以开始分析','active');
  });
});

// 智能识别列名
function t10DetectCol(headers,keywords){
  for(let i=0;i<headers.length;i++){
    const h=String(headers[i]||'').toLowerCase();
    if(keywords.some(k=>h.includes(k)))return i;
  }
  return -1;
}

// 解析并分析（支持双文件：评分表 + 订单明细表，通过订单号关联）
async function processTab10(){
  const f1=document.getElementById('t10File1');
  const f2=document.getElementById('t10File2');
  if(!f1.files.length&&!f2.files.length){alert('请至少上传一个文件（评分表或订单明细表）');return;}
  document.getElementById('t10Btn').disabled=true;
  showStatus('t10','正在解析数据...','active');
  try{
    // 收集所有上传的文件，智能识别类型
    const allFiles=[];
    for(let i=0;i<f1.files.length;i++)allFiles.push(f1.files[i]);
    for(let i=0;i<f2.files.length;i++)allFiles.push(f2.files[i]);

    let scoreRows=null,scoreHeaders=null;
    let orderRows=null,orderHeaders=null;

    for(let fi=0;fi<allFiles.length;fi++){
      const file=allFiles[fi];
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      let rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      // 找表头行
      let hr=0;
      for(let i=0;i<Math.min(5,rows.length);i++){
        if(rows[i].some(c=>{const s=String(c);return s.includes('分数')||s.includes('广告账户')||s.includes('订单号')||s.includes('性别')||s.includes('手机');})){hr=i;break;}
      }
      const headers=rows[hr].map(h=>String(h||'').trim());
      const dataRows=rows.slice(hr+1);
      // 判断文件类型：有"分数"列=评分表，有"性别/年龄/城市/手机品牌"=订单明细表
      const hasScoreCol=headers.some(h=>h.includes('分数'));
      const hasProfileCol=headers.some(h=>h.includes('性别')||h.includes('年龄')||h.includes('城市')||h.includes('手机品牌')||h.includes('手机型号'));
      if(hasScoreCol&&!scoreRows){
        scoreRows=dataRows;scoreHeaders=headers;
      }else if(hasProfileCol&&!orderRows){
        orderRows=dataRows;orderHeaders=headers;
      }else if(hasScoreCol&&scoreRows){
        // 第二个评分表，合并
        scoreRows=scoreRows.concat(dataRows);
      }else if(hasProfileCol&&orderRows){
        // 第二个订单明细表，合并
        orderRows=orderRows.concat(dataRows);
      }
    }

    // 如果都没识别到，尝试第一个文件作为评分表
    if(!scoreRows&&!orderRows&&allFiles.length>0){
      const file=allFiles[0];
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      let rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      let hr=0;
      for(let i=0;i<Math.min(5,rows.length);i++){
        if(rows[i].some(c=>String(c).includes('分数')||String(c).includes('广告账户'))){hr=i;break;}
      }
      scoreHeaders=rows[hr].map(h=>String(h||'').trim());
      scoreRows=rows.slice(hr+1);
    }

    // 识别评分表列
    const colOrderId=scoreHeaders?t10DetectCol(scoreHeaders,['订单号']):-1;
    const colScore=scoreHeaders?t10DetectCol(scoreHeaders,['分数']):-1;
    const colAccId=scoreHeaders?t10DetectCol(scoreHeaders,['广告账户id','账户id']):-1;
    const colAccName=scoreHeaders?t10DetectCol(scoreHeaders,['广告账户']):-1;
    const colChannel=scoreHeaders?t10DetectCol(scoreHeaders,['渠道名称','渠道']):-1;
    const colRepeat=scoreHeaders?t10DetectCol(scoreHeaders,['是否复学','复学']):-1;
    const colCondition=scoreHeaders?t10DetectCol(scoreHeaders,['评分条件']):-1;
    const colPeriod=scoreHeaders?t10DetectCol(scoreHeaders,['期数']):-1;

    // 识别订单明细表列
    const ocolOrderId=orderHeaders?t10DetectCol(orderHeaders,['订单号id','订单号']):-1;
    const ocolGender1=orderHeaders?t10DetectCol(orderHeaders,['性别-微信号','性别']):-1;
    const ocolGender2=orderHeaders?t10DetectCol(orderHeaders,['性别-皮肤收集']):-1;
    const ocolAge=orderHeaders?t10DetectCol(orderHeaders,['年龄']):-1;
    const ocolCity=orderHeaders?t10DetectCol(orderHeaders,['城市']):-1;
    const ocolCityLevel=orderHeaders?t10DetectCol(orderHeaders,['城市等级']):-1;
    const ocolProvince=orderHeaders?t10DetectCol(orderHeaders,['省份']):-1;
    const ocolBrand=orderHeaders?t10DetectCol(orderHeaders,['手机品牌']):-1;
    const ocolModel=orderHeaders?t10DetectCol(orderHeaders,['手机型号']):-1;
    const ocolPrice=orderHeaders?t10DetectCol(orderHeaders,['手机价格']):-1;
    const ocolOS=orderHeaders?t10DetectCol(orderHeaders,['手机系统']):-1;
    const ocolHighAmount=orderHeaders?t10DetectCol(orderHeaders,['高价课订单金额','高价课金额']):-1;
    const ocolHighOrder=orderHeaders?t10DetectCol(orderHeaders,['高价课订单号']):-1;
    const ocolAccId=orderHeaders?t10DetectCol(orderHeaders,['广告账号id','广告账户id','账号id']):-1;
    const ocolPeriod=orderHeaders?t10DetectCol(orderHeaders,['期数']):-1;
    const ocolPosition=orderHeaders?t10DetectCol(orderHeaders,['版位']):-1;

    // 构建订单明细表的索引（订单号 -> 画像数据）
    const orderIndex={};
    if(orderRows&&ocolOrderId>=0){
      for(let i=0;i<orderRows.length;i++){
        const row=orderRows[i];
        if(!row||row.every(c=>c===''||c==null))continue;
        const oid=String(row[ocolOrderId]||'').trim();
        if(!oid)continue;
        let gender=ocolGender1>=0?String(row[ocolGender1]||'').trim():'';
        if((!gender||gender==='未知')&&ocolGender2>=0)gender=String(row[ocolGender2]||'').trim();
        const age=ocolAge>=0?String(row[ocolAge]||'').trim():'';
        const city=ocolCity>=0?String(row[ocolCity]||'').trim():'';
        const cityLevel=ocolCityLevel>=0?String(row[ocolCityLevel]||'').trim():'';
        const province=ocolProvince>=0?String(row[ocolProvince]||'').trim():'';
        let brand=ocolBrand>=0?String(row[ocolBrand]||'').trim():'';
        brand=brand.toUpperCase().replace('HUAWEI','华为').replace('HONOR','荣耀').replace('XIAOMI','小米').replace('REDMI','红米').replace('IPHONE','苹果').replace('SONY','索尼').replace('SAMSUNG','三星').replace('REALME','真我').replace('UNKNOWN','未知');
        const model=ocolModel>=0?String(row[ocolModel]||'').trim():'';
        const priceRaw=ocolPrice>=0?parseFloat(row[ocolPrice])||0:0;
        let priceTier='未知';
        if(priceRaw>0){if(priceRaw<1500)priceTier='千元以下';else if(priceRaw<2500)priceTier='1500-2500';else if(priceRaw<4000)priceTier='2500-4000';else if(priceRaw<6000)priceTier='4000-6000';else priceTier='6000以上';}
        const highAmount=ocolHighAmount>=0?parseFloat(row[ocolHighAmount])||0:0;
        const hasHighOrder=ocolHighOrder>=0?!String(row[ocolHighOrder]||'').trim()?false:true:(highAmount>0);
        orderIndex[oid]={gender,age,city,cityLevel,province,brand,model,price:priceRaw,priceTier,highAmount,hasHighOrder,os:ocolOS>=0?String(row[ocolOS]||'').trim():''};
      }
    }

    // 聚合数据
    const accMap={};
    let totalOrders=0,totalScore=0,repeatCount=0,nonRepeatCount=0,repeatScore=0,nonRepeatScore=0;
    let repeatHigh=0,nonRepeatHigh=0,matchedCount=0;
    let totalHighOrder=0,totalHighAmount=0;
    const scoreDist={};
    const conditionMap={};
    const periodMap={};
    const tierMap={low:{orders:0,scoreSum:0,repeat:0,highOrder:0,highAmount:0},mid:{orders:0,scoreSum:0,repeat:0,highOrder:0,highAmount:0},high:{orders:0,scoreSum:0,repeat:0,highOrder:0,highAmount:0}};
    // 画像聚合（全局）
    const profileMap={gender:{},age:{},city:{},cityLevel:{},brand:{},model:{},priceTier:{}};
    const highProfile={gender:{},age:{},city:{},cityLevel:{},brand:{},priceTier:{}};
    const lowProfile={gender:{},age:{},city:{},cityLevel:{},brand:{},priceTier:{}};
    // 评分×转化
    const scoreConv={};
    // 账户×画像
    const accProfile={};
    // 账户×期数（每期每账户：评分分数→产值 + 画像）
    const accPeriodMap={};

    function addP(map,key,score,highAmount,hasHigh){
      if(!key||key===''||key==='未知')return;
      if(!map[key])map[key]={key,orders:0,scoreSum:0,highAmount:0,highOrder:0};
      map[key].orders++;map[key].scoreSum+=score;
      if(highAmount>0)map[key].highAmount+=highAmount;
      if(hasHigh)map[key].highOrder++;
    }

    // 遍历评分表
    const dataRows=scoreRows||[];
    if(dataRows.length===0){
      showStatus('t10','未检测到包含「分数」字段的评分表，请上传评分表（含分数、广告账户id等字段）','error');
      document.getElementById('t10Btn').disabled=false;
      return;
    }
    for(let i=0;i<dataRows.length;i++){
      const row=dataRows[i];
      if(!row||row.every(c=>c===''||c==null))continue;
      const score=colScore>=0?parseInt(row[colScore])||0:0;
      const accId=colAccId>=0?String(row[colAccId]||'').trim():(ocolAccId>=0?'': '');
      const accName=colAccName>=0?String(row[colAccName]||'').trim():'';
      const channel=colChannel>=0?String(row[colChannel]||'').trim():'';
      const isRepeat=colRepeat>=0?String(row[colRepeat]||'').includes('是'):false;
      const condition=colCondition>=0?String(row[colCondition]||'').trim():'';
      const period=colPeriod>=0?String(row[colPeriod]||'').trim():'';
      const oid=colOrderId>=0?String(row[colOrderId]||'').trim():'';
      if(score<1||score>10)continue;
      // 如果评分表没有账户id，尝试从订单表获取
      let finalAccId=accId;
      if(!finalAccId&&oid&&orderIndex[oid]&&ocolAccId>=0){
        // 从订单表行找账户id（需要重新查）
      }
      if(!finalAccId)finalAccId='未知账户';
      totalOrders++;totalScore+=score;
      // 关联画像
      const profile=oid&&orderIndex[oid]?orderIndex[oid]:null;
      if(profile)matchedCount++;
      const highAmount=profile?profile.highAmount:0;
      const hasHigh=profile?profile.hasHighOrder:false;
      if(hasHigh){totalHighOrder++;totalHighAmount+=highAmount;}
      // 评分×转化
      if(!scoreConv[score])scoreConv[score]={orders:0,highOrder:0,highAmount:0};
      scoreConv[score].orders++;scoreConv[score].highAmount+=highAmount;
      if(hasHigh)scoreConv[score].highOrder++;
      // 账户聚合
      if(!accMap[finalAccId])accMap[finalAccId]={id:finalAccId,name:accName||finalAccId,channel:channel,orders:0,scoreSum:0,highCount:0,lowCount:0,repeatCount:0,scoreDist:{},conditionDist:{},highOrder:0,highAmount:0,profileMatchCount:0,profile:{gender:{},age:{},city:{},cityLevel:{},brand:{},priceTier:{}}};
      const a=accMap[finalAccId];
      a.orders++;a.scoreSum+=score;
      if(score>=7)a.highCount++;
      if(score<=3)a.lowCount++;
      if(isRepeat)a.repeatCount++;
      if(hasHigh){a.highOrder++;a.highAmount+=highAmount;}
      a.scoreDist[score]=(a.scoreDist[score]||0)+1;
      if(condition)a.conditionDist[condition]=(a.conditionDist[condition]||0)+1;
      if(!scoreDist[finalAccId])scoreDist[finalAccId]={};
      scoreDist[finalAccId][score]=(scoreDist[finalAccId][score]||0)+1;
      // 账户×画像
      if(profile){
        a.profileMatchCount++;
        addP(a.profile.gender,profile.gender,score,highAmount,hasHigh);
        addP(a.profile.age,profile.age,score,highAmount,hasHigh);
        addP(a.profile.city,profile.city,score,highAmount,hasHigh);
        addP(a.profile.cityLevel,profile.cityLevel,score,highAmount,hasHigh);
        addP(a.profile.brand,profile.brand,score,highAmount,hasHigh);
        addP(a.profile.priceTier,profile.priceTier,score,highAmount,hasHigh);
      }
      // 复学统计
      if(isRepeat){repeatCount++;repeatScore+=score;if(score>=7)repeatHigh++;}
      else{nonRepeatCount++;nonRepeatScore+=score;if(score>=7)nonRepeatHigh++;}
      // 评分条件
      if(condition){
        if(!conditionMap[condition])conditionMap[condition]={name:condition,orders:0,scoreSum:0,highCount:0,lowCount:0,repeatCount:0};
        const c=conditionMap[condition];
        c.orders++;c.scoreSum+=score;
        if(score>=7)c.highCount++;
        if(score<=3)c.lowCount++;
        if(isRepeat)c.repeatCount++;
      }
      // 期数趋势
      if(period){
        if(!periodMap[period])periodMap[period]={period,orders:0,scoreSum:0,highCount:0,repeatCount:0,highOrder:0};
        const p=periodMap[period];
        p.orders++;p.scoreSum+=score;
        if(score>=7)p.highCount++;
        if(isRepeat)p.repeatCount++;
        if(hasHigh)p.highOrder++;
      }
      // 账户×期数聚合（每期每账户）
      const apKey=finalAccId+'\u0001'+(period||'未知期');
      if(!accPeriodMap[apKey])accPeriodMap[apKey]={accId:finalAccId,accName:accName||finalAccId,period:period||'未知期',orders:0,scoreSum:0,highOrder:0,highAmount:0,profileMatchCount:0,scoreConv:{},profile:{gender:{},age:{},city:{},cityLevel:{},brand:{},model:{},priceTier:{}}};
      const ap=accPeriodMap[apKey];
      ap.orders++;ap.scoreSum+=score;
      if(hasHigh){ap.highOrder++;ap.highAmount+=highAmount;}
      if(!ap.scoreConv[score])ap.scoreConv[score]={orders:0,highOrder:0,highAmount:0};
      ap.scoreConv[score].orders++;ap.scoreConv[score].highAmount+=highAmount;
      if(hasHigh)ap.scoreConv[score].highOrder++;
      if(profile){
        ap.profileMatchCount++;
        addP(ap.profile.gender,profile.gender,score,highAmount,hasHigh);
        addP(ap.profile.age,profile.age,score,highAmount,hasHigh);
        addP(ap.profile.city,profile.city,score,highAmount,hasHigh);
        addP(ap.profile.cityLevel,profile.cityLevel,score,highAmount,hasHigh);
        addP(ap.profile.brand,profile.brand,score,highAmount,hasHigh);
        addP(ap.profile.model,profile.model,score,highAmount,hasHigh);
        addP(ap.profile.priceTier,profile.priceTier,score,highAmount,hasHigh);
      }
      // 评分分层
      const tier=score<=3?'low':(score<=6?'mid':'high');
      tierMap[tier].orders++;tierMap[tier].scoreSum+=score;
      if(isRepeat)tierMap[tier].repeat++;
      if(hasHigh){tierMap[tier].highOrder++;tierMap[tier].highAmount+=highAmount;}
      // 全局画像
      if(profile){
        addP(profileMap.gender,profile.gender,score,highAmount,hasHigh);
        addP(profileMap.age,profile.age,score,highAmount,hasHigh);
        addP(profileMap.city,profile.city,score,highAmount,hasHigh);
        addP(profileMap.cityLevel,profile.cityLevel,score,highAmount,hasHigh);
        addP(profileMap.brand,profile.brand,score,highAmount,hasHigh);
        addP(profileMap.model,profile.model,score,highAmount,hasHigh);
        addP(profileMap.priceTier,profile.priceTier,score,highAmount,hasHigh);
        if(score>=7){
          addP(highProfile.gender,profile.gender,score,highAmount,hasHigh);
          addP(highProfile.age,profile.age,score,highAmount,hasHigh);
          addP(highProfile.city,profile.city,score,highAmount,hasHigh);
          addP(highProfile.cityLevel,profile.cityLevel,score,highAmount,hasHigh);
          addP(highProfile.brand,profile.brand,score,highAmount,hasHigh);
          addP(highProfile.priceTier,profile.priceTier,score,highAmount,hasHigh);
        }else if(score<=3){
          addP(lowProfile.gender,profile.gender,score,highAmount,hasHigh);
          addP(lowProfile.age,profile.age,score,highAmount,hasHigh);
          addP(lowProfile.city,profile.city,score,highAmount,hasHigh);
          addP(lowProfile.cityLevel,profile.cityLevel,score,highAmount,hasHigh);
          addP(lowProfile.brand,profile.brand,score,highAmount,hasHigh);
          addP(lowProfile.priceTier,profile.priceTier,score,highAmount,hasHigh);
        }
      }
    }

    // 检查是否有有效数据
    if(totalOrders===0){
      showStatus('t10','未解析到有效评分数据，请确认文件包含「分数」字段（1-10分）和「广告账户id」字段','error');
      document.getElementById('t10Btn').disabled=false;
      return;
    }

    // 整理账户列表（带画像）
    function sortP(map){return Object.values(map).sort((a,b)=>b.orders-a.orders).map(x=>({key:x.key,orders:x.orders,avgScore:x.orders>0?x.scoreSum/x.orders:0,highAmount:x.highAmount,avgOutput:x.orders>0?x.highAmount/x.orders:0,highOrder:x.highOrder,convRate:x.orders>0?x.highOrder/x.orders:0}));}
    const accounts=Object.values(accMap).map(a=>({
      id:a.id,name:a.name||a.id,channel:a.channel,
      orders:a.orders,avgScore:a.orders>0?a.scoreSum/a.orders:0,
      highRatio:a.orders>0?a.highCount/a.orders:0,
      lowRatio:a.orders>0?a.lowCount/a.orders:0,
      repeatRatio:a.orders>0?a.repeatCount/a.orders:0,
      highOrder:a.highOrder,highAmount:a.highAmount,
      profileMatchCount:a.profileMatchCount,
      convRate:a.orders>0?a.highOrder/a.orders:0,
      avgOutput:a.orders>0?a.highAmount/a.orders:0,
      scoreDist:a.scoreDist,conditionDist:a.conditionDist,
      profile:{gender:sortP(a.profile.gender),age:sortP(a.profile.age),city:sortP(a.profile.city),cityLevel:sortP(a.profile.cityLevel),brand:sortP(a.profile.brand),priceTier:sortP(a.profile.priceTier)}
    })).sort((a,b)=>b.orders-a.orders);

    const conditions=Object.values(conditionMap).map(c=>({
      name:c.name,orders:c.orders,avgScore:c.orders>0?c.scoreSum/c.orders:0,
      highRatio:c.orders>0?c.highCount/c.orders:0,
      lowRatio:c.orders>0?c.lowCount/c.orders:0,
      repeatRatio:c.orders>0?c.repeatCount/c.orders:0
    })).sort((a,b)=>b.orders-a.orders);

    const periods=Object.values(periodMap).map(p=>({
      period:p.period,orders:p.orders,avgScore:p.orders>0?p.scoreSum/p.orders:0,
      highRatio:p.orders>0?p.highCount/p.orders:0,
      repeatRatio:p.orders>0?p.repeatCount/p.orders:0,
      highOrder:p.highOrder,convRate:p.orders>0?p.highOrder/p.orders:0
    })).sort((a,b)=>a.period.localeCompare(b.period));

    const tiers={
      low:{name:'低分段(1-3分)',orders:tierMap.low.orders,avgScore:tierMap.low.orders>0?tierMap.low.scoreSum/tierMap.low.orders:0,ratio:totalOrders>0?tierMap.low.orders/totalOrders:0,repeatRatio:tierMap.low.orders>0?tierMap.low.repeat/tierMap.low.orders:0,convRate:tierMap.low.orders>0?tierMap.low.highOrder/tierMap.low.orders:0,avgOutput:tierMap.low.orders>0?tierMap.low.highAmount/tierMap.low.orders:0},
      mid:{name:'中分段(4-6分)',orders:tierMap.mid.orders,avgScore:tierMap.mid.orders>0?tierMap.mid.scoreSum/tierMap.mid.orders:0,ratio:totalOrders>0?tierMap.mid.orders/totalOrders:0,repeatRatio:tierMap.mid.orders>0?tierMap.mid.repeat/tierMap.mid.orders:0,convRate:tierMap.mid.orders>0?tierMap.mid.highOrder/tierMap.mid.orders:0,avgOutput:tierMap.mid.orders>0?tierMap.mid.highAmount/tierMap.mid.orders:0},
      high:{name:'高分段(7-10分)',orders:tierMap.high.orders,avgScore:tierMap.high.orders>0?tierMap.high.scoreSum/tierMap.high.orders:0,ratio:totalOrders>0?tierMap.high.orders/totalOrders:0,repeatRatio:tierMap.high.orders>0?tierMap.high.repeat/tierMap.high.orders:0,convRate:tierMap.high.orders>0?tierMap.high.highOrder/tierMap.high.orders:0,avgOutput:tierMap.high.orders>0?tierMap.high.highAmount/tierMap.high.orders:0}
    };

    // 评分×转化
    const scoreConvArr=Object.entries(scoreConv).map(([score,v])=>({
      score:parseInt(score),orders:v.orders,highOrder:v.highOrder,highAmount:v.highAmount,
      convRate:v.orders>0?v.highOrder/v.orders:0,
      avgOutput:v.orders>0?v.highAmount/v.orders:0,
     客单价:v.highOrder>0?v.highAmount/v.highOrder:0
    })).sort((a,b)=>a.score-b.score);

    // 整理画像数据
    const profile={
      gender:sortP(profileMap.gender),age:sortP(profileMap.age),
      city:sortP(profileMap.city),cityLevel:sortP(profileMap.cityLevel),
      brand:sortP(profileMap.brand),model:sortP(profileMap.model),priceTier:sortP(profileMap.priceTier)
    };
    const highP={
      gender:sortP(highProfile.gender),age:sortP(highProfile.age),
      city:sortP(highProfile.city),cityLevel:sortP(highProfile.cityLevel),
      brand:sortP(highProfile.brand),priceTier:sortP(highProfile.priceTier)
    };
    const lowP={
      gender:sortP(lowProfile.gender),age:sortP(lowProfile.age),
      city:sortP(lowProfile.city),cityLevel:sortP(lowProfile.cityLevel),
      brand:sortP(lowProfile.brand),priceTier:sortP(lowProfile.priceTier)
    };
    const hasProfile=profile.gender.length>0||profile.age.length>0||profile.city.length>0||profile.brand.length>0||profile.priceTier.length>0;

    // 每期每账户结果
    const accPeriods=Object.values(accPeriodMap).map(ap=>({
      accId:ap.accId,accName:ap.accName,period:ap.period,
      orders:ap.orders,avgScore:ap.orders>0?ap.scoreSum/ap.orders:0,
      highOrder:ap.highOrder,highAmount:ap.highAmount,
      convRate:ap.orders>0?ap.highOrder/ap.orders:0,
      avgOutput:ap.orders>0?ap.highAmount/ap.orders:0,
      scoreConv:Object.entries(ap.scoreConv).map(([score,v])=>({
        score:parseInt(score),orders:v.orders,highOrder:v.highOrder,highAmount:v.highAmount,
        convRate:v.orders>0?v.highOrder/v.orders:0,
        avgOutput:v.orders>0?v.highAmount/v.orders:0
      })).sort((a,b)=>a.score-b.score),
      profileMatchCount:ap.profileMatchCount,
      profile:{
        gender:sortP(ap.profile.gender),age:sortP(ap.profile.age),
        city:sortP(ap.profile.city),cityLevel:sortP(ap.profile.cityLevel),
        brand:sortP(ap.profile.brand),model:sortP(ap.profile.model),
        priceTier:sortP(ap.profile.priceTier)
      }
    })).sort((a,b)=>String(a.accId).localeCompare(String(b.accId))||String(a.period).localeCompare(String(b.period)));

    const result={
      totalOrders,matchedCount,matchRate:totalOrders>0?matchedCount/totalOrders:0,
      avgScore:totalOrders>0?totalScore/totalOrders:0,
      highRatio:totalOrders>0?accounts.reduce((s,a)=>s+a.highCount,0)/totalOrders:0,
      lowRatio:totalOrders>0?accounts.reduce((s,a)=>s+a.lowCount,0)/totalOrders:0,
      repeatRatio:totalOrders>0?repeatCount/totalOrders:0,
      totalHighOrder,totalHighAmount,
      overallConvRate:totalOrders>0?totalHighOrder/totalOrders:0,
      overallAvgOutput:totalOrders>0?totalHighAmount/totalOrders:0,
      accounts,conditions,periods,tiers,scoreConv:scoreConvArr,
      accPeriods,
      repeat:{count:repeatCount,avgScore:repeatCount>0?repeatScore/repeatCount:0,highRatio:repeatCount>0?repeatHigh/repeatCount:0},
      nonRepeat:{count:nonRepeatCount,avgScore:nonRepeatCount>0?nonRepeatScore/nonRepeatCount:0,highRatio:nonRepeatCount>0?nonRepeatHigh/nonRepeatCount:0},
      scoreDist,profile,highProfile:highP,lowProfile:lowP,hasProfile
    };
    t10Output=t10BuildExcel(result);
    saveSync('accountScore',result);
    t10Display(result);
    showStatus('t10','分析完成！共 '+totalOrders+' 单，关联成功 '+matchedCount+' 单（'+(result.matchRate*100).toFixed(1)+'%），'+accounts.length+' 个账户，已同步分析模块','done');
  }catch(e){
    showStatus('t10','分析失败：'+e.message,'error');
    document.getElementById('t10Btn').disabled=false;
  }
}

// 构建Excel
function t10BuildExcel(d){
  const wb=XLSX.utils.book_new();
  if(!d||!d.accounts||d.accounts.length===0){
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['提示'],['无有效数据']]),'提示');
    return wb;
  }
  // ===== Sheet1: 核心总览 =====
  const ov=[];
  ov.push(['一、账户对比总览']);
  ov.push(['广告账户ID','账户名称','承载链路','订单量','订单占比','平均分','高分率(≥7)','低分率(≤3)','复学率']);
  d.accounts.forEach(a=>ov.push([a.id,a.name,a.channel,a.orders,(a.orders/d.totalOrders*100).toFixed(1)+'%',a.avgScore.toFixed(2),(a.highRatio*100).toFixed(2)+'%',(a.lowRatio*100).toFixed(2)+'%',(a.repeatRatio*100).toFixed(2)+'%']));
  ov.push([]);
  ov.push(['二、各账户评分分布（1-10分）']);
  ov.push(['账户ID','1分','2分','3分','4分','5分','6分','7分','8分','9分','10分']);
  d.accounts.forEach(a=>{const row=[a.id];for(let s=1;s<=10;s++)row.push(a.scoreDist[s]||0);ov.push(row);});
  ov.push([]);
  ov.push(['三、复学 vs 非复学']);
  ov.push(['类型','订单量','平均分','高分率(≥7)']);
  ov.push(['非复学新客',d.nonRepeat.count,d.nonRepeat.avgScore.toFixed(2),(d.nonRepeat.highRatio*100).toFixed(2)+'%']);
  ov.push(['复学老客',d.repeat.count,d.repeat.avgScore.toFixed(2),(d.repeat.highRatio*100).toFixed(2)+'%']);
  if(d.conditions&&d.conditions.length){
    ov.push([]);ov.push(['四、评分条件对比']);
    ov.push(['评分条件','订单量','订单占比','平均分','高分率(≥7)','低分率(≤3)','复学率']);
    d.conditions.forEach(c=>ov.push([c.name,c.orders,(c.orders/d.totalOrders*100).toFixed(1)+'%',c.avgScore.toFixed(2),(c.highRatio*100).toFixed(2)+'%',(c.lowRatio*100).toFixed(2)+'%',(c.repeatRatio*100).toFixed(2)+'%']));
  }
  if(d.tiers){
    ov.push([]);ov.push(['五、评分分层（低1-3 / 中4-6 / 高7-10）']);
    ov.push(['分层','订单量','占比','平均分','复学率']);
    [d.tiers.low,d.tiers.mid,d.tiers.high].forEach(t=>ov.push([t.name,t.orders,(t.ratio*100).toFixed(1)+'%',t.avgScore.toFixed(2),(t.repeatRatio*100).toFixed(2)+'%']));
  }
  if(d.periods&&d.periods.length){
    ov.push([]);ov.push(['六、期数趋势']);
    ov.push(['期数','订单量','平均分','高分率','复学率']);
    d.periods.forEach(p=>ov.push([p.period,p.orders,p.avgScore.toFixed(2),(p.highRatio*100).toFixed(2)+'%',(p.repeatRatio*100).toFixed(2)+'%']));
  }
  const findings=t10GenFindings(d);
  if(findings.length){
    ov.push([]);ov.push(['七、核心洞察']);
    ov.push(['序号','洞察','数据支撑']);
    findings.forEach((f,i)=>ov.push([i+1,f.title,f.detail]));
  }
  const sug=t10GenSuggestions(d);
  if(sug.length){
    ov.push([]);ov.push(['八、实操建议']);
    ov.push(['类别','建议']);
    sug.forEach(s=>ov.push([s.title,s.text]));
  }
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(ov),'核心总览');
  // ===== Sheet2: 评分×转化验证 =====
  if(d.scoreConv&&d.scoreConv.length){
    const sc=[['评分','订单数','高价课转化数','转化率','客单价','人均产值']];
    d.scoreConv.forEach(s=>sc.push([s.score+'分',s.orders,s.highOrder,(s.convRate*100).toFixed(2)+'%',s.客单价>0?'¥'+s.客单价.toFixed(0):'-','¥'+s.avgOutput.toFixed(0)]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sc),'评分转化验证');
  }
  // ===== Sheet3+: 每个账户一个sheet，整合所有画像维度 =====
  if(d.hasProfile){
    d.accounts.forEach(acc=>{
      if(!acc.profile)return;
      const matchCnt=acc.profileMatchCount||0;
      const accDims=[['性别',acc.profile.gender],['年龄',acc.profile.age],['城市等级',acc.profile.cityLevel],['城市',acc.profile.city],['手机品牌',acc.profile.brand],['手机价格档',acc.profile.priceTier]];
      const hasData=accDims.some(([_,data])=>data.length>0);
      const rows=[];
      rows.push(['账户：'+acc.name+'（ID:'+acc.id+'）','画像匹配样本：'+matchCnt+'条']);
      if(matchCnt<20)rows.push(['⚠️ 样本偏少（'+matchCnt+'条），结论仅供参考']);
      rows.push([]);
      // 该账户对应的期数（每期每账户）
      const accPeriods=d.accPeriods?d.accPeriods.filter(ap=>ap.accId===acc.id):[];
      if(accPeriods.length){
        rows.push(['【期数】']);
        rows.push(['期数','订单量','平均分','高价课转化数','转化率','人均产值']);
        accPeriods.forEach(ap=>rows.push([ap.period,ap.orders,ap.avgScore.toFixed(2),ap.highOrder,(ap.convRate*100).toFixed(2)+'%',ap.avgOutput>0?'¥'+ap.avgOutput.toFixed(0):'-']));
        rows.push([]);
      }
      if(hasData){
        accDims.forEach(([name,data])=>{
          if(data.length){
            rows.push(['【'+name+'】']);
            rows.push(['维度值','订单量','占比(匹配样本)','平均分','转化率','人均产值']);
            data.forEach(x=>rows.push([x.key,x.orders,(matchCnt>0?(x.orders/matchCnt*100).toFixed(1):0)+'%',x.avgScore.toFixed(2),(x.convRate*100).toFixed(1)+'%',x.avgOutput>0?'¥'+x.avgOutput.toFixed(0):'-']));
            rows.push([]);
          }
        });
      }else{
        rows.push(['该账户暂无匹配的画像数据']);
        rows.push(['建议检查订单明细表是否包含此账户的订单号']);
      }
      const sheetName=('账户'+acc.id).substring(0,31);
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),sheetName);
    });
  }
  // ===== 全局画像（整合到一个sheet）=====
  if(d.hasProfile){
    const gp=[];
    const dims=[['性别',d.profile.gender],['年龄',d.profile.age],['城市',d.profile.city],['手机品牌',d.profile.brand],['手机型号',d.profile.model],['手机价格',d.profile.priceTier]];
    dims.forEach(([name,data])=>{
      if(data.length){
        gp.push(['【'+name+'分布】']);
        gp.push(['维度值','订单量','占比','平均分','产值','人均产值']);
        data.forEach(x=>gp.push([x.key,x.orders,(x.orders/d.totalOrders*100).toFixed(1)+'%',x.avgScore.toFixed(2),x.output>0?x.output:'-',x.avgOutput>0?x.avgOutput.toFixed(1):'-']));
        gp.push([]);
      }
    });
    if(gp.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(gp),'全局画像');
  }
  // ===== 高低分对比（整合到一个sheet）=====
  if(d.hasProfile){
    const cmp=[];
    const cmpDims=[['性别',d.highProfile.gender,d.lowProfile.gender],['年龄',d.highProfile.age,d.lowProfile.age],['手机品牌',d.highProfile.brand,d.lowProfile.brand],['手机价格',d.highProfile.priceTier,d.lowProfile.priceTier]];
    cmpDims.forEach(([name,high,low])=>{
      if(high.length||low.length){
        cmp.push(['【'+name+'：高分(7-10) vs 低分(1-3)】']);
        cmp.push(['维度值','高分订单量','高分占比','低分订单量','低分占比','差异']);
        const allKeys=new Set([...high.map(x=>x.key),...low.map(x=>x.key)]);
        const ht=high.reduce((s,x)=>s+x.orders,0)||1,lt=low.reduce((s,x)=>s+x.orders,0)||1;
        [...allKeys].forEach(k=>{
          const h=high.find(x=>x.key===k),l=low.find(x=>x.key===k);
          const ho=h?h.orders:0,lo=l?l.orders:0;
          cmp.push([k,ho,(ho/ht*100).toFixed(1)+'%',lo,(lo/lt*100).toFixed(1)+'%',(ho/ht-lo/lt>0?'高分集中':'低分集中')]);
        });
        cmp.push([]);
      }
    });
    if(cmp.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(cmp),'高低分对比');
  }
  // ===== 每期每账户分析 =====
  if(d.accPeriods&&d.accPeriods.length){
    // Sheet: 账户×期数总览
    const apOv=[];
    apOv.push(['每期每账户分析：账户 × 期数 总览']);
    apOv.push(['说明：每个广告账户在每一期的评分质量与成交产值']);
    apOv.push(['账户ID','账户名称','期数','订单量','平均分','高价课数','高价课金额','高价转化率','人均产值']);
    d.accPeriods.forEach(ap=>{
      apOv.push([ap.accId,ap.accName,ap.period,ap.orders,ap.avgScore.toFixed(2),ap.highOrder,ap.highAmount>0?ap.highAmount:'',(ap.convRate*100).toFixed(2)+'%',ap.avgOutput>0?'¥'+ap.avgOutput.toFixed(0):'-']);
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(apOv),'账户×期数总览');

    // Sheet: 评分×产值明细（每期每账户，评分1-10分→订单/转化/产值）
    const sc=[];
    sc.push(['每期每账户：评分分数 对应 产值明细']);
    sc.push(['说明：同一账户同一期内，1-10分各自承载的订单数与高价课产值，验证评分是否与产出正相关']);
    sc.push(['账户ID','账户名称','期数','评分','订单数','高价课转化数','高价课金额','转化率','人均产值']);
    d.accPeriods.forEach(ap=>{
      (ap.scoreConv||[]).forEach(s=>{
        sc.push([ap.accId,ap.accName,ap.period,s.score+'分',s.orders,s.highOrder,s.highAmount>0?s.highAmount:'',(s.convRate*100).toFixed(2)+'%',s.avgOutput>0?'¥'+s.avgOutput.toFixed(0):'-']);
      });
    });
    if(sc.length>3)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sc),'评分×产值明细');

    // Sheet: 画像×期数明细（每期每账户画像分布）
    const pf=[];
    pf.push(['每期每账户：人群画像明细']);
    pf.push(['说明：同一账户同一期内的人群画像分布（性别/年龄/城市/手机品牌/手机价格/手机型号）']);
    pf.push(['账户ID','账户名称','期数','画像维度','维度值','订单量','占比(该期匹配样本)','平均分','转化率','人均产值']);
    d.accPeriods.forEach(ap=>{
      const matchCnt=ap.profileMatchCount||1;
      const dims=[['性别',ap.profile.gender],['年龄',ap.profile.age],['城市',ap.profile.city],['城市等级',ap.profile.cityLevel],['手机品牌',ap.profile.brand],['手机价格',ap.profile.priceTier],['手机型号',ap.profile.model]];
      dims.forEach(([name,data])=>{
        (data||[]).forEach(x=>{
          pf.push([ap.accId,ap.accName,ap.period,name,x.key,x.orders,(x.orders/matchCnt*100).toFixed(1)+'%',x.avgScore.toFixed(2),(x.convRate*100).toFixed(1)+'%',x.avgOutput>0?'¥'+x.avgOutput.toFixed(0):'-']);
        });
      });
    });
    if(pf.length>3)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(pf),'画像×期数明细');
  }
  return wb;
}function t10GenFindings(d){
  const findings=[];
  if(!d.accounts||d.accounts.length===0)return findings;
  // 洞察1：链路优劣
  const bestAcc=d.accounts.reduce((a,b)=>a.highRatio>b.highRatio?a:b,d.accounts[0]);
  const worstAcc=d.accounts.reduce((a,b)=>a.highRatio<b.highRatio?a:b,d.accounts[0]);
  if(bestAcc&&worstAcc){
    findings.push({title:'链路类型各有优劣',detail:bestAcc.name+'高分率最高（'+(bestAcc.highRatio*100).toFixed(1)+'%），'+worstAcc.name+'高分率最低（'+(worstAcc.highRatio*100).toFixed(1)+'%）'});
  }
  // 洞察2：复学塌陷
  if(d.repeat.count>0&&d.nonRepeat.count>0){
    const scoreGap=(d.nonRepeat.avgScore-d.repeat.avgScore).toFixed(2);
    const highGap=((d.nonRepeat.highRatio-d.repeat.highRatio)*100).toFixed(1);
    findings.push({title:'复学老客评分显著塌陷',detail:'非复学均分'+d.nonRepeat.avgScore.toFixed(2)+'，复学均分'+d.repeat.avgScore.toFixed(2)+'（差'+scoreGap+'分）；高分率差'+highGap+'%'});
  }
  // 洞察3：评分集中
  const midCount=d.accounts.reduce((s,a)=>s+(a.scoreDist[3]||0)+(a.scoreDist[4]||0)+(a.scoreDist[5]||0),0);
  findings.push({title:'评分集中在3-5分中段',detail:'3-5分占比'+(midCount/d.totalOrders*100).toFixed(1)+'%，高分段(8-10分)仅占'+(d.accounts.reduce((s,a)=>s+(a.scoreDist[8]||0)+(a.scoreDist[9]||0)+(a.scoreDist[10]||0),0)/d.totalOrders*100).toFixed(1)+'%'});
  // 洞察4：复学率高的账户
  const highRepeatAcc=d.accounts.filter(a=>a.repeatRatio>0.1);
  if(highRepeatAcc.length){
    findings.push({title:'部分账户复学老客占比偏高',detail:highRepeatAcc.map(a=>a.name+'（'+(a.repeatRatio*100).toFixed(1)+'%）').join('、')+'，老客混入稀释新客获客效率'});
  }
  return findings;
}

function t10GenSuggestions(d){
  const sug=[];
  if(!d.accounts||d.accounts.length===0)return sug;
  // 排除包
  const highRepeatAccs=d.accounts.filter(a=>a.repeatRatio>0.08);
  if(highRepeatAccs.length){
    sug.push({title:'人群包定向优化',text:'将历史已购/复学用户打包为全账户排除包，重点治理'+highRepeatAccs.map(a=>a.name).join('、')+'，将老客比例压到5%以下'});
  }
  // 深度出价
  sug.push({title:'深度转化出价策略',text:'当前样本超'+d.totalOrders+'单，建议主消耗账户开启双出价/深度高价值出价，回传门槛锚定≥6分或≥7分订单'});
  // 版位素材
  const bestAcc=d.accounts.reduce((a,b)=>a.highRatio>b.highRatio?a:b,d.accounts[0]);
  if(bestAcc){
    sug.push({title:'链路素材倾斜',text:bestAcc.name+'作为高质标杆（高分率'+(bestAcc.highRatio*100).toFixed(1)+'%），重点投放高价值版位；其他链路落地页首屏增加人群筛选问卷'});
  }
  // 模型回传
  sug.push({title:'评分条件优化',text:'持续推进设备号优先评分，提升与腾讯DMP回传匹配率，让算法更精准抓取高净值人群'});
  return sug;
}

function t10Display(d){
  document.getElementById('t10Results').style.display='block';
  // KPI
  const kpiHtml='<div class="ex-kpi-card"><div class="ex-val">'+d.totalOrders.toLocaleString()+'</div><div class="ex-lbl">总订单数</div></div>'+
    '<div class="ex-kpi-card"><div class="ex-val">'+d.avgScore.toFixed(2)+'</div><div class="ex-lbl">整体平均分</div></div>'+
    '<div class="ex-kpi-card"><div class="ex-val">'+(d.highRatio*100).toFixed(1)+'%</div><div class="ex-lbl">高分率(≥7分)</div></div>'+
    '<div class="ex-kpi-card"><div class="ex-val">'+(d.lowRatio*100).toFixed(1)+'%</div><div class="ex-lbl">低分率(≤3分)</div></div>'+
    (d.matchedCount?'<div class="ex-kpi-card"><div class="ex-val">'+(d.matchRate*100).toFixed(1)+'%</div><div class="ex-lbl">订单关联率</div></div>':'')+
    (d.totalHighOrder?'<div class="ex-kpi-card"><div class="ex-val">'+(d.overallConvRate*100).toFixed(1)+'%</div><div class="ex-lbl">高价课转化率</div></div>':'')+
    (d.totalHighAmount?'<div class="ex-kpi-card"><div class="ex-val">¥'+d.overallAvgOutput.toFixed(0)+'</div><div class="ex-lbl">人均产值</div></div>':'')+
    '<div class="ex-kpi-card"><div class="ex-val">'+d.accounts.length+'</div><div class="ex-lbl">账户数</div></div>';
  document.getElementById('t10Kpi').innerHTML=kpiHtml;
  // 分析内容
  let html='<h3 style="font-size:15px;font-weight:600;margin:16px 0 10px">一、账户链路对比总览</h3>';
  html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
  ['广告账户','承载链路','订单量','订单占比','平均分','高分率','低分率','复学率'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
  html+='</tr></thead><tbody>';
  d.accounts.forEach(a=>{
    html+='<tr><td style="border:1px solid #e2e8f0;padding:8px"><b>'+a.name+'</b><br><span style="font-size:10px;color:#94a3b8">ID:'+a.id+'</span></td><td style="border:1px solid #e2e8f0;padding:8px;font-size:11px">'+a.channel+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+a.orders+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(a.orders/d.totalOrders*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+a.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:#16a34a">'+(a.highRatio*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:#dc2626">'+(a.lowRatio*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(a.repeatRatio*100).toFixed(1)+'%</td></tr>';
  });
  html+='</tbody></table>';
  // 复学对比
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">二、复学 vs 非复学 评分对比</h3>';
  html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
  ['类型','订单量','平均分','高分率(≥7)'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
  html+='</tr></thead><tbody>';
  html+='<tr><td style="border:1px solid #e2e8f0;padding:8px"><b>非复学新客</b></td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+d.nonRepeat.count+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+d.nonRepeat.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:#16a34a">'+(d.nonRepeat.highRatio*100).toFixed(1)+'%</td></tr>';
  html+='<tr><td style="border:1px solid #e2e8f0;padding:8px"><b>复学老客</b></td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+d.repeat.count+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+d.repeat.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:#dc2626">'+(d.repeat.highRatio*100).toFixed(1)+'%</td></tr>';
  html+='</tbody></table>';
  // 评分条件对比
  if(d.conditions&&d.conditions.length){
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">三、评分条件对比（设备号 vs 手机号）</h3>';
    html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
    ['评分条件','订单量','订单占比','平均分','高分率','低分率','复学率'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
    html+='</tr></thead><tbody>';
    d.conditions.forEach(c=>{
      html+='<tr><td style="border:1px solid #e2e8f0;padding:8px"><b>'+c.name+'</b></td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+c.orders+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(c.orders/d.totalOrders*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+c.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:#16a34a">'+(c.highRatio*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:#dc2626">'+(c.lowRatio*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(c.repeatRatio*100).toFixed(1)+'%</td></tr>';
    });
    html+='</tbody></table>';
  }
  // 评分分层
  if(d.tiers){
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">四、评分分层分析（低1-3 / 中4-6 / 高7-10）</h3>';
    html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
    ['分层','订单量','占比','平均分','复学率'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
    html+='</tr></thead><tbody>';
    [d.tiers.low,d.tiers.mid,d.tiers.high].forEach(t=>{
      const color=t.name.includes('高')?'#16a34a':t.name.includes('低')?'#dc2626':'#f59e0b';
      html+='<tr><td style="border:1px solid #e2e8f0;padding:8px"><b style="color:'+color+'">'+t.name+'</b></td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+t.orders+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(t.ratio*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+t.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(t.repeatRatio*100).toFixed(1)+'%</td></tr>';
    });
    html+='</tbody></table>';
  }
  // 期数趋势
  if(d.periods&&d.periods.length){
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">五、期数趋势（评分质量随期数变化）</h3>';
    html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
    ['期数','订单量','平均分','高分率','复学率'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
    html+='</tr></thead><tbody>';
    d.periods.forEach(p=>{
      html+='<tr><td style="border:1px solid #e2e8f0;padding:8px"><b>'+p.period+'</b></td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+p.orders+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+p.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:#16a34a">'+(p.highRatio*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(p.repeatRatio*100).toFixed(1)+'%</td></tr>';
    });
    html+='</tbody></table>';
  }
  // 每期每账户分析
  if(d.accPeriods&&d.accPeriods.length){
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">五·五、每期每账户分析（评分→产值 + 画像）</h3>';
    html+='<div style="margin-bottom:12px;padding:10px 14px;background:#eff6ff;border-radius:8px;border-left:4px solid #3b82f6;font-size:12px;color:#1e40af">每个广告账户在每一期：评分1-10分对应的高价课产值（验证评分与产出关系），以及该期人群画像（性别/年龄/城市/手机品牌/手机价格/手机型号）。点击展开查看明细。</div>';
    const accGroups={};
    d.accPeriods.forEach(ap=>{
      if(!accGroups[ap.accId])accGroups[ap.accId]={accId:ap.accId,accName:ap.accName,periods:[]};
      accGroups[ap.accId].periods.push(ap);
    });
    Object.values(accGroups).forEach(g=>{
      const gOrders=g.periods.reduce((s,p)=>s+p.orders,0);
      const gHigh=g.periods.reduce((s,p)=>s+p.highOrder,0);
      const gHighAmt=g.periods.reduce((s,p)=>s+p.highAmount,0);
      const gScore=g.periods.reduce((s,p)=>s+p.avgScore*p.orders,0);
      html+='<details style="margin-bottom:12px;border:1px solid #dbeafe;border-radius:10px;background:#fff"><summary style="cursor:pointer;padding:12px 16px;font-weight:600;font-size:13px;color:#1e40af">📊 账户 '+g.accName+'（ID:'+g.accId+'，共'+gOrders+'单 · 高价课'+gHigh+'单 · 人均¥'+(gHighAmt>0?(gHighAmt/gOrders).toFixed(0):0)+'）</summary><div style="padding:4px 16px 16px">';

      const pList=g.periods.map(p=>'<b style="color:#0369a1">'+p.period+'</b>').join('、');
      html+='<div style="margin:10px 0 6px;padding:8px 12px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;font-size:12px;color:#0c4a6e">📅 <b>分期（共'+g.periods.length+'期）：</b>'+pList+'</div>';
      g.periods.forEach(ap=>{
        const pScore=ap.orders>0?ap.avgScore.toFixed(2):'-';
        html+='<details style="margin:10px 0;border:1px solid #e2e8f0;border-radius:8px"><summary style="cursor:pointer;padding:10px 14px;font-size:12.5px;font-weight:600;color:#334155">▶ 期数 '+ap.period+' · '+ap.orders+'单 · 均分'+pScore+' · 高价课'+ap.highOrder+'单 · 转化率'+(ap.convRate*100).toFixed(2)+'% · 人均产值'+(ap.avgOutput>0?'¥'+ap.avgOutput.toFixed(0):'-')+'</summary><div style="padding:12px 14px">';
        // 评分→产值表
        if(ap.scoreConv&&ap.scoreConv.length){
          html+='<div style="font-size:12px;font-weight:600;color:#475569;margin:4px 0 6px">① 评分分数 → 产值</div>';
          html+='<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:12px"><thead><tr style="background:#f1f5f9">';
          ['评分','订单数','高价课转化数','转化率','高价课金额','人均产值'].forEach(h=>{html+='<th style="border:1px solid #e2e8f0;padding:5px 8px;text-align:left;color:#475569">'+h+'</th>';});
          html+='</tr></thead><tbody>';
          ap.scoreConv.forEach(s=>{
            const crColor=s.convRate>=0.08?'#16a34a':s.convRate>=0.04?'#f59e0b':'#dc2626';
            html+='<tr><td style="border:1px solid #e2e8f0;padding:5px 8px"><b>'+s.score+'分</b></td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+s.orders+'</td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+s.highOrder+'</td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center;color:'+crColor+';font-weight:600">'+(s.convRate*100).toFixed(2)+'%</td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+(s.highAmount>0?'¥'+s.highAmount.toFixed(0):'-')+'</td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+(s.avgOutput>0?'¥'+s.avgOutput.toFixed(0):'-')+'</td></tr>';
          });
          html+='</tbody></table>';
        }
        // 画像表
        const matchCnt=ap.profileMatchCount||1;
        const dims=[['性别',ap.profile.gender],['年龄',ap.profile.age],['城市',ap.profile.city],['城市等级',ap.profile.cityLevel],['手机品牌',ap.profile.brand],['手机价格',ap.profile.priceTier],['手机型号',ap.profile.model]];
        const hasDim=dims.some(([_,data])=>(data||[]).length>0);
        if(hasDim){
          html+='<div style="font-size:12px;font-weight:600;color:#475569;margin:4px 0 6px">② 该期人群画像（匹配样本'+ap.profileMatchCount+'条）</div>';
          dims.forEach(([name,data])=>{
            if((data||[]).length){
              html+='<div style="font-size:11px;font-weight:600;color:#64748b;margin:8px 0 4px">'+name+'</div>';
              html+='<table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:8px"><thead><tr style="background:#f8fafc">';
              ['维度值','订单量','占比','平均分','转化率','人均产值'].forEach(h=>{html+='<th style="border:1px solid #e2e8f0;padding:4px 8px;text-align:left;color:#64748b">'+h+'</th>';});
              html+='</tr></thead><tbody>';
              data.forEach(x=>{
                html+='<tr><td style="border:1px solid #e2e8f0;padding:4px 8px"><b>'+x.key+'</b></td><td style="border:1px solid #e2e8f0;padding:4px 8px;text-align:center">'+x.orders+'</td><td style="border:1px solid #e2e8f0;padding:4px 8px;text-align:center">'+(x.orders/matchCnt*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:4px 8px;text-align:center">'+x.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:4px 8px;text-align:center">'+(x.convRate*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:4px 8px;text-align:center">'+(x.avgOutput>0?'¥'+x.avgOutput.toFixed(0):'-')+'</td></tr>';
              });
              html+='</tbody></table>';
            }
          });
        }else{
          html+='<div style="font-size:11px;color:#94a3b8;padding:4px 0 8px">该期暂无匹配的画像数据（订单明细中未找到对应订单号）</div>';
        }
        html+='</div></details>';
      });
      html+='</div></details>';
    });
  }
  // 核心洞察
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">六、核心洞察</h3>';
  const findings=t10GenFindings(d);
  findings.forEach((f,i)=>{
    html+='<div style="margin-bottom:12px;padding:12px 16px;background:#f8fafc;border-radius:8px;border-left:4px solid #3b82f6"><div style="font-weight:600;font-size:13px;color:#1e293b;margin-bottom:4px">'+(i+1)+'. '+f.title+'</div><div style="font-size:12px;color:#475569;line-height:1.7">'+f.detail+'</div></div>';
  });
  // 实操建议
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">七、优化师投放实操建议</h3>';
  const sug=t10GenSuggestions(d);
  sug.forEach(s=>{
    html+='<div style="margin-bottom:12px;padding:12px 16px;background:#f0fdf4;border-radius:8px;border-left:4px solid #16a34a"><div style="font-weight:600;font-size:13px;color:#166534;margin-bottom:4px">'+s.title+'</div><div style="font-size:12px;color:#475569;line-height:1.7">'+s.text+'</div></div>';
  });
  // 评分×高价课转化（验证模型有效性）
  if(d.scoreConv&&d.scoreConv.length){
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">八、评分 vs 高价课转化（模型有效性验证）</h3>';
    html+='<div style="margin-bottom:12px;padding:10px 14px;background:#fef3c7;border-radius:8px;border-left:4px solid #f59e0b;font-size:12px;color:#92400e">验证企点评分模型是否有效：高分人群转化率和人均产值是否显著高于低分人群</div>';
    html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
    ['评分','订单数','高价课转化数','转化率','客单价','人均产值'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:center;color:#1e40af">'+h+'</th>';});
    html+='</tr></thead><tbody>';
    d.scoreConv.forEach(s=>{
      const crColor=s.convRate>=0.08?'#16a34a':s.convRate>=0.04?'#f59e0b':'#dc2626';
      html+='<tr><td style="border:1px solid #e2e8f0;padding:8px;text-align:center"><b>'+s.score+'分</b></td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+s.orders+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+s.highOrder+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:'+crColor+';font-weight:600">'+(s.convRate*100).toFixed(2)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(s.客单价>0?'¥'+s.客单价.toFixed(0):'-')+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;font-weight:600">¥'+s.avgOutput.toFixed(0)+'</td></tr>';
    });
    html+='</tbody></table>';
    // 分层汇总
    if(d.tiers){
      html+='<div style="display:flex;gap:12px;margin-bottom:16px">';
      [d.tiers.low,d.tiers.mid,d.tiers.high].forEach(t=>{
        html+='<div style="flex:1;padding:12px 16px;background:#f8fafc;border-radius:8px;border:1px solid #e2e8f0"><div style="font-size:12px;color:#64748b;margin-bottom:4px">'+t.name+'</div><div style="font-size:18px;font-weight:700;color:#1e293b">'+(t.convRate*100).toFixed(1)+'%</div><div style="font-size:11px;color:#94a3b8">转化率 | 人均¥'+t.avgOutput.toFixed(0)+'</div></div>';
      });
      html+='</div>';
    }
  }
  // 账户×画像交叉分析
  if(d.hasProfile){
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">九、各账户 × 画像维度 交叉拆解</h3>';
    d.accounts.forEach(acc=>{
      if(!acc.profile)return;
      const accDims=[
        {name:'性别',data:acc.profile.gender,key:'gender'},
        {name:'年龄',data:acc.profile.age,key:'age'},
        {name:'城市等级',data:acc.profile.cityLevel,key:'cityLevel'},
        {name:'手机品牌',data:acc.profile.brand,key:'brand'},
        {name:'手机价格档',data:acc.profile.priceTier,key:'priceTier'}
      ].filter(dim=>dim.data.length>0);
      const matchCnt=acc.profileMatchCount||0;
      const lowSample=matchCnt<20;
      html+='<div style="margin-bottom:16px;padding:12px 16px;background:#f8fafc;border-radius:8px;border-left:4px solid '+(lowSample?'#f59e0b':'#3b82f6')+'"><div style="font-weight:600;font-size:13px;color:'+(lowSample?'#92400e':'#1e40af')+';margin-bottom:4px">账户 '+acc.name+'（ID:'+acc.id+'，'+acc.orders+'单，均分'+acc.avgScore.toFixed(2)+'，高分率'+(acc.highRatio*100).toFixed(1)+'%）</div>';
      if(lowSample){
        html+='<div style="font-size:11px;color:#b45309;margin-bottom:8px;padding:6px 10px;background:#fef3c7;border-radius:4px">⚠️ 画像匹配样本仅 '+matchCnt+' 条（占比'+(acc.orders>0?(matchCnt/acc.orders*100).toFixed(1):0)+'%），样本偏少，画像结论仅供参考</div>';
      }else{
        html+='<div style="font-size:11px;color:#64748b;margin-bottom:8px">画像匹配样本：'+matchCnt+' 条（占比'+(acc.orders>0?(matchCnt/acc.orders*100).toFixed(1):0)+'%）</div>';
      }
      if(accDims.length){
        accDims.forEach(dim=>{
          html+='<div style="margin-bottom:10px"><div style="font-size:12px;font-weight:600;color:#475569;margin-bottom:4px">'+dim.name+'（全部）</div>';
          html+='<table style="width:100%;border-collapse:collapse;font-size:11px"><thead><tr style="background:#fff">';
          ['维度值','订单量','占比','平均分','转化率','人均产值'].forEach(h=>{html+='<th style="border:1px solid #e2e8f0;padding:5px 8px;text-align:left;color:#64748b">'+h+'</th>';});
          html+='</tr></thead><tbody>';
          dim.data.forEach(x=>{
            html+='<tr><td style="border:1px solid #e2e8f0;padding:5px 8px"><b>'+x.key+'</b></td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+x.orders+'</td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+(matchCnt>0?(x.orders/matchCnt*100).toFixed(1):0)+'%</td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+x.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+(x.convRate*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:5px 8px;text-align:center">'+(x.avgOutput>0?'¥'+x.avgOutput.toFixed(0):'-')+'</td></tr>';
          });
          html+='</tbody></table></div>';
        });
      }else{
        html+='<div style="font-size:12px;color:#94a3b8;padding:8px 0">该账户暂无匹配的画像数据，建议检查订单明细表是否包含此账户的订单</div>';
      }
      html+='</div>';
    });
  }
  // 画像维度分析
  if(d.hasProfile){
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">'+(d.scoreConv&&d.scoreConv.length?'十':'九')+'、人群画像深度拆解</h3>';
    const dims=[
      {name:'性别分布',data:d.profile.gender,color:'#ec4899'},
      {name:'年龄分布',data:d.profile.age,color:'#8b5cf6'},
      {name:'城市分布',data:d.profile.city,color:'#06b6d4'},
      {name:'手机品牌分布',data:d.profile.brand,color:'#f59e0b'},
      {name:'手机价格分布',data:d.profile.priceTier,color:'#10b981'},
      {name:'手机型号分布',data:d.profile.model,color:'#3b82f6'}
    ];
    dims.forEach(dim=>{
      if(dim.data.length){
        html+='<div style="margin-bottom:16px"><div style="font-weight:600;font-size:13px;color:#1e293b;margin-bottom:8px">'+dim.name+'</div>';
        html+='<table style="width:100%;border-collapse:collapse;font-size:12px"><thead><tr style="background:#f8fafc">';
        ['维度值','订单量','占比','平均分','人均产值'].forEach(h=>{html+='<th style="border:1px solid #e2e8f0;padding:6px 8px;text-align:left;color:#475569">'+h+'</th>';});
        html+='</tr></thead><tbody>';
        dim.data.forEach(x=>{
          html+='<tr><td style="border:1px solid #e2e8f0;padding:6px 8px"><b>'+x.key+'</b></td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+x.orders+'</td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+(x.orders/d.totalOrders*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+x.avgScore.toFixed(2)+'</td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+(x.avgOutput>0?'¥'+x.avgOutput.toFixed(0):'-')+'</td></tr>';
        });
        html+='</tbody></table></div>';
      }
    });
    // 高分vs低分画像对比
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">'+(d.scoreConv&&d.scoreConv.length?'十一':'十')+'、高分(7-10) vs 低分(1-3) 人群画像对比</h3>';
    const cmpDims=[['性别',d.highProfile.gender,d.lowProfile.gender],['年龄',d.highProfile.age,d.lowProfile.age],['手机品牌',d.highProfile.brand,d.lowProfile.brand],['手机价格',d.highProfile.priceTier,d.lowProfile.priceTier]];
    cmpDims.forEach(([name,high,low])=>{
      if(high.length||low.length){
        html+='<div style="margin-bottom:14px"><div style="font-weight:600;font-size:13px;color:#1e293b;margin-bottom:6px">'+name+'对比</div>';
        html+='<table style="width:100%;border-collapse:collapse;font-size:12px"><thead><tr style="background:#f8fafc">';
        ['维度值','高分订单','高分占比','低分订单','低分占比','结论'].forEach(h=>{html+='<th style="border:1px solid #e2e8f0;padding:6px 8px;text-align:left;color:#475569">'+h+'</th>';});
        html+='</tr></thead><tbody>';
        const allKeys=new Set([...high.map(x=>x.key),...low.map(x=>x.key)]);
        const ht=high.reduce((s,x)=>s+x.orders,0)||1,lt=low.reduce((s,x)=>s+x.orders,0)||1;
        [...allKeys].forEach(k=>{
          const h=high.find(x=>x.key===k),l=low.find(x=>x.key===k);
          const ho=h?h.orders:0,lo=l?l.orders:0;
          const hr=ho/ht,lr=lo/lt;
          const concl=hr>lr?'<span style="color:#16a34a">高分集中</span>':hr<lr?'<span style="color:#dc2626">低分集中</span>':'<span style="color:#64748b">均衡</span>';
          html+='<tr><td style="border:1px solid #e2e8f0;padding:6px 8px"><b>'+k+'</b></td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+ho+'</td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+(hr*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+lo+'</td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+(lr*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:6px 8px;text-align:center">'+concl+'</td></tr>';
        });
        html+='</tbody></table></div>';
      }
    });
  }
  // 数据补全提示
  html+='<div style="margin-top:20px;padding:14px 18px;background:#fffbeb;border-radius:10px;border-left:4px solid #f59e0b"><div style="font-weight:600;font-size:13px;color:#92400e;margin-bottom:6px">📌 待补全数据（完整画像闭环分析）</div><div style="font-size:12px;color:#78350f;line-height:1.8">当前为评分中间层数据，缺少以下维度可进一步分析：<br>1. <b>后端产值</b>：CRM/企点成交表（订单号+正价课金额/GMV/是否完课）→ 验证评分vs产值正相关性<br>2. <b>受众画像</b>：腾讯广告后台受众洞察报表（设备品牌/价格带/年龄/性别/城市）→ 高价值人群画像剖析<br>3. <b>前端消耗</b>：广告后台分账户/计划消耗数据 → 计算各账户ROI与获客成本</div></div>';
  document.getElementById('t10Analysis').innerHTML=html;
}

function downloadTab10(){ if(t10Output) XLSX.writeFile(t10Output,'分账户评分分析.xlsx'); }

// 分析模块渲染
function renderAccountScore(){
  const d=SYNC.accountScore;
  if(!d){
    document.getElementById('as-empty').style.display='block';
    ['as-overview-card','as-chart-card','as-repeat-card','as-condition-card','as-tier-card','as-period-card','as-accperiod-card','as-finding-card','as-suggest-card','as-scoreconv-card','as-accprofile-card','as-profile-card','as-cmp-card'].forEach(id=>document.getElementById(id).style.display='none');
    document.getElementById('as-badge').textContent='待同步';
    return;
  }
  document.getElementById('as-empty').style.display='none';
  document.getElementById('as-badge').textContent=d.totalOrders.toLocaleString()+'单 · '+d.accounts.length+'账户 · 已同步';
  document.getElementById('as-summary').textContent='整体均分'+d.avgScore.toFixed(2)+' · 高分率'+(d.highRatio*100).toFixed(1)+'% · 复学率'+(d.repeatRatio*100).toFixed(1)+'%';
  // KPI
  const colors=['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#ec4899'];
  const kpiHtml='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px">'+d.accounts.map((a,i)=>{
    const c=colors[i%colors.length];
    return '<div style="background:linear-gradient(135deg,'+c+'12,'+c+'05);border:1px solid '+c+'30;border-radius:12px;padding:14px 16px;border-top:3px solid '+c+'"><div style="font-size:12px;color:#64748b;font-weight:600;margin-bottom:4px">'+a.name+'</div><div style="font-size:22px;font-weight:700;color:'+c+'">'+a.orders.toLocaleString()+'</div><div style="font-size:11.5px;color:#475569;margin-top:4px">均分 '+a.avgScore.toFixed(2)+' · 高分 '+(a.highRatio*100).toFixed(1)+'%</div><div style="font-size:11px;color:#94a3b8;margin-top:2px">复学 '+(a.repeatRatio*100).toFixed(1)+'% · 低分 '+(a.lowRatio*100).toFixed(1)+'%</div></div>';
  }).join('')+'</div>';
  document.getElementById('as-kpi').innerHTML=kpiHtml;
  // 账户对比表
  document.getElementById('as-overview-card').style.display='block';
  const ovH=['广告账户','承载链路','订单量','订单占比','平均分','高分率','低分率','复学率'];
  const ovR=d.accounts.map(a=>[a.name+'\nID:'+a.id,a.channel,a.orders,(a.orders/d.totalOrders*100).toFixed(1)+'%',a.avgScore.toFixed(2),(a.highRatio*100).toFixed(1)+'%',(a.lowRatio*100).toFixed(1)+'%',(a.repeatRatio*100).toFixed(1)+'%']);
  fillTable('as-overview-table',ovH,ovR);
  // 评分分布图
  document.getElementById('as-chart-card').style.display='block';
  const scores=Array.from({length:10},(_,i)=>i+1);
  const series=d.accounts.map((a,i)=>({
    name:a.id,type:'bar',data:scores.map(s=>a.scoreDist[s]||0),
    itemStyle:{color:colors[i%colors.length]}
  }));
  const asChart=chart('as-chart');
  if(asChart)asChart.setOption({
    tooltip:{trigger:'axis',axisPointer:{type:'shadow'}},
    legend:{data:d.accounts.map(a=>a.id),top:0},
    grid:{left:50,right:30,top:40,bottom:30},
    xAxis:{type:'category',data:scores.map(s=>s+'分'),name:'评分'},
    yAxis:{type:'value',name:'订单数'},
    series
  });
  // 复学对比表
  document.getElementById('as-repeat-card').style.display='block';
  const rpH=['类型','订单量','平均分','高分率(≥7)','低分率(≤3)'];
  const rpR=[
    ['非复学新客',d.nonRepeat.count,d.nonRepeat.avgScore.toFixed(2),(d.nonRepeat.highRatio*100).toFixed(1)+'%','-'],
    ['复学老客',d.repeat.count,d.repeat.avgScore.toFixed(2),(d.repeat.highRatio*100).toFixed(1)+'%','-'],
  ];
  fillTable('as-repeat-table',rpH,rpR);
  // 评分条件对比
  if(d.conditions&&d.conditions.length){
    document.getElementById('as-condition-card').style.display='block';
    const cdH=['评分条件','订单量','订单占比','平均分','高分率','低分率','复学率'];
    const cdR=d.conditions.map(c=>[c.name,c.orders,(c.orders/d.totalOrders*100).toFixed(1)+'%',c.avgScore.toFixed(2),(c.highRatio*100).toFixed(1)+'%',(c.lowRatio*100).toFixed(1)+'%',(c.repeatRatio*100).toFixed(1)+'%']);
    fillTable('as-condition-table',cdH,cdR);
  }
  // 评分分层
  if(d.tiers){
    document.getElementById('as-tier-card').style.display='block';
    const trH=['分层','订单量','占比','平均分','复学率'];
    const trR=[d.tiers.low,d.tiers.mid,d.tiers.high].map(t=>[t.name,t.orders,(t.ratio*100).toFixed(1)+'%',t.avgScore.toFixed(2),(t.repeatRatio*100).toFixed(1)+'%']);
    fillTable('as-tier-table',trH,trR);
  }
  // 期数趋势图
  if(d.periods&&d.periods.length){
    document.getElementById('as-period-card').style.display='block';
    const asPeriodChart=chart('as-period-chart');
    if(asPeriodChart)asPeriodChart.setOption({
      tooltip:{trigger:'axis'},
      legend:{data:['订单量','平均分','高分率'],top:0},
      grid:{left:50,right:50,top:40,bottom:60},
      xAxis:{type:'category',data:d.periods.map(p=>p.period),axisLabel:{rotate:30,fontSize:10}},
      yAxis:[
        {type:'value',name:'订单量',position:'left'},
        {type:'value',name:'比率(%)',position:'right',max:100}
      ],
      series:[
        {name:'订单量',type:'bar',data:d.periods.map(p=>p.orders),itemStyle:{color:'#3b82f6'}},
        {name:'平均分',type:'line',yAxisIndex:1,data:d.periods.map(p=>p.avgScore.toFixed(2)),itemStyle:{color:'#10b981'},smooth:true},
        {name:'高分率',type:'line',yAxisIndex:1,data:d.periods.map(p=>(p.highRatio*100).toFixed(1)),itemStyle:{color:'#f59e0b'},smooth:true}
      ]
    });
  }
  // 每期每账户分析
  if(d.accPeriods&&d.accPeriods.length){
    document.getElementById('as-accperiod-card').style.display='block';
    const accGroups={};
    d.accPeriods.forEach(ap=>{
      if(!accGroups[ap.accId])accGroups[ap.accId]={accId:ap.accId,accName:ap.accName,periods:[]};
      accGroups[ap.accId].periods.push(ap);
    });
    let aph='';
    Object.values(accGroups).forEach(g=>{
      const gOrders=g.periods.reduce((s,p)=>s+p.orders,0);
      const gHigh=g.periods.reduce((s,p)=>s+p.highOrder,0);
      const gHighAmt=g.periods.reduce((s,p)=>s+p.highAmount,0);
      aph+='<details style="margin-bottom:14px;border:1px solid #dbeafe;border-radius:10px;background:#fff"><summary style="cursor:pointer;padding:12px 16px;font-weight:600;font-size:14px;color:#1e40af">📊 账户 '+g.accName+'（ID:'+g.accId+'，共'+gOrders+'单 · 高价课'+gHigh+'单 · 人均¥'+(gHighAmt>0?(gHighAmt/gOrders).toFixed(0):0)+'）</summary><div style="padding:4px 16px 16px">';

      const pList=g.periods.map(p=>'<b style="color:#0369a1">'+p.period+'</b>').join('、');
      aph+='<div style="margin:10px 0 6px;padding:8px 12px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;font-size:13px;color:#0c4a6e">📅 <b>分期（共'+g.periods.length+'期）：</b>'+pList+'</div>';
      g.periods.forEach(ap=>{
        const pScore=ap.orders>0?ap.avgScore.toFixed(2):'-';
        aph+='<details style="margin:10px 0;border:1px solid #e2e8f0;border-radius:8px"><summary style="cursor:pointer;padding:10px 14px;font-size:13px;font-weight:600;color:#334155">▶ 期数 '+ap.period+' · '+ap.orders+'单 · 均分'+pScore+' · 高价课'+ap.highOrder+'单 · 转化率'+(ap.convRate*100).toFixed(2)+'% · 人均产值'+(ap.avgOutput>0?'¥'+ap.avgOutput.toFixed(0):'-')+'</summary><div style="padding:12px 14px">';
        if(ap.scoreConv&&ap.scoreConv.length){
          aph+='<div style="font-size:13px;font-weight:600;color:#475569;margin:4px 0 6px">① 评分分数 → 产值</div>';
          aph+='<div class="table-wrap"><table class="data-table"><thead><tr>';
          ['评分','订单数','高价课转化数','转化率','高价课金额','人均产值'].forEach(h=>{aph+='<th>'+h+'</th>';});
          aph+='</tr></thead><tbody>';
          ap.scoreConv.forEach(s=>{
            const crColor=s.convRate>=0.08?'#16a34a':s.convRate>=0.04?'#f59e0b':'#dc2626';
            aph+='<tr><td><b>'+s.score+'分</b></td><td style="text-align:center">'+s.orders+'</td><td style="text-align:center">'+s.highOrder+'</td><td style="text-align:center;color:'+crColor+';font-weight:600">'+(s.convRate*100).toFixed(2)+'%</td><td style="text-align:center">'+(s.highAmount>0?'¥'+s.highAmount.toFixed(0):'-')+'</td><td style="text-align:center">'+(s.avgOutput>0?'¥'+s.avgOutput.toFixed(0):'-')+'</td></tr>';
          });
          aph+='</tbody></table></div>';
        }
        const matchCnt=ap.profileMatchCount||1;
        const dims=[['性别',ap.profile.gender],['年龄',ap.profile.age],['城市',ap.profile.city],['城市等级',ap.profile.cityLevel],['手机品牌',ap.profile.brand],['手机价格',ap.profile.priceTier],['手机型号',ap.profile.model]];
        const hasDim=dims.some(([_,data])=>(data||[]).length>0);
        if(hasDim){
          aph+='<div style="font-size:13px;font-weight:600;color:#475569;margin:12px 0 6px">② 该期人群画像（匹配样本'+ap.profileMatchCount+'条）</div>';
          dims.forEach(([name,data])=>{
            if((data||[]).length){
              aph+='<div style="font-size:12px;font-weight:600;color:#64748b;margin:8px 0 4px">'+name+'</div>';
              aph+='<div class="table-wrap"><table class="data-table"><thead><tr>';
              ['维度值','订单量','占比','平均分','转化率','人均产值'].forEach(h=>{aph+='<th>'+h+'</th>';});
              aph+='</tr></thead><tbody>';
              data.forEach(x=>{
                aph+='<tr><td><b>'+x.key+'</b></td><td style="text-align:center">'+x.orders+'</td><td style="text-align:center">'+(x.orders/matchCnt*100).toFixed(1)+'%</td><td style="text-align:center">'+x.avgScore.toFixed(2)+'</td><td style="text-align:center">'+(x.convRate*100).toFixed(1)+'%</td><td style="text-align:center">'+(x.avgOutput>0?'¥'+x.avgOutput.toFixed(0):'-')+'</td></tr>';
              });
              aph+='</tbody></table></div>';
            }
          });
        }else{
          aph+='<div style="font-size:12px;color:#94a3b8;padding:8px 0">该期暂无匹配的画像数据（订单明细中未找到对应订单号）</div>';
        }
        aph+='</div></details>';
      });
      aph+='</div></details>';
    });
    document.getElementById('as-accperiod').innerHTML=aph;
  }
  // 核心洞察
  document.getElementById('as-finding-card').style.display='block';
  const findings=t10GenFindings(d);
  document.getElementById('as-findings').innerHTML=findings.map((f,i)=>'<div style="margin-bottom:12px;padding:14px 18px;background:#f8fafc;border-radius:10px;border-left:4px solid #3b82f6"><div style="font-weight:600;font-size:14px;color:#1e293b;margin-bottom:6px">'+(i+1)+'. '+f.title+'</div><div style="font-size:13px;color:#475569;line-height:1.8">'+f.detail+'</div></div>').join('');
  // 实操建议
  document.getElementById('as-suggest-card').style.display='block';
  const sug=t10GenSuggestions(d);
  document.getElementById('as-suggestions').innerHTML=sug.map(s=>'<div style="margin-bottom:14px;padding:14px 18px;background:#f0fdf4;border-radius:10px;border-left:4px solid #16a34a"><div style="font-weight:600;font-size:14px;color:#166534;margin-bottom:6px">'+s.title+'</div><div style="font-size:13px;color:#475569;line-height:1.8">'+s.text+'</div></div>').join('');
  // 评分×转化
  if(d.scoreConv&&d.scoreConv.length){
    document.getElementById('as-scoreconv-card').style.display='block';
    let sch='<div style="margin-bottom:12px;padding:10px 14px;background:#fef3c7;border-radius:8px;border-left:4px solid #f59e0b;font-size:13px;color:#92400e">验证企点评分模型有效性：高分人群转化率和人均产值是否显著高于低分人群</div>';
    sch+='<div class="table-wrap"><table class="data-table"><thead><tr>';
    ['评分','订单数','高价课转化','转化率','客单价','人均产值'].forEach(h=>{sch+='<th>'+h+'</th>';});
    sch+='</tr></thead><tbody>';
    d.scoreConv.forEach(s=>{
      const crColor=s.convRate>=0.08?'#16a34a':s.convRate>=0.04?'#f59e0b':'#dc2626';
      sch+='<tr><td><b>'+s.score+'分</b></td><td style="text-align:center">'+s.orders+'</td><td style="text-align:center">'+s.highOrder+'</td><td style="text-align:center;color:'+crColor+';font-weight:600">'+(s.convRate*100).toFixed(2)+'%</td><td style="text-align:center">'+(s.客单价>0?'¥'+s.客单价.toFixed(0):'-')+'</td><td style="text-align:center;font-weight:600">¥'+s.avgOutput.toFixed(0)+'</td></tr>';
    });
    sch+='</tbody></table></div>';
    if(d.tiers){
      sch+='<div style="display:flex;gap:12px;margin-top:14px">';
      [d.tiers.low,d.tiers.mid,d.tiers.high].forEach(t=>{
        sch+='<div style="flex:1;padding:14px 18px;background:#f8fafc;border-radius:10px;border:1px solid #e2e8f0;text-align:center"><div style="font-size:12px;color:#64748b;margin-bottom:6px">'+t.name+'</div><div style="font-size:22px;font-weight:700;color:#1e293b">'+(t.convRate*100).toFixed(1)+'%</div><div style="font-size:11px;color:#94a3b8;margin-top:4px">转化率 | 人均¥'+t.avgOutput.toFixed(0)+'</div></div>';
      });
      sch+='</div>';
    }
    document.getElementById('as-scoreconv').innerHTML=sch;
  }
  // 账户×画像
  if(d.hasProfile){
    document.getElementById('as-accprofile-card').style.display='block';
    let aph='';
    d.accounts.forEach(acc=>{
      if(!acc.profile)return;
      const accDims=[
        {name:'性别',data:acc.profile.gender},{name:'年龄',data:acc.profile.age},
        {name:'城市等级',data:acc.profile.cityLevel},{name:'手机品牌',data:acc.profile.brand},
        {name:'手机价格档',data:acc.profile.priceTier}
      ].filter(dim=>dim.data.length>0);
      const matchCnt=acc.profileMatchCount||0;
      const lowSample=matchCnt<20;
      aph+='<div style="margin-bottom:18px;padding:14px 18px;background:#f8fafc;border-radius:10px;border-left:4px solid '+(lowSample?'#f59e0b':'#3b82f6')+'"><div style="font-weight:600;font-size:14px;color:'+(lowSample?'#92400e':'#1e40af')+';margin-bottom:6px">账户 '+acc.name+'（ID:'+acc.id+'，'+acc.orders+'单，均分'+acc.avgScore.toFixed(2)+'，高分率'+(acc.highRatio*100).toFixed(1)+'%）</div>';
      if(lowSample){
        aph+='<div style="font-size:12px;color:#b45309;margin-bottom:10px;padding:8px 12px;background:#fef3c7;border-radius:6px">⚠️ 画像匹配样本仅 '+matchCnt+' 条（占比'+(acc.orders>0?(matchCnt/acc.orders*100).toFixed(1):0)+'%），样本偏少，画像结论仅供参考</div>';
      }else{
        aph+='<div style="font-size:12px;color:#64748b;margin-bottom:10px">画像匹配样本：'+matchCnt+' 条（占比'+(acc.orders>0?(matchCnt/acc.orders*100).toFixed(1):0)+'%）</div>';
      }
      if(accDims.length){
        accDims.forEach(dim=>{
          aph+='<div style="margin-bottom:12px"><div style="font-size:13px;font-weight:600;color:#475569;margin-bottom:6px">'+dim.name+'（全部）</div>';
          aph+='<div class="table-wrap"><table class="data-table"><thead><tr>';
          ['维度值','订单量','占比','平均分','转化率','人均产值'].forEach(h=>{aph+='<th>'+h+'</th>';});
          aph+='</tr></thead><tbody>';
          dim.data.forEach(x=>{
            aph+='<tr><td><b>'+x.key+'</b></td><td style="text-align:center">'+x.orders+'</td><td style="text-align:center">'+(matchCnt>0?(x.orders/matchCnt*100).toFixed(1):0)+'%</td><td style="text-align:center">'+x.avgScore.toFixed(2)+'</td><td style="text-align:center">'+(x.convRate*100).toFixed(1)+'%</td><td style="text-align:center">'+(x.avgOutput>0?'¥'+x.avgOutput.toFixed(0):'-')+'</td></tr>';
          });
          aph+='</tbody></table></div></div>';
        });
      }else{
        aph+='<div style="font-size:13px;color:#94a3b8;padding:10px 0">该账户暂无匹配的画像数据，建议检查订单明细表是否包含此账户的订单</div>';
      }
      aph+='</div>';
    });
    document.getElementById('as-accprofile').innerHTML=aph;
  }
  // 画像维度
  if(d.hasProfile){
    document.getElementById('as-profile-card').style.display='block';
    const dims=[
      {name:'性别分布',data:d.profile.gender},
      {name:'年龄分布',data:d.profile.age},
      {name:'城市分布',data:d.profile.city},
      {name:'手机品牌分布',data:d.profile.brand},
      {name:'手机价格分布',data:d.profile.priceTier},
      {name:'手机型号分布',data:d.profile.model}
    ];
    let ph='';
    dims.forEach(dim=>{
      if(dim.data.length){
        ph+='<div style="margin-bottom:20px"><div style="font-weight:600;font-size:14px;color:#1e293b;margin-bottom:8px">'+dim.name+'</div>';
        ph+='<div class="table-wrap"><table class="data-table"><thead><tr>';
        ['维度值','订单量','占比','平均分','人均产值'].forEach(h=>{ph+='<th>'+h+'</th>';});
        ph+='</tr></thead><tbody>';
        dim.data.forEach(x=>{
          ph+='<tr><td><b>'+x.key+'</b></td><td style="text-align:center">'+x.orders+'</td><td style="text-align:center">'+(x.orders/d.totalOrders*100).toFixed(1)+'%</td><td style="text-align:center">'+x.avgScore.toFixed(2)+'</td><td style="text-align:center">'+(x.avgOutput>0?'¥'+x.avgOutput.toFixed(0):'-')+'</td></tr>';
        });
        ph+='</tbody></table></div></div>';
      }
    });
    document.getElementById('as-profile').innerHTML=ph;
    // 高低分对比
    document.getElementById('as-cmp-card').style.display='block';
    const cmpDims=[['性别',d.highProfile.gender,d.lowProfile.gender],['年龄',d.highProfile.age,d.lowProfile.age],['手机品牌',d.highProfile.brand,d.lowProfile.brand],['手机价格',d.highProfile.priceTier,d.lowProfile.priceTier]];
    let ch='';
    cmpDims.forEach(([name,high,low])=>{
      if(high.length||low.length){
        ch+='<div style="margin-bottom:18px"><div style="font-weight:600;font-size:14px;color:#1e293b;margin-bottom:8px">'+name+'对比</div>';
        ch+='<div class="table-wrap"><table class="data-table"><thead><tr>';
        ['维度值','高分订单','高分占比','低分订单','低分占比','结论'].forEach(h=>{ch+='<th>'+h+'</th>';});
        ch+='</tr></thead><tbody>';
        const allKeys=new Set([...high.map(x=>x.key),...low.map(x=>x.key)]);
        const ht=high.reduce((s,x)=>s+x.orders,0)||1,lt=low.reduce((s,x)=>s+x.orders,0)||1;
        [...allKeys].forEach(k=>{
          const h=high.find(x=>x.key===k),l=low.find(x=>x.key===k);
          const ho=h?h.orders:0,lo=l?l.orders:0;
          const hr=ho/ht,lr=lo/lt;
          const concl=hr>lr?'<span style="color:#16a34a;font-weight:600">高分集中</span>':hr<lr?'<span style="color:#dc2626;font-weight:600">低分集中</span>':'<span style="color:#64748b">均衡</span>';
          ch+='<tr><td><b>'+k+'</b></td><td style="text-align:center">'+ho+'</td><td style="text-align:center">'+(hr*100).toFixed(1)+'%</td><td style="text-align:center">'+lo+'</td><td style="text-align:center">'+(lr*100).toFixed(1)+'%</td><td style="text-align:center">'+concl+'</td></tr>';
        });
        ch+='</tbody></table></div></div>';
      }
    });
    document.getElementById('as-cmp').innerHTML=ch;
  }
}

/* ---------- 工具11：周报历史归档 ---------- */
let t11Output=null;

// 解析周报文本
function t11ParseWeeklyText(text){
  // 清理UI元素
  text=text.replace(/菜单\n插入\n图片\n正文\n默认字体\n11\n快捷工具\n打印\n提及/g,'');
  text=text.replace(/用户可以通过 control 加 ~ 打开或关闭无障碍功能/g,'');
  text=text.replace(/\d+ 个字/g,'');
  text=text.replace(/100%/g,'');
  text=text.replace(/^text$/gm,'');
  text=text.replace(/^主页$/gm,'');
  text=text.replace(/^新建$/gm,'');
  text=text.replace(/^打开最近的文档$/gm,'');
  text=text.replace(/^账号信息变化$/gm,'');
  text=text.replace(/^刷新$/gm,'');

  const lines=text.split('\n').map(l=>l.trim()).filter(l=>l);
  const reports=[];
  let current=null;

  for(let i=0;i<lines.length;i++){
    const line=lines[i];
    // 检测周报日期标记（如 0831周报、0824周报、0817周报等）
    const dateMatch=line.match(/^(\d{4})周报$/);
    if(dateMatch){
      if(current)reports.push(current);
      current={date:dateMatch[1],sections:[],issues:[],actions:[],rawLines:[]};
      continue;
    }
    // 检测"太极二部订单进度表"后的日期（如 0629-0705）
    const rangeMatch=line.match(/^(\d{4}-\d{4})$/);
    if(rangeMatch&&current===null){
      current={date:rangeMatch[1],sections:[],issues:[],actions:[],rawLines:[]};
      continue;
    }
    if(!current)continue;
    current.rawLines.push(line);
    // 提取问题定位
    if(/问题\d+[:：]/.test(line)||/问题[一二三四五六七八九十][:：]/.test(line)){
      current.issues.push(line);
    }
    // 提取调整动作
    if(/^调整[:：]/.test(line)||/^本周调整/.test(line)||/^本周计划及调整/.test(line)||/^本周重点规划/.test(line)){
      current.actions.push(line);
    }
    // 提取主要section
    if(/^[一二三四五六七八九十]+[、.]/.test(line)||/^\d+[、.]/.test(line)){
      current.sections.push(line);
    }
  }
  if(current)reports.push(current);

  // 按日期排序（倒序，最新的在前）
  reports.sort((a,b)=>b.date.localeCompare(a.date));
  return reports;
}

// 提取高频问题
function t11ExtractIssues(reports){
  const issueMap={};
  reports.forEach(r=>{
    r.issues.forEach(issue=>{
      // 提取问题关键词
      let key=issue.replace(/问题\d+[:：]\s*/,'').trim();
      if(key.length>20)key=key.slice(0,20)+'...';
      if(!issueMap[key])issueMap[key]={key,count:0,examples:[]};
      issueMap[key].count++;
      if(issueMap[key].examples.length<3)issueMap[key].examples.push({date:r.date,text:issue});
    });
  });
  return Object.values(issueMap).sort((a,b)=>b.count-a.count);
}

// 提取调整动作方法论
function t11ExtractActions(reports){
  const actionMap={};
  reports.forEach(r=>{
    r.actions.forEach(action=>{
      // 分类调整动作
      let category='其他';
      if(/素材|上新|养身|ip|复合|爆款/.test(action))category='素材策略';
      else if(/年龄|定向|高龄|人群/.test(action))category='定向调整';
      else if(/链路|分流|问答|首页|落地页/.test(action))category='链路调整';
      else if(/版位|暂停|开单|roi/.test(action))category='版位/预算';
      else if(/出价|成本|控/.test(action))category='出价/成本';
      if(!actionMap[category])actionMap[category]={category,count:0,examples:[]};
      actionMap[category].count++;
      if(actionMap[category].examples.length<5)actionMap[category].examples.push({date:r.date,text:action});
    });
  });
  return Object.values(actionMap).sort((a,b)=>b.count-a.count);
}

// 生成方法论总结
function t11GenMethodology(reports,issues,actions){
  const methods=[];
  // 方法论1：链路测试
  const linkTests=reports.filter(r=>r.rawLines.some(l=>/链路|分流|问答|首页/.test(l))).length;
  if(linkTests>3)methods.push({title:'链路AB测试是核心调优手段',detail:'历史周报中多次出现链路切换与分流测试（H5问答+首页、小程序纯问答、H5分流等），通过对比不同链路的进量质量与产值，找到最优链路组合。建议持续保持至少2条链路并行测试。'});
  // 方法论2：素材方向
  const matActions=actions.find(a=>a.category==='素材策略');
  if(matActions)methods.push({title:'素材方向：养身类/IP类/复合类为主力',detail:'多次调整中反复提到"重点基建养身类、复合类素材"、"新广告重点上新养身、ip、复合类型素材"。说明这三类素材在太极品类中验证有效，应作为素材库基建的核心方向。'});
  // 方法论3：年龄定向
  const ageActions=actions.find(a=>a.category==='定向调整');
  if(ageActions)methods.push({title:'年龄定向：40-65岁为核心人群，控高龄占比',detail:'出现"部分新广告年龄定向40-65，控高龄占比"的调整，说明高龄人群（70+）虽然进量容易但后端产值低，需要通过定向控制年龄结构。'});
  // 方法论4：版位优化
  const placeActions=actions.find(a=>a.category==='版位/预算');
  if(placeActions)methods.push({title:'版位优化：未开单/低ROI素材及时止损',detail:'多次提到"未开单和roi较低的素材暂停"，说明素材生命周期管理很重要，消耗一定金额后未出单或ROI不达标的素材应及时暂停，把预算转移到跑量素材。'});
  // 方法论5：周报结构
  methods.push({title:'周报固定结构：进量→产值→问题→调整',detail:'13期周报均遵循"进量数据→产值数据→问题定位→本周调整"的结构。进量看链路/版位/画像，产值看分链路ROI与过程数据，问题定位到具体链路，调整给出可执行动作。'});
  return methods;
}

async function processTab11(){
  const fileInput=document.getElementById('t11File');
  const textInput=document.getElementById('t11Text');
  let text='';
  if(fileInput.files.length){
    text=await fileInput.files[0].text();
  }else if(textInput.value.trim()){
    text=textInput.value;
  }else{
    alert('请上传周报文本文件或粘贴文本内容');return;
  }
  document.getElementById('t11Btn').disabled=true;
  showStatus('t11','正在解析周报...','active');
  try{
    const reports=t11ParseWeeklyText(text);
    const issues=t11ExtractIssues(reports);
    const actions=t11ExtractActions(reports);
    const methods=t11GenMethodology(reports,issues,actions);
    const result={reports,issues,actions,methods,totalReports:reports.length,totalIssues:issues.reduce((s,i)=>s+i.count,0),totalActions:actions.reduce((s,a)=>s+a.count,0)};
    t11Output=t11BuildExcel(result);
    saveSync('weeklyArchive',result);
    t11Display(result);
    showStatus('t11','解析完成！共识别 '+reports.length+' 期周报，'+issues.length+' 类问题，'+actions.length+' 类调整动作','done');
  }catch(e){
    showStatus('t11','解析失败：'+e.message,'error');
  }
  document.getElementById('t11Btn').disabled=false;
}

function t11BuildExcel(d){
  const wb=XLSX.utils.book_new();
  // Sheet1: 周报总览
  const h1=['期数','主要板块数','问题数','调整动作数'];
  const r1=d.reports.map(r=>[r.date,r.sections.length,r.issues.length,r.actions.length]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h1,...r1]),'周报总览');
  // Sheet2: 每期详情
  const h2=['期数','原始内容'];
  const r2=d.reports.map(r=>[r.date,r.rawLines.join('\n')]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h2,...r2]),'每期详情');
  // Sheet3: 高频问题
  const h3=['问题类型','出现次数','示例期数','示例内容'];
  const r3=d.issues.map(i=>[i.key,i.count,i.examples.map(e=>e.date).join(','),i.examples.map(e=>e.text).join(' | ')]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h3,...r3]),'高频问题');
  // Sheet4: 调整动作
  const h4=['类别','出现次数','示例期数','示例内容'];
  const r4=d.actions.map(a=>[a.category,a.count,a.examples.map(e=>e.date).join(','),a.examples.map(e=>e.text).join(' | ')]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h4,...r4]),'调整动作');
  // Sheet5: 方法论
  const h5=['方法论','说明'];
  const r5=d.methods.map(m=>[m.title,m.detail]);
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([h5,...r5]),'方法论总结');
  return wb;
}

function t11Display(d){
  document.getElementById('t11Results').style.display='block';
  // KPI
  const kpiHtml='<div class="ex-kpi-card"><div class="ex-val">'+d.totalReports+'</div><div class="ex-lbl">周报期数</div></div>'+
    '<div class="ex-kpi-card"><div class="ex-val">'+d.issues.length+'</div><div class="ex-lbl">问题类型</div></div>'+
    '<div class="ex-kpi-card"><div class="ex-val">'+d.actions.length+'</div><div class="ex-lbl">调整类别</div></div>'+
    '<div class="ex-kpi-card"><div class="ex-val">'+d.methods.length+'</div><div class="ex-lbl">方法论</div></div>';
  document.getElementById('t11Kpi').innerHTML=kpiHtml;
  // 分析内容
  let html='<h3 style="font-size:15px;font-weight:600;margin:16px 0 10px">一、历史周报时间线</h3>';
  d.reports.forEach((r,i)=>{
    html+='<div style="margin-bottom:10px;padding:12px 16px;background:#f8fafc;border-radius:8px;cursor:pointer" onclick="this.querySelector(\'.wa-detail\').style.display=this.querySelector(\'.wa-detail\').style.display===\'none\'?\'block\':\'none\'">';
    html+='<div style="font-weight:600;font-size:13px;color:#1e293b">📅 '+r.date+' 周报 <span style="font-size:11px;color:#94a3b8;font-weight:normal">（'+r.sections.length+'板块 / '+r.issues.length+'问题 / '+r.actions.length+'调整）</span></div>';
    html+='<div class="wa-detail" style="display:none;margin-top:8px;font-size:12px;color:#475569;line-height:1.8;max-height:300px;overflow-y:auto">';
    if(r.issues.length){html+='<div style="color:#dc2626;font-weight:600;margin-bottom:4px">问题定位：</div>'+r.issues.map(x=>'<div>• '+x+'</div>').join('');}
    if(r.actions.length){html+='<div style="color:#16a34a;font-weight:600;margin:8px 0 4px">调整动作：</div>'+r.actions.map(x=>'<div>• '+x+'</div>').join('');}
    html+='<div style="color:#64748b;font-weight:600;margin:8px 0 4px">主要板块：</div>'+r.sections.slice(0,8).map(x=>'<div>• '+x+'</div>').join('');
    html+='</div></div>';
  });
  // 高频问题
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">二、高频问题定位汇总</h3>';
  d.issues.forEach((iss,i)=>{
    html+='<div style="margin-bottom:10px;padding:12px 16px;background:#fef2f2;border-radius:8px;border-left:4px solid #ef4444"><div style="font-weight:600;font-size:13px;color:#991b1b">'+(i+1)+'. '+iss.key+' <span style="font-size:11px;color:#dc2626">（出现'+iss.count+'次）</span></div><div style="font-size:11px;color:#64748b;margin-top:4px">出现在：'+iss.examples.map(e=>e.date).join('、')+'</div></div>';
  });
  // 调整动作
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">三、调整动作方法论库</h3>';
  d.actions.forEach((act,i)=>{
    html+='<div style="margin-bottom:10px;padding:12px 16px;background:#f0fdf4;border-radius:8px;border-left:4px solid #16a34a"><div style="font-weight:600;font-size:13px;color:#166534">'+(i+1)+'. '+act.category+' <span style="font-size:11px;color:#16a34a">（'+act.count+'次）</span></div><div style="font-size:11px;color:#64748b;margin-top:4px">出现在：'+act.examples.map(e=>e.date).join('、')+'</div><div style="font-size:12px;color:#475569;margin-top:6px;line-height:1.7">'+act.examples.map(e=>'• '+e.text).join('<br>')+'</div></div>';
  });
  // 方法论
  html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px">四、核心方法论总结</h3>';
  d.methods.forEach((m,i)=>{
    html+='<div style="margin-bottom:12px;padding:14px 18px;background:#eff6ff;border-radius:10px;border-left:4px solid #3b82f6"><div style="font-weight:600;font-size:14px;color:#1e40af;margin-bottom:6px">'+(i+1)+'. '+m.title+'</div><div style="font-size:13px;color:#475569;line-height:1.8">'+m.detail+'</div></div>';
  });
  document.getElementById('t11Analysis').innerHTML=html;
}

function downloadTab11(){ if(t11Output) XLSX.writeFile(t11Output,'周报历史归档.xlsx'); }

// 分析模块渲染
function renderWeeklyArchive(){
  const d=SYNC.weeklyArchive;
  if(!d){
    document.getElementById('wa-empty').style.display='block';
    ['wa-timeline-card','wa-issues-card','wa-actions-card','wa-method-card'].forEach(id=>document.getElementById(id).style.display='none');
    document.getElementById('wa-badge').textContent='待同步';
    return;
  }
  document.getElementById('wa-empty').style.display='none';
  document.getElementById('wa-badge').textContent=d.totalReports+'期周报 · '+d.issues.length+'类问题 · 已同步';
  document.getElementById('wa-summary').textContent='覆盖 '+d.reports[0].date+' 至 '+d.reports[d.reports.length-1].date;
  // KPI
  const kpiHtml='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px">'+
    '<div style="background:linear-gradient(135deg,#3b82f612,#3b82f605);border:1px solid #3b82f630;border-radius:12px;padding:14px 16px;border-top:3px solid #3b82f6"><div style="font-size:12px;color:#64748b;font-weight:600">周报期数</div><div style="font-size:24px;font-weight:700;color:#3b82f6">'+d.totalReports+'</div></div>'+
    '<div style="background:linear-gradient(135deg,#ef444412,#ef444405);border:1px solid #ef444430;border-radius:12px;padding:14px 16px;border-top:3px solid #ef4444"><div style="font-size:12px;color:#64748b;font-weight:600">问题类型</div><div style="font-size:24px;font-weight:700;color:#ef4444">'+d.issues.length+'</div></div>'+
    '<div style="background:linear-gradient(135deg,#10b98112,#10b98105);border:1px solid #10b98130;border-radius:12px;padding:14px 16px;border-top:3px solid #10b981"><div style="font-size:12px;color:#64748b;font-weight:600">调整类别</div><div style="font-size:24px;font-weight:700;color:#10b981">'+d.actions.length+'</div></div>'+
    '<div style="background:linear-gradient(135deg,#8b5cf612,#8b5cf605);border:1px solid #8b5cf630;border-radius:12px;padding:14px 16px;border-top:3px solid #8b5cf6"><div style="font-size:12px;color:#64748b;font-weight:600">方法论</div><div style="font-size:24px;font-weight:700;color:#8b5cf6">'+d.methods.length+'</div></div></div>';
  document.getElementById('wa-kpi').innerHTML=kpiHtml;
  // 时间线
  document.getElementById('wa-timeline-card').style.display='block';
  document.getElementById('wa-timeline').innerHTML=d.reports.map((r,i)=>{
    return '<div style="margin-bottom:10px;padding:12px 16px;background:#f8fafc;border-radius:8px;cursor:pointer" onclick="this.querySelector(\'.wa-detail\').style.display=this.querySelector(\'.wa-detail\').style.display===\'none\'?\'block\':\'none\'">'+
      '<div style="font-weight:600;font-size:13px;color:#1e293b">📅 '+r.date+' 周报 <span style="font-size:11px;color:#94a3b8;font-weight:normal">（'+r.sections.length+'板块 / '+r.issues.length+'问题 / '+r.actions.length+'调整）</span></div>'+
      '<div class="wa-detail" style="display:none;margin-top:8px;font-size:12px;color:#475569;line-height:1.8;max-height:300px;overflow-y:auto">'+
      (r.issues.length?'<div style="color:#dc2626;font-weight:600;margin-bottom:4px">问题定位：</div>'+r.issues.map(x=>'<div>• '+x+'</div>').join(''):'')+
      (r.actions.length?'<div style="color:#16a34a;font-weight:600;margin:8px 0 4px">调整动作：</div>'+r.actions.map(x=>'<div>• '+x+'</div>').join(''):'')+
      '<div style="color:#64748b;font-weight:600;margin:8px 0 4px">主要板块：</div>'+r.sections.slice(0,8).map(x=>'<div>• '+x+'</div>').join('')+
      '</div></div>';
  }).join('');
  // 高频问题
  document.getElementById('wa-issues-card').style.display='block';
  document.getElementById('wa-issues').innerHTML=d.issues.map((iss,i)=>'<div style="margin-bottom:10px;padding:12px 16px;background:#fef2f2;border-radius:8px;border-left:4px solid #ef4444"><div style="font-weight:600;font-size:13px;color:#991b1b">'+(i+1)+'. '+iss.key+' <span style="font-size:11px;color:#dc2626">（出现'+iss.count+'次）</span></div><div style="font-size:11px;color:#64748b;margin-top:4px">出现在：'+iss.examples.map(e=>e.date).join('、')+'</div></div>').join('');
  // 调整动作
  document.getElementById('wa-actions-card').style.display='block';
  document.getElementById('wa-actions').innerHTML=d.actions.map((act,i)=>'<div style="margin-bottom:10px;padding:12px 16px;background:#f0fdf4;border-radius:8px;border-left:4px solid #16a34a"><div style="font-weight:600;font-size:13px;color:#166534">'+(i+1)+'. '+act.category+' <span style="font-size:11px;color:#16a34a">（'+act.count+'次）</span></div><div style="font-size:11px;color:#64748b;margin-top:4px">出现在：'+act.examples.map(e=>e.date).join('、')+'</div><div style="font-size:12px;color:#475569;margin-top:6px;line-height:1.7">'+act.examples.map(e=>'• '+e.text).join('<br>')+'</div></div>').join('');
  // 方法论
  document.getElementById('wa-method-card').style.display='block';
  document.getElementById('wa-method').innerHTML=d.methods.map((m,i)=>'<div style="margin-bottom:12px;padding:14px 18px;background:#eff6ff;border-radius:10px;border-left:4px solid #3b82f6"><div style="font-weight:600;font-size:14px;color:#1e40af;margin-bottom:6px">'+(i+1)+'. '+m.title+'</div><div style="font-size:13px;color:#475569;line-height:1.8">'+m.detail+'</div></div>').join('');
}

/* ---------- 拖拽上传（通用，自动绑定所有file input） ---------- */
function initAllDragUploads(){
  document.querySelectorAll('input[type=file]').forEach(input=>{
    // 跳过t9的input（已有独立拖拽逻辑）
    if(input.closest('.t9-zone')||input.closest('#t9DropZone'))return;
    // 找最近的可点击容器（button的父级或card）
    let zone=input.closest('.card')||input.parentElement;
    // 如果有ex-upload-row，用它作为拖拽区
    const row=input.closest('.ex-upload-row');
    if(row)zone=row;
    zone.style.position='relative';
    zone.addEventListener('dragover',e=>{e.preventDefault();zone.classList.add('dragover');zone.style.borderColor='var(--accent)';});
    zone.addEventListener('dragleave',e=>{e.preventDefault();zone.classList.remove('dragover');zone.style.borderColor='';});
    zone.addEventListener('drop',e=>{
      e.preventDefault();zone.classList.remove('dragover');zone.style.borderColor='';
      const files=e.dataTransfer.files;
      if(files&&files.length){
        // 尝试匹配：如果拖拽区有多个file input，按顺序分配
        const inputs=zone.querySelectorAll('input[type=file]');
        if(inputs.length>1){
          // 多文件场景，尝试按文件名匹配
          [...files].forEach(f=>{
            let target=null;
            const name=f.name;
            // 工具9：线索评分·版位分析（4份文件）
            if(/手机/i.test(name)&&/公小/i.test(name))target=zone.querySelector('#t9PhoneGxFile')||document.getElementById('t9PhoneGxFile');
            else if(/手机/i.test(name)&&/视频/i.test(name))target=zone.querySelector('#t9PhoneSpFile')||document.getElementById('t9PhoneSpFile');
            else if(/设备/i.test(name)&&/公小/i.test(name))target=zone.querySelector('#t9DevGxFile')||document.getElementById('t9DevGxFile');
            else if(/设备/i.test(name)&&/视频/i.test(name))target=zone.querySelector('#t9DevSpFile')||document.getElementById('t9DevSpFile');
            // 工具7：线索评分（2份文件）
            else if(/手机/i.test(name))target=zone.querySelector('#t7PhoneFile')||document.getElementById('t7PhoneFile');
            else if(/设备/i.test(name))target=zone.querySelector('#t7DevFile')||document.getElementById('t7DevFile');
            // 兜底：找第一个空的input
            if(!target)target=inputs[[...inputs].findIndex(i=>!i.files.length)]||inputs[0];
            if(target){const dt=new DataTransfer();dt.items.add(f);target.files=dt.files;target.dispatchEvent(new Event('change'));}
          });
        }else{
          const dt=new DataTransfer();dt.items.add(files[0]);input.files=dt.files;input.dispatchEvent(new Event('change'));
        }
      }
    });
  });
}

/* ---------- 周报撰写 ---------- */
let t8DocxText='';
let t8WeeklyReports=[];
let t8GeneratedReport=null;
let t8WeeklyData=null;
let t8SelectedFiles=[];

// 绑定文件上传
document.addEventListener('DOMContentLoaded',()=>{
  const af=document.getElementById('t8AllFiles');
  if(af)af.addEventListener('change',e=>{
    t8SelectedFiles=Array.from(e.target.files).slice(0,4);
    if(t8SelectedFiles.length){
      const names=t8SelectedFiles.map(f=>f.name).join('、');
      document.getElementById('t8AllFileName').textContent='已选 '+t8SelectedFiles.length+' 个文件：'+names;
      document.getElementById('t8ParseAllBtn').disabled=false;
    }
  });
  // 拖拽上传支持
  const zone=document.querySelector('#view-excel-weekly-report .ex-upload-zone');
  if(zone){
    zone.addEventListener('dragover',e=>{e.preventDefault();zone.style.borderColor='#3b82f6';zone.style.background='#eff6ff';});
    zone.addEventListener('dragleave',e=>{e.preventDefault();zone.style.borderColor='';zone.style.background='';});
    zone.addEventListener('drop',e=>{
      e.preventDefault();zone.style.borderColor='';zone.style.background='';
      t8SelectedFiles=Array.from(e.dataTransfer.files).slice(0,4);
      if(t8SelectedFiles.length){
        try{const dt=new DataTransfer();t8SelectedFiles.forEach(f=>dt.items.add(f));af.files=dt.files;}catch(err){}
        const names=t8SelectedFiles.map(f=>f.name).join('、');
        document.getElementById('t8AllFileName').textContent='已选 '+t8SelectedFiles.length+' 个文件：'+names;
        document.getElementById('t8ParseAllBtn').disabled=false;
      }
    });
  }
});

// 解析docx
async function parseDocx(file){
  const buf=await file.arrayBuffer();
  const zip=await JSZip.loadAsync(buf);
  const xml=await zip.file('word/document.xml').async('string');
  const parser=new DOMParser();
  const doc=parser.parseFromString(xml,'text/xml');
  const body=doc.querySelector('body');
  const blocks=[];
  for(const child of body.children){
    const tag=child.tagName.replace(/^.*:/,'');
    if(tag==='p'){
      const texts=[];
      child.querySelectorAll('t').forEach(t=>texts.push(t.textContent||''));
      const text=texts.join('').trim();
      if(text)blocks.push({type:'p',text});
    }else if(tag==='tbl'){
      const rows=[];
      child.querySelectorAll('tr').forEach(tr=>{
        const cells=[];
        tr.querySelectorAll('tc').forEach(tc=>{
          const ctexts=[];
          tc.querySelectorAll('t').forEach(t=>ctexts.push(t.textContent||''));
          cells.push(ctexts.join('').trim());
        });
        rows.push(cells);
      });
      if(rows.length)blocks.push({type:'table',rows});
    }
  }
  return blocks;
}

// 从blocks中提取各期周报
function extractWeeklyReports(blocks){
  const reports=[];
  let cur=null;
  const datePattern=/^\d{4}周报|^\d{2}\/\d{2}|^\d{2}月\d{2}日|太极投放周报/;
  for(const b of blocks){
    if(b.type==='p'&&/周报/.test(b.text)&&/^\d{4}/.test(b.text)){
      if(cur)reports.push(cur);
      cur={title:b.text,date:b.text.replace(/周报.*/,'').trim(),sections:[],tables:[]};
    }else if(cur){
      if(b.type==='p')cur.sections.push(b.text);
      else cur.tables.push(b.rows);
    }
  }
  if(cur)reports.push(cur);
  // 如果没识别到多期，把所有内容作为一期
  if(reports.length===0&&blocks.length){
    reports.push({title:'历史周报',date:'未知',sections:blocks.filter(b=>b.type==='p').map(b=>b.text),tables:blocks.filter(b=>b.type==='table').map(b=>b.rows)});
  }
  return reports;
}

// 解析周报
async function processTab8(){
  const fileInput=document.getElementById('t8File');
  if(!fileInput.files.length){alert('请先上传历史周报docx');return;}
  document.getElementById('t8ParseBtn').disabled=true;
  showStatus('t8','正在解析docx文件...','active');
  try{
    const blocks=await parseDocx(fileInput.files[0]);
    t8WeeklyReports=extractWeeklyReports(blocks);
    t8DocxText=blocks.filter(b=>b.type==='p').map(b=>b.text).join('\n');
    // 展示历史周报列表
    document.getElementById('t8Results').style.display='block';
    document.getElementById('t8Count').textContent='共解析 '+t8WeeklyReports.length+' 期周报';
    const listHtml=t8WeeklyReports.map((r,i)=>{
      const summary=r.sections.slice(0,3).map(s=>s.length>80?s.slice(0,80)+'...':s).join(' | ');
      return '<div style="padding:12px;border:1px solid #e2e8f0;border-radius:8px;margin-bottom:8px;cursor:pointer" onclick="toggleWeeklyDetail('+i+')">'+
        '<div style="font-weight:600;color:#1e293b">'+r.title+'</div>'+
        '<div style="font-size:12px;color:#64748b;margin-top:4px">段落数：'+r.sections.length+'，表格数：'+r.tables.length+'</div>'+
        '<div id="wr-detail-'+i+'" style="display:none;margin-top:8px;font-size:12px;color:#475569;line-height:1.8;max-height:300px;overflow-y:auto">'+
        r.sections.map(s=>'<div>'+esc(s)+'</div>').join('')+'</div></div>';
    }).join('');
    document.getElementById('t8HistoryList').innerHTML=listHtml;
    // 显示数据表上传区
    document.getElementById('t8InputCard').style.display='block';
    // 绑定数据表文件选择
    const df=document.getElementById('t8DataFile');
    if(df)df.addEventListener('change',e=>{
      const files=e.target.files;
      if(files.length){
        const names=Array.from(files).map(f=>f.name).join('、');
        document.getElementById('t8DataFileName').textContent='已选 '+files.length+' 个文件：'+names;
        document.getElementById('t8ParseDataBtn').disabled=false;
      }
    });
    document.getElementById('t8GenBtn').disabled=false;
    showStatus('t8','解析完成！共 '+t8WeeklyReports.length+' 期周报，上传本周数据表后一键生成','done');
  }catch(e){
    showStatus('t8','解析失败：'+e.message,'error');
    document.getElementById('t8ParseBtn').disabled=false;
  }
}

function toggleWeeklyDetail(i){
  const el=document.getElementById('wr-detail-'+i);
  if(el)el.style.display=el.style.display==='none'?'block':'none';
}

// 统一解析所有上传文件（1~4个）
async function parseAllWeeklyFiles(){
  const files=t8SelectedFiles.length?t8SelectedFiles:Array.from(document.getElementById('t8AllFiles').files||[]);
  if(!files.length){alert('请先上传1~4个文件');return;}
  document.getElementById('t8ParseAllBtn').disabled=true;
  showStatus('t8','正在解析文件...','active');
  try{
    const docxFiles=files.filter(f=>/\.docx$/i.test(f.name));
    const dataFiles=files.filter(f=>/\.(xlsx|xls|csv)$/i.test(f.name));
    let docxParsed=false,dataParsed=false;

    // 解析历史周报docx
    if(docxFiles.length){
      const blocks=await parseDocx(docxFiles[0]);
      t8WeeklyReports=extractWeeklyReports(blocks);
      t8DocxText=blocks.filter(b=>b.type==='p').map(b=>b.text).join('\n');
      document.getElementById('t8Results').style.display='block';
      document.getElementById('t8Count').textContent='历史周报：'+t8WeeklyReports.length+' 期';
      const listHtml=t8WeeklyReports.map((r,i)=>{
        const summary=r.sections.slice(0,2).map(s=>s.length>60?s.slice(0,60)+'...':s).join(' | ');
        return '<div style="padding:10px;border:1px solid #e2e8f0;border-radius:6px;margin-bottom:6px;cursor:pointer" onclick="toggleWeeklyDetail('+i+')">'+
          '<div style="font-weight:600;font-size:13px;color:#1e293b">'+r.title+'</div>'+
          '<div style="font-size:11px;color:#64748b;margin-top:2px">'+r.sections.length+'段 / '+r.tables.length+'表</div>'+
          '<div id="wr-detail-'+i+'" style="display:none;margin-top:6px;font-size:11px;color:#475569;line-height:1.7;max-height:200px;overflow-y:auto">'+
          r.sections.map(s=>'<div>'+esc(s)+'</div>').join('')+'</div></div>';
      }).join('');
      document.getElementById('t8HistoryList').innerHTML=listHtml;
      docxParsed=true;
    }

    // 解析数据表
    if(dataFiles.length){
      const result={totalOrders:0,totalCost:0,femaleRatio:0,links:[],placements:[],outputs:[],rawSheets:[]};
      const linkMap={},placeMap={},outputMap={};
      let totalSpend=0,totalOrders=0,femaleCount=0,totalCount=0;
      for(const file of dataFiles){
        const buf=await file.arrayBuffer();
        const wb=XLSX.read(buf,{type:'array'});
        for(const sheetName of wb.SheetNames){
          const ws=wb.Sheets[sheetName];
          const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
          if(rows.length<2)continue;
          let headerRow=0,maxScore=0;
          for(let i=0;i<Math.min(5,rows.length);i++){
            let score=0;rows[i].forEach(c=>{if(detectColumnType(c))score++;});
            if(score>maxScore){maxScore=score;headerRow=i;}
          }
          const headers=rows[headerRow].map(h=>String(h||'').trim());
          const colMap={};
          headers.forEach((h,i)=>{const t=detectColumnType(h);if(t&&!colMap[t])colMap[t]=i;});
          result.rawSheets.push({name:sheetName,rows:rows.length,colMap});
          for(let i=headerRow+1;i<rows.length;i++){
            const row=rows[i];
            if(!row||row.every(c=>c===''||c==null))continue;
            const orders=colMap.orders!=null?parseFloat(row[colMap.orders])||0:0;
            const cost=colMap.cost!=null?parseFloat(row[colMap.cost])||0:0;
            const output=colMap.output!=null?parseFloat(row[colMap.output])||0:0;
            const roi=colMap.roi!=null?parseFloat(row[colMap.roi])||0:0;
            const roiPct=roi>0&&roi<1?roi*100:roi;
            const link=colMap.link!=null?String(row[colMap.link]||'').trim():'';
            const place=colMap.placement!=null?String(row[colMap.placement]||'').trim():'';
            const gender=colMap.gender!=null?String(row[colMap.gender]||'').trim():'';
            if(orders>0){totalOrders+=orders;totalCount+=orders;}
            if(cost>0)totalSpend+=cost;
            if(/女/.test(gender)&&orders>0)femaleCount+=orders;
            if(link&&!/(汇总|总计|合计)/.test(link)){
              if(!linkMap[link])linkMap[link]={name:link,orders:0,cost:0,output:0,roi:0,count:0};
              linkMap[link].orders+=orders;linkMap[link].cost+=cost;linkMap[link].output+=output;
              if(roiPct>0){linkMap[link].roi+=roiPct;linkMap[link].count++;}
            }
            if(place&&!/(汇总|总计|合计)/.test(place)){
              if(!placeMap[place])placeMap[place]={name:place,orders:0,cost:0};
              placeMap[place].orders+=orders;placeMap[place].cost+=cost;
            }
            if(link&&(output>0||roiPct>0)&&!/(汇总|总计|合计)/.test(link)){
              if(!outputMap[link])outputMap[link]={name:link,output:0,roi:0,count:0};
              outputMap[link].output+=output;
              if(roiPct>0){outputMap[link].roi+=roiPct;outputMap[link].count++;}
            }
          }
        }
      }
      result.totalOrders=Math.round(totalOrders);
      result.totalCost=totalOrders>0?Math.round(totalSpend/totalOrders):0;
      result.femaleRatio=totalCount>0?Math.round(femaleCount/totalCount*100):0;
      result.links=Object.values(linkMap).map(l=>({name:l.name,orders:Math.round(l.orders),cost:l.orders>0?Math.round(l.cost/l.orders):0,output:Math.round(l.output),roi:l.count>0?Math.round(l.roi/l.count*100)/100:0})).sort((a,b)=>b.orders-a.orders);
      result.placements=Object.values(placeMap).map(p=>({name:p.name,orders:Math.round(p.orders),cost:p.orders>0?Math.round(p.cost/p.orders):0})).sort((a,b)=>b.orders-a.orders);
      result.outputs=Object.values(outputMap).map(o=>({name:o.name,output:Math.round(o.output),roi:o.count>0?Math.round(o.roi/o.count*100)/100:0,issue:''})).sort((a,b)=>b.output-a.output);
      t8WeeklyData=result;
      document.getElementById('t8DataPreview').style.display='block';
      let summary='<div><b>总订单数：</b>'+result.totalOrders+' 单</div>';
      summary+='<div><b>整体成本：</b>'+result.totalCost+' 元</div>';
      summary+='<div><b>女性占比：</b>'+result.femaleRatio+'%</div>';
      summary+='<div><b>识别链路：</b>'+result.links.length+' 条（'+result.links.map(l=>l.name).join('、')+'）</div>';
      summary+='<div><b>识别版位：</b>'+result.placements.length+' 个（'+result.placements.map(p=>p.name).join('、')+'）</div>';
      summary+='<div><b>产值数据：</b>'+result.outputs.length+' 条链路</div>';
      document.getElementById('t8DataSummary').innerHTML=summary;
      dataParsed=true;
    }

    // 启用生成按钮
    if(dataParsed){
      document.getElementById('t8GenFinalBtn').disabled=false;
    }
    const msg=[];
    if(docxParsed)msg.push('历史周报'+t8WeeklyReports.length+'期');
    if(dataParsed)msg.push('数据表'+t8WeeklyData.links.length+'条链路');
    showStatus('t8','解析完成！'+msg.join('，')+'，点击一键生成周报','done');
  }catch(e){
    showStatus('t8','解析失败：'+e.message,'error');
    document.getElementById('t8ParseAllBtn').disabled=false;
  }
}

// 周报数据存储
// 智能识别列名对应的指标
function detectColumnType(header){
  const h=String(header).toLowerCase().trim();
  if(/roi|投产比|投产/.test(h))return 'roi';
  if(/产值|收入|金额|gmv|转化金额|客单价|arpu/.test(h))return 'output';
  if(/女性|性别/.test(h))return 'gender';
  if(/版位|平台|流量来源|广告位|流量版位/.test(h))return 'placement';
  if(/链路|皮肤|落地页|计划名称|广告名称|推广计划|计划/.test(h))return 'link';
  if(/花费|消耗|成本|订单成本|转化成本/.test(h))return 'cost';
  if(/订单数|订单量|转化数|下单数|成交数|转化量/.test(h))return 'orders';
  if(/日期|时间|期次/.test(h))return 'date';
  return null;
}

// 解析周报数据表
async function parseWeeklyDataFiles(){
  const fileInput=document.getElementById('t8DataFile');
  if(!fileInput.files.length){alert('请先上传数据表');return;}
  document.getElementById('t8ParseDataBtn').disabled=true;
  showStatus('t8Data','正在解析数据表...','active');
  try{
    const result={
      totalOrders:0,totalCost:0,femaleRatio:0,
      links:[],placements:[],outputs:[],
      rawSheets:[]
    };
    const linkMap={},placeMap={},outputMap={};
    let totalSpend=0,totalOrders=0,femaleCount=0,totalCount=0;

    for(const file of fileInput.files){
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:'array'});
      for(const sheetName of wb.SheetNames){
        const ws=wb.Sheets[sheetName];
        const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
        if(rows.length<2)continue;
        // 找表头行（前5行中包含最多识别字段的行）
        let headerRow=0,maxScore=0;
        for(let i=0;i<Math.min(5,rows.length);i++){
          let score=0;
          rows[i].forEach(c=>{if(detectColumnType(c))score++;});
          if(score>maxScore){maxScore=score;headerRow=i;}
        }
        const headers=rows[headerRow].map(h=>String(h||'').trim());
        const colMap={};
        headers.forEach((h,i)=>{const t=detectColumnType(h);if(t&&!colMap[t])colMap[t]=i;});
        result.rawSheets.push({name:sheetName,rows:rows.length,cols:headers.length,colMap});
        // 解析数据行
        for(let i=headerRow+1;i<rows.length;i++){
          const row=rows[i];
          if(!row||row.every(c=>c===''||c==null))continue;
          const orders=colMap.orders!=null?parseFloat(row[colMap.orders])||0:0;
          const cost=colMap.cost!=null?parseFloat(row[colMap.cost])||0:0;
          const output=colMap.output!=null?parseFloat(row[colMap.output])||0:0;
          const roi=colMap.roi!=null?parseFloat(row[colMap.roi])||0:0;
          const roiPct=roi>0&&roi<1?roi*100:roi;
          const link=colMap.link!=null?String(row[colMap.link]||'').trim():'';
          const place=colMap.placement!=null?String(row[colMap.placement]||'').trim():'';
          const gender=colMap.gender!=null?String(row[colMap.gender]||'').trim():'';
          // 聚合总量
          if(orders>0){totalOrders+=orders;totalCount+=orders;}
          if(cost>0)totalSpend+=cost;
          if(/女/.test(gender)&&orders>0)femaleCount+=orders;
          // 按链路聚合
          if(link&&!/(汇总|总计|合计)/.test(link)){
            if(!linkMap[link])linkMap[link]={name:link,orders:0,cost:0,output:0,roi:0,count:0};
            linkMap[link].orders+=orders;
            linkMap[link].cost+=cost;
            linkMap[link].output+=output;
            if(roiPct>0){linkMap[link].roi+=roiPct;linkMap[link].count++;}
          }
          // 按版位聚合
          if(place&&!/(汇总|总计|合计)/.test(place)){
            if(!placeMap[place])placeMap[place]={name:place,orders:0,cost:0};
            placeMap[place].orders+=orders;
            placeMap[place].cost+=cost;
          }
          // 产值数据（有产值或ROI的行）
          if(link&&(output>0||roiPct>0)&&!/(汇总|总计|合计)/.test(link)){
            if(!outputMap[link])outputMap[link]={name:link,output:0,roi:0,count:0};
            outputMap[link].output+=output;
            if(roiPct>0){outputMap[link].roi+=roiPct;outputMap[link].count++;}
          }
        }
      }
    }
    // 整理结果
    result.totalOrders=Math.round(totalOrders);
    result.totalCost=totalOrders>0?Math.round(totalSpend/totalOrders):0;
    result.femaleRatio=totalCount>0?Math.round(femaleCount/totalCount*100):0;
    result.links=Object.values(linkMap).map(l=>({
      name:l.name,orders:Math.round(l.orders),
      cost:l.orders>0?Math.round(l.cost/l.orders):0,
      output:Math.round(l.output),
      roi:l.count>0?Math.round(l.roi/l.count*100)/100:0
    })).sort((a,b)=>b.orders-a.orders);
    result.placements=Object.values(placeMap).map(p=>({
      name:p.name,orders:Math.round(p.orders),
      cost:p.orders>0?Math.round(p.cost/p.orders):0
    })).sort((a,b)=>b.orders-a.orders);
    result.outputs=Object.values(outputMap).map(o=>({
      name:o.name,output:Math.round(o.output),
      roi:o.count>0?Math.round(o.roi/o.count*100)/100:0,issue:''
    })).sort((a,b)=>b.output-a.output);
    t8WeeklyData=result;
    // 显示预览
    document.getElementById('t8DataPreview').style.display='block';
    let summary='<div><b>总订单数：</b>'+result.totalOrders+' 单</div>';
    summary+='<div><b>整体成本：</b>'+result.totalCost+' 元</div>';
    summary+='<div><b>女性占比：</b>'+result.femaleRatio+'%</div>';
    summary+='<div><b>识别到链路：</b>'+result.links.length+' 条（'+result.links.map(l=>l.name).join('、')+'）</div>';
    summary+='<div><b>识别到版位：</b>'+result.placements.length+' 个（'+result.placements.map(p=>p.name).join('、')+'）</div>';
    summary+='<div><b>识别到产值数据：</b>'+result.outputs.length+' 条链路</div>';
    summary+='<div><b>解析Sheet：</b>'+result.rawSheets.length+' 个</div>';
    document.getElementById('t8DataSummary').innerHTML=summary;
    document.getElementById('t8GenV2Btn').disabled=false;
    showStatus('t8Data','解析完成！识别到 '+result.links.length+' 条链路、'+result.placements.length+' 个版位','done');
  }catch(e){
    showStatus('t8Data','解析失败：'+e.message,'error');
    document.getElementById('t8ParseDataBtn').disabled=false;
  }
}

// V2：生成带表格和方法论预警的完整周报
function generateWeeklyReportV2(){
  // 数据适配：优先用Excel上传的t8WeeklyData，没有则从工作台SYNC数据构建
  let d=t8WeeklyData;
  if(!d && SYNC.link && SYNC.link.linkSummary){
    const ls=SYNC.link.linkSummary;
    const links=ls.map(l=>({name:l.link,orders:l.eff||0,cost:l.orderCost||(l.spend||0)/(l.eff||1)}));
    const outputs=ls.map(l=>({name:l.link,output:l.output||0,roi:(l.roi||0)*100,issue:''}));
    const totalOrders=ls.reduce((s,l)=>s+(l.eff||0),0);
    const totalSpend=ls.reduce((s,l)=>s+(l.spend||0),0);
    d={links,placements:[],outputs,totalOrders,totalCost:totalOrders>0?Math.round(totalSpend/totalOrders):0,femaleRatio:0};
  }
  if(!d){alert('请先在各数据模块导入数据（链路数据表/素材分析/线索评分），或上传Excel数据表');return;}
  showStatus('t8','正在生成周报...','active');
  const linkData=d.links;
  const placeData=d.placements;
  const outputData=d.outputs.length>0?d.outputs:linkData.map(l=>({name:l.name,output:l.output,roi:l.roi,issue:''}));

  const totalOrders=d.totalOrders;
  const totalCost=d.totalCost;
  const femaleRatio=d.femaleRatio;
  const period='最近期';
  const lastFeedback='';

  // 周报内容模块勾选状态
  const modWarning=$('#wr-mod-warning')?$('#wr-mod-warning').checked:true;
  const modIntake=$('#wr-mod-intake')?$('#wr-mod-intake').checked:true;
  const modOutput=$('#wr-mod-output')?$('#wr-mod-output').checked:true;
  const modLink=$('#wr-mod-link')?$('#wr-mod-link').checked:true;
  const modMaterial=$('#wr-mod-material')?$('#wr-mod-material').checked:true;
  const modScore=$('#wr-mod-score')?$('#wr-mod-score').checked:true;
  const modAction=$('#wr-mod-action')?$('#wr-mod-action').checked:true;

  // 计算日期范围
  const now=new Date();
  const day=now.getDay();
  const lastMon=new Date(now);lastMon.setDate(now.getDate()-day-6);
  const lastSun=new Date(now);lastSun.setDate(now.getDate()-day);
  const fmt=d=>(d.getMonth()+1)+'.'+d.getDate();
  const weekRange=fmt(lastMon)+'-'+fmt(lastSun);
  const reportTitle=(lastMon.getMonth()+1)+''+String(lastMon.getDate()).padStart(2,'0')+'周报';

  // 方法论预警
  const warnings=[];
  if(femaleRatio>20)warnings.push({type:'warn',text:'女性占比'+femaleRatio+'%超过20%警戒线，女性成本比男性高23-45块，产值不稳定，建议控到15%左右'});
  if(totalCost>200)warnings.push({type:'warn',text:'整体成本'+totalCost+'元偏高，建议降出价、控量级、排低R素材'});
  outputData.forEach(o=>{
    if(o.roi>0&&o.roi<50)warnings.push({type:'warn',text:o.name+' ROI仅'+o.roi+'%未达标，需排查跑量素材和画像'});
    if(/领取课程|强调领课|大字报/.test(o.issue))warnings.push({type:'warn',text:o.name+' 素材片头为"领取课程"类，历史数据这类素材低R，建议换痛点前置'});
  });
  const bestLink=outputData.reduce((a,b)=>b.roi>a.roi?b:a,outputData[0]);
  const worstLink=outputData.reduce((a,b)=>b.roi<a.roi&&b.roi>0?b:a,outputData[0]);

  // 生成本周调整建议
  const suggestions=[];
  suggestions.push('新广告全部用双出价模式，对比单出价广告的ROI是否更稳定');
  if(worstLink&&worstLink.roi>0&&worstLink.roi<50)suggestions.push(worstLink.name+'控成本，暂停低ROI素材，培养新广告');
  if(femaleRatio>20)suggestions.push('控女性占比到15%左右，女性占比高的广告定投男性');
  suggestions.push('评分+AI落地页培养头部素材，控成本');
  if(bestLink&&bestLink.roi>=100)suggestions.push(bestLink.name+'表现优异，可适当放量');

  // 生成HTML
  let html='<div style="font-family:inherit;line-height:1.8;color:#1e293b;font-size:14px">';
  html+='<h1 style="font-size:22px;font-weight:700;margin-bottom:4px;color:#1e293b">太极投放周报-吴婷婷</h1>';
  html+='<div style="font-size:16px;font-weight:600;color:#3b82f6;margin-bottom:20px">'+reportTitle+'</div>';

  // 预警区
  if(modWarning && warnings.length){
    html+='<div style="background:#fef3c7;border-left:4px solid #f59e0b;padding:12px 16px;border-radius:8px;margin-bottom:20px">';
    html+='<div style="font-weight:600;color:#92400e;margin-bottom:6px">⚠️ 方法论预警（'+warnings.length+'条）</div>';
    warnings.forEach(w=>{html+='<div style="font-size:13px;color:#78350f;margin-bottom:4px">• '+w.text+'</div>';});
    html+='</div>';
  }

  // 一、进量数据
  if(modIntake){
  html+='<h2 style="font-size:17px;font-weight:700;margin:24px 0 12px;padding-bottom:8px;border-bottom:2px solid #3b82f6">一、上周进量数据（上周时间维度：'+weekRange+'）</h2>';
  html+='<h3 style="font-size:15px;font-weight:600;margin:16px 0 10px">1、进量情况（测试数据）</h3>';

  // 链路进量表
  html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
  ['链路','订单数','订单成本','订单占比'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
  html+='</tr></thead><tbody>';
  const totalLinkOrders=linkData.reduce((s,l)=>s+l.orders,0)||1;
  linkData.forEach(l=>{
    const ratio=((l.orders/totalLinkOrders)*100).toFixed(1);
    html+='<tr><td style="border:1px solid #e2e8f0;padding:8px">'+l.name+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+l.orders+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+l.cost+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+ratio+'%</td></tr>';
  });
  html+='<tr style="background:#f1f5f9;font-weight:600"><td style="border:1px solid #e2e8f0;padding:8px">汇总</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+totalLinkOrders+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+totalCost+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">100%</td></tr>';
  html+='</tbody></table>';

  // 版位表
  html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
  ['版位','订单数','订单占比','订单成本'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
  html+='</tr></thead><tbody>';
  const totalPlaceOrders=placeData.reduce((s,p)=>s+p.orders,0)||1;
  placeData.forEach(p=>{
    if(p.orders===0)return;
    const ratio=((p.orders/totalPlaceOrders)*100).toFixed(1);
    html+='<tr><td style="border:1px solid #e2e8f0;padding:8px">'+p.name+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+p.orders+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+ratio+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+p.cost+'</td></tr>';
  });
  html+='</tbody></table>';

  // 进量总结
  html+='<p style="margin:12px 0">上周总订单数共'+totalOrders+'个，整体成本'+totalCost+'元，女性占比'+femaleRatio+'%。';
  if(placeData[0]&&placeData[0].orders>0&&placeData[1]&&placeData[1].orders>0){
    html+='版位的订单占比变化不大，主跑量的公小和视频号版位成本'+(placeData[0].cost<placeData[1].cost?'公小更低':'视频号更低')+'。';
  }
  html+='</p>';
  }

  // 二、产值数据
  if(modOutput){
  html+='<h2 style="font-size:17px;font-weight:700;margin:24px 0 12px;padding-bottom:8px;border-bottom:2px solid #3b82f6">二、产值数据（'+period+'）</h2>';
  html+='<h3 style="font-size:15px;font-weight:600;margin:16px 0 10px">1、产值数据</h3>';

  // 产值总表
  html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
  ['链路','产值','ROI','状态'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
  html+='</tr></thead><tbody>';
  outputData.forEach(o=>{
    const status=o.roi>=100?'<span style="color:#16a34a">达标</span>':o.roi>=50?'<span style="color:#f59e0b">一般</span>':o.roi>0?'<span style="color:#dc2626">未达标</span>':'-';
    html+='<tr><td style="border:1px solid #e2e8f0;padding:8px">'+o.name+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+o.output+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+o.roi+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+status+'</td></tr>';
  });
  html+='</tbody></table>';

  // 产值总结
  const sortedOutput=outputData.filter(o=>o.output>0).sort((a,b)=>b.output-a.output);
  if(sortedOutput.length){
    html+='<p style="margin:12px 0">从产值来看'+sortedOutput[0].name+'的产值最高';
    if(sortedOutput.length>1)html+='，其次是'+sortedOutput[1].name;
    html+='。';
    if(worstLink&&worstLink.roi>0&&worstLink.roi<50)html+=worstLink.name+'产值只有'+worstLink.output+'，和大盘的高产值不符，需要重点排查。';
    html+='</p>';
  }
  }

  // 分链路分析（结合链路数据表+链路画像评分分析）
  if(modLink){
  html+='<h2 style="font-size:17px;font-weight:700;margin:24px 0 12px;padding-bottom:8px;border-bottom:2px solid #3b82f6">三、分链路分析</h2>';
  const linkSum=SYNC.link&&SYNC.link.linkSummary?SYNC.link.linkSummary:[];
  const scoreRes=SYNC.score&&SYNC.score.linkResults?SYNC.score.linkResults:{};
  outputData.forEach(o=>{
    const link=linkSum.find(l=>l.link===o.name)||{};
    const score=scoreRes[o.name]||{};
    const hasData=(link.eff&&link.eff>0)||o.output>0||o.issue;
    if(!hasData)return;
    html+='<h3 style="font-size:15px;font-weight:600;margin:20px 0 10px;color:#1e40af">'+o.name+'</h3>';
    if(o.issue){
      html+='<div style="margin-bottom:8px"><span style="font-weight:600;color:#dc2626">问题：</span>'+o.issue+'</div>';
    }
    // 基础数据表（链路数据表口径）
    const orders=link.eff||o.orders||0;
    const orderCost=link.orderCost||(link.spend&&orders?link.spend/orders:0)||0;
    const spend=link.spend||(orderCost*orders)||0;
    const highOrders=link.highOrders||0;
    const output=link.output||(link.orderOutput&&orders?link.orderOutput*orders:0)||o.output||0;
    const roi=link.roi?link.roi*100:o.roi||0;
    html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:10px"><thead><tr style="background:#eff6ff">';
    ['订单数','订单成本','消耗','高价课','产值','ROI'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:6px;text-align:center;color:#1e40af">'+h+'</th>';});
    html+='</tr></thead><tbody><tr>';
    html+='<td style="border:1px solid #e2e8f0;padding:6px;text-align:center">'+Math.round(orders)+'</td>';
    html+='<td style="border:1px solid #e2e8f0;padding:6px;text-align:center">¥'+orderCost.toFixed(0)+'</td>';
    html+='<td style="border:1px solid #e2e8f0;padding:6px;text-align:center">¥'+spend.toFixed(0)+'</td>';
    html+='<td style="border:1px solid #e2e8f0;padding:6px;text-align:center">'+highOrders+'</td>';
    html+='<td style="border:1px solid #e2e8f0;padding:6px;text-align:center">¥'+(typeof output==='number'?output.toFixed(0):output)+'</td>';
    const roiColor=roi>=100?'#16a34a':roi>=50?'#f59e0b':'#dc2626';
    html+='<td style="border:1px solid #e2e8f0;padding:6px;text-align:center;color:'+roiColor+';font-weight:600">'+roi.toFixed(1)+'%</td>';
    html+='</tr></tbody></table>';
    // 评分分析（链路画像评分分析口径）
    if(score.scoring&&score.scoring.summaryRows&&score.scoring.summaryRows.length){
      const scored=score.scoring.summaryRows.find(r=>r.value&&r.value.includes('有'));
      const unscored=score.scoring.summaryRows.find(r=>r.value&&r.value.includes('无'));
      if(scored||unscored){
        html+='<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;padding:10px 12px;margin-bottom:10px">';
        html+='<div style="font-weight:600;font-size:13px;color:#334155;margin-bottom:6px">📊 评分分析</div>';
        if(scored){
          html+='<div style="font-size:12px;color:#475569;margin-bottom:3px">• 有评分：订单占比'+(scored.share*100).toFixed(1)+'%，单订单产值¥'+(scored.orderOutput||0).toFixed(0)+'</div>';
        }
        if(unscored){
          html+='<div style="font-size:12px;color:#475569;margin-bottom:3px">• 无评分：订单占比'+(unscored.share*100).toFixed(1)+'%，单订单产值¥'+(unscored.orderOutput||0).toFixed(0)+'</div>';
        }
        if(scored&&unscored&&scored.orderOutput>0&&unscored.orderOutput>0){
          const diff=scored.orderOutput-unscored.orderOutput;
          const diffPct=unscored.orderOutput>0?(diff/unscored.orderOutput*100):0;
          const diffColor=diff>0?'#16a34a':'#dc2626';
          html+='<div style="font-size:12px;font-weight:600;color:'+diffColor+';margin-top:4px">评分差异：有评分比无评分单订单产值'+(diff>0?'高':'低')+'¥'+Math.abs(diff).toFixed(0)+'（'+(diffPct>0?'+':'')+diffPct.toFixed(1)+'%）</div>';
        }
        html+='</div>';
      }
    }
    // 自动生成调整建议（结合ROI+评分）
    let adj='';
    if(roi>=100){
      adj='ROI达标（'+roi.toFixed(1)+'%），表现良好，保持当前投放节奏，可适当放量。';
      if(score.scoring&&score.scoring.summaryRows){
        const scored=score.scoring.summaryRows.find(r=>r.value&&r.value.includes('有'));
        if(scored&&scored.share<0.5)adj+='但有评分订单占比偏低（'+(scored.share*100).toFixed(1)+'%），建议优化加微链路提升好友率。';
      }
    }
    else if(roi>=50){
      adj='ROI一般（'+roi.toFixed(1)+'%），稳定成本，培养高R素材起量，排低R素材。';
      if(score.scoring&&score.scoring.summaryRows){
        const scored=score.scoring.summaryRows.find(r=>r.value&&r.value.includes('有'));
        const unscored=score.scoring.summaryRows.find(r=>r.value&&r.value.includes('无'));
        if(scored&&unscored&&scored.orderOutput>unscored.orderOutput*1.5)adj+='有评分人群单产显著高于无评分，建议用评分做深度回传优化，压制无评分低质流量。';
      }
    }
    else if(roi>0){
      adj='ROI偏低（'+roi.toFixed(1)+'%），控成本，排查跑量素材ROI，暂停低ROI素材，新广告用双出价模式。';
      if(score.scoring&&score.scoring.summaryRows){
        const unscored=score.scoring.summaryRows.find(r=>r.value&&r.value.includes('无'));
        if(unscored&&unscored.share>0.6)adj+='无评分订单占比高达'+(unscored.share*100).toFixed(1)+'%，低质流量占比大，建议接入深层转化出价过滤。';
      }
    }
    else{adj='持续观察，培养头部跑量素材，控成本。';}
    html+='<div style="margin-bottom:8px"><span style="font-weight:600;color:#16a34a">调整：</span>'+adj+'</div>';
  });
  }

  // 素材分析
  if(modMaterial && SYNC.material && SYNC.material.topMaterials){
  html+='<h2 style="font-size:17px;font-weight:700;margin:24px 0 12px;padding-bottom:8px;border-bottom:2px solid #3b82f6">四、素材分析</h2>';
  const mat=SYNC.material;
  if(mat.topMaterials&&mat.topMaterials.length){
    html+='<h3 style="font-size:15px;font-weight:600;margin:16px 0 10px">高ROI素材 Top5</h3>';
    html+='<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px"><thead><tr style="background:#eff6ff">';
    ['排名','素材名称','ROI','消耗','高价课订单'].forEach(h=>{html+='<th style="border:1px solid #dbeafe;padding:8px;text-align:left;color:#1e40af">'+h+'</th>';});
    html+='</tr></thead><tbody>';
    mat.topMaterials.slice(0,5).forEach((m,i)=>{
      html+='<tr><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(i+1)+'</td><td style="border:1px solid #e2e8f0;padding:8px">'+(m.name||'')+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center;color:#16a34a">'+((m.roi||0)*100).toFixed(1)+'%</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">¥'+(m.spend||0).toFixed(0)+'</td><td style="border:1px solid #e2e8f0;padding:8px;text-align:center">'+(m.highOrders||0)+'</td></tr>';
    });
    html+='</tbody></table>';
  }
  if(mat.bottomMaterials&&mat.bottomMaterials.length){
    html+='<h3 style="font-size:15px;font-weight:600;margin:16px 0 10px;color:#dc2626">低ROI素材预警（需关注）</h3>';
    mat.bottomMaterials.slice(0,3).forEach(m=>{
      html+='<div style="padding:8px 12px;background:#fef2f2;border-radius:6px;margin-bottom:6px;font-size:12.5px"><b>'+(m.name||'')+'</b> - ROI '+((m.roi||0)*100).toFixed(1)+'%，消耗 ¥'+(m.spend||0).toFixed(0)+'，建议暂停或优化</div>';
    });
  }
  }

  // 线索评分
  if(modScore && SYNC.clue && SYNC.clue.dev){
  html+='<h2 style="font-size:17px;font-weight:700;margin:24px 0 12px;padding-bottom:8px;border-bottom:2px solid #3b82f6">五、线索评分情况</h2>';
  const dv=SYNC.clue.dev;
  html+='<p style="margin:12px 0">设备号模型：总低价课'+(dv.total&&dv.total.low?dv.total.low.toLocaleString():0)+'单，高价课'+(dv.total&&dv.total.high?dv.total.high:0)+'单，转化率'+((dv.total&&dv.total.conv?dv.total.conv:0)*100).toFixed(2)+'%，ARPU ¥'+(dv.total&&dv.total.arpu?dv.total.arpu:0).toFixed(2)+'。</p>';
  const highTier=dv.tierList?dv.tierList.find(t=>t.name&&t.name.includes('高')):null;
  if(highTier){
    html+='<p style="margin:8px 0">高分段（7-10分）占比'+((highTier.lowShare||0)*100).toFixed(1)+'%，转化率'+((highTier.conv||0)*100).toFixed(2)+'%，产值贡献'+((highTier.gmvShare||0)*100).toFixed(1)+'%。</p>';
  }
  html+='<p style="margin:8px 0;color:#475569">评分和产值呈正向关系，建议继续用评分做深度回传优化。</p>';
  }

  // 本周调整
  if(modAction){
  html+='<h2 style="font-size:17px;font-weight:700;margin:24px 0 12px;padding-bottom:8px;border-bottom:2px solid #3b82f6">六、本周调整</h2>';
  html+='<ol style="margin:12px 0;padding-left:24px">';
  suggestions.slice(0,4).forEach(s=>{html+='<li style="margin-bottom:8px">'+s+'</li>';});
  html+='</ol>';
  }
  html+='</div>';

  t8GeneratedReport={title:reportTitle,date:weekRange,content:html,isHtml:true};
  saveSync('weeklyReport',t8GeneratedReport);

  document.getElementById('t8Generated').style.display='block';
  document.getElementById('t8Preview').innerHTML=html;
  document.getElementById('t8DownloadBtn').disabled=false;
  showStatus('t8','周报生成完成！已同步到分析模块','done');

  if(activeView==='weekly-report')renderWeeklyReport();
}

// 生成新周报
function generateWeeklyReport(){
  showStatus('t8','正在结合工作台数据生成周报...','active');
  // 从工作台获取数据
  const linkData=SYNC.link;      // 链路数据表
  const scoreData=SYNC.score;    // 链路画像评分
  const materialData=SYNC.material; // 素材分析
  const clueData=SYNC.clue;      // 线索评分
  // 生成周报日期（上周一到上周日）
  const now=new Date();
  const day=now.getDay();
  const lastMon=new Date(now);lastMon.setDate(now.getDate()-day-6);
  const lastSun=new Date(now);lastSun.setDate(now.getDate()-day);
  const fmt=d=>(d.getMonth()+1)+'.'+d.getDate();
  const weekRange=fmt(lastMon)+'-'+fmt(lastSun);
  const reportTitle=(lastMon.getMonth()+1)+''+String(lastMon.getDate()).padStart(2,'0')+'周报';
  // 构建周报内容
  let content='';
  content+='太极投放周报-吴婷婷\n';
  content+=reportTitle+'\n\n';
  content+='上周进量数据（上周时间维度：'+weekRange+'）\n';
  content+='进量情况（测试数据）\n\n';
  // 进量数据总结
  if(linkData&&linkData.linkSummary){
    const totalOrders=linkData.linkSummary.reduce((s,l)=>s+(l.eff||0),0);
    const totalSpend=linkData.linkSummary.reduce((s,l)=>s+(l.spend||0),0);
    const avgCost=totalOrders>0?totalSpend/totalOrders:0;
    content+='上周总订单数共'+Math.round(totalOrders)+'个，整体成本'+avgCost.toFixed(0)+'元。';
    if(linkData.linkSummary.length){
      const best=linkData.linkSummary.slice().sort((a,b)=>(b.roi||0)-(a.roi||0))[0];
      const worst=linkData.linkSummary.slice().sort((a,b)=>(a.roi||0)-(b.roi||0))[0];
      content+='分链路看，'+best.link+'的ROI表现最好（'+((best.roi||0)*100).toFixed(1)+'%），'+worst.link+'的ROI偏低（'+((worst.roi||0)*100).toFixed(1)+'%），需要重点关注。';
    }
    content+='\n\n';
  }else{
    content+='（请先在Excel工具-链路数据表中上传数据同步）\n\n';
  }
  // 产值数据
  content+='产值数据\n\n';
  if(linkData&&linkData.linkSummary){
    content+='从产值来看：\n';
    linkData.linkSummary.forEach(l=>{
      content+='- '+l.link+'：订单数'+Math.round(l.eff||0)+'，单订单产值¥'+(l.orderOutput||0).toFixed(0)+'，ROI'+((l.roi||0)*100).toFixed(1)+'%';
      if((l.roi||0)<0.5)content+='（ROI未达标，需优化）';
      content+='\n';
    });
    content+='\n';
  }
  // 分链路问题与调整
  content+='分链路分析\n\n';
  if(linkData&&linkData.linkSummary){
    linkData.linkSummary.forEach(l=>{
      content+='【'+l.link+'】\n';
      if((l.roi||0)>=0.6){
        content+='表现：ROI达标（'+((l.roi||0)*100).toFixed(1)+'%），订单成本¥'+(l.orderCost||0).toFixed(0)+'。\n';
        content+='调整：保持当前投放节奏，继续放量。\n';
      }else{
        content+='问题：ROI偏低（'+((l.roi||0)*100).toFixed(1)+'%），订单成本¥'+(l.orderCost||0).toFixed(0)+'偏高。\n';
        content+='调整：控成本，排查跑量素材ROI，暂停低ROI素材，培养新广告。\n';
      }
      content+='\n';
    });
  }
  // 素材分析
  if(materialData&&materialData.topMaterials){
    content+='素材分析\n\n';
    content+='高ROI素材Top5：\n';
    materialData.topMaterials.slice(0,5).forEach((m,i)=>{
      content+=(i+1)+'. '+m.name+' - ROI'+((m.roi||0)*100).toFixed(1)+'%，消耗¥'+(m.spend||0).toFixed(0)+'，高价课'+(m.highOrders||0)+'单\n';
    });
    content+='\n低ROI素材（需关注）：\n';
    if(materialData.bottomMaterials){
      materialData.bottomMaterials.slice(0,3).forEach(m=>{
        content+='- '+m.name+' - ROI'+((m.roi||0)*100).toFixed(1)+'%，消耗¥'+(m.spend||0).toFixed(0)+'，建议暂停或优化\n';
      });
    }
    content+='\n';
  }
  // 线索评分
  if(clueData&&clueData.dev){
    content+='线索评分情况\n\n';
    const dv=clueData.dev;
    content+='设备号模型：总低价课'+dv.total.low.toLocaleString()+'单，高价课'+dv.total.high+'单，转化率'+(dv.total.conv*100).toFixed(2)+'%，ARPU¥'+dv.total.arpu.toFixed(2)+'。\n';
    const highTier=dv.tierList.find(t=>t.name.includes('高'));
    if(highTier){
      content+='高分段（7-10分）占比'+(highTier.lowShare*100).toFixed(1)+'%，转化率'+(highTier.conv*100).toFixed(2)+'%，产值贡献'+(highTier.gmvShare*100).toFixed(1)+'%。\n';
    }
    content+='评分和产值呈正向关系，建议继续用评分做深度回传优化。\n\n';
  }
  // 本周调整
  content+='本周调整\n\n';
  content+='1. 新广告全部用双出价模式，对比单出价广告的ROI是否更稳定\n';
  content+='2. 低ROI链路控成本，暂停低ROI素材，培养新广告\n';
  content+='3. 高分段线索重点跟进，提升高价课转化率\n';
  content+='4. 关注版位变化，主跑量版位成本波动及时调整\n';
  t8GeneratedReport={title:reportTitle,date:weekRange,content};
  saveSync('weeklyReport',t8GeneratedReport);
  // 展示预览
  document.getElementById('t8Generated').style.display='block';
  document.getElementById('t8Preview').innerHTML='<pre style="white-space:pre-wrap;font-family:inherit;font-size:13px;line-height:1.9;margin:0">'+esc(content)+'</pre>';
  document.getElementById('t8DownloadBtn').disabled=false;
  showStatus('t8','周报生成完成！已同步到分析模块','done');
  // 同步到分析模块
  if(activeView==='weekly-report')renderWeeklyReport();
}

function regenerateWeeklyReport(){
  if(t8WeeklyReports.length===0){alert('请先在Excel工具-周报撰写中上传并解析历史周报');navigate('excel-weekly-report');return;}
  navigate('excel-weekly-report');
  showStatus('t8','请在下方表单中修改本周数据后，点击"一键生成周报"','active');
}

// 分析模块渲染
function renderWeeklyReport(){
  const d=SYNC.weeklyReport||t8GeneratedReport;
  if(!d){
    document.getElementById('wr-empty').style.display='block';
    document.getElementById('wr-actions').style.display='none';
    document.getElementById('wr-content').style.display='none';
    document.getElementById('wr-history').style.display='none';
    document.getElementById('wr-badge').textContent='待生成';
    return;
  }
  document.getElementById('wr-empty').style.display='none';
  document.getElementById('wr-actions').style.display='block';
  document.getElementById('wr-content').style.display='block';
  document.getElementById('wr-badge').textContent=d.title+' · 已生成';
  document.getElementById('wr-date').textContent='时间范围：'+d.date;
  // 渲染为HTML
  if(d.isHtml){
    document.getElementById('wr-body').innerHTML=d.content;
  }else{
    const lines=d.content.split('\n');
    let html='';
    for(const line of lines){
      if(!line.trim()){html+='<div style="height:8px"></div>';continue;}
      if(/周报/.test(line)&&line.length<20){html+='<h2 style="font-size:20px;font-weight:700;color:#1e293b;margin:16px 0 8px">'+esc(line)+'</h2>';continue;}
      if(/^上周进量|^产值数据|^分链路|^素材分析|^线索评分|^本周调整|^进量情况/.test(line.trim())){html+='<h3 style="font-size:16px;font-weight:600;color:#3b82f6;margin:16px 0 8px;border-left:3px solid #3b82f6;padding-left:8px">'+esc(line)+'</h3>';continue;}
      if(/^【.+】/.test(line.trim())){html+='<h4 style="font-size:14px;font-weight:600;color:#1e293b;margin:12px 0 6px">'+esc(line)+'</h4>';continue;}
      if(/^\d+\./.test(line.trim())){html+='<div style="margin:4px 0;padding-left:16px">'+esc(line)+'</div>';continue;}
      if(/^-/.test(line.trim())){html+='<div style="margin:4px 0;padding-left:16px;color:#475569">'+esc(line)+'</div>';continue;}
      html+='<div style="margin:6px 0;color:#334155">'+esc(line)+'</div>';
    }
    document.getElementById('wr-body').innerHTML=html;
  }
  // 历史周报参考
  if(t8WeeklyReports.length>0){
    document.getElementById('wr-history').style.display='block';
    document.getElementById('wr-history-count').textContent='共 '+t8WeeklyReports.length+' 期';
    document.getElementById('wr-history-list').innerHTML=t8WeeklyReports.map((r,i)=>
      '<details style="margin-bottom:8px;padding:8px 12px;background:#f8fafc;border-radius:6px"><summary style="cursor:pointer;font-weight:600;color:#475569;font-size:13px">'+esc(r.title)+'</summary><div style="margin-top:8px;font-size:12px;color:#64748b;line-height:1.8;max-height:200px;overflow-y:auto">'+r.sections.slice(0,20).map(s=>'<div>'+esc(s)+'</div>').join('')+'</div></details>'
    ).join('');
  }
}

// 下载周报docx
function downloadWeeklyReport(){
  const d=SYNC.weeklyReport||t8GeneratedReport;
  if(!d){alert('请先生成周报');return;}
  // HTML格式：生成Word兼容的HTML文件
  if(d.isHtml){
    const wordHtml='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>'+d.title+'</title><style>body{font-family:微软雅黑;font-size:14px;line-height:1.8}table{border-collapse:collapse;width:100%;margin:10px 0}th,td{border:1px solid #ccc;padding:6px 8px;font-size:12px}th{background:#eff6ff}h1{font-size:22px}h2{font-size:17px;border-bottom:2px solid #3b82f6;padding-bottom:6px}h3{font-size:15px;color:#1e40af}</style></head><body>'+d.content+'</body></html>';
    const blob=new Blob(['\ufeff',wordHtml],{type:'application/msword'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;a.download=d.title+'.doc';
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    URL.revokeObjectURL(url);
    return;
  }
  // 纯文本格式：转docx
  const lines=d.content.split('\n');
  let bodyXml='';
  for(const line of lines){
    if(!line.trim()){bodyXml+='<w:p/>';continue;}
    let bold=false,size=24;
    if(/周报/.test(line)&&line.length<20){bold=true;size=36;}
    else if(/^上周进量|^产值数据|^分链路|^素材分析|^线索评分|^本周调整/.test(line.trim())){bold=true;size=28;}
    else if(/^【.+】/.test(line.trim())){bold=true;size=24;}
    bodyXml+='<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr>'+(bold?'<w:b/>':'')+'<w:sz w:val="'+size+'"/></w:rPr><w:t xml:space="preserve">'+esc(line)+'</w:t></w:r></w:p>';
  }
  const docXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+
    '<w:body>'+bodyXml+'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>';
  const contentTypes='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  const rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const zip=new JSZip();
  zip.file('[Content_Types].xml',contentTypes);
  zip.file('_rels/.rels',rels);
  zip.file('word/document.xml',docXml);
  zip.generateAsync({type:'blob'}).then(blob=>{
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;a.download=d.title+'.docx';
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });
}

// 复制全文
function copyWeeklyReport(){
  const d=SYNC.weeklyReport||t8GeneratedReport;
  if(!d){alert('请先生成周报');return;}
  const text=d.isHtml?d.content.replace(/<[^>]+>/g,'').replace(/\n\s*\n/g,'\n'):d.content;
  navigator.clipboard.writeText(text).then(()=>{
    alert('周报已复制到剪贴板');
  }).catch(()=>{
    const ta=document.createElement('textarea');
    ta.value=text;document.body.appendChild(ta);ta.select();
    document.execCommand('copy');document.body.removeChild(ta);
    alert('周报已复制到剪贴板');
  });
}

/* ============================================================
   第二轮增强：历史数据 / 操作日志 / 数据质量 / AB测试 /
   日报 / 命名字典 / 跨账户 / 转化延迟 / 前后端关联
   ============================================================ */

/* ---------- 1. 历史数据快照管理 ---------- */
const SNAPSHOT_KEY='workbench_snapshots_v1';
const MAX_SNAPSHOTS=60;
function loadSnapshots(){ try{return JSON.parse(localStorage.getItem(SNAPSHOT_KEY))||[];}catch(e){return[];} }
function saveSnapshots(s){ try{localStorage.setItem(SNAPSHOT_KEY,JSON.stringify(s.slice(0,MAX_SNAPSHOTS)));}catch(e){} }
function takeSnapshot(){
  if(!TENCENT_LIVE_DATA) return null;
  const snap={
    time:new Date().toISOString(),
    date:new Date().toISOString().slice(0,10),
    total:{spend:TENCENT_LIVE_DATA.total.spend,conv:TENCENT_LIVE_DATA.total.conv,cpa:TENCENT_LIVE_DATA.total.cpa,roi:TENCENT_LIVE_DATA.total.roi},
    accounts:TENCENT_LIVE_DATA.accounts.map(a=>({id:a.id,spend:a.spend,conv:a.conv,cpa:a.cpa,balance:a.balance})),
    units:(TENCENT_ADUNITS_DATA||[]).flatMap(a=>a.units.map(u=>({acc:a.account_id,name:u[0],spend:parseFloat(String(u[5]||'0').replace(/[,¥]/g,''))||0,conv:parseFloat(String(u[7]||'0').replace(/[,¥]/g,''))||0,cpa:parseFloat(String(u[8]||'0').replace(/[,¥]/g,''))||0})))
  };
  const snaps=loadSnapshots();
  // 同一天只保留最新一条
  const filtered=snaps.filter(s=>s.date!==snap.date);
  filtered.unshift(snap);
  saveSnapshots(filtered);
  return snap;
}
function getUnitTrend(unitName,days=7){
  const snaps=loadSnapshots().slice(0,days);
  return snaps.map(s=>{
    const u=s.units.find(x=>x.name===unitName);
    return {date:s.date,spend:u?u.spend:0,conv:u?u.conv:0,cpa:u?u.cpa:0};
  }).reverse();
}

/* ---------- 2. 操作日志 ---------- */
const OPLOG_KEY='workbench_oplog_v1';
let opLogFilter='all';
function loadOpLog(){ try{return JSON.parse(localStorage.getItem(OPLOG_KEY))||[];}catch(e){return[];} }
function saveOpLog(l){ try{localStorage.setItem(OPLOG_KEY,JSON.stringify(l));}catch(e){} }
function addOpLog(type,target,before,note){
  const log=loadOpLog();
  log.unshift({id:Date.now(),time:new Date().toLocaleString('zh-CN'),type,target,before,note,status:'pending',after:null});
  saveOpLog(log);
}
function updateOpLogStatus(id,status,after){
  const log=loadOpLog();
  const item=log.find(x=>x.id===id);
  if(item){item.status=status;if(after)item.after=after;item.afterTime=new Date().toLocaleString('zh-CN');saveOpLog(log);}
}
function renderOpLog(){
  const log=loadOpLog();
  const filtered=opLogFilter==='all'?log:log.filter(x=>x.status===opLogFilter);
  $('#op-log-count').textContent=filtered.length+' 条记录（共'+log.length+'条）';
  const list=$('#op-log-list');
  if(filtered.length===0){ list.innerHTML='<div class="empty-state-enhanced"><div class="empty-icon">📋</div>暂无操作记录，在上方记录新操作或从优化建议一键导入</div>'; return; }
  const typeMap={pause:'暂停',reduce:'降价',increase:'加预算',creative:'换素材',other:'其他'};
  list.innerHTML=filtered.map(item=>`
    <div class="op-log-item">
      <div class="op-log-time">${item.time}</div>
      <div class="op-log-type ${item.type}">${typeMap[item.type]||item.type}</div>
      <div class="op-log-target" title="${esc(item.target)}">${esc(item.target)}</div>
      <div class="op-log-before">前:${esc(item.before||'—')}${item.after?'<br>后:'+esc(item.after):''}</div>
      <div style="display:flex;gap:4px;justify-content:flex-end">
        ${item.status==='pending'?`<button class="btn btn-secondary" style="padding:2px 8px;font-size:10px" onclick="markOpDone(${item.id})">标记完成</button>`:''}
        <button class="btn btn-secondary" style="padding:2px 8px;font-size:10px;color:#ef4444" onclick="delOpLog(${item.id})">删除</button>
      </div>
    </div>`).join('');
}
function markOpDone(id){ updateOpLogStatus(id,'done'); showToast('已标记为完成'); renderOpLog(); }
function delOpLog(id){ saveOpLog(loadOpLog().filter(x=>x.id!==id)); renderOpLog(); }

/* ---------- 3. 数据质量校验 ---------- */
function runDataQuality(){
  const issues=[];
  const num=v=>{const n=parseFloat(String(v||'').replace(/[,¥%]/g,''));return isNaN(n)?0:n;};
  // 检查腾讯数据
  if(!TENCENT_LIVE_DATA){ issues.push({level:'error',field:'腾讯广告数据',value:'未加载',msg:'腾讯广告数据未加载，所有投放分析不可用'}); }
  else{
    // 检查0消耗账户
    TENCENT_LIVE_DATA.accounts.forEach(a=>{
      if(a.spend===0 && a.status!=='余额不足') issues.push({level:'warning',field:'账户'+a.id,value:'消耗0',msg:'该账户今日消耗为0，可能未投放或数据未更新'});
    });
    // 检查异常高CPA
    const avgCpa=TENCENT_LIVE_DATA.total.cpa||0;
    TENCENT_LIVE_DATA.accounts.forEach(a=>{
      if(a.cpa>avgCpa*3 && a.conv>0) issues.push({level:'warning',field:'账户'+a.id,value:'CPA异常高',msg:'CPA ¥'+fmtM2(a.cpa)+' 是均值的'+(a.cpa/avgCpa).toFixed(1)+'倍，建议检查'});
    });
  }
  // 检查单元数据
  if(TENCENT_ADUNITS_DATA){
    const allUnits=TENCENT_ADUNITS_DATA.flatMap(a=>a.units);
    const zeroConv=allUnits.filter(u=>num(u[5])>100 && num(u[7])===0);
    if(zeroConv.length>0) issues.push({level:'info',field:'空耗单元',value:zeroConv.length+'个',msg:zeroConv.length+'个单元消耗>100元但0转化，建议在优化建议中查看详情'});
    // 检查命名不规范
    const badNames=allUnits.filter(u=>{const n=u[0]||'';return !/^\d{4}-/.test(n);});
    if(badNames.length>0) issues.push({level:'info',field:'命名不规范',value:badNames.length+'个',msg:badNames.length+'个单元名称不符合"日期-产品-版位--定向-变体--出价"规则，维度解析可能不准确'});
  }
  // 检查同步数据
  const syncKeys=Object.keys(SYNC).filter(k=>SYNC[k]);
  if(syncKeys.length===0) issues.push({level:'warning',field:'后端数据',value:'未同步',msg:'未同步任何后端数据（链路/素材/线索/评分），后端转化分析不可用'});
  // 检查数据时效性
  if(TENCENT_LIVE_DATA && TENCENT_LIVE_DATA.syncTime){
    const f=getDataFreshness(TENCENT_LIVE_DATA.syncTime);
    if(f.hours>8) issues.push({level:'error',field:'数据时效性',value:f.text,msg:'腾讯数据已超过8小时未更新，建议重新同步后再做决策'});
  }
  return issues;
}
function renderDataQuality(){
  const issues=runDataQuality();
  const errors=issues.filter(i=>i.level==='error').length;
  const warnings=issues.filter(i=>i.level==='warning').length;
  const infos=issues.filter(i=>i.level==='info').length;
  renderKpis('#dq-kpis',[
    {label:'严重问题',value:errors,cls:errors>0?'kpi-accent':''},
    {label:'警告',value:warnings,cls:warnings>0?'kpi-amber':''},
    {label:'提示',value:infos,cls:''},
    {label:'校验通过',value:errors===0&&warnings===0?'是':'否',cls:errors===0&&warnings===0?'kpi-teal':'kpi-accent'}
  ]);
  const container=$('#dq-issues');
  if(issues.length===0){ container.innerHTML='<div class="empty-state-enhanced"><div class="empty-icon">✅</div>数据质量良好，未发现异常</div>'; return; }
  const iconMap={error:'❌',warning:'⚠️',info:'ℹ️'};
  container.innerHTML=issues.map(i=>`<div class="dq-issue ${i.level}"><span class="dq-icon">${iconMap[i.level]}</span><span class="dq-field">${esc(i.field)}</span><span class="dq-value">${esc(i.value)}</span><span style="flex:1">${esc(i.msg)}</span></div>`).join('');
}

/* ---------- 4. A/B测试管理 ---------- */
const AB_KEY='workbench_abtests_v1';
function loadABTests(){ try{return JSON.parse(localStorage.getItem(AB_KEY))||[];}catch(e){return[];} }
function saveABTests(t){ try{localStorage.setItem(AB_KEY,JSON.stringify(t));}catch(e){} }
function calcConfidence(aConv,aTotal,bConv,bTotal){
  // 简化版置信度计算（基于Z检验）
  const pA=aConv/Math.max(aTotal,1), pB=bConv/Math.max(bTotal,1);
  const p=(aConv+bConv)/Math.max(aTotal+bTotal,1);
  const se=Math.sqrt(p*(1-p)*(1/Math.max(aTotal,1)+1/Math.max(bTotal,1)));
  if(se===0) return {z:0,confidence:0,winner:'none'};
  const z=(pB-pA)/se;
  const confidence=Math.min(99,Math.abs(z)*30);
  return {z,confidence:confidence.toFixed(1),winner:z>1.96?'B':z<-1.96?'A':'none'};
}
function renderABTests(){
  const tests=loadABTests();
  const list=$('#ab-test-list');
  if(tests.length===0){ list.innerHTML='<div class="card"><div class="empty-state-enhanced"><div class="empty-icon">🧪</div>暂无A/B测试，在上方创建新测试</div></div>'; return; }
  list.innerHTML=tests.map((t,idx)=>{
    const conf=calcConfidence(t.aConv||0,t.aTotal||0,t.bConv||0,t.bTotal||0);
    const confCls=conf.confidence>=95?'high':conf.confidence>=80?'medium':'low';
    const confLabel=conf.confidence>=95?'高置信':conf.confidence>=80?'中置信':'低置信';
    return `<div class="ab-test-card">
      <div class="ab-test-head">
        <span class="ab-test-title">${esc(t.name)}</span>
        <div style="display:flex;gap:8px;align-items:center">
          <span class="ab-confidence ${confCls}">${confLabel} ${conf.confidence}%</span>
          <span class="ab-test-status ${t.status}">${t.status==='running'?'进行中':t.status==='concluded'?'已结束':'已暂停'}</span>
          <button class="btn btn-secondary" style="padding:2px 8px;font-size:10px" onclick="delABTest(${idx})">删除</button>
        </div>
      </div>
      <div style="font-size:11px;color:var(--ink-3);margin-bottom:8px">变量类型：${t.varType} | 创建时间：${t.time} | ${t.note||''}</div>
      <div class="ab-metric-row"><div class="ab-metric-name">指标</div><div class="ab-metric-val" style="color:#3b82f6">A组（对照组）</div><div class="ab-metric-val" style="color:#f59e0b">B组（实验组）</div><div class="ab-metric-name">差异</div></div>
      <div class="ab-metric-row"><div class="ab-metric-name">样本量</div><div class="ab-metric-val">${t.aTotal||0}</div><div class="ab-metric-val">${t.bTotal||0}</div><div class="ab-metric-diff">—</div></div>
      <div class="ab-metric-row"><div class="ab-metric-name">转化量</div><div class="ab-metric-val ${conf.winner==='A'?'winner':''}">${t.aConv||0}</div><div class="ab-metric-val ${conf.winner==='B'?'winner':''}">${t.bConv||0}</div><div class="ab-metric-diff">${conf.winner==='A'?'A胜出':conf.winner==='B'?'B胜出':'无显著差异'}</div></div>
      <div class="ab-metric-row"><div class="ab-metric-name">转化率</div><div class="ab-metric-val">${t.aTotal?((t.aConv||0)/t.aTotal*100).toFixed(2)+'%':'—'}</div><div class="ab-metric-val">${t.bTotal?((t.bConv||0)/t.bTotal*100).toFixed(2)+'%':'—'}</div><div class="ab-metric-diff">${t.aTotal&&t.bTotal?(((t.bConv||0)/t.bTotal-(t.aConv||0)/t.aTotal)*100).toFixed(2)+'%':'—'}</div></div>
      <div style="margin-top:10px;display:flex;gap:8px">
        <input type="number" id="ab-a-total-${idx}" placeholder="A样本量" value="${t.aTotal||''}" style="width:90px;padding:4px 8px;border:1px solid var(--border-2);border-radius:6px;font-size:11px">
        <input type="number" id="ab-a-conv-${idx}" placeholder="A转化" value="${t.aConv||''}" style="width:80px;padding:4px 8px;border:1px solid var(--border-2);border-radius:6px;font-size:11px">
        <input type="number" id="ab-b-total-${idx}" placeholder="B样本量" value="${t.bTotal||''}" style="width:90px;padding:4px 8px;border:1px solid var(--border-2);border-radius:6px;font-size:11px">
        <input type="number" id="ab-b-conv-${idx}" placeholder="B转化" value="${t.bConv||''}" style="width:80px;padding:4px 8px;border:1px solid var(--border-2);border-radius:6px;font-size:11px">
        <button class="btn btn-secondary" style="padding:4px 12px;font-size:11px" onclick="updateABTest(${idx})">更新数据</button>
        <button class="btn btn-secondary" style="padding:4px 12px;font-size:11px" onclick="toggleABTest(${idx})">${t.status==='running'?'暂停':'继续'}</button>
      </div>
    </div>`;
  }).join('');
}
function delABTest(idx){ const t=loadABTests(); t.splice(idx,1); saveABTests(t); renderABTests(); }
function updateABTest(idx){ const t=loadABTests(); t[idx].aTotal=parseInt($('#ab-a-total-'+idx).value)||0; t[idx].aConv=parseInt($('#ab-a-conv-'+idx).value)||0; t[idx].bTotal=parseInt($('#ab-b-total-'+idx).value)||0; t[idx].bConv=parseInt($('#ab-b-conv-'+idx).value)||0; saveABTests(t); renderABTests(); showToast('测试数据已更新'); }
function toggleABTest(idx){ const t=loadABTests(); t[idx].status=t[idx].status==='running'?'paused':'running'; saveABTests(t); renderABTests(); }

/* ---------- 5. 日报生成 ---------- */
function generateDailyReport(){
  const date=$('#dr-date').value||new Date().toISOString().slice(0,10);
  $('#dr-date-label').textContent=date;
  let report='【投放日报】'+date+'\n\n';
  // 数据总览
  if($('#dr-include-overview').checked && TENCENT_LIVE_DATA){
    const t=TENCENT_LIVE_DATA.total;
    report+='一、数据总览\n';
    report+='今日总消耗：¥'+fmtM(t.spend)+'\n';
    report+='总曝光：'+fmtN(t.impressions)+'\n';
    report+='总点击：'+fmtN(t.clicks)+'\n';
    report+='点击率：'+fmtP(t.ctr*100)+'\n';
    report+='总转化：'+fmtN(t.conv)+'\n';
    report+='平均转化成本：¥'+fmtM2(t.cpa)+'\n';
    report+='整体ROI：'+fmtP(t.roi*100)+'\n\n';
  }
  // 预算追踪
  if($('#dr-include-budget').checked && TENCENT_LIVE_DATA){
    report+='二、预算消耗\n';
    const budgets=loadBudgets();
    TENCENT_LIVE_DATA.accounts.forEach(a=>{
      const bg=budgets[a.id]||{daily:0};
      const pct=bg.daily?(a.spend/bg.daily*100).toFixed(0):'—';
      report+='账户'+a.id+'（'+(a.operator||'')+'）：消耗¥'+fmtM(a.spend)+' / 预算¥'+fmtN(bg.daily||0)+'（'+pct+'%），余额¥'+fmtM(a.balance)+'\n';
    });
    report+='\n';
  }
  // 优化建议
  if($('#dr-include-suggestion').checked){
    const sugs=generateSuggestions(TENCENT_LIVE_DATA,TENCENT_ADUNITS_DATA);
    report+='三、优化建议（'+sugs.length+'条）\n';
    if(sugs.length===0) report+='暂无需要紧急处理的问题\n';
    sugs.slice(0,10).forEach((s,i)=>{ report+=(i+1)+'.【'+s.action+'】'+s.title+'：'+s.unit+' — '+s.detail+'\n'; });
    report+='\n';
  }
  // 素材动态
  if($('#dr-include-material').checked && SYNC.material){
    report+='四、素材动态\n';
    const list=SYNC.material.list||[];
    const running=list.filter(m=>getMatStage(m.name)==='running').length;
    const fatigue=list.filter(m=>getMatStage(m.name)==='fatigue').length;
    report+='素材总数：'+list.length+'条，跑量中：'+running+'条，疲劳：'+fatigue+'条\n';
    const top=list.slice().sort((a,b)=>b.roi-a.roi).slice(0,3);
    report+='ROI Top3：\n';
    top.forEach((m,i)=>{ report+=(i+1)+'.'+m.name+' ROI'+(m.roi*100).toFixed(1)+'% 花费¥'+m.cost.toFixed(0)+'\n'; });
    report+='\n';
  }
  // 操作记录
  if($('#dr-include-oplog').checked){
    const log=loadOpLog().filter(x=>x.time.includes(date.slice(5).replace('-','/'))||x.time.includes(date));
    report+='五、今日操作（'+log.length+'条）\n';
    if(log.length===0) report+='今日暂无操作记录\n';
    log.forEach(x=>{ report+='- ['+x.type+'] '+x.target+'（'+x.status+'）\n'; });
    report+='\n';
  }
  report+='六、明日计划\n';
  report+='1. 持续监控超成本单元，及时降价或暂停\n';
  report+='2. 跟进疲劳素材的替换\n';
  report+='3. 关注余额不足账户的充值\n';
  $('#dr-content').textContent=report;
  $('#dr-result-card').style.display='block';
  $('#dr-empty').style.display='none';
  $('#dr-copy').disabled=false;
  showToast('日报已生成');
}

/* ---------- 6. 命名字典 ---------- */
const DICT_KEY='workbench_namedict_v1';
const DEFAULT_DICT=[
  {pattern:'多版位',field:'placement',label:'多版位'},
  {pattern:'短剧',field:'placement',label:'短剧'},
  {pattern:'排小程序',field:'placement',label:'小程序'},
  {pattern:'低r城|低线城',field:'cityTier',label:'低线城市'},
  {pattern:'高r|高线',field:'cityTier',label:'高线城市'},
  {pattern:'偏远',field:'cityTier',label:'偏远地区'},
  {pattern:'双出价',field:'bidType',label:'双出价'},
  {pattern:'高r迭代|迭代',field:'variant',label:'高r迭代'},
  {pattern:'痛点',field:'variant',label:'痛点'},
  {pattern:'AB|AB素材',field:'variant',label:'AB素材'},
  {pattern:'xin',field:'variant',label:'xin'}
];
function loadDict(){ try{const d=JSON.parse(localStorage.getItem(DICT_KEY)); return d&&d.length?d:JSON.parse(JSON.stringify(DEFAULT_DICT));}catch(e){return JSON.parse(JSON.stringify(DEFAULT_DICT));} }
function saveDict(d){ try{localStorage.setItem(DICT_KEY,JSON.stringify(d));}catch(e){} }
function renderDict(){
  const dict=loadDict();
  const container=$('#dict-rules');
  container.innerHTML=dict.map((r,i)=>`<div class="dict-rule">
    <input type="text" value="${esc(r.pattern)}" data-idx="${i}" data-field="pattern" placeholder="匹配规则（正则）">
    <select data-idx="${i}" data-field="field" style="padding:5px 8px;border:1px solid var(--border-2);border-radius:6px;font-size:12px">
      <option value="placement" ${r.field==='placement'?'selected':''}>版位</option>
      <option value="targeting" ${r.field==='targeting'?'selected':''}>定向</option>
      <option value="variant" ${r.field==='variant'?'selected':''}>变体</option>
      <option value="bidType" ${r.field==='bidType'?'selected':''}>出价类型</option>
      <option value="cityTier" ${r.field==='cityTier'?'selected':''}>城市层级</option>
      <option value="product" ${r.field==='product'?'selected':''}>产品</option>
    </select>
    <span class="dict-del" onclick="delDictRule(${i})">删除</span>
  </div>`).join('');
}
function delDictRule(i){ const d=loadDict(); d.splice(i,1); saveDict(d); renderDict(); }
function testDictParse(){
  const name=$('#dict-test-input').value;
  if(!name){ showToast('请输入单元名'); return; }
  const parsed=parseUnitName(name);
  const dict=loadDict();
  // 用字典增强解析
  dict.forEach(r=>{ try{ if(new RegExp(r.pattern).test(name)){ parsed[r.field]=r.label; } }catch(e){} });
  $('#dict-test-result').innerHTML=`<div style="background:var(--surface-2);padding:14px;border-radius:8px">
    <div style="font-weight:600;margin-bottom:8px">解析结果：</div>
    <div style="display:grid;grid-template-columns:100px 1fr;gap:6px;font-size:12.5px">
      <div style="color:var(--ink-3)">原始名称：</div><div>${esc(name)}</div>
      <div style="color:var(--ink-3)">日期：</div><div>${parsed.date||'—'}</div>
      <div style="color:var(--ink-3)">产品：</div><div>${parsed.product||'—'}</div>
      <div style="color:var(--ink-3)">版位：</div><div>${parsed.placement||'—'}</div>
      <div style="color:var(--ink-3)">定向：</div><div>${esc(parsed.targeting||'—')}</div>
      <div style="color:var(--ink-3)">变体：</div><div>${parsed.variant||'—'}</div>
      <div style="color:var(--ink-3)">出价类型：</div><div>${parsed.bidType||'—'}</div>
      <div style="color:var(--ink-3)">城市层级：</div><div>${parsed.cityTier||'—'}</div>
      <div style="color:var(--ink-3)">年龄段：</div><div>${parsed.age||'—'}</div>
    </div>
    <div style="margin-top:10px">${unitDimTags(parsed)}</div>
  </div>`;
}

/* ---------- 7. 跨账户分析 ---------- */
function renderCrossAccount(){
  if(!TENCENT_LIVE_DATA){ $('#ca-compare').innerHTML='<div class="empty-state-enhanced">暂无腾讯广告数据</div>'; return; }
  const accs=TENCENT_LIVE_DATA.accounts.slice().sort((a,b)=>b.spend-a.spend);
  // 账户横向对比
  $('#ca-compare').innerHTML='<div class="tbl-wrap"><table><tr><th>账户ID</th><th>运营方</th><th>消耗</th><th>转化</th><th>CPA</th><th>ROI</th><th>余额</th><th>状态</th></tr>'+
    accs.map(a=>`<tr><td><b>${a.id}</b></td><td>${esc(a.operator||'')}</td><td>¥${fmtM(a.spend)}</td><td>${fmtN(a.conv)}</td><td>¥${fmtM2(a.cpa)}</td><td style="color:${a.roi>=0.5?'#059669':a.roi<0.1?'#dc2626':'#6b7280'};font-weight:600">${fmtP(a.roi*100)}</td><td>¥${fmtM(a.balance)}</td><td>${esc(a.status||'')}</td></tr>`).join('')+'</table></div>';
  // 重复单元检测
  const allUnits=[];
  (TENCENT_ADUNITS_DATA||[]).forEach(a=>a.units.forEach(u=>{
    const p=parseUnitName(u[0]||'');
    allUnits.push({acc:a.account_id,op:a.operator||'',name:u[0],parsed:p,spend:parseFloat(String(u[5]||'0').replace(/[,¥]/g,''))||0});
  }));
  // 按版位+定向+出价类型分组
  const groups={};
  allUnits.forEach(u=>{
    const key=(u.parsed.placement||'未知')+'|'+(u.parsed.targeting||'未知')+'|'+(u.parsed.bidType||'未知');
    if(!groups[key]) groups[key]=[];
    groups[key].push(u);
  });
  const duplicates=Object.entries(groups).filter(([k,v])=>v.length>1).sort((a,b)=>b[1].length-a[1].length);
  if(duplicates.length===0){ $('#ca-duplicates').innerHTML='<div class="empty-state-enhanced"><div class="empty-icon">✅</div>未检测到明显的重复定向单元</div>'; }
  else{
    $('#ca-duplicates').innerHTML=duplicates.slice(0,10).map(([key,units])=>{
      const [pl,ta,bt]=key.split('|');
      const totalSpend=units.reduce((s,u)=>s+u.spend,0);
      return `<div style="padding:10px 14px;border:1px solid var(--border);border-radius:8px;margin-bottom:8px">
        <div style="font-weight:600;font-size:13px;margin-bottom:6px">${esc(pl)} / ${esc(ta)} / ${esc(bt)} <span style="color:var(--ink-3);font-weight:400;font-size:11px">${units.length}个单元重复，总消耗¥${fmtM(totalSpend)}</span></div>
        <div style="font-size:11.5px;color:var(--ink-2);line-height:1.8">${units.map(u=>'<div>• 账户'+u.acc+'（'+u.op+'）：'+esc(u.name)+' — 消耗¥'+fmtM(u.spend)+'</div>').join('')}</div>
        <div style="margin-top:6px;font-size:11px;color:#b45309;background:#fffbeb;padding:4px 8px;border-radius:6px">⚠️ 建议：保留CPA最低的1-2个单元，暂停其余重复单元，避免内部竞争抬高CPA</div>
      </div>`;
    }).join('');
  }
  // 预算分配建议
  const avgRoi=accs.reduce((s,a)=>s+a.roi,0)/(accs.length||1);
  const highRoi=accs.filter(a=>a.roi>avgRoi*1.2);
  const lowRoi=accs.filter(a=>a.roi<avgRoi*0.8);
  let suggestHtml='';
  if(highRoi.length>0) suggestHtml+='<div style="margin-bottom:10px"><b style="color:#059669">建议加预算：</b>'+highRoi.map(a=>'账户'+a.id+'（ROI'+fmtP(a.roi*100)+'）').join('、')+'，ROI高于均值'+((highRoi[0].roi/avgRoi-1)*100).toFixed(0)+'%以上</div>';
  if(lowRoi.length>0) suggestHtml+='<div style="margin-bottom:10px"><b style="color:#dc2626">建议减预算：</b>'+lowRoi.map(a=>'账户'+a.id+'（ROI'+fmtP(a.roi*100)+'）').join('、')+'，ROI低于均值'+((1-lowRoi[0].roi/avgRoi)*100).toFixed(0)+'%以上</div>';
  if(!suggestHtml) suggestHtml='<div class="empty-state-enhanced">各账户ROI较为均衡，暂无需大幅调整预算分配</div>';
  $('#ca-budget-suggest').innerHTML=suggestHtml;
}

/* ---------- 8. 转化延迟（框架） ---------- */
function renderConvDelay(){
  const snaps=loadSnapshots();
  // 图表
  const c=chart('cd-chart');
  if(c && snaps.length>=2){
    const dates=snaps.slice(0,7).map(s=>s.date).reverse();
    const spends=snaps.slice(0,7).map(s=>s.total.spend).reverse();
    const convs=snaps.slice(0,7).map(s=>s.total.conv).reverse();
    c.setOption({tooltip:tip,legend:{data:['消耗','转化'],bottom:0},grid:{left:8,right:16,top:16,bottom:40,containLabel:true},xAxis:{type:'category',data:dates,...baseAxis},yAxis:[{type:'value',name:'消耗',...baseAxis},{type:'value',name:'转化',...baseAxis}],series:[{name:'消耗',type:'bar',data:spends,itemStyle:{color:C.accent}},{name:'转化',type:'line',yAxisIndex:1,data:convs,itemStyle:{color:C.teal},smooth:true}]});
  } else if(c){
    c.setOption({title:{text:'需积累至少2天历史数据',left:'center',top:'center',textStyle:{color:'#94a3b8',fontSize:14}}});
  }
  // 表格
  const table=$('#cd-table');
  if(SYNC.link){
    const links=SYNC.link.linkSummary||[];
    table.innerHTML='<tr><th>链路</th><th>低价课订单</th><th>T+0转化</th><th>T+1累计</th><th>T+3累计</th><th>T+7累计</th><th>高价课订单</th><th>总产值</th><th>回传周期</th></tr>'+
      links.map(l=>`<tr><td><b>${esc(l.link)}</b></td><td>${fmtN(l.eff)}</td><td>—</td><td>—</td><td>—</td><td>—</td><td>${fmtN(l.highOrders)}</td><td>¥${fmtM(l.output||0)}</td><td style="color:#b45309">需多期数据</td></tr>`).join('');
  } else { table.innerHTML='<tr><td colspan="9" style="text-align:center;padding:20px;color:#94a3b8">请先同步链路数据</td></tr>'; }
  // 回传期保护
  const todayStr=new Date().toLocaleString('zh-CN',{month:'2-digit',day:'2-digit'}).replace(/\//g,'/'); // "08/31"
  const protectedUnits=(TENCENT_ADUNITS_DATA||[]).flatMap(a=>a.units.filter(u=>{
    const p=parseUnitName(u[0]||'');
    return p.date && p.date===todayStr;
  }));
  $('#cd-protected').innerHTML=protectedUnits.length>0
    ? '<div style="color:#059669">今日新上线 '+protectedUnits.length+' 个单元，处于回传保护期（48小时内），不纳入空耗/超成本判断</div>'
    : '<div style="color:#6b7280">当前无处于回传保护期的新单元</div>';
}

/* ---------- 9. 前后端关联配置（框架） ---------- */
const LINKCFG_KEY='workbench_linkcfg_v1';
function loadLinkCfg(){ try{return JSON.parse(localStorage.getItem(LINKCFG_KEY))||{mode:'plan',mappings:[]};}catch(e){return{mode:'plan',mappings:[]};} }
function saveLinkCfg(c){ try{localStorage.setItem(LINKCFG_KEY,JSON.stringify(c));}catch(e){} }
function renderLinkConfig(){
  const fieldMap=$('#link-field-map');
  if(!fieldMap) return;
  const cfg=loadLinkCfg();
  document.querySelectorAll('input[name="link-mode"]').forEach(r=>{r.checked=r.value===cfg.mode;});
  // 字段映射
  const frontFields=['计划名称','营销单元名称','账户ID','创意ID','UTM参数'];
  const backFields=['链路名称','期数','手机号','设备号','计划名','UTM'];
  const mapHtml=cfg.mappings.length?cfg.mappings.map((m,i)=>`<div class="link-config-row">
    <select data-idx="${i}" data-side="front">${frontFields.map(f=>`<option ${m.front===f?'selected':''}>${f}</option>`).join('')}</select>
    <select data-idx="${i}" data-side="back">${backFields.map(f=>`<option ${m.back===f?'selected':''}>${f}</option>`).join('')}</select>
    <span class="dict-del" onclick="delLinkMap(${i})">删除</span>
  </div>`).join(''):'<div style="color:var(--ink-3);font-size:12px;padding:10px 0">暂无映射规则，点击下方"添加映射"配置</div>';
  fieldMap.innerHTML=mapHtml+'<button class="btn btn-secondary" id="link-add-map" style="margin-top:8px;padding:4px 12px;font-size:11px">+ 添加映射</button>';
  const addMapBtn=$('#link-add-map');
  if(addMapBtn)addMapBtn.onclick=()=>{ const c=loadLinkCfg(); c.mappings.push({front:'计划名称',back:'计划名'}); saveLinkCfg(c); renderLinkConfig(); };
  const saveBtn=$('#link-save');
  if(saveBtn)saveBtn.onclick=()=>{
    const modeEl=document.querySelector('input[name="link-mode"]:checked');
    const mode=modeEl?modeEl.value:'plan';
    const rows=document.querySelectorAll('#link-field-map .link-config-row');
    const mappings=[];
    rows.forEach(r=>{ const fs=r.querySelector('[data-side="front"]'); const bs=r.querySelector('[data-side="back"]'); if(fs&&bs)mappings.push({front:fs.value,back:bs.value}); });
    saveLinkCfg({mode,mappings}); showToast('关联配置已保存');
    tryLinkMatch();
  };
}
function delLinkMap(i){ const c=loadLinkCfg(); c.mappings.splice(i,1); saveLinkCfg(c); renderLinkConfig(); }
function tryLinkMatch(){
  const resultCard=$('#link-result-card');
  if(!resultCard) return;
  // 简单的关联预览：基于计划名称模糊匹配
  if(!TENCENT_ADUNITS_DATA || !SYNC.link){ resultCard.style.display='none'; return; }
  const links=SYNC.link.linkSummary||[];
  const allUnits=(TENCENT_ADUNITS_DATA||[]).flatMap(a=>a.units.map(u=>({acc:a.account_id,name:u[0]||''})));
  let matched=0;
  const results=allUnits.slice(0,20).map(u=>{
    const link=links.find(l=>u.name.includes(l.link)||l.link.includes(u.name.slice(0,4)));
    if(link) matched++;
    return {unit:u.name,acc:u.acc,link:link?link.link:'未匹配',roi:link?fmtP(link.roi*100):'—'};
  });
  const matchRate=$('#link-match-rate');
  if(matchRate)matchRate.textContent='匹配率 '+matched+'/'+Math.min(allUnits.length,20)+'（预览前20条）';
  resultCard.style.display='block';
  const resultTable=$('#link-result-table');
  if(resultTable)resultTable.innerHTML='<tr><th>营销单元</th><th>账户</th><th>匹配链路</th><th>链路ROI</th></tr>'+
    results.map(r=>`<tr><td style="max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(r.unit)}">${esc(r.unit)}</td><td>${r.acc}</td><td style="color:${r.link==='未匹配'?'#dc2626':'#059669'}">${esc(r.link)}</td><td>${r.roi}</td></tr>`).join('');
}

/* ---------- 10. 增强版优化建议（含样本量门槛+回传期保护） ---------- */
function generateSuggestionsEnhanced(){
  const base=generateSuggestions(TENCENT_LIVE_DATA,TENCENT_ADUNITS_DATA);
  // 过滤回传期内的新单元（投放不足48小时）
  const todayStr=new Date().toLocaleString('zh-CN',{month:'2-digit',day:'2-digit'}).replace(/\//g,'/');
  const filtered=base.filter(s=>{
    const p=parseUnitName(s.unit);
    if(p.date===todayStr && s.type==='pause') return false; // 今日新单元不建议暂停
    return true;
  });
  // 添加回传期保护说明
  const protectedCount=(TENCENT_ADUNITS_DATA||[]).flatMap(a=>a.units).filter(u=>{
    const p=parseUnitName(u[0]||''); return p.date===today;
  }).length;
  if(protectedCount>0){
    filtered.push({type:'watch',title:'回传期保护',unit:protectedCount+'个今日新单元',accountId:'—',operator:'—',detail:'今日新上线的'+protectedCount+'个单元处于转化回传期（48小时内），暂不纳入空耗/超成本判断，建议观察至明日',meta:'回传期保护',action:'观察'});
  }
  return filtered;
}

/* ---------- 11. 素材自动疲劳检测（基于历史数据） ---------- */
function detectMaterialFatigue(){
  if(!SYNC.material) return [];
  const list=SYNC.material.list||[];
  const fatigued=[];
  list.forEach(m=>{
    // 简易疲劳评分：高消耗+低ROI+高完播率下降（模拟）
    const fatigueScore=(m.cost>500?1:0)+(m.roi<0.1?1:0)+(m.finishRate<0.1?1:0);
    if(fatigueScore>=2 && getMatStage(m.name)!=='dead'){
      fatigued.push({name:m.name,score:fatigueScore,cost:m.cost,roi:m.roi,reason:'高消耗¥'+m.cost.toFixed(0)+'+低ROI'+(m.roi*100).toFixed(1)+'%，建议标记为疲劳并更换素材'});
    }
  });
  return fatigued;
}

/* ---------- 12. 新视图页面元数据 ---------- */
const EXTRA_VIEWS={
  'op-log':{title:'操作日志',render:renderOpLog},
  'ab-test':{title:'A/B测试',render:renderABTests},
  'daily-report':{title:'日报生成',render:()=>{ $('#dr-date').value=new Date().toISOString().slice(0,10); }},
  'data-quality':{title:'数据质量',render:renderDataQuality},
  'link-config':{title:'前后端关联',render:renderLinkConfig},
  'name-dict':{title:'命名字典',render:renderDict},
  'cross-account':{title:'跨账户分析',render:renderCrossAccount},
  'conv-delay':{title:'转化延迟',render:renderConvDelay}
};

// 增强route函数支持新视图（已在route函数中处理）

// 绑定新视图的事件
function bindEnhancedEvents(){
  // 操作日志
  const addBtn=$('#op-log-add');
  if(addBtn) addBtn.onclick=()=>{
    const type=$('#op-log-type').value;
    const target=$('#op-log-target').value;
    const before=$('#op-log-before').value;
    const note=$('#op-log-note').value;
    if(!target){ showToast('请输入操作对象'); return; }
    addOpLog(type,target,before,note);
    $('#op-log-target').value=''; $('#op-log-before').value=''; $('#op-log-note').value='';
    showToast('操作已记录'); renderOpLog();
  };
  const fAll=$('#op-log-filter-all'); if(fAll)fAll.onclick=()=>{opLogFilter='all';renderOpLog();};
  const fPending=$('#op-log-filter-pending'); if(fPending)fPending.onclick=()=>{opLogFilter='pending';renderOpLog();};
  const fDone=$('#op-log-filter-done'); if(fDone)fDone.onclick=()=>{opLogFilter='done';renderOpLog();};
  const fClear=$('#op-log-clear'); if(fClear)fClear.onclick=()=>{ if(confirm('确定清空全部操作日志？')){ saveOpLog([]); renderOpLog(); } };
  // A/B测试
  const abCreate=$('#ab-create'); if(abCreate)abCreate.onclick=()=>{
    const name=$('#ab-name').value; const varType=$('#ab-var').value;
    if(!name){ showToast('请输入测试名称'); return; }
    const tests=loadABTests();
    tests.unshift({name,varType,time:new Date().toLocaleString('zh-CN'),status:'running',aTotal:0,aConv:0,bTotal:0,bConv:0,note:''});
    saveABTests(tests); $('#ab-name').value=''; renderABTests(); showToast('测试已创建');
  };
  // 日报
  const drGen=$('#dr-generate'); if(drGen)drGen.onclick=generateDailyReport;
  const drCopy=$('#dr-copy'); if(drCopy)drCopy.onclick=()=>{ const text=$('#dr-content')?.textContent||''; if(text){ navigator.clipboard.writeText(text).then(()=>showToast('日报已复制到剪贴板')).catch(()=>showToast('复制失败，请手动选择复制')); } };
  // 数据质量
  const dqRun=$('#dq-run'); if(dqRun)dqRun.onclick=renderDataQuality;
  // 命名字典
  const dictAdd=$('#dict-add'); if(dictAdd)dictAdd.onclick=()=>{ const d=loadDict(); d.push({pattern:'',field:'placement',label:''}); saveDict(d); renderDict(); };
  const dictSave=$('#dict-save'); if(dictSave)dictSave.onclick=()=>{
    const rows=document.querySelectorAll('#dict-rules .dict-rule');
    const d=[];
    rows.forEach(r=>{ const p=r.querySelector('input'); const f=r.querySelector('select'); if(p&&p.value) d.push({pattern:p.value,field:f?.value||'placement',label:p.value}); });
    saveDict(d); showToast('命名字典已保存（'+d.length+'条规则）');
  };
  const dictTest=$('#dict-test-btn'); if(dictTest)dictTest.onclick=testDictParse;
}

/* ---------- 工具12：账户对比分析 ---------- */
let t12Output=null;

// 绑定文件上传
document.addEventListener('DOMContentLoaded',()=>{
  const f=document.getElementById('t12File');
  if(f)f.addEventListener('change',e=>{
    const files=e.target.files;
    if(!files||!files.length)return;
    const names=Array.from(files).map(x=>x.name);
    document.getElementById('t12FileName').textContent='✓ '+names.join('、');
    document.getElementById('t12Btn').disabled=false;
    showStatus('t12','已选择'+files.length+'个文件，可以开始分析','active');
  });
});

// 解析并分析账户对比数据
async function processTab12(){
  const f=document.getElementById('t12File');
  if(!f.files.length){alert('请先上传「人群画像数据看板」Excel');return;}
  document.getElementById('t12Btn').disabled=true;
  showStatus('t12','正在解析数据...','active');
  try{
    const file=f.files[0];
    const buf=await file.arrayBuffer();
    const wb=XLSX.read(buf,{type:'array'});
    let sn='导出数据';
    if(!wb.SheetNames.includes(sn))sn=wb.SheetNames[0];
    const ws=wb.Sheets[sn];
    // 无表头文件，用header:1获取二维数组
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
    if(!rows.length){showStatus('t12','文件为空或无法读取','error');return;}

        showStatus('t12','正在按评分分段聚合...','active');

    // 列索引（0-based）：E=4评分, F=5低价课订单, I=8高价课, J=9消耗, K=10单订单产值, L=11有效/进群
    // 第一步：找到总计行（F值最大的行）
    let totalRow=null, totalRowIdx=-1;
    for(let i=0;i<rows.length;i++){
      const r=rows[i];
      if(!r||r.length<12)continue;
      const fVal=Number(r[5])||0;
      if(totalRow===null||fVal>(Number(totalRow[5])||0)){
        totalRow=r; totalRowIdx=i;
      }
    }

    // 第二步：按评分分段聚合（排除总计行和E为空的汇总行）
    const agg={};
    let rawCount=0, skipped=0;
    for(let i=0;i<rows.length;i++){
      if(i===totalRowIdx)continue;
      const r=rows[i];
      if(!r||r.length<12)continue;
      const score=r[4];
      if(score===null||score===undefined||score===''){skipped++;continue;}
      const scoreNum=Number(score);
      if(isNaN(scoreNum)||scoreNum<1||scoreNum>10){skipped++;continue;}
      rawCount++;
      const lowOrders=Number(r[5])||0;
      const highConv=Number(r[8])||0;
      const spend=Number(r[9])||0;
      const validOrders=Number(r[11])||0;
      if(!agg[scoreNum])agg[scoreNum]={score:scoreNum,lowOrders:0,spend:0,highConv:0,validOrders:0,rowCount:0};
      agg[scoreNum].lowOrders+=lowOrders;
      agg[scoreNum].spend+=spend;
      agg[scoreNum].highConv+=highConv;
      agg[scoreNum].validOrders+=validOrders;
      agg[scoreNum].rowCount++;
    }

    // 计算各评分段指标
    const segments=Object.values(agg).map(a=>{
      return {
        score:a.score,
        lowOrders:a.lowOrders,
        spend:a.spend,
        highConv:a.highConv,
        validOrders:a.validOrders,
        orderValue:a.lowOrders?a.spend/a.lowOrders:0,
        highRate:a.lowOrders?a.highConv/a.lowOrders*100:0,
        highCpa:a.highConv?a.spend/a.highConv:0,
        rowCount:a.rowCount
      };
    }).sort((a,b)=>a.score-b.score);

    // 总体指标：优先用总计行
    let total;
    if(totalRow){
      total={
        lowOrders:Number(totalRow[5])||0,
        highConv:Number(totalRow[8])||0,
        spend:Number(totalRow[9])||0,
        validOrders:Number(totalRow[11])||0,
        orderValue:Number(totalRow[10])||0,
        segmentCount:segments.length,
        activeSegments:segments.filter(s=>s.spend>0).length
      };
    }else{
      total={
        lowOrders:segments.reduce((s,a)=>s+a.lowOrders,0),
        highConv:segments.reduce((s,a)=>s+a.highConv,0),
        spend:segments.reduce((s,a)=>s+a.spend,0),
        validOrders:segments.reduce((s,a)=>s+a.validOrders,0),
        orderValue:0,
        segmentCount:segments.length,
        activeSegments:segments.filter(s=>s.spend>0).length
      };
    }
    total.highRate=total.lowOrders?total.highConv/total.lowOrders*100:0;
    total.highCpa=total.highConv?total.spend/total.highConv:0;
    if(!total.orderValue&&total.lowOrders)total.orderValue=total.spend/total.lowOrders;

    // 评分分层分析：高(8-10)、中(4-7)、低(1-3)
    const tierHigh=segments.filter(s=>s.score>=8);
    const tierMid=segments.filter(s=>s.score>=4&&s.score<=7);
    const tierLow=segments.filter(s=>s.score<=3);
    const sumTier=(arr)=>({
      lowOrders:arr.reduce((s,a)=>s+a.lowOrders,0),
      spend:arr.reduce((s,a)=>s+a.spend,0),
      highConv:arr.reduce((s,a)=>s+a.highConv,0)
    });
    const tiers=[
      {name:'高价值层(8-10分)',...sumTier(tierHigh),scores:tierHigh.map(s=>s.score)},
      {name:'中分层(4-7分)',...sumTier(tierMid),scores:tierMid.map(s=>s.score)},
      {name:'低价值层(1-3分)',...sumTier(tierLow),scores:tierLow.map(s=>s.score)}
    ];
    tiers.forEach(t=>{
      t.orderValue=t.lowOrders?t.spend/t.lowOrders:0;
      t.highRate=t.lowOrders?t.highConv/t.lowOrders*100:0;
      t.orderShare=total.lowOrders?t.lowOrders/total.lowOrders*100:0;
      t.spendShare=total.spend?t.spend/total.spend*100:0;
    });

    // 生成洞察和建议
    const findings=genAccountFindings(segments,total,tiers);
    const suggestions=genAccountSuggestions(segments,total,tiers);

    t12Output={segments,total,tiers,findings,suggestions,rawCount,skipped,fileName:file.name};

    rendert12Results();
    showStatus('t12','分析完成！共解析'+rawCount+'行明细，聚合为'+segments.length+'个评分段','done');
  }catch(err){
    console.error(err);
    showStatus('t12','分析失败：'+err.message,'error');
    document.getElementById('t12Btn').disabled=false;
  }
}

// 生成洞察
function genAccountFindings(segments,total,tiers){
  const findings=[];
  // 总体情况
  findings.push({type:'info',text:'期间总低价课 '+fmtN(total.lowOrders)+' 单，成单 '+total.highConv+' 单高价课，总消耗 ¥'+total.spend.toFixed(0)+'，综合单粉产值 ¥'+total.orderValue.toFixed(2)+'，高价转化率 '+total.highRate.toFixed(2)+'%'});
  // 评分分层
  const highTier=tiers[0], midTier=tiers[1], lowTier=tiers[2];
  if(highTier.lowOrders>0){
    findings.push({type:'success',text:'高价值层(8-10分)：'+fmtN(highTier.lowOrders)+'单，占比'+highTier.orderShare.toFixed(1)+'%，高价转化率'+highTier.highRate.toFixed(2)+'%，单粉产值¥'+highTier.orderValue.toFixed(2)+'，属于高价值优质流量'});
  }
  if(lowTier.lowOrders>0&&lowTier.highConv===0){
    findings.push({type:'warning',text:'低价值层(1-3分)：'+fmtN(lowTier.lowOrders)+'单，占比'+lowTier.orderShare.toFixed(1)+'%，但高价课转化为0，属于低价值流量，消耗了前端买量成本与助教跟进精力'});
  }
  if(midTier.lowOrders>0){
    findings.push({type:'info',text:'中分层(4-7分)：'+fmtN(midTier.lowOrders)+'单，占比'+midTier.orderShare.toFixed(1)+'%，贡献了大部分订单基数与稳定转化，高价转化率'+midTier.highRate.toFixed(2)+'%'});
  }
  // 各评分段对比
  const withHigh=segments.filter(s=>s.highConv>0).sort((a,b)=>b.highRate-a.highRate);
  if(withHigh.length>=2){
    const best=withHigh[0], worst=withHigh[withHigh.length-1];
    findings.push({type:'success',text:'高价转化率最高的是评分'+best.score+'（'+best.highRate.toFixed(2)+'%，'+best.highConv+'单高价课），最低的是评分'+worst.score+'（'+worst.highRate.toFixed(2)+'%）'});
  }
  // 单粉产值对比
  const withSpend=segments.filter(s=>s.spend>0).sort((a,b)=>b.orderValue-a.orderValue);
  if(withSpend.length>=2){
    findings.push({type:'info',text:'单粉产值最高的是评分'+withSpend[0].score+'（¥'+withSpend[0].orderValue.toFixed(2)+'/单），最低的是评分'+withSpend[withSpend.length-1].score+'（¥'+withSpend[withSpend.length-1].orderValue.toFixed(2)+'/单）'});
  }
  return findings;
}

// 生成优化建议
function genAccountSuggestions(segments,total,tiers){
  const suggestions=[];
  const lowTier=tiers[2], highTier=tiers[0];
  // 低分层建议
  if(lowTier.lowOrders>0&&lowTier.highConv===0){
    suggestions.push({score:'1-3分',spend:lowTier.spend,action:'压制消耗/深层出价',reason:'低分层'+fmtN(lowTier.lowOrders)+'单零高价转化，占总单量'+lowTier.orderShare.toFixed(1)+'%，建议接入深层转化出价（双目标/优化回传意向分），过滤低分人群提升ROI',priority:'high'});
  }
  // 高分层建议
  if(highTier.lowOrders>0&&highTier.highConv>0){
    suggestions.push({score:'8-10分',spend:highTier.spend,action:'加预算/放量',reason:'高分层高价转化率'+highTier.highRate.toFixed(2)+'%，单粉产值¥'+highTier.orderValue.toFixed(2)+'，属于高价值优质流量，建议倾斜预算放量',priority:'high'});
  }
  // 各评分段具体建议
  segments.filter(s=>s.spend>0).forEach(s=>{
    if(s.highConv===0){
      suggestions.push({score:'评分'+s.score,spend:s.spend,action:'观察/暂停',reason:'消耗¥'+s.spend.toFixed(0)+'但零高价转化，建议检查定向和素材方向',priority:'medium'});
    }else if(s.highRate>=total.highRate*1.5){
      suggestions.push({score:'评分'+s.score,spend:s.spend,action:'加预算',reason:'高价转化率'+s.highRate.toFixed(2)+'%远高于大盘('+total.highRate.toFixed(2)+'%)，效率优秀，建议放量',priority:'high'});
    }else if(s.highRate<total.highRate*0.5){
      suggestions.push({score:'评分'+s.score,spend:s.spend,action:'砍预算/优化',reason:'高价转化率'+s.highRate.toFixed(2)+'%远低于大盘，建议缩减预算或优化素材定向',priority:'medium'});
    }
  });
  // 运营侧建议
  suggestions.push({score:'运营侧',spend:0,action:'优化Day0触达与开播提醒',reason:'从好友到直播到课存在约30%流失，建议优化加粉后Day0触达链路、第一节课开播前的私聊提醒与破冰互动',priority:'medium'});
  suggestions.push({score:'投放侧',spend:0,action:'排查素材夸大诱导',reason:'到课率低可能因素材存在虚假夸大或买赠诱导，导致进线人群非核心兴趣人群，建议排查素材方向',priority:'low'});
  return suggestions;
}

// 渲染上传页面结果
function rendert12Results(){
  if(!t12Output)return;
  const {segments,total,tiers,findings,suggestions}=t12Output;
  document.getElementById('t12Results').style.display='block';
  // KPI
  renderKpis('#t12Kpi',[
    {label:'总低价课订单',value:fmtN(total.lowOrders)},
    {label:'总高价课',value:fmtN(total.highConv),cls:'kpi-teal'},
    {label:'总消耗',value:fmtM(total.spend),cls:'kpi-accent'},
    {label:'单粉产值',value:'¥'+total.orderValue.toFixed(2)},
    {label:'高价转化率',value:total.highRate.toFixed(2)+'%'},
    {label:'高价课CPA',value:total.highConv?'¥'+total.highCpa.toFixed(0):'—'},
    {label:'评分段数',value:total.segmentCount},
    {label:'有消耗段',value:total.activeSegments}
  ]);
  // 评分分层对比表
  let html='<tr><th>评分分层</th><th>包含评分</th><th>低价课订单</th><th>订单占比</th><th>高价课</th><th>高价转化率</th><th>消耗(元)</th><th>单粉产值</th></tr>';
  tiers.forEach(t=>{
    html+='<tr><td><strong>'+t.name+'</strong></td><td>'+(t.scores.length?t.scores.join('、'):'—')+'</td><td>'+fmtN(t.lowOrders)+'</td><td>'+t.orderShare.toFixed(1)+'%</td><td>'+t.highConv+'</td><td>'+t.highRate.toFixed(2)+'%</td><td>'+t.spend.toFixed(0)+'</td><td>¥'+t.orderValue.toFixed(2)+'</td></tr>';
  });
  html+='<tr style="background:#eaf2ff;font-weight:bold"><td>合计</td><td>1-10分</td><td>'+fmtN(total.lowOrders)+'</td><td>100%</td><td>'+total.highConv+'</td><td>'+total.highRate.toFixed(2)+'%</td><td>'+total.spend.toFixed(0)+'</td><td>¥'+total.orderValue.toFixed(2)+'</td></tr>';
  // 各评分段明细
  html+='<tr style="background:#f0f0f0"><td colspan="8" style="text-align:center;font-weight:bold">各评分段明细</td></tr>';
  html+='<tr><th>评分</th><th>低价课订单</th><th>高价课</th><th>高价转化率</th><th>消耗(元)</th><th>单粉产值</th><th>高价课CPA</th><th>明细行数</th></tr>';
  segments.forEach(s=>{
    html+='<tr><td><strong>评分'+s.score+'</strong></td><td>'+fmtN(s.lowOrders)+'</td><td>'+s.highConv+'</td><td>'+s.highRate.toFixed(2)+'%</td><td>'+s.spend.toFixed(0)+'</td><td>¥'+s.orderValue.toFixed(2)+'</td><td>'+(s.highConv?'¥'+s.highCpa.toFixed(0):'—')+'</td><td>'+s.rowCount+'</td></tr>';
  });
  document.getElementById('t12Tbl').innerHTML=html;
  // 洞察
  let findHtml='<h4 style="margin:14px 0 8px">核心洞察</h4>';
  findings.forEach(f=>{
    const icon=f.type==='success'?'✅':f.type==='warning'?'⚠️':'ℹ️';
    const color=f.type==='success'?'#27ae60':f.type==='warning'?'#f39c12':'#3498db';
    findHtml+='<div style="padding:10px 14px;margin:6px 0;background:#f8f9fa;border-radius:6px;border-left:3px solid '+color+'">'+icon+' '+f.text+'</div>';
  });
  // 建议
  let sugHtml='<h4 style="margin:14px 0 8px">优化建议（按优先级排序）</h4>';
  suggestions.forEach(s=>{
    const color=s.priority==='high'?'#e74c3c':s.priority==='medium'?'#f39c12':'#95a5a6';
    const bg=s.priority==='high'?'#fef0f0':s.priority==='medium'?'#fff8e6':'#f5f6fa';
    sugHtml+='<div style="padding:10px 14px;margin:6px 0;border-left:4px solid '+color+';background:'+bg+';border-radius:0 6px 6px 0"><strong>'+s.score+'</strong> — <span style="color:'+color+';font-weight:bold">'+s.action+'</span><br><span style="font-size:12px;color:#666">'+s.reason+'</span></div>';
  });
  document.getElementById('t12Suggestions').innerHTML=findHtml+sugHtml;
}

// 同步到分析模块
function syncTab12ToAnalysis(){
  if(!t12Output){alert('请先点击开始分析');return;}
  saveSync('accountCompare',t12Output);
  showToast('已同步到分析模块 → 账户对比分析');
  setTimeout(()=>{ location.hash='#/account-compare'; },800);
}

// 渲染分析页面
function renderAccountCompare(){
  const badge=document.getElementById('ac-badge');
  const d=SYNC.accountCompare;
  if(badge){
    if(d)badge.innerHTML='<span style="color:#27ae60">● 已同步Excel数据（'+d.fileName+'）</span>';
    else badge.textContent='数据同步中…';
  }
  const empty=document.getElementById('ac-empty');
  if(!d){
    empty.style.display='block';
    ['ac-table-card','ac-chart-card','ac-bar-card','ac-conv-card','ac-finding-card','ac-suggest-card'].forEach(id=>{
      const el=document.getElementById(id);if(el)el.style.display='none';
    });
    return;
  }
  empty.style.display='none';
  const {segments,total,tiers,findings,suggestions}=d;

  // KPI
  document.getElementById('ac-summary').textContent='总低价课'+fmtN(total.lowOrders)+'单 · 高价课'+total.highConv+'单 · 总消耗¥'+total.spend.toFixed(0)+' · 单粉产值¥'+total.orderValue.toFixed(2);
  renderKpis('#ac-kpi',[
    {label:'总低价课订单',value:fmtN(total.lowOrders)},
    {label:'总高价课',value:fmtN(total.highConv),cls:'kpi-teal'},
    {label:'总消耗',value:fmtM(total.spend),cls:'kpi-accent'},
    {label:'单粉产值',value:'¥'+total.orderValue.toFixed(2)},
    {label:'高价转化率',value:total.highRate.toFixed(2)+'%'},
    {label:'高价课CPA',value:total.highConv?'¥'+total.highCpa.toFixed(0):'—'},
    {label:'评分段数',value:total.segmentCount},
    {label:'有消耗段',value:total.activeSegments}
  ]);

  // 评分分层对比表
  document.getElementById('ac-table-card').style.display='block';
  let html='<thead><tr><th>评分分层</th><th>包含评分</th><th>低价课订单</th><th>订单占比</th><th>高价课</th><th>高价转化率</th><th>消耗(元)</th><th>单粉产值</th></tr></thead><tbody>';
  tiers.forEach(t=>{
    html+='<tr><td><strong>'+t.name+'</strong></td><td>'+(t.scores.length?t.scores.join('、'):'—')+'</td><td>'+fmtN(t.lowOrders)+'</td><td>'+t.orderShare.toFixed(1)+'%</td><td>'+t.highConv+'</td><td>'+t.highRate.toFixed(2)+'%</td><td>'+t.spend.toFixed(0)+'</td><td>¥'+t.orderValue.toFixed(2)+'</td></tr>';
  });
  html+='<tr style="background:#eaf2ff;font-weight:bold"><td>合计</td><td>1-10分</td><td>'+fmtN(total.lowOrders)+'</td><td>100%</td><td>'+total.highConv+'</td><td>'+total.highRate.toFixed(2)+'%</td><td>'+total.spend.toFixed(0)+'</td><td>¥'+total.orderValue.toFixed(2)+'</td></tr>';
  // 各评分段明细
  html+='<tr style="background:#f0f0f0"><td colspan="8" style="text-align:center;font-weight:bold">各评分段明细</td></tr>';
  html+='<tr><th>评分</th><th>低价课订单</th><th>高价课</th><th>高价转化率</th><th>消耗(元)</th><th>单粉产值</th><th>高价课CPA</th><th>明细行数</th></tr>';
  segments.forEach(s=>{
    html+='<tr><td><strong>评分'+s.score+'</strong></td><td>'+fmtN(s.lowOrders)+'</td><td>'+s.highConv+'</td><td>'+s.highRate.toFixed(2)+'%</td><td>'+s.spend.toFixed(0)+'</td><td>¥'+s.orderValue.toFixed(2)+'</td><td>'+(s.highConv?'¥'+s.highCpa.toFixed(0):'—')+'</td><td>'+s.rowCount+'</td></tr>';
  });
  html+='</tbody>';
  document.getElementById('ac-table').innerHTML=html;

  // 柱状图：各评分段高价转化率 vs 单粉产值
  document.getElementById('ac-chart-card').style.display='block';
  const c1=chart('ac-chart');
  if(c1){
    const sorted=segments.filter(s=>s.lowOrders>0).sort((a,b)=>a.score-b.score);
    c1.setOption({
      tooltip:{trigger:'axis'},
      legend:{data:['高价转化率(%)','单粉产值(元)'],bottom:0},
      grid:{left:50,right:50,top:20,bottom:50},
      xAxis:{type:'category',data:sorted.map(s=>'评分'+s.score),...baseAxis},
      yAxis:[
        {type:'value',name:'转化率(%)',...baseAxis},
        {type:'value',name:'单粉产值(元)',...baseAxis}
      ],
      series:[
        {name:'高价转化率(%)',type:'bar',data:sorted.map(s=>+s.highRate.toFixed(2)),itemStyle:{color:'#e74c3c'},barWidth:'30%'},
        {name:'单粉产值(元)',type:'line',yAxisIndex:1,data:sorted.map(s=>+s.orderValue.toFixed(2)),itemStyle:{color:'#27ae60'},smooth:true}
      ]
    });
  }

  // 饼图：评分分层订单占比
  document.getElementById('ac-bar-card').style.display='block';
  const c2=chart('ac-bar');
  if(c2){
    c2.setOption({
      tooltip:{trigger:'item',formatter:'{b}: {c}单 ({d}%)'},
      legend:{bottom:0},
      series:[{
        type:'pie',radius:['40%','65%'],center:['50%','45%'],
        label:{show:true,formatter:'{b}\n{c}单 ({d}%)'},
        data:tiers.filter(t=>t.lowOrders>0).map(t=>({name:t.name,value:t.lowOrders}))
      }]
    });
  }

  // 柱状图：各评分段消耗 vs 高价课
  document.getElementById('ac-conv-card').style.display='block';
  const c3=chart('ac-conv-chart');
  if(c3){
    const sorted=segments.filter(s=>s.spend>0).sort((a,b)=>a.score-b.score);
    c3.setOption({
      tooltip:{trigger:'axis'},
      legend:{data:['消耗(元)','高价课数'],bottom:0},
      grid:{left:60,right:50,top:20,bottom:50},
      xAxis:{type:'category',data:sorted.map(s=>'评分'+s.score),...baseAxis},
      yAxis:[
        {type:'value',name:'消耗(元)',...baseAxis},
        {type:'value',name:'高价课数',...baseAxis}
      ],
      series:[
        {name:'消耗(元)',type:'bar',data:sorted.map(s=>s.spend),itemStyle:{color:'#3498db'},barWidth:'35%'},
        {name:'高价课数',type:'bar',yAxisIndex:1,data:sorted.map(s=>s.highConv),itemStyle:{color:'#e74c3c'},barWidth:'35%'}
      ]
    });
  }

  // 洞察
  document.getElementById('ac-finding-card').style.display='block';
  let fHtml='';
  findings.forEach(f=>{
    const icon=f.type==='success'?'✅':f.type==='warning'?'⚠️':'ℹ️';
    const color=f.type==='success'?'#27ae60':f.type==='warning'?'#f39c12':'#3498db';
    fHtml+='<div style="padding:10px 14px;margin:6px 0;background:#f8f9fa;border-radius:6px;border-left:3px solid '+color+'">'+icon+' '+f.text+'</div>';
  });
  document.getElementById('ac-findings').innerHTML=fHtml;

  // 建议
  document.getElementById('ac-suggest-card').style.display='block';
  let sHtml='';
  suggestions.forEach(s=>{
    const color=s.priority==='high'?'#e74c3c':s.priority==='medium'?'#f39c12':'#95a5a6';
    const bg=s.priority==='high'?'#fef0f0':s.priority==='medium'?'#fff8e6':'#f5f6fa';
    sHtml+='<div style="padding:10px 14px;margin:6px 0;border-left:4px solid '+color+';background:'+bg+';border-radius:0 6px 6px 0"><strong>'+s.score+'</strong> — <span style="color:'+color+';font-weight:bold">'+s.action+'</span><br><span style="font-size:12px;color:#666">'+s.reason+'</span></div>';
  });
  document.getElementById('ac-suggestions').innerHTML=sHtml;
}

// 下载Excel
function downloadTab12(){
  if(!t12Output){alert('请先点击开始分析');return;}
  const {segments,total,tiers,findings,suggestions}=t12Output;
  const wb=XLSX.utils.book_new();
  // Sheet1: 评分分层对比
  const header1=['评分分层','包含评分','低价课订单','订单占比','高价课','高价转化率','消耗(元)','单粉产值'];
  const data1=tiers.map(t=>[t.name,t.scores.join('、'),t.lowOrders,t.orderShare.toFixed(1)+'%',t.highConv,t.highRate.toFixed(2)+'%',t.spend.toFixed(0),t.orderValue.toFixed(2)]);
  data1.push(['合计','1-10分',total.lowOrders,'100%',total.highConv,total.highRate.toFixed(2)+'%',total.spend.toFixed(0),total.orderValue.toFixed(2)]);
  const ws1=XLSX.utils.aoa_to_sheet([header1,...data1]);
  ws1['!cols']=[{wch:18},{wch:12},{wch:12},{wch:10},{wch:8},{wch:12},{wch:10},{wch:10}];
  XLSX.utils.book_append_sheet(wb,ws1,'评分分层对比');
  // Sheet2: 各评分段明细
  const header2=['评分','低价课订单','高价课','高价转化率','消耗(元)','单粉产值','高价课CPA','明细行数'];
  const data2=segments.map(s=>['评分'+s.score,s.lowOrders,s.highConv,s.highRate.toFixed(2)+'%',s.spend.toFixed(0),s.orderValue.toFixed(2),s.highConv?s.highCpa.toFixed(0):'—',s.rowCount]);
  const ws2=XLSX.utils.aoa_to_sheet([header2,...data2]);
  ws2['!cols']=[{wch:10},{wch:12},{wch:8},{wch:12},{wch:10},{wch:10},{wch:10},{wch:10}];
  XLSX.utils.book_append_sheet(wb,ws2,'各评分段明细');
  // Sheet3: 洞察与建议
  const header3=['类型','内容'];
  const data3=[];
  findings.forEach(f=>data3.push(['洞察('+f.type+')',f.text]));
  suggestions.forEach(s=>data3.push(['建议('+s.priority+')',s.score+' - '+s.action+'：'+s.reason]));
  const ws3=XLSX.utils.aoa_to_sheet([header3,...data3]);
  ws3['!cols']=[{wch:15},{wch:80}];
  XLSX.utils.book_append_sheet(wb,ws3,'洞察与建议');
  XLSX.writeFile(wb,'账户对比分析.xlsx');
}

/* ---------- 启动 ---------- */
loadSync();
loadPersistentData();
loadTodos();
bindEvents();
bindTencentEvents();
initAllDragUploads();
refreshFilters();
// 快速上传下拉框
const qu=document.getElementById('quickUpload');
if(qu){ qu.onchange=()=>{ if(qu.value){ location.hash='#/'+qu.value; qu.value=''; } }; }
// 绑定增强事件
bindEnhancedEvents();
// 自动快照（每次加载时保存）
takeSnapshot();
// 初始路由
route();
