'use strict';
// 화면 동작: 버튼 처리, Worker 호출, 표·그래프 그리기
const $ = id => document.getElementById(id);
const LEVELS = [
  {level: '33', label: '대분류', n: 33, file: './io-2023-33.xlsx', size: '0.6MB'},
  {level: '83', label: '중분류', n: 83, file: './io-2023-83.xlsx', size: '1.2MB'},
  {level: '165', label: '소분류', n: 165, file: './io-2023-165.xlsx', size: '3.1MB'},
  {level: '380', label: '기본부문', n: 380, file: './io-2023-380.xlsx', size: '12.3MB'},
];
const LV = Object.fromEntries(LEVELS.map(l => [l.level, l]));
const FD = ['9111', '9112', '9121', '9122', '9131', '9132', '9140'];
const FD_NAMES = {'9111': '민간소비', '9112': '정부소비', '9121': '민간투자', '9122': '정부투자',
                  '9131': '재고증감', '9132': '귀중품', '9140': '수출'};

const won = x => x === null || x === undefined ? '' : Number(x).toLocaleString('ko-KR', {maximumFractionDigits: 0});
const won1 = x => x === null || x === undefined ? '' : Number(x).toLocaleString('ko-KR', {maximumFractionDigits: 1});
const pct = x => x === null || x === undefined ? '' : Number(x).toFixed(2);
const coef = x => x === null || x === undefined ? '' : Number(x).toFixed(4);
const raw = x => Number(x).toLocaleString('ko-KR', {maximumFractionDigits: 6});

let worker = null, nextId = 0;
const requests = new Map();
const loaded = {};          // level → {sectors, sheets, residuals}
let auxReady = false, basicList = null, basket = [], simResult = null, current = '';
const cache = {};           // `${fn}:${level}` → 결과

function el(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') e.className = v; else if (k.startsWith('on')) e[k] = v; else e.setAttribute(k, v);
  }
  for (const k of kids) e.append(k instanceof Node ? k : document.createTextNode(k ?? ''));
  return e;
}
function svg(tag, attrs = {}, text) {
  const e = document.createElementNS('http://www.w3.org/2000/svg', tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  if (text !== undefined) e.textContent = text;
  return e;
}
function shortError(e) {
  const s = String(e && e.message || e);
  const lines = s.trim().split('\n').filter(Boolean);
  return lines[lines.length - 1].replace(/^\w*Error:\s*/, '');
}

// ---------------------------------------------------------------- Worker
function startWorker() {
  worker = new Worker('./worker.js');
  worker.onmessage = ({data: d}) => {
    if (d.type === 'status') { $('status').textContent = d.text; return; }
    const p = requests.get(d.id); if (!p) return;
    requests.delete(d.id);
    d.ok ? p.resolve(d.result) : p.reject(Error(d.error));
  };
  worker.onerror = e => {
    for (const p of requests.values()) p.reject(Error(e.message || 'Worker 실행 오류'));
    requests.clear();
  };
}
function call(type, payload, transfer = []) {
  if (!worker) startWorker();
  return new Promise((resolve, reject) => {
    const id = ++nextId; requests.set(id, {resolve, reject});
    worker.postMessage({id, type, payload}, transfer);
  });
}
async function cached(fn, args) {
  const key = fn + ':' + JSON.stringify(args);
  if (!(key in cache)) cache[key] = await call(fn, args);
  return cache[key];
}
async function fetchBuffer(url) {
  // 저장소 최상위에 없으면 data/ 폴더에서 다시 찾는다
  const tries = [url, './data/' + url.slice(2)];
  for (const u of tries) {
    const r = await fetch(u);
    if (r.ok) return r.arrayBuffer();
  }
  throw Error(url + ' 다운로드 실패: 최상위와 data 폴더 모두에서 파일을 찾지 못했습니다.');
}

// ---------------------------------------------------------------- 불러오기
function renderLoadGrid() {
  $('load-grid').replaceChildren(...LEVELS.map(l => {
    const st = loaded[l.level];
    const b = el('button', {type: 'button', id: 'load-' + l.level},
      st ? l.label + ' 불러오기 완료' : l.label + '(' + l.n + ') 불러오기');
    b.disabled = !!st || busy;
    b.onclick = () => loadLevels([l.level]);
    return el('div', {class: 'load-card' + (st ? ' done' : '')},
      el('strong', {}, l.label + ' ' + l.n + '부문'), el('span', {}, 'Excel ' + l.size), b);
  }));
  $('load-all').disabled = busy || LEVELS.every(l => loaded[l.level]);
}
let busy = false;
async function ensureAux() {
  if (auxReady) return;
  const [map, main] = await Promise.all([fetchBuffer('./sector-map.xlsx'), fetchBuffer('./main-tables-2023.xlsx')]);
  await call('load_aux', {map, main}, [map, main]);
  basicList = await call('basic_list', []);
  auxReady = true;
}
async function loadLevels(levels) {
  if (busy) return;
  busy = true; renderLoadGrid();
  try {
    await ensureAux();
    for (const lv of levels) {
      if (loaded[lv]) continue;
      const l = LV[lv];
      $('status').textContent = l.label + ' Excel 내려받는 중…';
      const buffer = await fetchBuffer(l.file);
      const info = await call('load_level', {level: lv, label: l.label, buffer}, [buffer]);
      loaded[lv] = info;
      for (const k of Object.keys(cache)) delete cache[k];
      renderValidation();
      refreshLevelSelect(lv);
    }
    $('status').textContent = '준비 완료 · 불러온 분류: ' +
      LEVELS.filter(l => loaded[l.level]).map(l => l.label + '(' + l.n + ')').join(', ') + ' · 계수 검증 통과';
  } catch (e) {
    $('status').textContent = '불러오기 실패: ' + shortError(e) + ' — 새로고침 후 다시 시도하세요.';
  } finally {
    busy = false; renderLoadGrid();
  }
}
function renderValidation() {
  const out = {};
  for (const l of LEVELS) if (loaded[l.level]) {
    out[l.label + '(' + l.n + ')'] = {
      '시트 수': loaded[l.level].sheets.length,
      '총산출(백만원)': loaded[l.level].totalOutput,
      '국내 생산이 없는 부문': loaded[l.level].zeroOutput,
      '최대 절대 잔차': loaded[l.level].residuals,
    };
  }
  $('validation').textContent = JSON.stringify(out, null, 2) + '\n모든 잔차가 허용값(1e-6) 미만이면 통과입니다.';
}
function refreshLevelSelect(prefer) {
  const sel = $('level');
  sel.replaceChildren(...LEVELS.filter(l => loaded[l.level]).map(l => el('option', {value: l.level}, l.label + ' (' + l.n + '부문)')));
  sel.disabled = false;
  sel.value = prefer && loaded[prefer] ? prefer : (current && loaded[current] ? current : sel.options[0].value);
  setLevel(sel.value);
  // 비교 영역
  $('compare-search').disabled = false; $('compare-sector').disabled = false;
  filterCompare(); renderCompare();
}
$('level').onchange = () => setLevel($('level').value);
$('load-all').onclick = () => loadLevels(LEVELS.map(l => l.level));

async function setLevel(lv) {
  if (!loaded[lv]) return;
  const changed = current !== lv;
  current = lv;
  for (const id of ['sheet', 'row', 'row-search', 'col-search', 'structure-search', 'linkage-filter', 'linkage-search',
                    'fd-search', 'sector-search', 'sector', 'mode', 'item', 'amount', 'add', 'run', 'aggregate',
                    'heat-kind', 'price-search', 'price-sector', 'price-type', 'price-pct', 'price-add', 'wage-pct', 'price-run']) $(id).disabled = false;
  if (changed) {
    $('sheet').replaceChildren(...loaded[lv].sheets.map(s => el('option', {value: s}, s)));
    $('row-search').value = ''; $('row').replaceChildren();
    $('sector-search').value = '';
    fillSectors();
    $('aggregate').replaceChildren(el('option', {value: ''}, '집계하지 않음'),
      ...LEVELS.filter(l => Number(l.n) < Number(LV[lv].n)).map(l => el('option', {value: l.level}, l.label + '(' + l.n + ')로 묶어 보기')));
    basket = []; renderBasket(); invalidate();
    $('detail').hidden = true;
    $('price-search').value = ''; fillPriceSectors(); priceBasket = []; renderPriceBasket(); $('price-results').hidden = true;
  }
  renderPresets();
  await Promise.all([explore(), renderStructure(), renderLinkage(), renderFinal(), renderHeatmap()]).catch(e => {
    $('status').textContent = shortError(e);
  });
}

// ---------------------------------------------------------------- 공통 표
function table(container, columns, rows, opts = {}) {
  const state = container._state || (container._state = {key: opts.sort || null, dir: -1});
  const draw = () => {
    let data = rows.slice();
    if (opts.filter) data = data.filter(opts.filter);
    if (state.key) {
      const col = columns.find(c => c.key === state.key);
      data.sort((a, b) => {
        const x = a[state.key], y = b[state.key];
        const r = col && col.num ? (x - y) : String(x).localeCompare(String(y), 'ko');
        return r * state.dir;
      });
    }
    const head = el('tr', {}, ...columns.map(c => {
      const th = el('th', {class: c.num ? 'num sortable' : 'sortable', title: '눌러서 정렬'},
        c.label + (state.key === c.key ? (state.dir < 0 ? ' ▼' : ' ▲') : ''));
      th.onclick = () => { if (state.key === c.key) state.dir *= -1; else { state.key = c.key; state.dir = c.num ? -1 : 1; } draw(); };
      return th;
    }));
    const body = data.slice(0, opts.limit || data.length).map(r => {
      const tr = el('tr', {}, ...columns.map(c => el('td', {class: c.num ? 'num' : ''}, c.fmt ? c.fmt(r[c.key], r) : r[c.key])));
      if (opts.onRow) { tr.className = 'clickable'; tr.tabIndex = 0; tr.onclick = () => opts.onRow(r);
        tr.onkeydown = e => { if (e.key === 'Enter') opts.onRow(r); }; }
      return tr;
    });
    const dl = el('button', {type: 'button', class: 'small csv', title: '이 표를 CSV로 내려받기'}, 'CSV');
    dl.onclick = () => downloadCsv((opts.file || 'io-2023-table') + '.csv',
      [columns.map(c => c.label), ...data.map(r => columns.map(c => r[c.key] ?? ''))]);
    const t = el('table', {}, el('caption', {}, (opts.caption || '') + ' · ' + data.length + '행 ', dl), el('thead', {}, head), el('tbody', {}, ...body));
    container.replaceChildren(t);
  };
  container._redraw = draw;
  draw();
}
const match = q => r => !q || (r.code + ' ' + r.name).toLowerCase().includes(q.toLowerCase());

// ---------------------------------------------------------------- 1. 원자료
let exploreSeq = 0, rowChoices = [], cells = [];
async function explore() {
  const seq = ++exploreSeq;
  const d = await call('explore', [current, $('sheet').value, $('row').value || '']);
  if (seq !== exploreSeq) return;
  rowChoices = d.rows; filterRows(d.selected); cells = d.cells;
  $('unit').textContent = d.title + ' · 단위: ' + d.unit;
  $('explore-caption').textContent = LV[current].label + ' / ' + $('sheet').value + ' / 행: ' + d.selected + ' ' + d.name;
  renderCells();
}
function filterRows(selected) {
  const q = $('row-search').value.trim();
  const list = rowChoices.filter(match(q));
  $('row').replaceChildren(...list.map(r => el('option', {value: r.code}, r.code + ' ' + r.name)));
  if (selected && list.some(r => r.code === selected)) $('row').value = selected;
}
function renderCells() {
  const q = $('col-search').value.trim();
  $('cells').replaceChildren(...cells.filter(match(q)).map(c =>
    el('tr', {}, el('td', {}, c.code), el('td', {}, c.name),
      el('td', {class: 'num'}, c.value === null ? '빈 셀' : typeof c.value === 'number' ? raw(c.value) : c.value))));
}
$('sheet').onchange = () => { $('row-search').value = ''; $('row').replaceChildren(); explore(); };
$('row').onchange = explore;
$('row-search').oninput = () => { filterRows(); if ($('row').value) explore(); else { $('cells').replaceChildren(); $('explore-caption').textContent = '일치하는 행이 없습니다.'; } };
$('col-search').oninput = renderCells;

// ---------------------------------------------------------------- 2. 산업구조
async function renderStructure() {
  const lv = current;
  const d = await cached('structure', [lv]);
  if (lv !== current) return;
  const t = d.totals;
  $('structure-cards').replaceChildren(...[
    ['총산출', won(t.output), '백만원'], ['부가가치(GDP 기준)', won(t.valueAdded), '백만원'],
    ['중간투입률', pct(t.midRate), '%'], ['부가가치율', pct(t.vaRate), '%'],
    ['국산화율', pct(t.localRate), '%'], ['수입의존도', pct(t.importDep), '%'],
    ['수출률', pct(t.exportRate), '%'], ['수입률', pct(t.importRate), '%'],
  ].map(([k, v, u]) => el('div', {class: 'card'}, k, el('strong', {}, v), el('span', {}, u))));
  const cols = [
    {key: 'code', label: '코드'}, {key: 'name', label: '부문'},
    {key: 'output', label: '총산출', num: true, fmt: won}, {key: 'outputShare', label: '산출비중', num: true, fmt: pct},
    {key: 'valueAdded', label: '부가가치', num: true, fmt: won}, {key: 'vaRate', label: '부가가치율', num: true, fmt: pct},
    {key: 'midRate', label: '중간투입률', num: true, fmt: pct}, {key: 'localRate', label: '국산화율', num: true, fmt: pct},
    {key: 'importDep', label: '수입의존도', num: true, fmt: pct}, {key: 'exportRate', label: '수출률', num: true, fmt: pct},
    {key: 'importRate', label: '수입률', num: true, fmt: pct}, {key: 'midDemandRate', label: '중간수요율', num: true, fmt: pct},
  ];
  const box = $('structure-table'); box._state = null;
  table(box, cols, d.rows, {file: 'structure-' + lv, sort: 'output', caption: LV[lv].label + ' 부문별 산업구조 (총산출·부가가치: 백만원, 비율: %)',
    filter: r => match($('structure-search').value.trim())(r), onRow: r => showDetail(r.code)});
}
$('structure-search').oninput = () => $('structure-table')._redraw && $('structure-table')._redraw();

function miniBars(rows, valueKey, labelFn, unit) {
  const max = Math.max(...rows.map(r => Math.abs(r[valueKey])), 1);
  return el('div', {}, ...rows.map(r => el('div', {class: 'bar'}, labelFn(r) + ' · ' + won(r[valueKey]) + unit,
    el('div', {class: 'track'}, (() => { const f = el('div', {class: 'fill'}); f.style.width = (Math.abs(r[valueKey]) / max * 100) + '%'; return f; })()))));
}
async function showDetail(code) {
  const lv = current;
  try {
    const d = await call('sector_detail', [lv, code]);
    const m = d.multipliers;
    const box = $('detail'); box.hidden = false;
    const kids = [
      el('h3', {}, '부문 상세: ' + d.code + ' ' + d.name + ' (' + LV[lv].label + ')'),
      el('div', {class: 'cards'}, ...[
        ['총산출', won(d.output), '백만원'], ['중간투입', won(d.midInput), '백만원'], ['부가가치', won(d.valueAdded), '백만원'],
        ['수입', won(d.imports), '백만원'], ['생산유발계수(열합)', coef(m.production), '원/원'],
        ['부가가치유발계수', coef(m.valueAdded), '원/원'], ['수입유발계수', coef(m.imports), '원/원'],
        ['영향력 / 감응도', coef(m.backward) + ' / ' + coef(m.forward), '무단위'],
      ].map(([k, v, u]) => el('div', {class: 'card'}, k, el('strong', {}, v), el('span', {}, u)))),
      el('div', {class: 'grid2'},
        el('div', {}, el('h4', {}, '주요 투입처 상위 10 (이 부문이 사 오는 중간재, 백만원)'),
          miniBars(d.inputs, 'value', r => r.name + ' (투입계수 ' + coef(r.coef) + ')', '')),
        el('div', {}, el('h4', {}, '주요 판매처 상위 10 (이 부문 생산물을 중간재로 쓰는 부문, 백만원)'),
          miniBars(d.sales, 'value', r => r.name, ''))),
      el('div', {class: 'grid2'},
        el('div', {}, el('h4', {}, '최종수요 구성 (백만원)'), miniBars(d.finalDemand, 'value', r => r.name, '')),
        el('div', {}, el('h4', {}, '부가가치 구성 (백만원)'), miniBars(d.vaParts, 'value', r => r.name, ''))),
    ];
    if (d.trend) {
      const years = d.trend.years;
      const rows = Object.entries(d.trend.series).filter(([, v]) => v);
      kids.push(el('h4', {}, '소속 대분류 「' + d.trend.name + '」 2020~2023년 추이 (주요표)'),
        el('div', {class: 'scroll'}, el('table', {}, el('thead', {}, el('tr', {}, el('th', {}, '지표'), ...years.map(y => el('th', {class: 'num'}, y + '년')))),
          el('tbody', {}, ...rows.map(([k, v]) => el('tr', {}, el('td', {}, k), ...v.map(x => el('td', {class: 'num'}, Number(x).toFixed(2)))))))));
    }
    box.replaceChildren(...kids);
    box.scrollIntoView({behavior: 'smooth', block: 'start'});
  } catch (e) { $('status').textContent = shortError(e); }
}

// ---------------------------------------------------------------- 3. 연관효과
let linkRows = [];
async function renderLinkage() {
  const lv = current;
  const d = await cached('linkage', [lv]);
  if (lv !== current) return;
  linkRows = d.rows;
  drawScatter();
  const cols = [
    {key: 'code', label: '코드'}, {key: 'name', label: '부문'},
    {key: 'backward', label: '영향력계수(후방)', num: true, fmt: coef}, {key: 'forward', label: '감응도계수(전방)', num: true, fmt: coef},
    {key: 'group', label: '유형'},
    {key: 'production', label: '생산승수', num: true, fmt: coef}, {key: 'valueAdded', label: '부가가치승수', num: true, fmt: coef},
    {key: 'imports', label: '수입승수', num: true, fmt: coef},
    {key: 'jobs', label: '취업유발계수', num: true, fmt: v => v == null ? '' : pct(v)},
  ];
  const box = $('linkage-table'); box._state = null;
  table(box, cols, linkRows, {file: 'linkage-' + lv, sort: 'backward', caption: LV[lv].label + ' 영향력·감응도계수와 승수 (계수 무단위, 취업유발계수 명/10억원)',
    filter: r => (!$('linkage-filter').value || r.group === $('linkage-filter').value) && match($('linkage-search').value.trim())(r),
    onRow: r => { showDetail(r.code); }});
}
$('linkage-filter').onchange = () => { $('linkage-table')._redraw && $('linkage-table')._redraw(); drawScatter(); };
$('linkage-search').oninput = () => { $('linkage-table')._redraw && $('linkage-table')._redraw(); drawScatter(); };
function drawScatter() {
  const W = 760, H = 440, P = {l: 56, r: 16, t: 16, b: 46};
  const xs = linkRows.map(r => r.backward), ys = linkRows.map(r => r.forward);
  const x0 = Math.min(0.5, ...xs) - 0.02, x1 = Math.max(1.5, ...xs) + 0.02;
  const y0 = Math.min(0.3, ...ys) - 0.02, y1 = Math.max(...ys) * 1.04;
  const sx = v => P.l + (v - x0) / (x1 - x0) * (W - P.l - P.r);
  const sy = v => H - P.b - (v - y0) / (y1 - y0) * (H - P.t - P.b);
  const s = svg('svg', {viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': '영향력계수와 감응도계수 산점도'});
  s.append(svg('rect', {x: sx(1), y: P.t, width: W - P.r - sx(1), height: sy(1) - P.t, class: 'quad'}));
  for (const t of [0.6, 0.8, 1, 1.2, 1.4]) if (t > x0 && t < x1) {
    s.append(svg('line', {x1: sx(t), x2: sx(t), y1: P.t, y2: H - P.b, class: t === 1 ? 'axis1' : 'grid'}));
    s.append(svg('text', {x: sx(t), y: H - P.b + 16, 'text-anchor': 'middle', class: 'tick'}, t.toFixed(1)));
  }
  const ystep = y1 - y0 > 3 ? 1 : 0.5;
  for (let t = Math.ceil(y0 / ystep) * ystep; t < y1; t += ystep) {
    s.append(svg('line', {x1: P.l, x2: W - P.r, y1: sy(t), y2: sy(t), class: Math.abs(t - 1) < 1e-9 ? 'axis1' : 'grid'}));
    s.append(svg('text', {x: P.l - 6, y: sy(t) + 4, 'text-anchor': 'end', class: 'tick'}, t.toFixed(1)));
  }
  if (y0 < 1 && y1 > 1) s.append(svg('line', {x1: P.l, x2: W - P.r, y1: sy(1), y2: sy(1), class: 'axis1'}));
  s.append(svg('text', {x: (W + P.l) / 2, y: H - 8, 'text-anchor': 'middle', class: 'label'}, '영향력계수 (후방연쇄효과) →'));
  s.append(svg('text', {x: 14, y: (H - P.b) / 2, transform: `rotate(-90 14 ${(H - P.b) / 2})`, 'text-anchor': 'middle', class: 'label'}, '감응도계수 (전방연쇄효과) →'));
  s.append(svg('text', {x: W - P.r - 6, y: P.t + 16, 'text-anchor': 'end', class: 'qlabel'}, '핵심(주도)산업'));
  const q = $('linkage-search').value.trim(), g = $('linkage-filter').value;
  const top = new Set([...linkRows].sort((a, b) => (b.backward + b.forward) - (a.backward + a.forward)).slice(0, 8).map(r => r.code));
  for (const r of linkRows) {
    const on = (!g || r.group === g) && match(q)(r);
    const c = svg('circle', {cx: sx(r.backward), cy: sy(r.forward), r: on ? 5 : 3, class: on ? 'dot' : 'dot off'});
    c.append(svg('title', {}, `${r.code} ${r.name}\n영향력 ${coef(r.backward)} · 감응도 ${coef(r.forward)}`));
    c.onclick = () => showDetail(r.code);
    s.append(c);
    if (on && (top.has(r.code) || (q && linkRows.filter(match(q)).length <= 8)))
      s.append(svg('text', {x: sx(r.backward) + 7, y: sy(r.forward) - 6, class: 'dlabel'}, r.name));
  }
  $('scatter').replaceChildren(s);
}

// ---------------------------------------------------------------- 4. 최종수요
let fdData = null;
async function renderFinal() {
  const lv = current;
  const d = await cached('final_demand', [lv]);
  if (lv !== current) return;
  fdData = d;
  const cols = [
    {key: 'name', label: '최종수요 항목'},
    {key: 'domestic', label: '국산 최종수요', num: true, fmt: won}, {key: 'imported', label: '수입 최종수요', num: true, fmt: won},
    {key: 'production', label: '생산유발액', num: true, fmt: won}, {key: 'pCoef', label: '생산유발계수', num: true, fmt: coef},
    {key: 'pShare', label: '생산유발 구성비', num: true, fmt: pct},
    {key: 'valueAdded', label: '부가가치유발액', num: true, fmt: won}, {key: 'vCoef', label: '부가가치유발계수', num: true, fmt: coef},
    {key: 'imports', label: '수입유발액', num: true, fmt: won}, {key: 'mCoef', label: '수입유발계수', num: true, fmt: coef},
    {key: 'jobs', label: '취업유발인원(명)', num: true, fmt: v => v == null ? '' : won(v)},
    {key: 'jCoef', label: '취업유발계수(명/10억원)', num: true, fmt: v => v == null ? '' : pct(v)},
  ];
  const box = $('fd-items'); box._state = {key: null, dir: -1};
  table(box, cols, d.items, {file: 'final-demand-' + lv, caption: LV[lv].label + ' 최종수요 항목별 유발효과 (금액: 백만원, 계수: 원/원, 구성비: %)'});
  // 계수 비교 막대
  const items = d.items.filter(i => i.code !== '9190' && i.code !== '9131');
  const max = Math.max(...items.map(i => i.pCoef));
  $('fd-bars').replaceChildren(el('h3', {}, '항목별 유발계수 비교 (최종수요 1원당)'),
    ...items.map(i => el('div', {class: 'bar'}, i.name + ' · 생산 ' + coef(i.pCoef) + ' / 부가가치 ' + coef(i.vCoef) + ' / 수입 ' + coef(i.mCoef),
      el('div', {class: 'track'}, (() => { const f = el('div', {class: 'fill'}); f.style.width = (i.pCoef / max * 100) + '%'; return f; })()),
      el('div', {class: 'track thin'}, (() => { const f = el('div', {class: 'fill va'}); f.style.width = (i.vCoef / max * 100) + '%'; return f; })()),
      el('div', {class: 'track thin'}, (() => { const f = el('div', {class: 'fill im'}); f.style.width = (i.mCoef / max * 100) + '%'; return f; })()))),
    el('p', {class: 'note'}, '굵은 막대: 생산유발계수, 가는 막대: 부가가치유발계수(초록)·수입유발계수(주황). 재고증감은 부호가 섞여 있어 제외했습니다.' + (lv !== '33' ? ' 취업유발은 소속 대분류 취업계수를 적용한 근사치입니다.' : '')));
  const dcols = [{key: 'code', label: '코드'}, {key: 'name', label: '부문'}, {key: 'total', label: '생산유발액 합계', num: true, fmt: won},
    ...FD.map(k => ({key: k, label: FD_NAMES[k] + '(%)', num: true, fmt: pct}))];
  const dbox = $('fd-dep'); dbox._state = null;
  table(dbox, dcols, d.dependency, {file: 'dependency-' + lv, sort: '9140', caption: LV[lv].label + ' 부문별 생산유발의존도 (%)',
    filter: r => match($('fd-search').value.trim())(r)});
}
$('fd-search').oninput = () => $('fd-dep')._redraw && $('fd-dep')._redraw();

// ---------------------------------------------------------------- 5. 시뮬레이션
function fillSectors() {
  const q = $('sector-search').value.trim();
  const list = loaded[current].sectors.filter(match(q));
  const opts = list.map(s => el('option', {value: s.code}, s.code + ' ' + s.name));
  if ($('mode').value === 'pct' && !q) opts.unshift(el('option', {value: '*'}, '★ 전체 산업 (선택한 최종수요 항목 전체를 같은 비율로)'));
  $('sector').replaceChildren(...opts);
}
$('sector-search').oninput = fillSectors;
function invalidate() { simResult = null; $('results').hidden = true; }
function syncMode() {
  const pctMode = $('mode').value === 'pct';
  $('item-label').hidden = !pctMode;
  $('amount-label').textContent = pctMode ? '변화율 (%)' : '변화액 (백만원)';
  $('amount-note').textContent = pctMode
    ? '선택한 부문의 기존 국산 최종수요(선택 항목) × 변화율이 충격액이 됩니다. 예: 수출 −10%.'
    : '1,000백만원 = 10억원. 감소 충격은 음수(예: −1000)로 입력하세요.';
  if (pctMode && Number($('amount').value) === 1000) $('amount').value = -10;
  if (!pctMode && Number($('amount').value) === -10) $('amount').value = 1000;
}
$('mode').onchange = () => { syncMode(); fillSectors(); };
const shockKey = () => basket.map(s => s.pct != null ? {code: s.code, pct: s.pct, item: s.item} : {code: s.code, amount: s.amount});
function renderBasket() {
  $('basket').replaceChildren(...basket.map((s, i) => {
    const b = el('button', {type: 'button', class: 'small'}, '삭제');
    b.onclick = () => { basket.splice(i, 1); invalidate(); renderBasket(); };
    const what = s.pct != null ? s.itemName + ' ' + (s.pct > 0 ? '+' : '') + s.pct + '%' : (s.amount > 0 ? '+' : '') + won1(s.amount) + '백만원';
    return el('li', {}, s.code + ' ' + s.name + ' / ' + what + ' ', b);
  }));
  $('sim-status').textContent = basket.length ? basket.length + '개 입력 준비(' + LV[current].label + '). 계산 버튼을 누르세요.' : '시나리오를 추가하세요.';
}
$('scenario').onsubmit = e => {
  e.preventDefault();
  const amount = Number($('amount').value);
  const s = $('sector').value === '*' ? {code: '*', name: '전체 산업'} : loaded[current].sectors.find(s => s.code === $('sector').value);
  if (!s || !Number.isFinite(amount) || amount === 0) { $('sim-status').textContent = '부문과 0이 아닌 값을 확인하세요.'; return; }
  if ($('mode').value === 'pct') {
    if (amount < -100) { $('sim-status').textContent = '감소율은 −100%보다 작을 수 없습니다.'; return; }
    basket.push({...s, pct: amount, item: $('item').value, itemName: $('item').selectedOptions[0].textContent});
  } else if (s.code === '*') { $('sim-status').textContent = '전체 산업 충격은 비율 입력 방식에서만 쓸 수 있습니다.'; return; }
  else basket.push({...s, amount});
  invalidate(); renderBasket();
};
$('clear').onclick = () => { basket = []; invalidate(); renderBasket(); };
$('aggregate').onchange = invalidate;
$('run').onclick = () => runSimulation();
async function runSimulation() {
  if (!basket.length) { $('sim-status').textContent = '시나리오를 먼저 추가하세요.'; return; }
  const snapshot = JSON.stringify({shocks: shockKey(), aggregate: $('aggregate').value});
  $('run').disabled = true; $('sim-status').textContent = '계산 중…';
  try {
    const r = await call('simulate', [current, snapshot]);
    if (snapshot !== JSON.stringify({shocks: shockKey(), aggregate: $('aggregate').value})) {
      $('sim-status').textContent = '입력이 바뀌었습니다. 다시 계산하세요.'; return;
    }
    simResult = r; renderResult();
    $('sim-status').textContent = '계산 완료. 아래 결과는 현재 시나리오(' + LV[current].label + ')에 대응합니다.';
  } catch (e) { invalidate(); $('sim-status').textContent = '계산 실패: ' + shortError(e); }
  finally { $('run').disabled = false; }
}
function renderResult() {
  const r = simResult; $('results').hidden = false;
  const jo = x => (x / 1e6).toLocaleString('ko-KR', {maximumFractionDigits: 2}) + '조원';
  const sign = x => x >= 0 ? '증가' : '감소';
  const top = (r.aggregate || r.rows).filter(x => x.production !== 0).slice(0, 3).map(x => x.name + '(' + jo(x.production) + ')');
  const vp = r.vaParts || {};
  const lab = r.aggregateLevel ? LV[r.aggregateLevel].label : LV[r.level].label;
  $('sim-summary').textContent = '【결과 요약】 국산품 최종수요가 ' + jo(Math.abs(r.demand)) + ' ' + sign(r.demand) + '하면, 경제 전체 생산은 ' +
    jo(Math.abs(r.production)) + ' ' + sign(r.production) + '합니다(생산승수 ' + coef(r.multiplier) + '). 부가가치는 ' + jo(Math.abs(r.valueAdded)) + ' ' + sign(r.valueAdded) +
    '하여 2023년 전체 부가가치의 ' + Math.abs(r.vaShareOfTotal).toFixed(3) + '%에 해당하고, 이 가운데 가계 근로소득(피용자보수)이 ' +
    (vp['9610'] / r.valueAdded * 100).toFixed(1) + '%, 순생산세(세수)가 ' + (vp['9630'] / r.valueAdded * 100).toFixed(1) + '%입니다. ' +
    (r.jobs !== undefined ? '취업자는 약 ' + won(Math.abs(r.jobs)) + '명 ' + sign(r.jobs) + '합니다. ' : '') +
    '수입은 ' + jo(Math.abs(r.imports)) + ' ' + sign(r.imports) + '하여 효과 일부가 해외로 누출됩니다. 영향이 가장 큰 ' + lab + ' 부문은 ' + top.join(', ') + '입니다.';
  $('va-cards').replaceChildren(...[['9610', '피용자보수', '가계 근로소득'], ['9621', '영업잉여', '기업 이윤·자영업 소득'],
    ['9622', '고정자본소모', '감가상각'], ['9630', '순생산세', '세수(보조금 공제)']].map(([k, nm, u]) =>
    el('div', {class: 'card'}, nm, el('strong', {}, won1(vp[k])), el('span', {}, '백만원 · ' + u + ' · ' + (vp[k] / r.valueAdded * 100).toFixed(1) + '%'))));
  const cards = [['최종수요 변화', won1(r.demand), '백만원'], ['생산유발액', won1(r.production), '백만원'],
    ['부가가치유발액', won1(r.valueAdded), '백만원'], ['수입유발액', won1(r.imports), '백만원'],
    ['생산승수', coef(r.multiplier), '생산유발액 ÷ 최종수요 변화'], ['부가가치승수', coef(r.vaMultiplier), '부가가치유발액 ÷ 최종수요 변화']];
  if (r.jobs !== undefined) cards.push(['취업유발인원', won1(r.jobs), '명'], ['고용유발인원(임금근로자)', won1(r.wageJobs), '명']);
  $('cards').replaceChildren(...cards.map(([k, v, u]) => el('div', {class: 'card'}, k, el('strong', {}, v), el('span', {}, u))));
  const mixed = r.rows.some(x => x.shock > 0) && r.rows.some(x => x.shock < 0);
  const conv = (r.resolved || []).filter(x => x.pct != null).map(x =>
    x.code + ' ' + x.name + ' ' + x.item + ' ' + won(x.base) + '백만원 × ' + x.pct + '% = ' + won1(x.amount) + '백만원').join('; ');
  $('emp-note').textContent = (conv ? '비율 입력 환산: ' + conv + '. ' : '') + (r.approxEmployment
    ? '취업·고용유발인원은 각 부문이 속한 대분류의 2023년 취업·고용계수(명/10억원)를 적용한 근사치입니다.'
    : '취업·고용유발인원은 2023년 대분류 취업·고용계수(명/10억원)를 적용했습니다.') +
    (mixed ? ' 증가와 감소가 섞인 시나리오이므로 승수는 순변화 기준이며, 부문별 결과를 함께 해석하세요.' : '');
  const rows = r.aggregate || r.rows;
  const unitLabel = r.aggregate ? LV[r.aggregateLevel].label + '로 집계' : LV[r.level].label;
  const max = Math.max(...rows.map(x => Math.abs(x.production)), 1);
  $('bars').replaceChildren(...rows.slice(0, 15).map(x => {
    const f = el('div', {class: 'fill' + (x.production < 0 ? ' neg' : '')}); f.style.width = (Math.abs(x.production) / max * 100) + '%';
    return el('div', {class: 'bar'}, x.name + ' · ' + won1(x.production) + '백만원 (직접 ' + won1(x.shock) + ' + 간접 ' + won1(x.indirect) + ')',
      el('div', {class: 'track'}, f));
  }));
  const cols = [{key: 'code', label: '코드'}, {key: 'name', label: '부문'}, {key: 'shock', label: '직접(입력)', num: true, fmt: won1},
    {key: 'production', label: '생산유발', num: true, fmt: won1}, {key: 'indirect', label: '간접 생산', num: true, fmt: won1},
    {key: 'valueAdded', label: '부가가치유발', num: true, fmt: won1}, {key: 'imports', label: '수입유발', num: true, fmt: won1}];
  if (r.jobs !== undefined) cols.push({key: 'jobs', label: '취업(명)', num: true, fmt: won1}, {key: 'wageJobs', label: '고용(명)', num: true, fmt: won1});
  const box = $('result-table'); box._state = null;
  table(box, cols, rows, {file: 'simulation-table', caption: '부문별 유발효과 — ' + unitLabel + ', 금액 백만원', sort: null});
  $('download').textContent = '전체 결과 CSV 다운로드 (' + unitLabel + ' ' + rows.length + '개 부문)';
}
function downloadCsv(name, rows) {
  const url = URL.createObjectURL(new Blob(['\ufeff' + rows.map(r => r.map(csvCell).join(',')).join('\r\n')], {type: 'text/csv;charset=utf-8'}));
  const a = el('a', {href: url, download: name}); document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function csvCell(v) { let s = String(v ?? ''); if (/^[=+@-]/.test(s) && typeof v !== 'number') s = "'" + s; return '"' + s.replaceAll('"', '""') + '"'; }
$('download').onclick = () => {
  if (!simResult) return;
  const rows = simResult.aggregate || simResult.rows;
  const head = ['부문코드', '부문명', '직접_백만원', '생산유발_백만원', '간접생산_백만원', '부가가치유발_백만원', '수입유발_백만원', '취업유발_명', '고용유발_명'];
  const body = rows.map(r => [r.code, r.name, r.shock, r.production, r.indirect, r.valueAdded, r.imports, r.jobs ?? '', r.wageJobs ?? '']);
  const url = URL.createObjectURL(new Blob(['﻿' + [head, ...body].map(r => r.map(csvCell).join(',')).join('\r\n')], {type: 'text/csv;charset=utf-8'}));
  const a = el('a', {href: url, download: 'io-2023-simulation-' + (simResult.aggregateLevel || simResult.level) + '.csv'}); a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

// ---------------------------------------------------------------- 6. 분류 수준 비교
function filterCompare() {
  if (!basicList) return;
  const q = $('compare-search').value.trim();
  const prev = $('compare-sector').value;
  const list = basicList.filter(r => !q || (r.code + ' ' + r.name + ' ' + r.path).includes(q));
  $('compare-sector').replaceChildren(...list.map(r => el('option', {value: r.code}, r.code + ' ' + r.name)));
  if (list.some(r => r.code === prev)) $('compare-sector').value = prev;
}
$('compare-search').oninput = () => { filterCompare(); renderCompare(); };
$('compare-sector').onchange = renderCompare;
async function renderCompare() {
  const code = $('compare-sector').value;
  if (!code) { $('compare-table').replaceChildren(); return; }
  try {
    const d = await call('compare', [code]);
    const b = basicList.find(r => r.code === code);
    $('compare-path').textContent = '분류 경로: ' + b.path + ' > ' + b.name;
    const cols = [{key: 'label', label: '분류'}, {key: 'code', label: '코드'}, {key: 'name', label: '해당 부문'},
      {key: 'output', label: '총산출(백만원)', num: true, fmt: v => v === undefined ? '미로드' : won(v)},
      {key: 'production', label: '생산승수', num: true, fmt: v => v === undefined ? '미로드' : coef(v)},
      {key: 'valueAdded', label: '부가가치승수', num: true, fmt: v => v === undefined ? '미로드' : coef(v)},
      {key: 'imports', label: '수입승수', num: true, fmt: v => v === undefined ? '미로드' : coef(v)},
      {key: 'own', label: '자기부문 유발', num: true, fmt: v => v === undefined ? '미로드' : coef(v)},
      {key: 'backward', label: '영향력계수', num: true, fmt: v => v === undefined ? '미로드' : coef(v)},
      {key: 'forward', label: '감응도계수', num: true, fmt: v => v === undefined ? '미로드' : coef(v)},
      {key: 'jobs', label: '취업유발계수(명/10억원)', num: true, fmt: v => v === undefined ? '미로드' : pct(v)}];
    const box = $('compare-table'); box._state = {key: null, dir: -1};
    table(box, cols, d.rows, {file: 'compare-' + code, caption: '기본부문 ' + code + ' ' + b.name + ' — 소속 부문의 분류별 승수 (승수·계수: 무단위, 총산출: 백만원)'});
  } catch (e) { $('compare-table').replaceChildren(el('p', {}, shortError(e))); }
}

renderLoadGrid();

// ---------------------------------------------------------------- 대표 시나리오
const PRESETS = [
  {label: '반도체(집적회로) 수출 10% 감소', shocks: [{basic: '3102', pct: -10, item: '9140'}]},
  {label: '자동차 수출 10% 감소 (관세 충격)', shocks: [{basic: '4011', pct: -10, item: '9140'}, {basic: '4012', pct: -10, item: '9140'}]},
  {label: '정부 SOC 투자 1조원 확대', shocks: [{basic: '5111', amount: 500000}, {basic: '5112', amount: 300000}, {basic: '5113', amount: 200000}]},
  {label: '관광·외식 소비 10% 증가', shocks: [{basic: '5811', pct: 10, item: '9111'}, {basic: '5812', pct: 10, item: '9111'}, {basic: '5820', pct: 10, item: '9111'}]},
  {label: '수출 전체 5% 감소 (세계 경기 둔화)', shocks: [{code: '*', pct: -5, item: '9140'}]},
  {label: '민간소비 전체 1% 증가', shocks: [{code: '*', pct: 1, item: '9111'}]},
];
const PRICE_PRESETS = [
  {label: '국제유가(원유 수입가격) 10% 상승', imports: [{basic: '0621', pct: 10}]},
  {label: 'LNG 수입가격 20% 상승', imports: [{basic: '0622', pct: 20}]},
  {label: '전기요금(발전) 10% 인상', domestic: ['4501', '4502', '4503', '4504', '4505', '4506'].map(b => ({basic: b, pct: 10}))},
  {label: '임금 5% 상승', wage: 5},
];
const ITEM_NAMES = {'9140': '수출', '9111': '민간소비지출', '9112': '정부소비지출', '9121': '민간고정자본형성', '9122': '정부고정자본형성', '9190': '최종수요계'};
function finestLevel() { return ['380', '165', '83', '33'].find(l => loaded[l]); }
function toLevel(basic, lv) {
  const b = basicList && basicList.find(r => r.code === basic);
  return b ? b.levels[lv] : null;
}
function mapShocks(list, lv) {
  const out = [], seen = new Map();
  for (const sh of list) {
    const code = sh.code === '*' ? '*' : toLevel(sh.basic, lv);
    if (!code) continue;
    const name = code === '*' ? '전체 산업' : (loaded[lv].sectors.find(s => s.code === code) || {}).name;
    const key = code + '|' + (sh.item || '') + '|' + (sh.pct != null ? 'p' : 'a');
    if (seen.has(key)) {  // 상위 분류로 묶이면 금액은 합산, 비율은 한 번만 적용
      if (sh.amount != null) seen.get(key).amount += sh.amount;
      continue;
    }
    const o = sh.pct != null ? {code, name, pct: sh.pct, item: sh.item, itemName: ITEM_NAMES[sh.item]} : {code, name, amount: sh.amount};
    seen.set(key, o); out.push(o);
  }
  return out;
}
function renderPresets() {
  const ok = !!basicList;
  $('presets').replaceChildren(...PRESETS.map(p => { const b = el('button', {type: 'button', class: 'preset'}, p.label); b.disabled = !ok; b.onclick = () => runPreset(p); return b; }));
  $('price-presets').replaceChildren(...PRICE_PRESETS.map(p => { const b = el('button', {type: 'button', class: 'preset'}, p.label); b.disabled = !ok; b.onclick = () => runPricePreset(p); return b; }));
}
async function useFinest() {
  const lv = finestLevel();
  if ($('level').value !== lv) { $('level').value = lv; await setLevel(lv); }
  return lv;
}
async function runPreset(p) {
  const lv = await useFinest();
  basket = mapShocks(p.shocks, lv); renderBasket(); invalidate();
  $('aggregate').value = lv === '33' ? '' : '33';
  $('sim-status').textContent = '「' + p.label + '」 시나리오를 ' + LV[lv].label + ' 기준으로 계산합니다' + (lv !== '380' ? ' (기본부문을 불러오면 더 정밀합니다)' : '') + '…';
  await runSimulation();
  $('simulation').scrollIntoView({behavior: 'smooth', block: 'start'});
}

// ---------------------------------------------------------------- 물가파급효과
let priceBasket = [];
function fillPriceSectors() {
  const q = $('price-search').value.trim();
  $('price-sector').replaceChildren(...loaded[current].sectors.filter(match(q)).map(s => el('option', {value: s.code}, s.code + ' ' + s.name)));
}
$('price-search').oninput = fillPriceSectors;
function renderPriceBasket() {
  $('price-basket').replaceChildren(...priceBasket.map((p, i) => {
    const b = el('button', {type: 'button', class: 'small'}, '삭제');
    b.onclick = () => { priceBasket.splice(i, 1); renderPriceBasket(); };
    return el('li', {}, (p.type === 'imports' ? '수입품 ' : '국내 가격 고정 ') + p.code + ' ' + p.name + ' ' + (p.pct > 0 ? '+' : '') + p.pct + '% ', b);
  }));
}
$('price-add').onclick = () => {
  const pctv = Number($('price-pct').value), s = loaded[current].sectors.find(s => s.code === $('price-sector').value);
  if (!s || !Number.isFinite(pctv) || pctv === 0) { $('price-status').textContent = '부문과 0이 아닌 변화율을 확인하세요.'; return; }
  priceBasket.push({type: $('price-type').value, code: s.code, name: s.name, pct: pctv}); renderPriceBasket();
};
$('price-clear').onclick = () => { priceBasket = []; $('wage-pct').value = 0; renderPriceBasket(); $('price-results').hidden = true; $('price-status').textContent = ''; };
$('price-run').onclick = () => runPrice();
async function runPrice(label) {
  const payload = {imports: priceBasket.filter(p => p.type === 'imports').map(p => ({code: p.code, pct: p.pct})),
                   domestic: priceBasket.filter(p => p.type === 'domestic').map(p => ({code: p.code, pct: p.pct})),
                   wage: Number($('wage-pct').value) || 0};
  $('price-status').textContent = '계산 중…';
  try {
    const r = await call('price_effect', [current, JSON.stringify(payload)]);
    $('price-results').hidden = false;
    const top = r.rows.filter(x => !x.exogenous).slice(0, 3).map(x => x.name + '(' + x.price.toFixed(2) + '%)');
    const contrib = [...r.rows].sort((a, b) => b.cpiContrib - a.cpiContrib).slice(0, 3).map(x => x.name);
    $('price-summary').textContent = '【결과 요약】 ' + (label ? '「' + label + '」의 경우 ' : '입력한 비용 인상이 ') + '모두 가격에 전가되면 소비자물가는 약 ' +
      r.cpi.toFixed(2) + '%, 생산자물가는 약 ' + r.ppi.toFixed(2) + '% 오릅니다(' + LV[current].label + ' 기준). 가격이 가장 많이 오르는 부문은 ' + top.join(', ') +
      '이고, 소비자물가 상승에 가장 크게 기여하는 부문은 ' + contrib.join(', ') + '입니다.';
    $('price-cards').replaceChildren(...[['소비자물가 영향', r.cpi.toFixed(3), '% (민간소비 가중 근사)'], ['생산자물가 영향', r.ppi.toFixed(3), '% (총산출 가중)']]
      .map(([k, v, u]) => el('div', {class: 'card'}, k, el('strong', {}, v), el('span', {}, u))));
    const max = Math.max(...r.rows.map(x => Math.abs(x.price)), 1e-9);
    $('price-bars').replaceChildren(...r.rows.slice(0, 15).map(x => {
      const f = el('div', {class: 'fill' + (x.exogenous ? ' im' : '')}); f.style.width = (Math.abs(x.price) / max * 100) + '%';
      return el('div', {class: 'bar'}, x.name + ' · ' + x.price.toFixed(3) + '%' + (x.exogenous ? ' (외생 지정)' : ''), el('div', {class: 'track'}, f));
    }));
    const cols = [{key: 'code', label: '코드'}, {key: 'name', label: '부문'}, {key: 'price', label: '가격 변화율(%)', num: true, fmt: v => v.toFixed(4)},
      {key: 'consWeight', label: '민간소비 비중(%)', num: true, fmt: pct}, {key: 'cpiContrib', label: '소비자물가 기여(%p)', num: true, fmt: v => v.toFixed(4)},
      {key: 'exogenous', label: '외생 지정', fmt: v => v ? '예' : ''}];
    const box = $('price-table'); box._state = null;
    table(box, cols, r.rows, {file: 'price-effect-' + current, sort: 'price', caption: LV[current].label + ' 부문별 가격 파급효과 (%)'});
    $('price-status').textContent = '계산 완료 (' + LV[current].label + ' 기준).';
  } catch (e) { $('price-status').textContent = '계산 실패: ' + shortError(e); }
}
async function runPricePreset(p) {
  const lv = await useFinest();
  priceBasket = [];
  for (const [type, list] of [['imports', p.imports || []], ['domestic', p.domestic || []]]) {
    const seen = new Set();
    for (const it of list) {
      const code = toLevel(it.basic, lv);
      if (!code || seen.has(code)) continue; seen.add(code);
      priceBasket.push({type, code, name: loaded[lv].sectors.find(s => s.code === code).name, pct: it.pct});
    }
  }
  $('wage-pct').value = p.wage || 0;
  renderPriceBasket();
  await runPrice(p.label + (lv !== '380' ? ' — ' + LV[lv].label + '로 묶인 부문 전체에 적용되어 과대추정될 수 있음' : ''));
  $('price').scrollIntoView({behavior: 'smooth', block: 'start'});
}

// ---------------------------------------------------------------- 히트맵
$('heat-kind').onchange = () => renderHeatmap();
async function renderHeatmap() {
  const lv = current, box = $('heatmap');
  if (Number(LV[lv].n) > 83) { box.replaceChildren(el('p', {class: 'note'}, '히트맵은 대분류·중분류를 선택했을 때 표시됩니다.')); return; }
  const d = await cached('heatmap', [lv, $('heat-kind').value]);
  if (lv !== current) return;
  const n = d.codes.length, cell = n <= 33 ? 18 : 8, L = n <= 33 ? 44 : 36, T = n <= 33 ? 44 : 36;
  const W = L + n * cell + 10, H = T + n * cell + 10;
  const s = svg('svg', {viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': '투입계수 히트맵', style: 'max-width:' + W * 1.4 + 'px'});
  const mx = d.max || 1;
  d.matrix.forEach((row, i) => row.forEach((v, j) => {
    if (v <= 0) return;
    const t = Math.sqrt(v / mx);
    const r = svg('rect', {x: L + j * cell, y: T + i * cell, width: cell - 0.5, height: cell - 0.5, fill: `rgba(20,92,112,${(0.06 + 0.94 * t).toFixed(3)})`});
    r.append(svg('title', {}, d.names[i] + ' → ' + d.names[j] + '\n투입계수 ' + v.toFixed(4)));
    s.append(r);
  }));
  const fs = n <= 33 ? 9 : 5;
  d.codes.forEach((c, k) => {
    s.append(svg('text', {x: L - 3, y: T + k * cell + cell * 0.7, 'text-anchor': 'end', 'font-size': fs, class: 'tick'}, c));
    s.append(svg('text', {x: L + k * cell + cell / 2, y: T - 4, 'text-anchor': 'start', 'font-size': fs, class: 'tick',
      transform: `rotate(-60 ${L + k * cell + cell / 2} ${T - 4})`}, c));
  });
  box.replaceChildren(s, el('p', {class: 'note'}, '행(세로): 투입하는(공급) 부문, 열(가로): 투입받는(구매) 부문. 색이 진할수록 투입계수가 큽니다(제곱근 척도, 최댓값 ' + mx.toFixed(3) + '). 코드 이름은 위 산업구조 표에서 확인할 수 있습니다.'));
}
