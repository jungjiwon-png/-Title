// 별도 작업 공간(Web Worker)에서 Pyodide(Python)를 실행한다.
const BASE = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
const ALLOWED = ['explore', 'structure', 'sector_detail', 'linkage', 'final_demand',
                 'simulate', 'compare', 'basic_list', 'price_effect', 'heatmap',
                 'trend_all', 'supply_risk', 'import_items', 'pair_compare'];
let py = null;
let chain = Promise.resolve();

async function ensurePython() {
  if (py) return;
  postMessage({type: 'status', text: 'Python 실행 환경(Pyodide) 준비 중…'});
  importScripts(BASE + 'pyodide.js');
  py = await loadPyodide({indexURL: BASE});
  postMessage({type: 'status', text: 'numpy·openpyxl 설치 중…'});
  await py.loadPackage(['numpy', 'micropip']);
  await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
  const response = await fetch('./model.py');
  if (!response.ok) throw Error('model.py를 불러올 수 없습니다.');
  await py.runPythonAsync(await response.text());
}

async function handle({id, type, payload}) {
  try {
    await ensurePython();
    let result;
    if (type === 'load_level') {
      postMessage({type: 'status', text: payload.label + ' Excel 읽기와 계수 검증 중…'});
      py.globals.set('level_input', payload.level);
      py.globals.set('raw_input', new Uint8Array(payload.buffer));
      result = py.runPython('load_level(level_input, bytes(raw_input.to_py()))');
      py.runPython('del raw_input');
    } else if (type === 'load_aux') {
      postMessage({type: 'status', text: '부문분류표·주요표 읽는 중…'});
      py.globals.set('map_input', new Uint8Array(payload.map));
      py.globals.set('main_input', new Uint8Array(payload.main));
      result = py.runPython('load_aux(bytes(map_input.to_py()), bytes(main_input.to_py()))');
      py.runPython('del map_input, main_input');
    } else if (ALLOWED.includes(type)) {
      py.globals.set('args_input', JSON.stringify(payload || []));
      result = py.runPython(type + '(*__import__("json").loads(args_input))');
    } else {
      throw Error('알 수 없는 요청: ' + type);
    }
    postMessage({id, ok: true, result: JSON.parse(result)});
  } catch (error) {
    postMessage({id, ok: false, error: String(error && error.message || error)});
  }
}
onmessage = e => { chain = chain.then(() => handle(e.data)); };
