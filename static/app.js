/* ============================================================
   Config Builder — app.js
   Lógica drag & drop + comunicación con Flask API
   ============================================================ */

let sheetsData    = [];   // hojas del informe real
let currentSheet  = null; // hoja activa
let mappings      = {};   // { key: [ {key, archivo, codigos, valor, celda} ] }
let manualValues  = {};   // { key: [{ op, valor }] }
let mappingSheets = {};   // { key: 'GASTOS OPERATIVOS' }  — hoja destino de cada item
let filesData     = {};   // { nombre_archivo: { tipo, items } }
let configRaw     = null; // config_gastos.json completo
let folders       = [];

// ── Utilidades ──────────────────────────────────────────────
const fmt = v => v == null ? '' : '$' + Number(v).toLocaleString('es-CO', {maximumFractionDigits: 0});
const norm = s => s.replace(/\s+/g, ' ').trim();

function fileKey(filename) {
  // Genera una clave corta a partir del nombre de archivo
  const f = filename.toUpperCase();
  if (f.includes('ESTADO DE RESULTADOS')) {
    const m = f.match(/ESTADO DE RESULTADOS_(.+?)_[A-Z]{3}_/);
    return m ? 'ER_' + m[1].replace(/\s+/g,'_') : 'ER';
  }
  if (f.startsWith('51355001')) {
    const m = f.match(/51355001_(.+?)_[A-Z]{3}_/);
    return m ? 'AUX_' + m[1].replace(/\s+/g,'_') : 'AUX';
  }
  if (f.startsWith('5105')) {
    const m = f.match(/5105_(.+?)_[A-Z]{3}_/);
    return m ? '5105_' + m[1].replace(/\s+/g,'_') : '5105';
  }
  return filename.split('.')[0].replace(/\s+/g,'_');
}

// ── Inicialización ───────────────────────────────────────────
async function init() {
  showLog('Cargando hojas del informe...', true);

  // Cargar hojas y config en paralelo
  const [sheetsRes, configRes, foldersRes] = await Promise.all([
    fetch('/api/sheets').then(r => r.json()),
    fetch('/api/config').then(r => r.json()),
    fetch('/api/folders').then(r => r.json())
  ]);

  sheetsData = sheetsRes;
  configRaw  = configRes;
  folders    = foldersRes;

  // Poblar mappings desde config actual
  initMappingsFromConfig(configRaw);

  // Poblar selector de carpetas
  populateFolderSelects();

  // Renderizar pestañas
  renderTabs();

  // Cargar primera hoja
  if (sheetsData.length > 0) selectSheet(0);

  hideLog();
  updateStats();
}

function initMappingsFromConfig(config) {
  mappings      = {};
  manualValues  = {};
  mappingSheets = {};
  if (!config || !config.items) return;
  for (const item of config.items) {
    const codeKey = item.buscar_por === 'B'
      ? `B:${item.codigo_a.replace(/^B:/, '')}`
      : norm(item.codigo_a);
    // Si el item tiene fila_dest (nuevo formato) úsala directamente
    // Si no, busca el número de fila en sheetsData para migrar al nuevo formato
    let filaNum = item.fila_dest || null;
    if (!filaNum) {
      for (const sheet of sheetsData) {
        const match = sheet.filas.find(f => {
          const fk = f.tiene_codigo ? norm(f.codigo) : `B:${f.descripcion}`;
          return fk === codeKey && (!item.hoja || sheet.nombre.trim() === item.hoja.trim());
        });
        if (match) { filaNum = match.fila; break; }
      }
    }
    const key = filaNum ? `${codeKey}|${filaNum}` : codeKey;
    mappings[key] = (item.sources || []).map(src => ({
      key       : src.key,
      archivo   : src.key,
      codigos   : src.codes || [],
      sin_filtro: src.sin_filtro || false
    }));
    if (item.hoja) mappingSheets[key] = item.hoja;
    if (item.valor_fijo != null) {
      const vf = item.valor_fijo;
      if (Array.isArray(vf))           manualValues[key] = vf;
      else if (typeof vf === 'object') manualValues[key] = [vf];
      else                             manualValues[key] = [{ op: '+', valor: vf }];
    }
  }
}

function populateFolderSelects() {
  // Selector panel derecho: muestra nombre + cantidad de archivos
  const selPanel = document.querySelector('.folder-select');
  if (selPanel) {
    selPanel.innerHTML = '<option value="">-- Selecciona carpeta --</option>';
    folders.forEach(f => {
      const opt = document.createElement('option');
      opt.value       = f.nombre;
      opt.textContent = `${f.nombre}  (${f.n_archivos} archivos)`;
      selPanel.appendChild(opt);
    });
  }

  // Selector "Ejecutar mes": solo nombre del mes
  const selRun = document.getElementById('run-folder');
  if (selRun) {
    selRun.innerHTML = '<option value="">-- Selecciona --</option>';
    folders.forEach(f => {
      const opt = document.createElement('option');
      opt.value       = f.nombre;
      opt.textContent = f.nombre;
      selRun.appendChild(opt);
    });
  }
}

// ── Pestañas ─────────────────────────────────────────────────
function renderTabs() {
  const nav = document.getElementById('tabs-nav');
  nav.innerHTML = '';
  sheetsData.forEach((sheet, i) => {
    const configured = sheet.filas.filter(f => {
      const ck = f.tiene_codigo ? norm(f.codigo) : `B:${f.descripcion}`;
      return (mappings[`${ck}|${f.fila}`] || mappings[ck] || []).length > 0;
    }).length;
    const btn = document.createElement('button');
    btn.className   = 'tab-btn' + (i === 0 ? ' active' : '');
    btn.textContent = sheet.nombre;
    if (configured > 0) {
      const badge = document.createElement('span');
      badge.className   = 'badge';
      badge.textContent = configured;
      btn.appendChild(badge);
    }
    btn.onclick = () => selectSheet(i);
    nav.appendChild(btn);
  });
}

function selectSheet(index) {
  currentSheet = sheetsData[index];
  document.querySelectorAll('.tab-btn').forEach((b, i) => {
    b.classList.toggle('active', i === index);
  });
  renderDestRows();
}

// ── Filas destino ─────────────────────────────────────────────
function renderDestRows() {
  const container = document.getElementById('sheet-content');
  container.innerHTML = '';
  if (!currentSheet) return;

  currentSheet.filas.forEach(fila => {
    // Clave única: código (o B:desc) + número de fila para evitar colisiones con códigos duplicados
    const codeKey = fila.tiene_codigo ? norm(fila.codigo) : `B:${fila.descripcion}`;
    const key = `${codeKey}|${fila.fila}`;

    const sources = mappings[key] || [];
    const hasData = sources.length > 0;

    const div = document.createElement('div');
    div.className   = `dest-row ${hasData ? 'has-sources' : 'no-sources'}`;
    div.dataset.key = key;

    const safeKey     = key.replace(/[^a-z0-9]/gi,'_');
    const codigoLabel = fila.tiene_codigo ? fila.codigo : `— ${fila.descripcion}`;
    const descLabel   = fila.tiene_codigo ? fila.descripcion : '';

    div.innerHTML = `
      <div class="dest-row-header">
        <span class="dest-codigo ${fila.tiene_codigo ? '' : 'sin-codigo'}">${codigoLabel}</span>
        <span class="dest-desc">${descLabel}</span>
        <span class="dest-status ${hasData ? 'status-green' : 'status-gray'}"
              data-tip="${hasData ? sources.length + ' fuente(s)' : 'Sin configurar'}"></span>
      </div>
      <div class="dest-sources" id="sources-${safeKey}">
        ${hasData ? '' : '<span class="drop-hint">Arrastra aquí los códigos fuente</span>'}
      </div>
      <div class="dest-manual" id="manual-${safeKey}"></div>
    `;

    // Drag over
    const dropZone = div.querySelector('.dest-sources');
    dropZone.addEventListener('dragover', e => {
      e.preventDefault();
      div.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', () => div.classList.remove('drag-over'));
    dropZone.addEventListener('drop', e => {
      e.preventDefault();
      div.classList.remove('drag-over');
      const data = JSON.parse(e.dataTransfer.getData('application/json'));
      addSourceToRow(key, data);
    });

    // Renderizar operaciones fijas
    renderManualOps(key, div.querySelector(`#manual-${safeKey}`));

    container.appendChild(div);

    // Renderizar chips existentes
    if (hasData) renderChips(key, sources, dropZone);
  });
}

function renderManualOps(rowKey, container) {
  if (!container) return;
  const ops = manualValues[rowKey] || [];
  container.innerHTML = '';

  ops.forEach((op, idx) => {
    const row = document.createElement('div');
    row.className = 'manual-op-row';
    row.innerHTML = `
      <label>Op ${idx + 1}:</label>
      <select class="manual-op-sel">
        <option value="+" ${op.op==='+' ? 'selected':''}>+  sumar</option>
        <option value="-" ${op.op==='-' ? 'selected':''}>−  restar</option>
        <option value="*" ${op.op==='*' ? 'selected':''}>×  multiplicar</option>
        <option value="/" ${op.op==='/' ? 'selected':''}>÷  dividir</option>
      </select>
      <input type="number" class="manual-input" value="${op.valor}" placeholder="0">
      <button class="btn-op-remove" title="Eliminar">×</button>
    `;
    const sel    = row.querySelector('.manual-op-sel');
    const input  = row.querySelector('.manual-input');
    const btnDel = row.querySelector('.btn-op-remove');

    // Cambiar operador: solo actualiza el op, nunca elimina
    sel.addEventListener('change', () => {
      if (!manualValues[rowKey]) manualValues[rowKey] = [];
      if (!manualValues[rowKey][idx]) manualValues[rowKey][idx] = { op: '+', valor: 0 };
      manualValues[rowKey][idx].op = sel.value;
      // Re-renderizar chips para actualizar el total sin tocar el DOM del manual
      const safeKey  = rowKey.replace(/[^a-z0-9]/gi, '_');
      const dropZone = document.getElementById(`sources-${safeKey}`);
      const sources  = mappings[rowKey] || [];
      if (dropZone) {
        if (sources.length > 0) renderChips(rowKey, sources, dropZone);
      }
      updateStats();
    });

    // Cambiar valor: guarda o elimina si queda vacío
    input.addEventListener('change', () => {
      const v = parseFloat(input.value);
      if (!isNaN(v) && v !== 0) {
        manualValues[rowKey][idx].valor = v;
      } else {
        manualValues[rowKey].splice(idx, 1);
        if (manualValues[rowKey].length === 0) delete manualValues[rowKey];
        refreshRow(rowKey);
      }
      updateStats();
    });
    btnDel.addEventListener('click', () => {
      manualValues[rowKey].splice(idx, 1);
      if (manualValues[rowKey].length === 0) delete manualValues[rowKey];
      refreshRow(rowKey);
      updateStats();
    });

    container.appendChild(row);
  });

  // Botón agregar operación
  const btnAdd = document.createElement('button');
  btnAdd.className   = 'btn-op-add';
  btnAdd.textContent = '+ Agregar operación fija';
  btnAdd.addEventListener('click', () => {
    if (!manualValues[rowKey]) manualValues[rowKey] = [];
    manualValues[rowKey].push({ op: '+', valor: 0 });
    if (currentSheet) mappingSheets[rowKey] = currentSheet.nombre;
    refreshRow(rowKey);
  });
  container.appendChild(btnAdd);
}

function renderChips(rowKey, sources, container) {
  container.innerHTML = '';
  let total = 0;
  sources.forEach((src, idx) => {
    const chip = document.createElement('span');
    chip.className = 'source-chip';
    const codigos = Array.isArray(src.codigos) ? src.codigos.join(', ') : (src.codigo || '');
    const val = src.valor != null ? Number(src.valor) : 0;
    total += val;
    // Usar createElement en lugar de innerHTML para evitar que caracteres especiales
    // en rowKey (apóstrofes, guiones, etc.) rompan atributos onclick inline
    const codeSpan = document.createElement('span');
    codeSpan.className = 'chip-code';
    codeSpan.textContent = codigos;
    const fileSpan = document.createElement('span');
    fileSpan.className = 'chip-file';
    fileSpan.textContent = src.key || src.archivo;
    const removeSpan = document.createElement('span');
    removeSpan.className = 'chip-remove';
    removeSpan.textContent = '×';
    removeSpan.addEventListener('click', () => removeSource(rowKey, idx));
    chip.appendChild(codeSpan);
    chip.appendChild(fileSpan);
    if (val) {
      const valSpan = document.createElement('span');
      valSpan.className = 'chip-val';
      valSpan.textContent = fmt(val);
      chip.appendChild(valSpan);
    }
    chip.appendChild(removeSpan);
    container.appendChild(chip);
  });

  // Aplicar operaciones fijas encadenadas
  const ops = manualValues[rowKey] || [];
  const opLabel = { '+': '+', '-': '−', '*': '×', '/': '÷' };
  ops.forEach(mv => {
    if (!mv.valor) return;
    const chipFijo = document.createElement('span');
    chipFijo.className = 'source-chip chip-fijo';
    chipFijo.innerHTML = `<span class="chip-code">${opLabel[mv.op]} Fijo</span><span class="chip-val">${fmt(mv.valor)}</span>`;
    container.appendChild(chipFijo);
    if      (mv.op === '+') total = total + mv.valor;
    else if (mv.op === '-') total = total - mv.valor;
    else if (mv.op === '*') total = total * mv.valor;
    else if (mv.op === '/') total = mv.valor !== 0 ? total / mv.valor : total;
  });

  if (total !== 0) {
    const tot = document.createElement('span');
    tot.className = 'chip-total';
    tot.textContent = `Σ ${fmt(total)}`;
    container.appendChild(tot);
  }
}

function addSourceToRow(rowKey, chipData) {
  if (!mappings[rowKey]) mappings[rowKey] = [];
  // Registrar hoja destino
  if (currentSheet) mappingSheets[rowKey] = currentSheet.nombre;

  // Evitar duplicado exacto
  const existe = mappings[rowKey].some(s =>
    (s.codigos || []).join(',') === (chipData.codigos || [chipData.codigo]).join(',') &&
    s.key === chipData.key
  );
  if (existe) return;

  mappings[rowKey].push({
    key    : chipData.key,
    archivo: chipData.archivo,
    codigos: chipData.codigos || [chipData.codigo],
    valor  : chipData.valor,
    celda  : chipData.celda
  });

  // Re-renderizar la fila
  refreshRow(rowKey);
  updateStats();
}

function removeSource(rowKey, idx) {
  if (!mappings[rowKey]) return;
  mappings[rowKey].splice(idx, 1);
  if (mappings[rowKey].length === 0) delete mappings[rowKey];
  refreshRow(rowKey);
  updateStats();
}

function refreshRow(rowKey) {
  const safeKey  = rowKey.replace(/[^a-z0-9]/gi, '_');
  const dropZone = document.getElementById(`sources-${safeKey}`);
  const manualDiv = document.getElementById(`manual-${safeKey}`);
  const destRow  = dropZone?.closest('.dest-row');
  if (!dropZone || !destRow) return;

  const sources = mappings[rowKey] || [];
  const hasData = sources.length > 0;
  destRow.className = `dest-row ${hasData ? 'has-sources' : 'no-sources'}`;

  const statusDot = destRow.querySelector('.dest-status');
  if (statusDot) {
    statusDot.className  = `dest-status ${hasData ? 'status-green' : 'status-gray'}`;
    statusDot.dataset.tip = hasData ? sources.length + ' fuente(s)' : 'Sin configurar';
  }

  if (hasData) {
    renderChips(rowKey, sources, dropZone);
  } else {
    dropZone.innerHTML = '<span class="drop-hint">Arrastra aquí los códigos fuente</span>';
  }

  if (manualDiv) renderManualOps(rowKey, manualDiv);
}

// ── Panel derecho: acordeón de archivos ──────────────────────
async function loadFolder(folderName) {
  if (!folderName) return;
  const accordion = document.getElementById('accordion-wrap');
  accordion.innerHTML = '<p style="padding:20px;color:#888;font-size:.85rem;">Cargando archivos...</p>';

  const data = await fetch(`/api/files/${encodeURIComponent(folderName)}`).then(r => r.json());
  filesData  = data;
  accordion.innerHTML = '';

  Object.entries(data).forEach(([nombre, info]) => {
    const item = buildAccordionItem(nombre, info);
    accordion.appendChild(item);
  });

  if (Object.keys(data).length === 0) {
    accordion.innerHTML = '<p style="padding:20px;color:#888;font-size:.85rem;">No se encontraron archivos procesables.</p>';
  }
}

function buildAccordionItem(nombre, info) {
  const wrapper = document.createElement('div');
  wrapper.className = 'accordion-item';

  const header = document.createElement('div');
  header.className = 'accordion-header';
  header.innerHTML = `
    <span>${nombre}</span>
    <span style="font-size:.75rem;color:#666;">${info.items.length} códigos &nbsp;<span class="arrow">▶</span></span>
  `;

  const body = document.createElement('div');
  body.className = 'accordion-body';

  // Buscador
  const search = document.createElement('input');
  search.type        = 'text';
  search.placeholder = 'Buscar código o descripción...';
  search.className   = 'chip-search';

  const chipsWrap = document.createElement('div');
  renderDraggableChips(chipsWrap, nombre, info);

  search.addEventListener('input', () => {
    const q = search.value.toLowerCase();
    chipsWrap.querySelectorAll('.draggable-chip').forEach(chip => {
      const txt = chip.textContent.toLowerCase();
      chip.style.display = txt.includes(q) ? '' : 'none';
    });
  });

  body.appendChild(search);
  body.appendChild(chipsWrap);

  header.addEventListener('click', () => {
    header.classList.toggle('open');
    body.classList.toggle('open');
  });

  wrapper.appendChild(header);
  wrapper.appendChild(body);
  return wrapper;
}

function renderDraggableChips(container, nombreArchivo, info) {
  container.innerHTML = '';
  const key = fileKey(nombreArchivo);

  info.items.forEach(item => {
    if (item.error) return;
    const chip = document.createElement('div');
    chip.className = 'draggable-chip';
    chip.draggable = true;
    chip.innerHTML = `
      <span class="dc-code">${item.codigo}</span>
      <span class="dc-desc">${item.descripcion || ''}</span>
      <span class="dc-val">${fmt(item.valor)}</span>
    `;

    const dragData = {
      key     : key,
      archivo : nombreArchivo,
      codigos : [item.codigo],
      codigo  : item.codigo,
      descripcion: item.descripcion,
      valor   : item.valor,
      celda   : item.celda
    };

    chip.addEventListener('dragstart', e => {
      e.dataTransfer.setData('application/json', JSON.stringify(dragData));
      e.dataTransfer.effectAllowed = 'copy';
    });

    container.appendChild(chip);
  });
}

// ── Guardar config ───────────────────────────────────────────
async function saveConfig() {
  const btn = document.getElementById('btn-save');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Guardando...';

  // Construir nuevo config
  const newConfig = buildConfig();

  const res = await fetch('/api/config/save', {
    method : 'POST',
    headers: { 'Content-Type': 'application/json' },
    body   : JSON.stringify(newConfig)
  }).then(r => r.json());

  btn.disabled = false;
  btn.textContent = '💾 Guardar config';

  if (res.ok) {
    showLog(`✅ Config guardado. Backup en: ${res.backup}`, true);
    setTimeout(hideLog, 4000);
  } else {
    showLog(`❌ Error: ${res.error}`, true);
  }
}

function buildConfig() {
  // Reconstruir file_sources desde los mappings actuales
  const fileSources = {};
  const items = [];

  // Mantener file_sources existentes del config original
  if (configRaw && configRaw.file_sources) {
    Object.assign(fileSources, configRaw.file_sources);
  }

  // Agregar nuevas fuentes de los mappings
  Object.entries(mappings).forEach(([rowKey, sources]) => {
    // rowKey puede ser "CODIGO|filaNum" o "B:desc|filaNum"
    const pipIdx  = rowKey.lastIndexOf('|');
    const codigoA = pipIdx >= 0 ? rowKey.slice(0, pipIdx) : rowKey;
    const filaNum = pipIdx >= 0 ? parseInt(rowKey.slice(pipIdx + 1)) : null;

    const configSources = sources.map(src => {
      if (!fileSources[src.key]) {
        fileSources[src.key] = { prefijo: inferPrefijo(src.archivo || src.key), entidad: inferEntidad(src.archivo || src.key) };
      }
      const entry = { key: src.key, codes: src.codigos || [src.codigo] };
      if (src.sin_filtro) entry.sin_filtro = true;
      return entry;
    });

    const entry = { codigo_a: codigoA, sources: configSources };
    if (codigoA.startsWith('B:'))  entry.buscar_por = 'B';
    if (filaNum)                   entry.fila_dest  = filaNum;
    if (mappingSheets[rowKey])     entry.hoja       = mappingSheets[rowKey];
    if (manualValues[rowKey])      entry.valor_fijo = manualValues[rowKey];
    items.push(entry);
  });

  return {
    _comentario: configRaw?._comentario || 'Generado por Config Builder',
    file_sources: fileSources,
    items
  };
}

function inferPrefijo(archivo) {
  const f = (archivo || '').toUpperCase();
  if (f.startsWith('ESTADO DE RESULTADOS')) return 'ESTADO DE RESULTADOS';
  if (f.startsWith('51355001'))              return '51355001';
  if (f.startsWith('5105'))                  return '5105';
  return archivo.split('_')[0];
}

function inferEntidad(archivo) {
  const m = (archivo || '').match(/^[^_]+_(.+?)_[A-Z]{3}_/i);
  return m ? m[1] : '';
}

// ── Ejecutar mes ─────────────────────────────────────────────
async function runMes() {
  const folder = document.getElementById('run-folder').value;
  if (!folder) { alert('Selecciona una carpeta de mes primero.'); return; }

  const btn = document.getElementById('btn-run');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Ejecutando...';
  showLog(`▶ Ejecutando procesar_todo.py ${folder} ...\n`, true);

  const res = await fetch(`/api/run/${encodeURIComponent(folder)}`, { method: 'POST' })
    .then(r => r.json());

  btn.disabled = false;
  btn.textContent = '▶ Ejecutar mes';

  const log = document.getElementById('log-panel');
  log.textContent += (res.stdout || '') + (res.stderr || '');
  if (res.error) log.textContent += '\n❌ ' + res.error;
  log.scrollTop = log.scrollHeight;
}

// ── Log ──────────────────────────────────────────────────────
function showLog(msg, clear = false) {
  const log = document.getElementById('log-panel');
  log.classList.add('visible');
  if (clear) log.textContent = msg + '\n';
  else       log.textContent += msg + '\n';
  log.scrollTop = log.scrollHeight;
}
function hideLog() {
  // No ocultar automáticamente para que el usuario lea
}

// ── Stats ────────────────────────────────────────────────────
function updateStats() {
  const total       = Object.values(sheetsData).reduce((a, s) => a + s.filas.length, 0);
  const configurados = Object.keys(mappings).length;
  const el = document.getElementById('stats');
  if (el) el.textContent = `${configurados} / ${total} ítems configurados`;
}

// ── Evento cambio de carpeta ─────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  init();

  document.getElementById('folder-select').addEventListener('change', e => {
    document.getElementById('run-folder').value = e.target.value;
    loadFolder(e.target.value);
  });
});
