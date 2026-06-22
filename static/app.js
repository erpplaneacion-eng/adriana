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
let selectedChips   = new Map(); // chipId → dragData (selección múltiple)
let isDirty         = false; // cambios pendientes por guardar
// ponytail: previewOverrides eliminado — Σ = chips + valor_fijo siempre

// ── Utilidades ──────────────────────────────────────────────
const fmt = v => v == null ? '' : '$' + Number(v).toLocaleString('es-CO', {maximumFractionDigits: 0});
const norm = s => s.replace(/\s+/g, ' ').trim();
const markDirty = () => { isDirty = true; };
const clearDirty = () => { isDirty = false; };

function makeRowKey(sheetName, codeKey, filaNum) {
  if (sheetName && filaNum) return `${sheetName}::${codeKey}|${filaNum}`;
  if (filaNum) return `${codeKey}|${filaNum}`;
  return codeKey;
}

function parseRowKey(rowKey) {
  const hasSheet = rowKey.includes('::');
  const parts = hasSheet ? rowKey.split('::') : [null, rowKey];
  const sheetName = hasSheet ? parts[0] : null;
  const raw = hasSheet ? parts.slice(1).join('::') : parts[1];
  const pipIdx = raw.lastIndexOf('|');
  const codigoA = pipIdx >= 0 ? raw.slice(0, pipIdx) : raw;
  const filaNum = pipIdx >= 0 ? parseInt(raw.slice(pipIdx + 1), 10) : null;
  return { sheetName, codigoA, filaNum };
}


// MES_RE: 3-letter abbreviation OR full Spanish month name
const _MES_RE = '[A-Z]{3,13}';
function fileKey(filename) {
  const f = filename.toUpperCase();
  const re = (pfx) => new RegExp(`${pfx}_(.+?)_${_MES_RE}_`);
  if (f.includes('ESTADO DE RESULTADOS')) {
    const m = f.match(/ESTADO DE RESULTADOS_(.+?)_[A-Z]{3,13}_/);
    return m ? 'ER_' + m[1].replace(/\s+/g,'_') : 'ER';
  }
  if (f.startsWith('51355001')) {
    const m = f.match(re('51355001'));
    return m ? 'AUX_' + m[1].replace(/\s+/g,'_') : 'AUX';
  }
  if (f.startsWith('5105')) {
    const m = f.match(re('5105'));
    return m ? '5105_' + m[1].replace(/\s+/g,'_') : '5105';
  }
  if (f.startsWith('7205')) {
    const m = f.match(re('7205'));
    return m ? '7205_' + m[1].replace(/\s+/g,'_') : '7205';
  }
  if (f.startsWith('7105')) {
    const m = f.match(re('7105'));
    return m ? '7105_' + m[1].replace(/\s+/g,'_') : '7105';
  }
  return filename.split('.')[0].replace(/\s+/g,'_');
}

// ── Selección múltiple de chips ──────────────────────────────
function updateSelectionBar() {
  const bar = document.getElementById('selection-bar');
  if (!bar) return;
  const n = selectedChips.size;
  bar.style.display = n > 0 ? 'flex' : 'none';
  bar.querySelector('.sel-count').textContent =
    `${n} chip${n !== 1 ? 's' : ''} seleccionado${n !== 1 ? 's' : ''} — arrastra cualquiera al destino`;
}

function clearChipSelection() {
  selectedChips.clear();
  document.querySelectorAll('.source-row.chip-selected, .draggable-chip.chip-selected')
    .forEach(c => c.classList.remove('chip-selected'));
  updateSelectionBar();
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

  clearDirty();
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
    let sheetName = item.hoja || null;
    if (!filaNum) {
      for (const sheet of sheetsData) {
        const match = sheet.filas.find(f => {
          const fk = f.tiene_codigo ? norm(f.codigo) : `B:${f.descripcion}`;
          return fk === codeKey && (!item.hoja || sheet.nombre.trim() === item.hoja.trim());
        });
        if (match) {
          filaNum = match.fila;
          sheetName = sheet.nombre;
          break;
        }
      }
    }
    const key = makeRowKey(sheetName, codeKey, filaNum);
    mappings[key] = (item.sources || []).map(src => ({
      key       : src.key,
      archivo   : src.key,
      codigos   : src.codes || [],
      sin_filtro: src.sin_filtro || false,
      seccion   : src.seccion   || null,
      op        : src.op        || '+'
    }));
    if (sheetName) mappingSheets[key] = sheetName;
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
      if (f.es_titulo) return false;
      const ck = f.tiene_codigo ? norm(f.codigo) : `B:${f.descripcion}`;
      const key = makeRowKey(sheet.nombre, ck, f.fila);
      return (mappings[key] || mappings[`${ck}|${f.fila}`] || mappings[ck] || []).length > 0;
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
    // Título de sección: separador visual, no editable
    if (fila.es_titulo) {
      const sep = document.createElement('div');
      sep.className = 'section-title';
      sep.textContent = fila.descripcion;
      container.appendChild(sep);
      return;
    }

    // Fila con referencia interna (cross-sheet del mismo libro): solo lectura
    if (fila.es_interno) {
      const div = document.createElement('div');
      div.className = 'dest-row interno-row';
      const codigoLabel = fila.tiene_codigo ? fila.codigo : `— ${fila.descripcion}`;
      const descLabel   = fila.tiene_codigo ? fila.descripcion : '';
      div.innerHTML = `
        <div class="dest-row-header">
          <span class="dest-codigo ${fila.tiene_codigo ? '' : 'sin-codigo'}">${codigoLabel}</span>
          <span class="dest-desc">${descLabel}</span>
          <span class="dest-status status-interno" data-tip="Referencia interna al libro"></span>
        </div>
        <div class="interno-badge">Calculado desde otra pestaña del INFORME</div>
      `;
      container.appendChild(div);
      return;
    }

    // Clave única: hoja + código (o B:desc) + número de fila para evitar colisiones entre hojas
    const codeKey = fila.tiene_codigo ? norm(fila.codigo) : `B:${fila.descripcion}`;
    const key = makeRowKey(currentSheet.nombre, codeKey, fila.fila);

    const sources = mappings[key] || mappings[`${codeKey}|${fila.fila}`] || [];
    const hasData = sources.length > 0 || !!(manualValues[key] && manualValues[key].length > 0);

    const div = document.createElement('div');
    div.className   = `dest-row ${hasData ? 'has-sources' : 'no-sources'}`;
    div.dataset.key = key;

    const safeKey     = key.replace(/[^a-z0-9]/gi,'_');
    const codigoLabel = fila.tiene_codigo ? fila.codigo : `— ${fila.descripcion}`;
    const descLabel   = fila.tiene_codigo ? fila.descripcion : '';
    const statusClass = hasData ? 'status-green' : 'status-gray';
    const statusTip   = hasData ? sources.length + ' fuente(s)' : 'Sin configurar';

    div.innerHTML = `
      <div class="dest-row-header">
        <span class="dest-codigo ${fila.tiene_codigo ? '' : 'sin-codigo'}">${codigoLabel}</span>
        <span class="dest-desc">${descLabel}</span>
        <span class="dest-status ${statusClass}"
              data-tip="${statusTip}"></span>
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
      const raw = JSON.parse(e.dataTransfer.getData('application/json'));
      if (Array.isArray(raw)) {
        raw.forEach(data => addSourceToRow(key, data));
        clearChipSelection();
      } else {
        addSourceToRow(key, raw);
      }
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
      markDirty();
      const safeKey  = rowKey.replace(/[^a-z0-9]/gi, '_');
      const dropZone = document.getElementById(`sources-${safeKey}`);
      const sources  = mappings[rowKey] || [];
      if (dropZone) renderChips(rowKey, sources, dropZone);
      updateStats();
    });

    // Cambiar valor: guarda o elimina si queda vacío
    input.addEventListener('change', () => {
      const v = parseFloat(input.value);
      if (!isNaN(v) && v !== 0) {
        manualValues[rowKey][idx].valor = v;
        markDirty();
        const safeKey  = rowKey.replace(/[^a-z0-9]/gi, '_');
        const dropZone = document.getElementById(`sources-${safeKey}`);
        const sources  = mappings[rowKey] || [];
        if (dropZone) renderChips(rowKey, sources, dropZone);
      } else {
        manualValues[rowKey].splice(idx, 1);
        if (manualValues[rowKey].length === 0) delete manualValues[rowKey];
        markDirty();
        refreshRow(rowKey);
      }
      updateStats();
    });
    btnDel.addEventListener('click', () => {
      manualValues[rowKey].splice(idx, 1);
      if (manualValues[rowKey].length === 0) delete manualValues[rowKey];
      markDirty();
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
    markDirty();
    refreshRow(rowKey);
  });
  container.appendChild(btnAdd);
}

function renderChips(rowKey, sources, container) {
  container.innerHTML = '';
  let total = 0;
  sources.forEach((src, idx) => {
    const op  = src.op || '+';
    const val = src.valor != null ? Number(src.valor) : 0;
    if      (op === '+') total += val;
    else if (op === '-') total -= val;
    else if (op === '*') total *= val;
    else if (op === '/') total = val !== 0 ? total / val : total;

    // Detectar duplicado: mismo archivo resuelto + mismos códigos + misma sección que un chip anterior
    const isDup = src._resolvedKey != null && sources.some((other, oi) =>
      oi < idx
      && other._resolvedKey === src._resolvedKey
      && (other.codigos || []).join(',') === (src.codigos || []).join(',')
      && (other.seccion || null) === (src.seccion || null)
    );

    const chip = document.createElement('span');
    chip.className = 'source-chip' + (isDup ? ' chip-duplicate' : '');
    if (isDup) chip.title = '⚠ Duplicado: mismo archivo y código que otro chip en esta fila';
    const codigos = Array.isArray(src.codigos) ? src.codigos.join(', ') : (src.codigo || '');

    // Botón de operador (+/−) — solo visible si hay más de un chip o si ya es −
    const opBtn = document.createElement('button');
    opBtn.className = `chip-op-btn ${op === '-' ? 'op-minus' : 'op-plus'}`;
    opBtn.textContent = op === '-' ? '−' : '+';
    opBtn.title = 'Clic para cambiar a ' + (op === '-' ? 'suma' : 'resta');
    opBtn.addEventListener('click', e => {
      e.stopPropagation();
      mappings[rowKey][idx].op = op === '-' ? '+' : '-';
      markDirty();
      refreshRow(rowKey);
      updateStats();
    });

    const codeSpan = document.createElement('span');
    codeSpan.className = 'chip-code';
    codeSpan.textContent = codigos;
    const fileSpan = document.createElement('span');
    fileSpan.className = 'chip-file';
    fileSpan.textContent = src.key || src.archivo;
    if (src.seccion) {
      const secSpan = document.createElement('span');
      secSpan.className = 'chip-sec';
      secSpan.textContent = src.seccion;
      fileSpan.appendChild(secSpan);
    }
    const removeSpan = document.createElement('span');
    removeSpan.className = 'chip-remove';
    removeSpan.textContent = '×';
    removeSpan.addEventListener('click', () => removeSource(rowKey, idx));

    chip.appendChild(opBtn);
    chip.appendChild(codeSpan);
    chip.appendChild(fileSpan);
    if (val) {
      const valSpan = document.createElement('span');
      valSpan.className = 'chip-val';
      valSpan.textContent = fmt(val);
      chip.appendChild(valSpan);
    }
    if (src.op_jk) {
      const opSpan = document.createElement('span');
      opSpan.className = 'chip-creditos';
      const c = Number(src.credito || 0);
      opSpan.textContent = c
        ? `J-K (${fmt(src.debito || 0)} - ${fmt(c)})`
        : 'J-K';
      chip.appendChild(opSpan);
    }
    chip.appendChild(removeSpan);
    container.appendChild(chip);
  });

  // Aplicar operaciones fijas encadenadas (rastrear por separado para el override)
  const ops = manualValues[rowKey] || [];
  const opLabel = { '+': '+', '-': '−', '*': '×', '/': '÷' };
  let manualTotal = 0;
  ops.forEach(mv => {
    if (!mv.valor) return;
    const chipFijo = document.createElement('span');
    chipFijo.className = 'source-chip chip-fijo';
    chipFijo.innerHTML = `<span class="chip-code">${opLabel[mv.op]} Fijo</span><span class="chip-val">${fmt(mv.valor)}</span>`;
    container.appendChild(chipFijo);
    if      (mv.op === '+') manualTotal += mv.valor;
    else if (mv.op === '-') manualTotal -= mv.valor;
    else if (mv.op === '*') manualTotal *= mv.valor;
    else if (mv.op === '/') manualTotal = mv.valor !== 0 ? manualTotal / mv.valor : manualTotal;
  });
  total += manualTotal;

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

  // Evitar duplicado exacto (mismo archivo + mismo codigo + misma sección)
  const existe = mappings[rowKey].some(s =>
    (s.codigos || []).join(',') === (chipData.codigos || [chipData.codigo]).join(',') &&
    s.key === chipData.key &&
    (s.seccion || null) === (chipData.seccion || null)
  );
  if (existe) return;

  // Corregir clave: si el archivo es un filename completo, reconstruir desde inferEntidad+inferPrefijo
  const _ref      = chipData.archivo || chipData.key;
  const _entidad  = inferEntidad(_ref);
  const _prefijo  = inferPrefijo(_ref);
  const _savedKey = _entidad ? `${_prefijo}_${_entidad.replace(/\s+/g,'_')}` : chipData.key;

  mappings[rowKey].push({
    key     : _savedKey,
    archivo : chipData.archivo,
    codigos : chipData.codigos || [chipData.codigo],
    valor   : chipData.valor,
    op_jk   : !!chipData.op_jk,
    debito  : chipData.debito,
    credito : chipData.credito,

    celda   : chipData.celda,
    seccion : chipData.seccion || null,
    op      : '+'
  });

  markDirty();
  refreshRow(rowKey);
  updateStats();
}

function removeSource(rowKey, idx) {
  if (!mappings[rowKey]) return;
  mappings[rowKey].splice(idx, 1);
  if (mappings[rowKey].length === 0) delete mappings[rowKey];
  markDirty();
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
  const hasData = sources.length > 0 || !!(manualValues[rowKey] && manualValues[rowKey].length > 0);
  destRow.className = `dest-row ${hasData ? 'has-sources' : 'no-sources'}`;

  const statusDot = destRow.querySelector('.dest-status');
  if (statusDot) {
    statusDot.className   = `dest-status ${hasData ? 'status-green' : 'status-gray'}`;
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
  clearChipSelection();
  const accordion = document.getElementById('accordion-wrap');
  accordion.innerHTML = '<p style="padding:20px;color:#888;font-size:.85rem;">Cargando archivos...</p>';

  const data = await fetch(`/api/files/${encodeURIComponent(folderName)}`).then(r => r.json());

  if (document.getElementById('folder-select').value !== folderName) return;

  // Recargar config desde DB para restaurar mappings guardados antes de actualizar valores
  const configRes = await fetch('/api/config').then(r => r.json());
  configRaw = configRes;
  initMappingsFromConfig(configRes);
  clearDirty();

  filesData = data;
  actualizarValoresEnMappings(data);

  accordion.innerHTML = '';
  Object.entries(data).forEach(([nombre, info]) => {
    const item = buildAccordionItem(nombre, info);
    accordion.appendChild(item);
  });

  if (Object.keys(data).length === 0) {
    accordion.innerHTML = '<p style="padding:20px;color:#888;font-size:.85rem;">No se encontraron archivos procesables.</p>';
  }
}

function actualizarValoresEnMappings(data) {
  // lookup: { fileKey: { codigo: { valor, debito, credito, op_jk } } }
  // lookupSec: { fileKey: { "seccion::codigo": { ... } } }
  const lookup = {};
  const lookupSec = {};
  Object.entries(data).forEach(([nombre, info]) => {
    const fk = fileKey(nombre);
    if (!lookup[fk]) lookup[fk] = {};
    if (!lookupSec[fk]) lookupSec[fk] = {};
    // ER: leer_er() solo lee "99 GENERAL" (primera sección) → primera ocurrencia gana.
    // Resto (5105, 7205, 7105, AUX): leer_aux() suma todas las ocurrencias → acumular.
    const esER = nombre.toUpperCase().startsWith('ESTADO DE RESULTADOS');
    (info.items || []).forEach(item => {
      if (!item.error && !item.separador && item.codigo) {
        const cod   = String(item.codigo).trim();
        const entry = { valor: Number(item.valor) || 0, debito: Number(item.debito) || 0, credito: Number(item.credito) || 0, op_jk: !!item.op_jk };
        if (esER) {
          if (!(cod in lookup[fk])) lookup[fk][cod] = entry;
        } else {
          if (lookup[fk][cod]) {
            lookup[fk][cod].valor   += entry.valor;
            lookup[fk][cod].debito  += entry.debito;
            lookup[fk][cod].credito += entry.credito;
            if (entry.op_jk) lookup[fk][cod].op_jk = true;
          } else {
            lookup[fk][cod] = entry;
          }
        }
        // Con sección: siempre guardar para acceso preciso (sin acumular — sección:código es única)
        if (item.seccion) lookupSec[fk][`${item.seccion}::${cod}`] = entry;
      }
    });
  });

  // Alias 1: claves cortas del config (ej. "ER_BUGA26") → datos del fileKey() real.
  if (configRaw && configRaw.file_sources) {
    Object.entries(configRaw.file_sources).forEach(([shortKey, info]) => {
      if (lookup[shortKey]) return;
      const pref = (info.prefijo || '').toUpperCase();
      const ent  = (info.entidad || '').toUpperCase();
      for (const nombre of Object.keys(data)) {
        const f = nombre.toUpperCase();
        if (f.startsWith(pref + '_') && f.includes('_' + ent + '_')) {
          const fk = fileKey(nombre);
          if (lookup[fk]) {
            lookup[shortKey]    = lookup[fk];
            lookupSec[shortKey] = lookupSec[fk] || {};
          }
          break;
        }
      }
    });
  }

  // Alias 2: claves sin año (ej. "7205_UT_ALIM_YUMBO") → archivo con año más reciente.
  // Cubre chips arrastrados desde archivos con nombre distinto en sesiones anteriores.
  const allLookupKeys = Object.keys(lookup);
  Object.values(mappings).forEach(sources => {
    sources.forEach(src => {
      if (lookup[src.key]) return; // ya resuelto
      const matches = allLookupKeys.filter(k => k.startsWith(src.key));
      if (matches.length) {
        const best = matches.sort().reverse()[0]; // año más reciente primero
        lookup[src.key]    = lookup[best];
        lookupSec[src.key] = lookupSec[best] || {};
      }
    });
  });

  // Mapa objeto-lookup → clave canónica (para detectar chips duplicados)
  const lookupToKey = new Map();
  Object.entries(lookup).forEach(([k, v]) => { if (!lookupToKey.has(v)) lookupToKey.set(v, k); });

  // Resetear: evita que el mes anterior contamine el Σ cuando se cambia de carpeta
  Object.values(mappings).forEach(sources => {
    sources.forEach(src => { src.valor = 0; src.debito = 0; src.credito = 0; src.op_jk = false; src._resolvedKey = null; });
  });

  // Actualizar valor, debito, credito y op_jk en cada source
  Object.values(mappings).forEach(sources => {
    sources.forEach(src => {
      const fileData = lookup[src.key];
      if (!fileData) return; // archivo no está en esta carpeta
      src._resolvedKey = lookupToKey.get(fileData) || src.key;
      if (src.sin_filtro) {
        const vals = Object.values(fileData);
        src.valor   = vals.reduce((a, e) => a + (Number(e.valor)   || 0), 0);
        src.debito  = vals.reduce((a, e) => a + (Number(e.debito)  || 0), 0);
        src.credito = vals.reduce((a, e) => a + (Number(e.credito) || 0), 0);
        src.op_jk   = vals.some(e => e.op_jk);
      } else {
        const fileDataSec = lookupSec[src.key] || {};
        let totalValor = 0, totalDebito = 0, totalCredito = 0, anyOpJk = false;
        (src.codigos || []).forEach(cod => {
          const codStr = String(cod).trim();
          let entry;
          // Prioridad: sección específica > código simple
          if (src.seccion) entry = fileDataSec[`${src.seccion}::${codStr}`];
          if (entry == null) entry = fileData[codStr];
          // Retrocompatibilidad: código con formato antiguo "010102           GERENCIA"
          if (entry == null) {
            const numPart = codStr.match(/^(\d+)/);
            if (numPart) {
              if (src.seccion) entry = fileDataSec[`${src.seccion}::${numPart[1]}`];
              if (entry == null) entry = fileData[numPart[1]];
            }
          }
          if (entry != null) {
            totalValor   += Number(entry.valor)   || 0;
            totalDebito  += Number(entry.debito)  || 0;
            totalCredito += Number(entry.credito) || 0;
            if (entry.op_jk) anyOpJk = true;
          }
        });
        src.valor   = totalValor;
        src.debito  = totalDebito;
        src.credito = totalCredito;
        src.op_jk   = anyOpJk;
      }
    });
  });

  // Re-renderizar la hoja activa para mostrar los valores/totales actualizados
  if (currentSheet) renderDestRows();
}

function buildAccordionItem(nombre, info) {
  const wrapper = document.createElement('div');
  wrapper.className = 'accordion-item';

  const nTotal   = info.items.filter(i => !i.error && i.negrita).length;
  const nDetalle = info.items.filter(i => !i.error && !i.negrita).length;

  const header = document.createElement('div');
  header.className = 'accordion-header';
  header.innerHTML = `
    <span>${nombre}</span>
    <span style="font-size:.75rem;color:#666;">${info.items.length} filas &nbsp;<span class="arrow">▶</span></span>
  `;

  const body = document.createElement('div');
  body.className = 'accordion-body';

  // Filtros: Todos / Solo detalle / Solo totales
  const filterBar = document.createElement('div');
  filterBar.className = 'chip-filter-bar';
  filterBar.innerHTML = `
    <button class="chip-filter active" data-filter="all">Todos (${info.items.length})</button>
    <button class="chip-filter" data-filter="detalle">Detalle (${nDetalle})</button>
    <button class="chip-filter" data-filter="total">Totales (${nTotal})</button>
  `;

  // Buscador
  const search = document.createElement('input');
  search.type        = 'text';
  search.placeholder = 'Buscar código o descripción...';
  search.className   = 'chip-search';

  const chipsWrap = document.createElement('div');
  chipsWrap.className = 'source-rows-wrap';
  renderDraggableChips(chipsWrap, nombre, info);

  function applyFilters() {
    const q      = search.value.toLowerCase();
    const active = filterBar.querySelector('.chip-filter.active').dataset.filter;
    chipsWrap.querySelectorAll('.source-row').forEach(row => {
      const esTotal    = row.dataset.negrita === 'true';
      const matchFilter = active === 'all'
        || (active === 'total'   &&  esTotal)
        || (active === 'detalle' && !esTotal);
      const matchSearch = !q || row.textContent.toLowerCase().includes(q);
      row.style.display = matchFilter && matchSearch ? '' : 'none';
    });
    // Los separadores siempre visibles (no son filas de datos)
    chipsWrap.querySelectorAll('.source-row-sep').forEach(sep => {
      sep.style.display = '';
    });
  }

  filterBar.querySelectorAll('.chip-filter').forEach(btn => {
    btn.addEventListener('click', () => {
      filterBar.querySelectorAll('.chip-filter').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      applyFilters();
    });
  });

  search.addEventListener('input', applyFilters);

  body.appendChild(filterBar);
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
  const fk = fileKey(nombreArchivo);

  // Cabecera de tabla
  const hdr = document.createElement('div');
  hdr.className = 'source-row-header';
  hdr.innerHTML = `
    <span class="sr-code">Código</span>
    <span class="sr-desc">Descripción</span>
    <span class="sr-val">Valor</span>
  `;
  container.appendChild(hdr);

  info.items.forEach(item => {
    if (item.error) return;

    // Fila separador de sección (ej: "PAEBUGA08 CONTRATO UT BUGA 2026")
    if (item.separador) {
      const sep = document.createElement('div');
      sep.className = 'source-row-sep';
      sep.textContent = item.descripcion;
      container.appendChild(sep);
      return;
    }

    const esTotal  = item.negrita === true;
    const indentPx = 8 + (item.indent || 0) * 14;

    const row = document.createElement('div');
    row.className        = 'source-row' + (esTotal ? ' chip-total' : '');
    row.dataset.negrita  = String(esTotal);
    row.draggable        = true;
    row.style.paddingLeft = `${indentPx}px`;

    const codeEl = document.createElement('span');
    codeEl.className   = 'sr-code';
    codeEl.textContent = item.codigo;

    const descEl = document.createElement('span');
    descEl.className   = 'sr-desc';
    descEl.textContent = item.descripcion || '';

    const valEl = document.createElement('span');
    valEl.className   = 'sr-val';
    valEl.textContent = fmt(item.valor);
    if (item.op_jk) {
      const opBadge = document.createElement('span');
      opBadge.className = 'sr-creditos';
      const c = Number(item.credito || 0);
      opBadge.textContent = c ? `J-K (${fmt(c)})` : 'J-K';
      opBadge.title = 'Valor calculado como Debitos - Creditos';
      valEl.appendChild(opBadge);
    }
    row.appendChild(codeEl);
    row.appendChild(descEl);
    row.appendChild(valEl);

    const dragData = {
      key        : fk,
      archivo    : nombreArchivo,
      codigos    : [item.codigo],
      codigo     : item.codigo,
      descripcion: item.descripcion,
      valor      : item.valor,
      op_jk      : !!item.op_jk,
      debito     : item.debito,
      credito    : item.credito,

      celda      : item.celda,
      seccion    : item.seccion || null
    };

    const chipId = `${fk}::${item.seccion || ''}::${item.codigo}`;

    // Click → seleccionar/deseleccionar
    row.addEventListener('click', () => {
      if (selectedChips.has(chipId)) {
        selectedChips.delete(chipId);
        row.classList.remove('chip-selected');
      } else {
        selectedChips.set(chipId, dragData);
        row.classList.add('chip-selected');
      }
      updateSelectionBar();
    });

    row.addEventListener('dragstart', e => {
      let payload;
      if (selectedChips.size > 0 && selectedChips.has(chipId)) {
        payload = Array.from(selectedChips.values());
        const ghost = document.createElement('div');
        ghost.textContent = `${payload.length} filas`;
        ghost.style.cssText = 'background:#3b82f6;color:#fff;padding:5px 12px;border-radius:12px;font-size:.8rem;position:fixed;top:-200px;left:-200px';
        document.body.appendChild(ghost);
        e.dataTransfer.setDragImage(ghost, 0, 0);
        setTimeout(() => ghost.remove(), 0);
      } else {
        payload = dragData;
      }
      e.dataTransfer.setData('application/json', JSON.stringify(payload));
      e.dataTransfer.effectAllowed = 'copy';
    });

    container.appendChild(row);
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
    clearDirty();
    updateStats();
    showLog('✅ Config guardado correctamente.', true);
    setTimeout(hideLog, 4000);
    return true;
  } else {
    showLog(`❌ Error: ${res.error}`, true);
    return false;
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
    // rowKey puede ser "HOJA::CODIGO|filaNum", "CODIGO|filaNum" o "B:desc|filaNum"
    const parsed  = parseRowKey(rowKey);
    const codigoA = parsed.codigoA;
    const filaNum = parsed.filaNum;

    const configSources = sources.map(src => {
      // Reconstruir clave con entidad desde el archivo completo (evita fallbacks genéricos como "5105" sin entidad)
      const ref     = src.archivo || src.key;
      const entidad = inferEntidad(ref);
      const prefijo = inferPrefijo(ref);
      const savedKey = entidad ? `${prefijo}_${entidad.replace(/\s+/g,'_')}` : src.key;
      if (!fileSources[savedKey]) {
        fileSources[savedKey] = { prefijo, entidad };
      }
      const entry = { key: savedKey, codes: src.codigos || [src.codigo] };
      if (src.sin_filtro)           entry.sin_filtro = true;
      if (src.seccion)              entry.seccion    = src.seccion;
      if (src.op && src.op !== '+') entry.op         = src.op;
      return entry;
    });

    const entry = { codigo_a: codigoA, sources: configSources };
    if (codigoA.startsWith('B:'))  entry.buscar_por = 'B';
    if (filaNum)                   entry.fila_dest  = filaNum;
    if (mappingSheets[rowKey] || parsed.sheetName) entry.hoja = mappingSheets[rowKey] || parsed.sheetName;
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
  if (f.startsWith('7205'))                  return '7205';
  if (f.startsWith('7105'))                  return '7105';
  return archivo.split('_')[0];
}

function inferEntidad(archivo) {
  // Caso normal: filename con _MES_YYYY (ej: "7205_CONS ALIM CALI_MAR_2026.xls")
  const m = (archivo || '').match(/^[^_]+_(.+?)_[A-Z]{3,13}_\d{4}/i);
  if (m) return m[1];
  // Fallback: chip key sin mes (ej: "7205_CONS_ALIM_CALI2026") → convierte _ a espacios
  const m2 = (archivo || '').match(/^(?:ER|AUX|5105|7205|7105|51355001)_(.+)$/i);
  if (m2) return m2[1].replace(/_/g, ' ');
  return '';
}

// ── Ejecutar mes ─────────────────────────────────────────────
async function runMes() {
  const folder = document.getElementById('run-folder').value;
  if (!folder) { alert('Selecciona una carpeta de mes primero.'); return; }

  if (isDirty) {
    showLog('ℹ️ Cambios sin guardar detectados. Guardando antes de ejecutar...', true);
    const okSave = await saveConfig();
    if (!okSave) {
      showLog('❌ No se pudo guardar la configuración. Ejecución cancelada.', false);
      return;
    }
  }

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
  if (el) {
    const dirtyTxt = isDirty ? ' • cambios sin guardar' : '';
    el.textContent = `${configurados} / ${total} ítems configurados${dirtyTxt}`;
  }
}

// ── Subida de archivos al servidor ──────────────────────────────
function toggleUpload() {
  const body  = document.getElementById('upload-body');
  const arrow = document.getElementById('upload-arrow');
  const open  = body.style.display !== 'none';
  body.style.display = open ? 'none' : 'flex';
  arrow.textContent  = open ? '▶' : '▼';
}

async function uploadFiles() {
  const folder = document.getElementById('upload-folder').value.trim().toUpperCase();
  const input  = document.getElementById('upload-files');
  const status = document.getElementById('upload-status');
  const btn    = document.getElementById('btn-upload');

  if (!folder) { status.textContent = 'Escribe el nombre de la carpeta (ej: ABRIL)'; return; }
  if (!input.files.length) { status.textContent = 'Selecciona los archivos XLS primero'; return; }

  btn.disabled = true;
  status.textContent = 'Subiendo...';

  const formData = new FormData();
  for (const f of input.files) formData.append('files', f);

  try {
    const res = await fetch(`/api/upload/${encodeURIComponent(folder)}`, {
      method: 'POST', body: formData
    }).then(r => r.json());

    if (res.ok) {
      status.style.color = '#28a745';
      status.textContent = `OK: ${res.total} archivo(s) subidos a ${res.folder}/`;
      input.value = '';
      const foldersRes = await fetch('/api/folders').then(r => r.json());
      folders = foldersRes;
      populateFolderSelects();
    } else {
      status.style.color = '#dc3545';
      status.textContent = `Error: ${res.error}`;
    }
  } catch (e) {
    status.style.color = '#dc3545';
    status.textContent = `Error de red: ${e.message}`;
  }
  btn.disabled = false;
}

async function subirInformeBase(input) {
  if (!input.files.length) return;
  const archivo = input.files[0];
  if (!archivo.name.toLowerCase().endsWith('.xlsx')) {
    alert('El archivo debe ser .xlsx');
    input.value = '';
    return;
  }
  if (!confirm(`¿Reemplazar el INFORME REAL base con "${archivo.name}"?\nEsto sobreescribirá el archivo actual en el servidor.`)) {
    input.value = '';
    return;
  }
  const formData = new FormData();
  formData.append('file', archivo);
  try {
    const res = await fetch('/api/upload/informe', { method: 'POST', body: formData }).then(r => r.json());
    if (res.ok) {
      alert('✅ ' + res.mensaje);
    } else {
      alert('❌ Error: ' + res.error);
    }
  } catch (e) {
    alert('❌ Error de red: ' + e.message);
  }
  input.value = '';
}

// ── Preview antes de ejecutar ────────────────────────────────
async function previewMes() {
  const folder = document.getElementById('run-folder').value;
  if (!folder) { alert('Selecciona una carpeta de mes primero.'); return; }

  const btn = document.getElementById('btn-preview');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>';

  try {
    const res = await fetch(`/api/preview/${encodeURIComponent(folder)}`).then(r => r.json());
    btn.disabled = false;
    btn.textContent = '🔍 Vista previa';

    if (!res.ok) { alert('Error al calcular preview: ' + (res.error || 'desconocido')); return; }

    const items = res.items || [];
    const advertencias = items.filter(i => i.advertencias && i.advertencias.length > 0);

    // Agrupar por hoja
    const porHoja = {};
    items.forEach(i => {
      if (!porHoja[i.hoja]) porHoja[i.hoja] = [];
      porHoja[i.hoja].push(i);
    });

    let html = `<div class="prev-header">
      <strong>Vista previa — ${folder}</strong>
      <span class="prev-stats">${items.length} ítems · ${advertencias.length > 0 ? `<span class="prev-warn">⚠️ ${advertencias.length} advertencias</span>` : '✅ sin advertencias'}</span>
    </div>`;

    for (const [hoja, filas] of Object.entries(porHoja)) {
      html += `<div class="prev-hoja">${hoja}</div>`;
      html += `<table class="prev-table"><thead><tr><th>Fila</th><th>Código</th><th>Valor</th><th></th></tr></thead><tbody>`;
      filas.forEach(f => {
        const warn = f.advertencias && f.advertencias.length > 0;
        const cls  = warn ? ' class="prev-row-warn"' : '';
        const adv  = warn ? `<span title="${f.advertencias.join(', ')}">⚠️</span>` : '';
        html += `<tr${cls}><td>${f.fila || '—'}</td><td>${f.codigo_a}</td>
          <td class="prev-val">$${(f.valor_total||0).toLocaleString('es-CO')}</td>
          <td>${adv}</td></tr>`;
      });
      html += `</tbody></table>`;
    }

    document.getElementById('prev-body').innerHTML = html;
    document.getElementById('prev-modal').style.display = 'flex';
  } catch (e) {
    btn.disabled = false;
    btn.textContent = '🔍 Vista previa';
    alert('Error de red: ' + e.message);
  }
}

function closePrevModal() {
  document.getElementById('prev-modal').style.display = 'none';
}

async function confirmarEjecutar() {
  closePrevModal();
  await runMes();
}

// ── Historial ─────────────────────────────────────────────────
let historialData = [];

async function loadHistorial() {
  try {
    historialData = await fetch('/api/historial').then(r => r.json());
    renderHistorial();
  } catch (e) {
    document.getElementById('hist-body').textContent = 'Error cargando historial.';
  }
}

function toggleHistorial() {
  const body  = document.getElementById('hist-body');
  const arrow = document.getElementById('hist-arrow');
  const open  = body.style.display !== 'none';
  body.style.display = open ? 'none' : 'block';
  arrow.textContent  = open ? '▶' : '▼';
  if (!open && historialData.length === 0) loadHistorial();
}

function renderHistorial() {
  const el = document.getElementById('hist-body');
  if (!historialData.length) { el.innerHTML = '<p style="padding:8px;color:#888">Sin ejecuciones registradas.</p>'; return; }

  let html = `<table class="hist-table"><thead><tr>
    <th>Fecha</th><th>Mes</th><th>Estado</th><th>Ítems</th><th></th>
  </tr></thead><tbody>`;

  historialData.slice(0, 10).forEach((ej, idx) => {
    const estado = ej.exito ? '✅' : '❌';
    const fecha  = ej.fecha ? ej.fecha.slice(0, 16) : '';
    html += `<tr>
      <td>${fecha}</td>
      <td>${ej.mes}</td>
      <td>${estado}</td>
      <td>${ej.valores ? ej.valores.length : 0}</td>
      <td>
        <button class="btn btn-sm btn-primary" onclick="toggleDetalle(${idx})">Detalle</button>
        ${idx < historialData.length - 1
          ? `<button class="btn btn-sm" style="background:#6c757d;color:#fff;margin-left:4px"
               onclick="compararEjecuciones(${historialData[idx].id},${historialData[idx+1].id},'${historialData[idx].mes}','${historialData[idx+1].mes}')">Comparar</button>`
          : ''}
      </td>
    </tr>
    <tr id="det-${idx}" style="display:none">
      <td colspan="5"><div class="hist-detail">${renderDetalle(ej)}</div></td>
    </tr>`;
  });

  html += '</tbody></table>';
  el.innerHTML = html;
}

function toggleDetalle(idx) {
  const row = document.getElementById(`det-${idx}`);
  row.style.display = row.style.display === 'none' ? 'table-row' : 'none';
}

function renderDetalle(ej) {
  if (!ej.valores || !ej.valores.length) return '<em>Sin datos</em>';
  const porHoja = {};
  ej.valores.forEach(v => { if (!porHoja[v.hoja]) porHoja[v.hoja] = []; porHoja[v.hoja].push(v); });
  let html = '';
  for (const [hoja, vals] of Object.entries(porHoja)) {
    html += `<strong>${hoja}</strong><table class="prev-table" style="margin-bottom:6px"><thead><tr><th>Fila</th><th>Código</th><th>Valor</th></tr></thead><tbody>`;
    vals.forEach(v => {
      html += `<tr><td>${v.fila}</td><td>${v.codigo_a}</td><td class="prev-val">$${(v.valor_total||0).toLocaleString('es-CO')}</td></tr>`;
    });
    html += '</tbody></table>';
  }
  return html;
}

async function compararEjecuciones(idA, idB, mesA, mesB) {
  try {
    const diff = await fetch(`/api/historial/diff?a=${idA}&b=${idB}`).then(r => r.json());
    if (diff.error) { alert('Error: ' + diff.error); return; }

    const cambios = diff.filter(d => d.delta !== null && Math.abs(d.delta) > 0.01);
    const soloA   = diff.filter(d => d.valor_b === null);
    const soloB   = diff.filter(d => d.valor_a === null);

    let html = `<div class="prev-header"><strong>Comparación: ${mesA} vs ${mesB}</strong>
      <span class="prev-stats">${cambios.length} diferencias · ${soloA.length} solo en ${mesA} · ${soloB.length} solo en ${mesB}</span>
    </div>`;

    if (cambios.length) {
      html += `<div class="prev-hoja">Valores cambiados</div>
        <table class="prev-table"><thead><tr><th>Hoja</th><th>Código</th>
          <th>${mesA}</th><th>${mesB}</th><th>Delta</th></tr></thead><tbody>`;
      cambios.forEach(d => {
        const pos = d.delta > 0;
        html += `<tr>
          <td>${d.hoja}</td><td>${d.codigo_a}</td>
          <td class="prev-val">$${d.valor_a.toLocaleString('es-CO')}</td>
          <td class="prev-val">$${d.valor_b.toLocaleString('es-CO')}</td>
          <td class="${pos ? 'delta-pos' : 'delta-neg'}">${pos ? '▲' : '▼'} $${Math.abs(d.delta).toLocaleString('es-CO')}</td>
        </tr>`;
      });
      html += '</tbody></table>';
    } else {
      html += '<p style="padding:12px;color:#28a745">✅ Los valores son idénticos entre ambas ejecuciones.</p>';
    }

    document.getElementById('prev-body').innerHTML = html;
    document.getElementById('prev-modal').style.display = 'flex';
  } catch (e) {
    alert('Error de red: ' + e.message);
  }
}

// ── Evento cambio de carpeta ─────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  init();

  document.getElementById('folder-select').addEventListener('change', e => {
    document.getElementById('run-folder').value = e.target.value;
    loadFolder(e.target.value);
  });
});
