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
let selectedChips = new Map(); // chipId → dragData (selección múltiple)

// ── Utilidades ──────────────────────────────────────────────
const fmt = v => v == null ? '' : '$' + Number(v).toLocaleString('es-CO', {maximumFractionDigits: 0});
const norm = s => s.replace(/\s+/g, ' ').trim();

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
  if (f.startsWith('7205')) {
    const m = f.match(/7205_(.+?)_[A-Z]{3}_/);
    return m ? '7205_' + m[1].replace(/\s+/g,'_') : '7205';
  }
  if (f.startsWith('7105')) {
    const m = f.match(/7105_(.+?)_[A-Z]{3}_/);
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
    const hasData = sources.length > 0;

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
    const op  = src.op || '+';
    const val = src.valor != null ? Number(src.valor) : 0;
    if      (op === '+') total += val;
    else if (op === '-') total -= val;
    else if (op === '*') total *= val;
    else if (op === '/') total = val !== 0 ? total / val : total;

    const chip = document.createElement('span');
    chip.className = 'source-chip';
    const codigos = Array.isArray(src.codigos) ? src.codigos.join(', ') : (src.codigo || '');

    // Botón de operador (+/−) — solo visible si hay más de un chip o si ya es −
    const opBtn = document.createElement('button');
    opBtn.className = `chip-op-btn ${op === '-' ? 'op-minus' : 'op-plus'}`;
    opBtn.textContent = op === '-' ? '−' : '+';
    opBtn.title = 'Clic para cambiar a ' + (op === '-' ? 'suma' : 'resta');
    opBtn.addEventListener('click', e => {
      e.stopPropagation();
      mappings[rowKey][idx].op = op === '-' ? '+' : '-';
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

  // Evitar duplicado exacto (mismo archivo + mismo codigo + misma sección)
  const existe = mappings[rowKey].some(s =>
    (s.codigos || []).join(',') === (chipData.codigos || [chipData.codigo]).join(',') &&
    s.key === chipData.key &&
    (s.seccion || null) === (chipData.seccion || null)
  );
  if (existe) return;

  mappings[rowKey].push({
    key     : chipData.key,
    archivo : chipData.archivo,
    codigos : chipData.codigos || [chipData.codigo],
    valor   : chipData.valor,

    celda   : chipData.celda,
    seccion : chipData.seccion || null,
    op      : '+'
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
  filesData  = data;

  // Actualizar src.valor en los mappings guardados usando los datos recién cargados
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
  // lookup simple: { fileKey: { codigo: valor } } — primera aparición gana (evita sobrescritura entre secciones)
  // lookupSec: { fileKey: { "seccion::codigo": valor } } — valor exacto por sección
  const lookup = {};
  const lookupSec = {};
  Object.entries(data).forEach(([nombre, info]) => {
    const fk = fileKey(nombre);
    if (!lookup[fk]) lookup[fk] = {};
    if (!lookupSec[fk]) lookupSec[fk] = {};
    (info.items || []).forEach(item => {
      if (!item.error && !item.separador && item.codigo) {
        const cod = String(item.codigo).trim();
        // Simple: primera ocurrencia gana (99 GENERAL aparece primero en el archivo)
        if (!(cod in lookup[fk])) lookup[fk][cod] = item.valor;
        // Con sección: siempre guardar para acceso preciso
        if (item.seccion) lookupSec[fk][`${item.seccion}::${cod}`] = item.valor;
      }
    });
  });

  // Para cada source en todos los mappings, calcular el valor sumando los códigos
  Object.values(mappings).forEach(sources => {
    sources.forEach(src => {
      const fileData = lookup[src.key];
      if (!fileData) return; // archivo no está en esta carpeta, no tocar
      if (src.sin_filtro) {
        src.valor = Object.values(fileData).reduce((a, v) => a + (Number(v) || 0), 0);
      } else {
        const fileDataSec = lookupSec[src.key] || {};
        src.valor = (src.codigos || []).reduce((sum, cod) => {
          const codStr = String(cod).trim();
          let v;
          // Prioridad: sección específica > código simple (evita colisión entre secciones)
          if (src.seccion) v = fileDataSec[`${src.seccion}::${codStr}`];
          if (v == null) v = fileData[codStr];
          // Retrocompatibilidad: código con formato antiguo "010102           GERENCIA"
          if (v == null) {
            const numPart = codStr.match(/^(\d+)/);
            if (numPart) {
              if (src.seccion) v = fileDataSec[`${src.seccion}::${numPart[1]}`];
              if (v == null) v = fileData[numPart[1]];
            }
          }
          return sum + (v != null ? Number(v) : 0);
        }, 0);
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
    // rowKey puede ser "HOJA::CODIGO|filaNum", "CODIGO|filaNum" o "B:desc|filaNum"
    const parsed  = parseRowKey(rowKey);
    const codigoA = parsed.codigoA;
    const filaNum = parsed.filaNum;

    const configSources = sources.map(src => {
      if (!fileSources[src.key]) {
        fileSources[src.key] = { prefijo: inferPrefijo(src.archivo || src.key), entidad: inferEntidad(src.archivo || src.key) };
      }
      const entry = { key: src.key, codes: src.codigos || [src.codigo] };
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

// ── Evento cambio de carpeta ─────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  init();

  document.getElementById('folder-select').addEventListener('change', e => {
    document.getElementById('run-folder').value = e.target.value;
    loadFolder(e.target.value);
  });
});
