/**
 * ------------------------------------------------------------------
 * CONFIGURACIÓN Y MAPEO DE HOJAS
 * ------------------------------------------------------------------
 */

// IDs extraídos de las URLs proporcionadas
const MAIN_SPREADSHEET_ID = '1MC76eZZt7qiso2M8LMz777_xJnzrl_ZpZptDZBnPlDo'; // Solicitudes, Historicos, Anexos, Solicitudes activos
const PERMISSIONS_SPREADSHEET_ID = '1zcZZGe_93ytWXtCF1kmk_Y8zc5b5cL1xH34i7v1w01k'; // Permisos, Usuarios filtro
const CLIENTS_SPREADSHEET_ID = '1hHWPJF9KSC0opplpCNgHRNkW6CLf7StXG2Y31m6yUpo'; // Clientes
const SEDES_SPREADSHEET_ID = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM'; // Sedes
const ACTIVOS_SPREADSHEET_ID = '1JU8c1MidgV4DRFg6W-GxZ2tHkfKNqGt1_cR5VDTehC4'; // Activos (ejemplo)

// Configuración para saber en qué Spreadsheet buscar cada tabla
const SHEET_CONFIG = {
  'Solicitudes': MAIN_SPREADSHEET_ID,
  'Estados historico': MAIN_SPREADSHEET_ID,
  'Observaciones historico': MAIN_SPREADSHEET_ID,
  'Solicitudes anexos': MAIN_SPREADSHEET_ID,
  'Permisos': PERMISSIONS_SPREADSHEET_ID,
  'Usuarios filtro': PERMISSIONS_SPREADSHEET_ID,
  'Clientes': CLIENTS_SPREADSHEET_ID,
  'Sedes': SEDES_SPREADSHEET_ID,
  'Solicitudes activos': MAIN_SPREADSHEET_ID,
  'Activos': ACTIVOS_SPREADSHEET_ID,
};

/**
 * ------------------------------------------------------------------
 * OPTIMIZACIÓN: Reuso de Spreadsheet abiertos (memoria por ejecución)
 * ------------------------------------------------------------------
 */
const __SS_MEMO = {}; // spreadsheetId -> Spreadsheet
function _openSS(spreadsheetId) {
  if (!__SS_MEMO[spreadsheetId]) __SS_MEMO[spreadsheetId] = SpreadsheetApp.openById(spreadsheetId);
  return __SS_MEMO[spreadsheetId];
}

/**
 * OPTIMIZACIÓN: memo de info de hoja por ejecución
 */
const __SHEET_INFO_MEMO = {}; // sheetName -> info
function _normHeader(x) { return String(x || "").trim().toLowerCase(); }

function _getSheetInfo(sheetName) {
  if (__SHEET_INFO_MEMO[sheetName]) return __SHEET_INFO_MEMO[sheetName];

  const spreadsheetId = SHEET_CONFIG[sheetName];
  if (!spreadsheetId) throw new Error(`Configuración no encontrada para la tabla: ${sheetName}`);

  const ss = _openSS(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const info = { sheet: null, headers: [], headersNorm: [], indexByNorm: {}, lastRow: 0, lastCol: 0 };
    __SHEET_INFO_MEMO[sheetName] = info;
    return info;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    const info = { sheet, headers: [], headersNorm: [], indexByNorm: {}, lastRow, lastCol };
    __SHEET_INFO_MEMO[sheetName] = info;
    return info;
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const headersNorm = headers.map(_normHeader);
  const indexByNorm = {};
  headersNorm.forEach((h, i) => { if (h && indexByNorm[h] === undefined) indexByNorm[h] = i; });

  const info = { sheet, headers, headersNorm, indexByNorm, lastRow, lastCol };
  __SHEET_INFO_MEMO[sheetName] = info;
  return info;
}

function _findColIndex(headersNorm, candidateNames) {
  for (let i = 0; i < candidateNames.length; i++) {
    const idx = headersNorm.indexOf(_normHeader(candidateNames[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

function _rowToObject(headers, row) {
  const obj = {};
  for (let i = 0; i < headers.length; i++) {
    let v = row[i];
    if (v instanceof Date) v = v.toISOString();
    obj[headers[i]] = v;
  }
  return obj;
}

function _mergeRowRuns(rowNumsSorted) {
  const runs = [];
  if (!rowNumsSorted.length) return runs;
  let s = rowNumsSorted[0], p = rowNumsSorted[0];
  for (let i = 1; i < rowNumsSorted.length; i++) {
    const cur = rowNumsSorted[i];
    if (cur === p + 1) {
      p = cur;
    } else {
      runs.push([s, p]);
      s = p = cur;
    }
  }
  runs.push([s, p]);
  return runs;
}

function _fetchRowRuns(sheet, runs, lastCol) {
  const rows = [];
  runs.forEach(([start, end]) => {
    const num = end - start + 1;
    const block = sheet.getRange(start, 1, num, lastCol).getValues();
    for (let i = 0; i < block.length; i++) rows.push(block[i]);
  });
  return rows;
}

/**
 * Busca una fila por valor exacto en alguna columna candidata y devuelve {rowNum, obj}
 */
function _findRowObjectByKey(sheetName, keyValue, colCandidates) {
  const { sheet, headers, headersNorm, lastRow, lastCol } = _getSheetInfo(sheetName);
  if (!sheet || lastRow < 2) return null;

  const k = String(keyValue || "").trim();
  if (!k) return null;

  for (let c = 0; c < colCandidates.length; c++) {
    const idx = _findColIndex(headersNorm, [colCandidates[c]]);
    if (idx === -1) continue;

    const colVals = sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < colVals.length; i++) {
      const v = String(colVals[i][0]).trim();
      if (v === k) {
        const rowNum = i + 2;
        const row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
        return { rowNum, obj: _rowToObject(headers, row) };
      }
    }
  }

  return null;
}

/**
 * Devuelve hijos filtrados por FK (sin leer toda la hoja completa)
 */
function _getChildrenFast(sheetName, parentKeys) {
  const { sheet, headers, headersNorm, lastRow, lastCol } = _getSheetInfo(sheetName);
  if (!sheet || lastRow < 2) return [];

  const fkCandidates = ['ID Solicitudes', 'ID Solicitud', 'ID Solicitudes '];
  const idxFk = _findColIndex(headersNorm, fkCandidates);
  if (idxFk === -1) return [];

  const idxDate = _findColIndex(headersNorm, ['Fecha Actualización', 'Fecha Actualizacion', 'Fecha', 'FechaCambio']);

  const set = new Set((parentKeys || []).map(x => String(x || "").trim()).filter(Boolean));
  if (!set.size) return [];

  // Solo leemos la columna FK (barato)
  const fkVals = sheet.getRange(2, idxFk + 1, lastRow - 1, 1).getValues();
  const rowNums = [];
  for (let i = 0; i < fkVals.length; i++) {
    const fk = String(fkVals[i][0]).trim();
    if (fk && set.has(fk)) rowNums.push(i + 2);
  }
  if (!rowNums.length) return [];

  rowNums.sort((a, b) => a - b);
  const runs = _mergeRowRuns(rowNums);

  // Traemos solo filas necesarias (por bloques contiguos)
  const rows = _fetchRowRuns(sheet, runs, lastCol);

  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const obj = _rowToObject(headers, r);

    let ts = 0;
    if (idxDate !== -1) {
      const dv = r[idxDate];
      ts = (dv instanceof Date) ? dv.getTime() : (Date.parse(String(dv)) || 0);
    }
    obj.__ts = ts;
    out.push(obj);
  }

  // Orden por fecha desc
  out.sort((a, b) => (b.__ts || 0) - (a.__ts || 0));
  out.forEach(o => delete o.__ts);

  return out;
}

/**
 * ------------------------------------------------------------------
 * Sirve el HTML principal (SPA)
 * ------------------------------------------------------------------
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('G4S Ticket Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ------------------------------------------------------------------
 * RESOLVER EMAIL (✅ evita fallos cuando Session.getActiveUser() viene vacío)
 * ------------------------------------------------------------------
 */
function _resolveCallerEmail(request) {
  const active = Session.getActiveUser().getEmail();
  if (active) return String(active).toLowerCase().trim();

  const p = request?.payload || {};
  const fromClient = p.__clientEmail || p.clientEmail || request?.clientEmail || "";
  const email = String(fromClient).toLowerCase().trim();
  if (email && email.includes("@")) return email;

  return "";
}

/**
 * Router Principal de la API
 */
function apiHandler(request) {
  const userEmail = _resolveCallerEmail(request);
  const { endpoint, payload } = request || {};

  console.log(`🔒 [API CHECK] Endpoint: ${endpoint} | ActiveUser: ${Session.getActiveUser().getEmail()} | Resuelto: ${userEmail}`);

  try {
    if (!userEmail) throw new Error("No se pudo verificar la identidad del usuario.");

    switch (endpoint) {
      case 'getUserContext': return getUserContext(userEmail);
      case 'getRequests': return getRequests(userEmail);
      case 'getRequestDetail': return getRequestDetail(userEmail, payload);
      case 'createRequest': return createRequest(userEmail, payload);

      case 'uploadAnexo': return uploadAnexo(userEmail, payload);
      case 'getAnexoDownload': return getAnexoDownload(userEmail, payload);

      // ✅ Activos por QR
      case 'createSolicitudActivo': return createSolicitudActivo(userEmail, payload);
      case 'getSolicitudActivos': return getSolicitudActivos(userEmail, payload);
      case 'getActivosCatalog': return getActivosCatalog(userEmail);
      case 'getActivoByQr': return getActivoByQr(userEmail, payload);

      default: throw new Error(`Endpoint desconocido: ${endpoint}`);
    }

  } catch (err) {
    console.error(`❌ ERROR DE SEGURIDAD/EJECUCIÓN: ${err.message}`, err);
    return { error: true, message: "Error procesando su solicitud. Contacte al administrador." };
  }
}

/**
 * ------------------------------------------------------------------
 * ✅ CACHE SOLO PARA DETALLE
 * ------------------------------------------------------------------
 */
const DETAIL_CACHE_VER = "v3"; // cambia versión para “barrer” cache viejo
function _detailCacheKey(email, id) {
  const e = String(email || "").toLowerCase().trim();
  const rid = String(id || "").trim();
  return `detail_${DETAIL_CACHE_VER}_${Utilities.base64Encode(e)}_${rid}`;
}
function _invalidateDetailCache(email, id) {
  try { CacheService.getScriptCache().remove(_detailCacheKey(email, id)); } catch (e) {}
}

/**
 * ------------------------------------------------------------------
 * SUBIR ANEXOS EN FORMULARIO SOLICITUD
 * ------------------------------------------------------------------
 */

function _sanitizeFileName(name) {
  const n = String(name || 'anexo').trim();
  return n
    .replace(/[/\\]/g, '_')
    .replace(/[<>:"|?*]/g, '')
    .replace(/\s+/g, ' ')
    .slice(0, 120) || 'anexo';
}

function _ensureFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function _ensurePathFromRoot(root, parts) {
  let current = root;
  parts.forEach(p => { current = _ensureFolder(current, p); });
  return current;
}

function uploadAnexo(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  const solicitudId = String(payload?.solicitudId || '').trim();
  const tipoAnexo = String(payload?.tipoAnexo || 'Archivo').trim();
  const fileName = _sanitizeFileName(payload?.fileName || 'anexo');
  const mimeType = String(payload?.mimeType || 'application/octet-stream').trim();
  const base64 = String(payload?.base64 || '').trim();

  if (!solicitudId) throw new Error("solicitudId requerido");
  if (!base64) throw new Error("base64 requerido");

  // ✅ header rápido (sin leer toda la hoja)
  const headerFound = _findSolicitudHeaderFast(solicitudId);
  if (!headerFound) throw new Error("Solicitud padre no encontrada.");
  const header = headerFound.obj;

  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede']));
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) {
      throw new Error("No tiene permisos para anexar archivos a este ticket.");
    }
  }

  const maxBytes = 8 * 1024 * 1024; // 8MB
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > maxBytes) throw new Error("Archivo demasiado grande (máx 8MB).");

  const root = _getRootFolderForFiles();
  const folder = _ensurePathFromRoot(root, ['Info', 'Clientes', 'Anexos']);

  const shortId = solicitudId.replace(/-/g, '').slice(0, 8);
  const rand = Math.floor(Math.random() * 900000) + 100000;

  const hasExt = fileName.includes('.') ? fileName.split('.').pop() : '';
  const safeBase = fileName.replace(/\.[^.]+$/, '');
  const finalName = hasExt
    ? `${shortId}.${tipoAnexo}.${rand}.${safeBase}.${hasExt}`
    : `${shortId}.${tipoAnexo}.${rand}.${safeBase}`;

  const blob = Utilities.newBlob(bytes, mimeType, finalName);
  const file = folder.createFile(blob);

  const storedPath = `/Info/Clientes/Anexos/${file.getName()}`;

  const now = new Date();
  const anexoUuid = Utilities.getUuid();

  const row = {
    "ID Solicitudes anexos": anexoUuid,
    "ID Solicitudes": solicitudId,
    "Tipo anexo": tipoAnexo,
    "Nombre": safeBase || file.getName(),
    "Archivo": storedPath,
    "Archivo ": storedPath,
    "Usuario Actualización": email,
    "Usuario Actualizacion": email,
    "Fecha Actualización": now,
    "Fecha Actualizacion": now
  };

  appendDataToSheet('Solicitudes anexos', row);

  _invalidateDetailCache(email, solicitudId);
  return { success: true, anexoId: anexoUuid, fileName: file.getName(), path: storedPath };
}

/**
 * ------------------------------------------------------------------
 * HELPERS GENÉRICOS (se mantienen para compatibilidad)
 * ------------------------------------------------------------------
 */

function getDataFromSheet(sheetName) {
  const { sheet, headers, lastRow, lastCol } = _getSheetInfo(sheetName);
  if (!sheet || lastRow < 2 || lastCol < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  if (values.length < 2) return [];

  const data = values.slice(1);
  return data.map(row => _rowToObject(headers, row));
}

function appendDataToSheet(sheetName, objectData) {
  const spreadsheetId = SHEET_CONFIG[sheetName];
  const ss = _openSS(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Hoja ${sheetName} no encontrada.`);

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) throw new Error(`La hoja ${sheetName} está vacía (sin cabeceras).`);

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowArray = headers.map(header => {
    const val = objectData[header];
    return val === undefined || val === null ? "" : val;
  });

  sheet.appendRow(rowArray);

  // Invalida detalle si aplica
  const possibleTicketId = objectData["ID Solicitud"] || objectData["ID Solicitudes"] || "";
  if (possibleTicketId) _invalidateDetailCache(String(objectData["Usuario Actualización"] || ""), possibleTicketId);

  return { success: true };
}

/**
 * ------------------------------------------------------------------
 * HELPERS ROBUSTOS (encabezados / rutas)
 * ------------------------------------------------------------------
 */
function _getField(row, candidateNames) {
  if (!row) return "";
  const keys = Object.keys(row);
  for (let i = 0; i < candidateNames.length; i++) {
    const c = candidateNames[i];

    if (row[c] !== undefined && row[c] !== null && row[c] !== "") return row[c];

    const k = keys.find(x => String(x).trim() === String(c).trim());
    if (k && row[k] !== undefined && row[k] !== null && row[k] !== "") return row[k];
  }
  return "";
}

function _normalizePath(path) {
  if (!path) return "";
  let p = String(path).trim();
  if (/^https?:\/\//i.test(p)) return p;
  p = p.replace(/\\/g, '/');
  p = p.replace(/^\/+/, '');
  p = p.replace(/\/+/g, '/');
  return p;
}

function _getRootFolderForFiles() {
  const file = DriveApp.getFileById(MAIN_SPREADSHEET_ID);
  const parents = file.getParents();
  if (parents.hasNext()) return parents.next();
  return DriveApp.getRootFolder();
}

function _resolveDriveFileFromAppSheetPath(pathValue) {
  const p = _normalizePath(pathValue);

  if (/^https?:\/\//i.test(p)) {
    return { kind: "url", url: p };
  }

  const root = _getRootFolderForFiles();
  const parts = p.split('/').filter(Boolean);
  const filename = parts.pop();

  try {
    let current = root;
    parts.forEach(folderName => {
      const it = current.getFoldersByName(folderName);
      if (!it.hasNext()) throw new Error(`Carpeta no encontrada: ${folderName}`);
      current = it.next();
    });

    const files = current.getFilesByName(filename);
    if (files.hasNext()) return { kind: "file", file: files.next() };

  } catch (e) {
    // fallback
  }

  const safeName = filename.replace(/"/g, '\\"');
  const q = `name = "${safeName}" and trashed = false`;
  const it2 = DriveApp.searchFiles(q);
  if (it2.hasNext()) return { kind: "file", file: it2.next() };

  throw new Error(`Archivo no encontrado en Drive: ${filename}`);
}

/**
 * Header rápido de Solicitudes por: ID Solicitud / Ticket G4S / Ticket Cliente
 */
function _findSolicitudHeaderFast(key) {
  return _findRowObjectByKey('Solicitudes', key, [
    'ID Solicitud', 'ID Solicitudes',
    'Ticket G4S',
    'Ticket Cliente', 'Ticket (Opcional)'
  ]);
}

/**
 * ------------------------------------------------------------------
 * LÓGICA DE NEGOCIO
 * ------------------------------------------------------------------
 */

function getUserContext(email) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `ctx_it_v4_${Utilities.base64Encode(email)}`;
  const cachedData = cache.get(cacheKey);
  if (cachedData) return JSON.parse(cachedData);

  try {
    let context = {
      email: email,
      role: 'Usuario',
      allowedClientIds: [],
      clientNames: {},
      isValidUser: false,
      isAdmin: false
    };

    const allPermissions = getDataFromSheet('Permisos');
    const userData = allPermissions.find(row => String(row['Correo']).toLowerCase() === email.toLowerCase());

    if (userData) {
      context.isValidUser = true;
      const rol = (userData['Rol_Asignado'] || '').trim().toLowerCase();
      if (rol === 'administrador') {
        context.role = 'Administrador';
        context.isAdmin = true;
      }
    } else {
      console.warn(`⚠️ Usuario IT ${email} no encontrado en tabla Permisos.`);
    }

    if (!context.isValidUser) return context;

    const allRelations = getDataFromSheet('Usuarios filtro');
    const myRelations = allRelations.filter(row => String(row['Usuario']).toLowerCase() === email.toLowerCase());

    const assignedClientIds = [];
    myRelations.forEach(row => {
      const id = row['Cliente'];
      if (id) assignedClientIds.push(String(id));
    });

    if (assignedClientIds.length > 0) {
      const allSedes = getDataFromSheet('Sedes');
      const mySedes = allSedes.filter(sede => assignedClientIds.includes(String(_getField(sede, ['ID Cliente', 'Id Cliente', 'Cliente']))));

      mySedes.forEach(sede => {
        const idSede = String(_getField(sede, ['ID Sede', 'Id Sede', 'Sede', 'IDSede'])).trim();
        const nombreSede =
          _getField(sede, ['Nombre', 'Nombre_Sede', 'Nombre sede', 'Nombre Sede', 'Sede', 'Label']) ||
          idSede;

        if (idSede) {
          context.allowedClientIds.push(idSede);
          context.clientNames[idSede] = String(nombreSede).trim() || idSede;
        }
      });
    }

    cache.put(cacheKey, JSON.stringify(context), 600);
    return context;

  } catch (e) {
    console.error("Error getUserContext", e);
    throw e;
  }
}

function getRequests(email) {
  const t0 = Date.now();
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  try {
    // Nota: aquí mantenemos getDataFromSheet por compatibilidad.
    // Si esta pantalla también está lenta, lo optimizamos igual que detalle.
    let allRows = getDataFromSheet('Solicitudes');
    let filteredRows = [];

    if (context.isAdmin) {
      filteredRows = allRows;
    } else {
      if (context.allowedClientIds.length === 0) return { data: [], total: 0 };
      filteredRows = allRows.filter(row => context.allowedClientIds.includes(String(_getField(row, ['ID Sede']))));
    }

    filteredRows.sort((a, b) => {
      const dateA = new Date(_getField(a, ['Fecha creación cliente', 'Fecha creacion cliente'])).getTime() || 0;
      const dateB = new Date(_getField(b, ['Fecha creación cliente', 'Fecha creacion cliente'])).getTime() || 0;
      return dateB - dateA;
    });

    console.log(`⚡ [PERF] getRequests: ${Date.now() - t0}ms | total=${filteredRows.length}`);
    return { data: filteredRows, total: filteredRows.length };

  } catch (e) {
    console.error("Error getRequests", e);
    throw new Error("Error obteniendo datos.");
  }
}

function getRequestDetail(email, { id }) {
  const t0 = Date.now();
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado");
  if (!id) throw new Error("ID requerido");

  const rid = String(id).trim();
  const cache = CacheService.getScriptCache();
  const ck = _detailCacheKey(email, rid);

  const cached = cache.get(ck);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      console.log(`⚡ [PERF] getRequestDetail (cache): ${Date.now() - t0}ms`);
      return parsed;
    } catch (e) {}
  }

  // Header rápido (1 fila)
  const headerFound = _findSolicitudHeaderFast(rid);
  if (!headerFound) throw new Error("Ticket no encontrado.");
  const header = headerFound.obj;

  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede'])).trim();
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) {
      throw new Error("No tiene permisos para ver este ticket.");
    }
  }

  const parentKeys = [
    rid,
    String(_getField(header, ['Ticket G4S'])),
    String(_getField(header, ['Ticket Cliente', 'Ticket (Opcional)']))
  ].filter(x => x && x !== "undefined" && x !== "null").map(x => String(x).trim());

  // Hijos rápidos (solo filas del ticket)
  const services = _getChildrenFast('Observaciones historico', parentKeys);
  const history = _getChildrenFast('Estados historico', parentKeys);
  const documents = _getChildrenFast('Solicitudes anexos', parentKeys);

  const result = { header, services, history, documents };

  const json = JSON.stringify(result);
  if (json.length < 90000) cache.put(ck, json, 30);

  console.log(`⚡ [PERF] getRequestDetail (full): ${Date.now() - t0}ms | obs=${services.length} hist=${history.length} anex=${documents.length}`);
  return result;
}

function createRequest(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  if (!payload?.idSede || !payload?.solicitud || !payload?.observacion) {
    throw new Error("Faltan campos obligatorios.");
  }

  if (!context.isAdmin && !context.allowedClientIds.includes(String(payload.idSede))) {
    throw new Error("No tiene permisos para esta sede.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("Servidor ocupado, intente de nuevo.");
  }

  try {
    const now = new Date();
    const uuid = Utilities.getUuid();

    const allSedes = getDataFromSheet('Sedes');
    const sedeInfo = allSedes.find(s => String(_getField(s, ['ID Sede', 'Id Sede', 'Sede'])).trim() === String(payload.idSede).trim());
    const idCliente = sedeInfo ? _getField(sedeInfo, ['ID Cliente', 'Id Cliente', 'Cliente']) : null;

    let letraInicial = "X";
    if (idCliente) {
      const allClientes = getDataFromSheet('Clientes');
      const clienteInfo = allClientes.find(c => String(_getField(c, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim() === String(idCliente).trim());
      if (clienteInfo) {
        const nombreCorto = _getField(clienteInfo, ['Nombre corto', 'Nombre_Corto', 'RazonSocial', 'Razón Social']) || "G";
        letraInicial = String(nombreCorto).trim().charAt(0).toUpperCase();
      }
    }

    const ss = _openSS(MAIN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Solicitudes');
    const nextRow = sheet.getLastRow() + 1;

    const rand = Math.floor(Math.random() * 90) + 10;
    const ticketG4S = `${letraInicial}${1000000 + nextRow}${rand}`;

    const newRow = {
      "ID Solicitud": uuid,
      "Ticket G4S": ticketG4S,
      "Fecha creación cliente": now,
      "Estado": "Abierto",
      "ID Sede": String(payload.idSede).trim(),
      "Ticket Cliente": payload.ticketCliente || "",
      "Clasificación": payload.clasificacion,
      "Prioridad Solicitud": payload.prioridad,
      "Solicitud": payload.solicitud,
      "Observación": payload.observacion,
      "Usuario Actualización": email
    };

    appendDataToSheet('Solicitudes', newRow);
    _invalidateDetailCache(email, uuid);

    const returnRow = {
      ...newRow,
      "Fecha creación cliente": (newRow["Fecha creación cliente"] instanceof Date)
        ? newRow["Fecha creación cliente"].toISOString()
        : newRow["Fecha creación cliente"]
    };

    return {
      success: true,
      solicitudId: uuid,
      ticketG4S: ticketG4S,
      GeneratedTicket: ticketG4S,
      Status: "Success",
      Rows: [returnRow],
      row: returnRow
    };

  } catch (e) {
    console.error("Error createRequest", e);
    throw new Error("Error guardando el ticket: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * ------------------------------------------------------------------
 * ✅ DESCARGA DE ANEXOS (Drive) SEGÚN RUTA GUARDADA EN LA TABLA
 * ------------------------------------------------------------------
 */
function getAnexoDownload(email, { anexoId }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");
  if (!anexoId) throw new Error("anexoId requerido");

  const found = _findRowObjectByKey('Solicitudes anexos', anexoId, [
    'ID Solicitudes anexos', 'ID Solicitud anexos', 'ID Anexo', 'ID', 'ID Solicitudes anexos '
  ]);
  if (!found) throw new Error("Anexo no encontrado.");
  const row = found.obj;

  const parentKey = _getField(row, ['ID Solicitudes', 'ID Solicitud']);
  const headerFound = _findSolicitudHeaderFast(parentKey);
  if (!headerFound) throw new Error("No se pudo validar la solicitud padre del anexo.");
  const header = headerFound.obj;

  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede'])).trim();
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) {
      throw new Error("No tiene permisos para descargar este anexo.");
    }
  }

  const pathValue = _getField(row, ['Archivo', 'Archivo ', 'Foto', 'Dibujo', 'QR']) || "";
  if (!pathValue) throw new Error("Este anexo no tiene archivo asociado.");

  const resolved = _resolveDriveFileFromAppSheetPath(pathValue);

  if (resolved.kind === "url") {
    const fileNameFromRow = _getField(row, ['Nombre']) || "Anexo";
    return { mode: "url", fileName: fileNameFromRow, url: resolved.url };
  }

  const file = resolved.file;
  const blob = file.getBlob();
  const mimeType = blob.getContentType() || "application/octet-stream";

  const originalName = file.getName() || "Anexo";
  const ext = (originalName.includes('.') ? originalName.split('.').pop() : "");
  let friendly = _getField(row, ['Nombre']) || originalName;
  if (ext && !String(friendly).toLowerCase().endsWith("." + ext.toLowerCase())) {
    if (!String(friendly).includes('.')) friendly = `${friendly}.${ext}`;
  }

  const maxBytes = 8 * 1024 * 1024;
  const bytes = blob.getBytes();

  if (bytes.length > maxBytes) {
    const url = `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    return { mode: "url", fileName: friendly, url: url, note: "Archivo grande: usando enlace de Drive." };
  }

  const base64 = Utilities.base64Encode(bytes);
  return { mode: "base64", fileName: friendly, mimeType: mimeType, base64: base64 };
}

/**
 * ------------------------------------------------------------------
 * ✅ SOLICITUDES ACTIVOS (QR)
 * ------------------------------------------------------------------
 */
function createSolicitudActivo(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  const solicitudId = String(payload?.solicitudId || payload?.IDSolicitudes || payload?.idSolicitud || '').trim();
  const qr = String(payload?.qrSerial || payload?.qr || payload?.QR || '').trim();
  const idActivo = String(payload?.idActivo || payload?.activoId || payload?.IDActivo || '').trim();

  const observaciones = String(payload?.observaciones || payload?.novedades || '').trim();
  const dibujoBase64 = String(payload?.dibujoBase64 || '').trim();

  if (!solicitudId) throw new Error("solicitudId requerido");
  if (!qr) throw new Error("QR requerido");
  if (!idActivo) throw new Error("ID Activo requerido");

  const headerFound = _findSolicitudHeaderFast(solicitudId);
  if (!headerFound) throw new Error("Solicitud padre no encontrada.");
  const header = headerFound.obj;

  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede'])).trim();
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) {
      throw new Error("No tiene permisos para asociar activos a este ticket.");
    }
  }

  let dibujoPath = "";
  if (dibujoBase64) {
    const bytes = Utilities.base64Decode(dibujoBase64);
    const maxBytes = 2 * 1024 * 1024;
    if (bytes.length > maxBytes) throw new Error("Dibujo demasiado grande (máx 2MB).");

    const root = _getRootFolderForFiles();
    const folder = _ensurePathFromRoot(root, ['Info', 'Clientes', 'Activos']);

    const short = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    const rand = Math.floor(Math.random() * 900000) + 100000;
    const fileName = `${short}.Dibujo.${rand}.png`;

    const blob = Utilities.newBlob(bytes, 'image/png', fileName);
    folder.createFile(blob);

    dibujoPath = `/Info/Clientes/Activos/${fileName}`;
  }

  const now = new Date();
  const rowId = Utilities.getUuid();

  const row = {
    "ID Solicitudes activos": rowId,
    "ID Solicitudes": solicitudId,
    "QR": qr,
    "ID Activo": idActivo,
    "Observaciones": observaciones,
    "Dibujo": dibujoPath,
    "Usuario Actualización": email,
    "Fecha Actualización": now
  };

  appendDataToSheet('Solicitudes activos', row);
  _invalidateDetailCache(email, solicitudId);

  return { success: true, activoRowId: rowId, dibujoPath };
}

function getSolicitudActivos(email, { solicitudId }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");
  const sid = String(solicitudId || '').trim();
  if (!sid) throw new Error("solicitudId requerido");

  const headerFound = _findSolicitudHeaderFast(sid);
  if (!headerFound) throw new Error("Solicitud no encontrada.");
  const header = headerFound.obj;

  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede'])).trim();
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) {
      throw new Error("No tiene permisos para ver activos de este ticket.");
    }
  }

  const rows = _getChildrenFast('Solicitudes activos', [sid]);

  return { data: rows, total: rows.length };
}

/**
 * ------------------------------------------------------------------
 * ✅ CATÁLOGO DE ACTIVOS + BÚSQUEDA POR QR
 * ------------------------------------------------------------------
 */
function getActivosCatalog(email) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  const cache = CacheService.getScriptCache();
  const key = "activos_catalog_v2";
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const rows = getDataFromSheet('Activos');

  const mapped = rows.map(r => {
    const idActivo = String(_getField(r, ['ID Activo', 'Id Activo', 'ID', 'Id'])).trim();
    const nombreActivo = String(_getField(r, ['Nombre Activo', 'Nombre', 'Activo'])).trim();
    const qrSerial = String(_getField(r, ['QR Serial', 'QR', 'Qr', 'Codigo QR'])).trim();
    const nombreUbicacion = String(_getField(r, ['Nombre Ubicacion', 'Ubicación', 'Ubicacion', 'Ubic'])).trim();
    const estadoActivo = String(_getField(r, ['Estado Activo', 'Estado', 'Condicion'])).trim();
    const funcionamiento = String(_getField(r, ['Funcionamiento', 'Funciona', 'Operativo'])).trim();

    return { idActivo, nombreActivo, qrSerial, nombreUbicacion, estadoActivo, funcionamiento };
  }).filter(x => x.idActivo || x.qrSerial || x.nombreActivo);

  const res = { data: mapped, total: mapped.length };
  cache.put(key, JSON.stringify(res), 600);
  return res;
}

function getActivoByQr(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  const q = String(payload?.qr || payload?.qrSerial || '').trim();
  if (!q) throw new Error("qr requerido");

  const rows = getDataFromSheet('Activos');
  const found = rows.find(r => String(_getField(r, ['QR Serial', 'QR', 'Qr', 'Codigo QR'])).trim() === q);

  if (!found) return { found: false };

  return {
    found: true,
    activo: {
      idActivo: String(_getField(found, ['ID Activo', 'Id Activo', 'ID', 'Id'])).trim(),
      nombreActivo: String(_getField(found, ['Nombre Activo', 'Nombre', 'Activo'])).trim(),
      qrSerial: q,
      nombreUbicacion: String(_getField(found, ['Nombre Ubicacion', 'Ubicación', 'Ubicacion', 'Ubic'])).trim(),
      estadoActivo: String(_getField(found, ['Estado Activo', 'Estado', 'Condicion'])).trim(),
      funcionamiento: String(_getField(found, ['Funcionamiento', 'Funciona', 'Operativo'])).trim()
    }
  };
}
