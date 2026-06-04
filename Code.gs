/**
 * ==================================================================
 * CONFIGURACIÓN DE TABLAS
 * ==================================================================
 */

const MAIN_SPREADSHEET_ID        = '1m1XnSzqlTqvKgvn9yLK8CKw_n1MmnUgCVKhrFrEJe4Q'; 
const PERMISSIONS_SPREADSHEET_ID = '18Ur6__aI2xjDjpA_yl-oqMz7dNh1h8Ds1lhYFA7LB14'; 
const CLIENTS_SPREADSHEET_ID     = '1GleT24bySpMAPd5lgRPvS0zQ4FsI8u04GGaTT05jLSo';
const SEDES_SPREADSHEET_ID       = '1GEXwm13dOw8ExF_DxJ7BaL7nqdKP7rB35bH0dkEdktU'; 
const ACTIVOS_SPREADSHEET_ID     = '1-n3WCFErmS7aOPR3l4VHmKdGh0EtvodyhEuahvWPUwg'; 

// HOJAS MIGRADAS DE BIGQUERY A GOOGLE SHEETS
const PISOS_SPREADSHEET_ID       = '1WWA6xqMZtKHrTvacqP6KUlZLjfPoEVnZdfYpudR8DRw'; 
const DISPOSITIVOS_SPREADSHEET_ID = '1r07LqMwpJyKQ2DvT7f9elteay4r80MBJl88NSiVvV2o';

// Mapeo 
const SHEET_CONFIG = {
  'Solicitudes': MAIN_SPREADSHEET_ID,
  'Estados historico': MAIN_SPREADSHEET_ID,
  'Observaciones historico': MAIN_SPREADSHEET_ID,
  'Estados': MAIN_SPREADSHEET_ID,
  'Solicitudes anexos': MAIN_SPREADSHEET_ID,
  'Permisos': PERMISSIONS_SPREADSHEET_ID,
  'Usuarios filtro': PERMISSIONS_SPREADSHEET_ID,
  'Clientes': CLIENTS_SPREADSHEET_ID,
  'Sedes': SEDES_SPREADSHEET_ID,
  'Solicitudes activos': MAIN_SPREADSHEET_ID,
  'Activos': ACTIVOS_SPREADSHEET_ID,
  'Pisos': PISOS_SPREADSHEET_ID,
  'Dispositivos': DISPOSITIVOS_SPREADSHEET_ID
};

/**
 * ------------------------------------------------------------------
 * OPTIMIZACIÓN: Apertura de Hojas de cálculo por ID o URL
 * ------------------------------------------------------------------
 */
const __SS_MEMO = {};
function _openSS(idOrUrl) {
  if (!__SS_MEMO[idOrUrl]) {
    if (String(idOrUrl).includes("docs.google.com/spreadsheets")) {
      __SS_MEMO[idOrUrl] = SpreadsheetApp.openByUrl(idOrUrl);
    } else {
      __SS_MEMO[idOrUrl] = SpreadsheetApp.openById(idOrUrl);
    }
  }
  return __SS_MEMO[idOrUrl];
}

/**
 * Helper concurrencia
 */
function _withLock(callback) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
    const result = callback();
    SpreadsheetApp.flush(); 
    return result;
  } catch (e) {
    console.error("Error de Lock/Concurrencia:", e);
    throw new Error("El servidor está ocupado. Intente de nuevo en unos segundos.");
  } finally {
    lock.releaseLock();
  }
}

const __SHEET_INFO_MEMO = {};
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

function _findRowObjectByKey(sheetName, keyValue, colCandidates) {
  const { sheet, headers, headersNorm, lastRow, lastCol } = _getSheetInfo(sheetName);
  if (!sheet || lastRow < 2) return null;

  const k = String(keyValue || "").trim();
  if (!k) return null;
  
  for (let c = 0; c < colCandidates.length; c++) {
    const idx = headersNorm.indexOf(_normHeader(colCandidates[c]));
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

function _getChildrenFast(sheetName, parentKeys) {
  const { sheet, headers, headersNorm, lastRow, lastCol } = _getSheetInfo(sheetName);
  if (!sheet || lastRow < 2) return [];

  const fkCandidates = ['ID Solicitudes', 'ID Solicitud', 'ID Solicitudes '];
  const idxFk = _findColIndex(headersNorm, fkCandidates);
  if (idxFk === -1) return [];
  const idxDate = _findColIndex(headersNorm, ['Fecha Actualización', 'Fecha Actualizacion', 'Fecha', 'FechaCambio']);
  const set = new Set((parentKeys || []).map(x => String(x || "").trim()).filter(Boolean));
  if (!set.size) return [];
  const fkVals = sheet.getRange(2, idxFk + 1, lastRow - 1, 1).getValues();
  const rowNums = [];
  for (let i = 0; i < fkVals.length; i++) {
    const fk = String(fkVals[i][0]).trim();
    if (fk && set.has(fk)) rowNums.push(i + 2);
  }
  if (!rowNums.length) return [];

  rowNums.sort((a, b) => a - b);
  const runs = _mergeRowRuns(rowNums);
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
  out.sort((a, b) => (b.__ts || 0) - (a.__ts || 0));
  out.forEach(o => delete o.__ts);
  return out;
}

function doGet(e) {
  if (e.parameter && e.parameter.v === 'archivo' && e.parameter.id) {
    return _renderFileView(e.parameter.id);
  }
  const template = HtmlService.createTemplateFromFile('Index');
  template.scriptUrl = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .setTitle('G4S Ticket Tracker')
    .setFaviconUrl('https://www.g4s.com/favicon.ico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _resolveCallerEmail(request) {
  const active = Session.getActiveUser().getEmail();
  if (active) return String(active).toLowerCase().trim();
  const p = request?.payload || {};
  const fromClient = p.__clientEmail || p.clientEmail || request?.clientEmail || "";
  const email = String(fromClient).toLowerCase().trim();
  if (email && email.includes("@")) return email;
  return "";
}

function getData(action, payload) {
  return apiHandler({ endpoint: 'getAssetsData', payload: { action: action, payload: payload } });
}

function apiHandler(request) {
  const userEmail = _resolveCallerEmail(request);
  const { endpoint, payload } = request || {};
  console.log(`🔒 [API CHECK] Endpoint: ${endpoint} | ActiveUser: ${Session.getActiveUser().getEmail()} | Resuelto: ${userEmail}`);
  try {
    if (!userEmail) throw new Error("No se pudo verificar la identidad del usuario.");
    switch (endpoint) {
      case 'getUserContext': return getUserContext(userEmail, payload?.ignoreCache);
      case 'getRequests': return getRequests(userEmail);
      case 'getRequestDetail': return getRequestDetail(userEmail, payload);
      case 'createRequest': return createRequest(userEmail, payload);
      case 'uploadAnexo': return uploadAnexo(userEmail, payload);
      case 'createSolicitudActivo': return createSolicitudActivo(userEmail, payload);
      case 'getAnexoDownload': return getAnexoDownload(userEmail, payload);
      case 'getSolicitudActivos': return getSolicitudActivos(userEmail, payload);
      case 'getActivosCatalog': return getActivosCatalog(userEmail);
      case 'getActivoByQr': return getActivoByQr(userEmail, payload);
      case 'getClassificationOptions': return getClassificationOptions(userEmail);
      case 'getBatchRequestDetails': return getBatchRequestDetails(userEmail, payload);
      case 'getAssetsData': return getAssetsData(userEmail, payload);
      default: throw new Error(`Endpoint desconocido: ${endpoint}`);
    }
  } catch (err) {
    console.error(`❌ ERROR DE SEGURIDAD/EJECUCIÓN: ${err.message}`, err);
    return { error: true, message: "Error procesando su solicitud. Contacte al administrador." };
  }
}

const DETAIL_CACHE_VER = "v3";
function _detailCacheKey(email, id) {
  const e = String(email || "").toLowerCase().trim();
  const rid = String(id || "").trim();
  return `detail_${DETAIL_CACHE_VER}_${Utilities.base64Encode(e)}_${rid}`;
}
function _invalidateDetailCache(email, id) {
  try { CacheService.getScriptCache().remove(_detailCacheKey(email, id)); } catch (e) {}
}

function _sanitizeFileName(name) {
  const n = String(name || 'anexo').trim();
  return n.normalize("NFD").replace(/[\u0300-\u036f]/g, "") 
          .replace(/[/\\]/g, '_')
          .replace(/[<>:"|?*]/g, '')
          .replace(/\s+/g, '_') 
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

  const IMAGES_FOLDER_ID = '1tzYk9jiQ7Lp_bSZylMn0vuzfWZt4xTHb'; 
  const DOCS_FOLDER_ID   = '1-CBsinL67dJUPfr8WXtKP1wM93B6zPDX';

  const solicitudId = String(payload?.solicitudId || '').trim();
  const tipoAnexo = String(payload?.tipoAnexo || 'Archivo').trim();
  const fileNameInput = String(payload?.fileName || 'anexo');
  const mimeType = String(payload?.mimeType || 'application/octet-stream').trim();
  const base64 = String(payload?.base64 || '').trim();

  if (!solicitudId) throw new Error("solicitudId requerido");
  if (!base64) throw new Error("base64 requerido");
  
  const headerFound = _findRowObjectByKey('Solicitudes', solicitudId, ['ID Solicitud', 'ID Solicitudes']);
  if (!headerFound) throw new Error("Solicitud padre no encontrada.");
  
  const maxBytes = 10 * 1024 * 1024;
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > maxBytes) throw new Error("Archivo demasiado grande (máx 10MB).");
  
  const safeFileName = _sanitizeFileName(fileNameInput).replace(/\s+/g, '_'); 
  const shortId = solicitudId.replace(/-/g, '').slice(0, 8);
  const rand = Math.floor(Math.random() * 900000) + 100000;
  
  const extMatch = safeFileName.match(/\.([0-9a-z]+)$/i);
  const ext = extMatch ? extMatch[1] : (mimeType.includes('image') ? 'jpg' : 'pdf');
  const baseName = safeFileName.replace(/\.[^/.]+$/, "").replace(/\./g, "_");
  const finalName = `${shortId}_${tipoAnexo}_${rand}_${baseName}.${ext}`;
  const blob = Utilities.newBlob(bytes, mimeType, finalName);

  return _withLock(() => {
    let file;
    let storedPath;
    if (tipoAnexo === 'Foto' || tipoAnexo === 'Dibujo' || mimeType.startsWith('image/')) {
        try {
          const targetFolder = DriveApp.getFolderById(IMAGES_FOLDER_ID);
          file = targetFolder.createFile(blob);
          storedPath = `${targetFolder.getName()}/${finalName}`;
        } catch (e) {
          throw new Error("No se pudo acceder a la carpeta de Imágenes definida. Verifique el ID.");
        }
    } else {
        try {
          const targetFolder = DriveApp.getFolderById(DOCS_FOLDER_ID);
          file = targetFolder.createFile(blob);
          storedPath = `https://drive.google.com/file/d/${file.getId()}/view`;
        } catch (e) {
          throw new Error("No se pudo acceder a la carpeta de Documentos definida. Verifique el ID.");
        }
    }

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch(e2) {}
    }

    const anexoUuid = Utilities.getUuid();
    const now = new Date();
    const row = {
      "ID Solicitudes anexos": anexoUuid,
      "ID Solicitudes": solicitudId,
      "Tipo anexo": tipoAnexo,
      "Nombre": safeFileName,
      "Usuario Actualización": email,
      "Fecha Actualización": now
    };
    if (tipoAnexo === 'Foto') { row['Foto'] = storedPath; } 
    else if (tipoAnexo === 'Dibujo') { row['Dibujo'] = storedPath; } 
    else { row['Archivo'] = storedPath; }

    appendDataToSheet('Solicitudes anexos', row);
    _invalidateDetailCache(email, solicitudId);
    return { success: true, anexoId: anexoUuid, fileName: file.getName(), path: storedPath };
  });
}

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
  if (lastCol === 0) throw new Error(`La hoja ${sheetName} está vacía.`);

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowArray = headers.map(header => {
    let val = objectData[header];
    if (val === undefined) {
       const cleanHeader = String(header).trim().toLowerCase();
       const foundKey = Object.keys(objectData).find(k => String(k).trim().toLowerCase() === cleanHeader);
       if (foundKey) val = objectData[foundKey];
    }
    return val === undefined || val === null ? "" : val;
  });
  sheet.appendRow(rowArray);
  
  const possibleTicketId = objectData["ID Solicitud"] || objectData["ID Solicitudes"] || "";
  if (possibleTicketId) _invalidateDetailCache(String(objectData["Usuario Actualización"] || ""), possibleTicketId);
  return { success: true };
}

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
  const p = String(pathValue || "").trim();
  if (/file\/d\/([^/]+)/.test(p)) { return { kind: "url", url: p }; }
  if (/id=([^&]+)/.test(p)) {
     const id = p.match(/id=([^&]+)/)[1];
     return { kind: "url", url: `https://drive.google.com/file/d/${id}/view` };
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
  } catch (e) {}

  const safeName = filename.replace(/"/g, '\\"');
  const q = `name = "${safeName}" and trashed = false`;
  const it2 = DriveApp.searchFiles(q);
  if (it2.hasNext()) return { kind: "file", file: it2.next() };
  throw new Error(`Archivo no encontrado: ${filename}`);
}

function _findSolicitudHeaderFast(key) {
  return _findRowObjectByKey('Solicitudes', key, [
    'ID Solicitud', 'ID Solicitudes', 'Ticket G4S', 'Ticket Cliente', 'Ticket (Opcional)'
  ]);
}

function getUserContext(email, ignoreCache = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `ctx_it_v6_${Utilities.base64Encode(email)}`;
  if (!ignoreCache) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData);
  }

  try {
    let context = {
      email: email,
      role: 'Usuario',
      allowedClientIds: [],
      allowedCustomerIds: [],
      clientNames: {},
      assignedCustomerNames: [],
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
    context.allowedCustomerIds = assignedClientIds;

    if (assignedClientIds.length > 0) {
      const allClientes = getDataFromSheet('Clientes');
      const myClients = allClientes.filter(c => 
        assignedClientIds.includes(String(_getField(c, ['ID Cliente', 'Id Cliente', 'Cliente'])))
      );
      myClients.forEach(c => {
        const clientName = _getField(c, ['Nombre cliente', 'Nombre Cliente', 'Nombre', 'RazonSocial']);
        if (clientName) context.assignedCustomerNames.push(String(clientName).trim());
      });
      const allSedes = getDataFromSheet('Sedes');
      const mySedes = allSedes.filter(sede => assignedClientIds.includes(String(_getField(sede, ['ID Cliente', 'Id Cliente', 'Cliente']))));
      mySedes.forEach(sede => {
        const idSede = String(_getField(sede, ['ID Sede', 'Id Sede', 'Sede', 'IDSede'])).trim();
        const nombreSede = _getField(sede, ['Nombre', 'Nombre_Sede', 'Nombre sede', 'Nombre Sede', 'Sede', 'Label']) || idSede;
        if (idSede) {
          context.allowedClientIds.push(idSede);
          context.clientNames[idSede] = String(nombreSede).trim() || idSede;
        }
      });
    }
    cache.put(cacheKey, JSON.stringify(context), 350);
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
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado");
  if (!id) throw new Error("ID requerido");

  const rid = String(id).trim();
  const cache = CacheService.getScriptCache();
  const ck = _detailCacheKey(email, rid);
  const cached = cache.get(ck);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

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
  const services = _getChildrenFast('Observaciones historico', parentKeys);
  const history = _getChildrenFast('Estados historico', parentKeys);
  const documents = _getChildrenFast('Solicitudes anexos', parentKeys);
  
  // Buscar activos vinculados usando el método nativo optimizado
  const activos = _getChildrenFast('Solicitudes activos', parentKeys);
  
  const result = { header, services, history, documents, activos };
  const json = JSON.stringify(result);
  if (json.length < 90000) cache.put(ck, json, 30);
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

  return _withLock(() => {
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
      "Estado": "Creado",
      "ID Sede": String(payload.idSede).trim(),
      "Ticket Cliente": payload.ticketCliente || "",
      "Clasificación Solicitud": payload.clasificacion, 
      "Clasificación": payload.tipoServicio,
      "Técnicos Clientes": "Por disponibilidad", 
      "Prioridad Solicitud": payload.prioridad,
      "Solicitud": payload.solicitud,
      "Observación": payload.observacion,
      "Usuario Actualización": email
    };
    appendDataToSheet('Solicitudes', newRow);
    SpreadsheetApp.flush();

    try { enviarAppSheetAPI('Solicitudes', newRow); } catch (e) { console.warn("AppSheet Sync advertencia:", e); }

    try {
      const historyRow = {
        "ID Estado": Utilities.getUuid(),
        "ID Solicitudes": uuid,
        "Estado actual": "Creado",
        "Usuario Actualización": email,
        "Fecha Actualización": now
      };
      appendDataToSheet('Estados historico', historyRow);
    } catch (e) { console.warn("Historial falló:", e); }

    _invalidateDetailCache(email, uuid);
    const returnRow = {
      ...newRow,
      "Fecha creación cliente": (newRow["Fecha creación cliente"] instanceof Date) ? newRow["Fecha creación cliente"].toISOString() : newRow["Fecha creación cliente"]
    };
    return { success: true, solicitudId: uuid, ticketG4S: ticketG4S, GeneratedTicket: ticketG4S, Status: "Success", Rows: [returnRow], row: returnRow };
  });
}

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
  if (pathValue.includes("drive.google.com")) {
     return { mode: "url", url: pathValue, fileName: _getField(row, ['Nombre']) };
  }
  const resolved = _resolveDriveFileFromAppSheetPath(pathValue);
  if (resolved.kind === "url") {
    return { mode: "url", fileName: _getField(row, ['Nombre']) || "Anexo", url: resolved.url };
  }
  const file = resolved.file;
  return { mode: "url", url: `https://drive.google.com/file/d/${file.getId()}/view`, fileName: file.getName() };
}

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

  return _withLock(() => {
    let dibujoPath = "";
    if (dibujoBase64) {
      const bytes = Utilities.base64Decode(dibujoBase64);
      const root = _getRootFolderForFiles();
      const folder = _ensurePathFromRoot(root, ['Info', 'Clientes', 'Activos']);
      const short = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
      const rand = Math.floor(Math.random() * 900000) + 100000;
      const fileName = `${short}.Dibujo.${rand}.png`;
      const blob = Utilities.newBlob(bytes, 'image/png', fileName);
      const file = folder.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
      dibujoPath = `https://drive.google.com/uc?export=view&id=${file.getId()}`;
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
  });
}

function getSolicitudActivos(email, { solicitudId }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");
  const sid = String(solicitudId || '').trim();
  if (!sid) throw new Error("solicitudId requerido");
  const rows = _getChildrenFast('Solicitudes activos', [sid]);
  return { data: rows, total: rows.length };
}

function getActivosCatalog(email) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");
  const cache = CacheService.getScriptCache();
  const key = "activos_catalog_v2";
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  const rows = getDataFromSheet('Activos');
  const mapped = rows.map(r => {
    return {
      idActivo: String(_getField(r, ['ID Activo'])).trim(),
      nombreActivo: String(_getField(r, ['Nombre Activo'])).trim(),
      qrSerial: String(_getField(r, ['QR Serial'])).trim(),
      nombreUbicacion: String(_getField(r, ['Nombre Ubicacion'])).trim(),
      estadoActivo: String(_getField(r, ['Estado Activo'])).trim(),
      funcionamiento: String(_getField(r, ['Funcionamiento'])).trim()
    };
  }).filter(x => x.idActivo || x.qrSerial);
  const res = { data: mapped, total: mapped.length };
  cache.put(key, JSON.stringify(res), 600);
  return res;
}

function getActivoByQr(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");
  const q = String(payload?.qr || '').trim();
  if (!q) throw new Error("qr requerido");
  const rows = getDataFromSheet('Activos');
  const found = rows.find(r => String(_getField(r, ['QR Serial', 'QR', 'Qr', 'Codigo QR'])).trim() === q);
  if (!found) return { found: false };
  return {
    found: true,
    activo: {
      idActivo: String(_getField(found, ['ID Activo'])).trim(),
      nombreActivo: String(_getField(found, ['Nombre Activo'])).trim(),
      qrSerial: q,
      nombreUbicacion: String(_getField(found, ['Nombre Ubicacion'])).trim(),
      estadoActivo: String(_getField(found, ['Estado Activo'])).trim(),
      funcionamiento: String(_getField(found, ['Funcionamiento'])).trim()
    }
  };
}

function getBatchRequestDetails(email, { ids }) {
  const t0 = Date.now();
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado");
  if (!ids || !Array.isArray(ids) || ids.length === 0) return {};
  const targetIds = new Set(ids.map(x => String(x).trim()));
  const allServices = getDataFromSheet('Observaciones historico');
  const allHistory = getDataFromSheet('Estados historico');
  const allDocs = getDataFromSheet('Solicitudes anexos');
  const allActivos = getDataFromSheet('Solicitudes activos');

  const result = {};
  targetIds.forEach(id => { result[id] = { services: [], history: [], documents: [], activos: [] }; });
  
  const findParentIdInRow = (row) => {
    if (!row) return "";
    const candidates = ['idsolicitud', 'idsolicitudes', 'ticketg4s', 'ticketcliente'];
    const keys = Object.keys(row);
    for (const key of keys) {
      const cleanKey = String(key).toLowerCase().replace(/[^a-z0-9]/g, '');
      if (candidates.includes(cleanKey)) {
        const val = row[key];
        if (val !== undefined && val !== null && val !== "") return String(val).trim();
      }
    }
    return "";
  };

  const groupByParentSmart = (rows, targetSet, targetKeyInResult) => {
    rows.forEach(row => {
      const parentId = findParentIdInRow(row);
      if (parentId && targetSet.has(parentId)) {
        if (!result[parentId][targetKeyInResult]) result[parentId][targetKeyInResult] = [];
        result[parentId][targetKeyInResult].push(row);
      }
    });
  };

  groupByParentSmart(allServices, targetIds, 'services');
  groupByParentSmart(allHistory, targetIds, 'history');
  groupByParentSmart(allDocs, targetIds, 'documents');
  groupByParentSmart(allActivos, targetIds, 'activos');
  console.log(`⚡ [BATCH SMART] Procesados ${ids.length} tickets. Tiempo: ${Date.now() - t0}ms`);
  return result;
}

function getClassificationOptions(email) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");
  return ["Visita técnica", "Visita comercial"];
}

function _renderFileView(anexoId) {
  try {
    const found = _findRowObjectByKey('Solicitudes anexos', anexoId, [
      'ID Solicitudes anexos', 'ID Solicitud anexos', 'ID Anexo', 'ID'
    ]);
    if (!found) return HtmlService.createHtmlOutput("<h1>Archivo no encontrado en la base de datos.</h1>").setFaviconUrl('https://www.g4s.com/favicon.ico');
    const row = found.obj;
    const pathValue = _getField(row, ['Archivo', 'Archivo ', 'Foto', 'Dibujo', 'QR']) || "";
    const fileName = _getField(row, ['Nombre']) || "Archivo_G4S";
    let file = null;

    if (pathValue.includes("drive.google.com") || pathValue.includes("/d/")) {
        const idMatch = pathValue.match(/\/d\/([a-zA-Z0-9_-]+)/) || pathValue.match(/id=([a-zA-Z0-9_-]+)/);
        if (idMatch && idMatch[1]) {
            try { file = DriveApp.getFileById(idMatch[1]); } catch(e) {}
        }
    } else {
        const parts = pathValue.split('/');
        const exactFileName = parts[parts.length - 1]; 
        if (exactFileName) {
            const filesIt = DriveApp.getFilesByName(exactFileName);
            if (filesIt.hasNext()) { file = filesIt.next(); }
        }
    }

    if (!file) {
       return HtmlService.createHtmlOutput(`
         <div style='font-family:sans-serif;text-align:center;padding:40px;'>
           <h1>Archivo no encontrado en Drive</h1>
           <p>No se pudo localizar el archivo físico: <b>${fileName}</b></p>
         </div>
       `).setFaviconUrl('https://www.g4s.com/favicon.ico');
    }

    if (file.getSize() > 8 * 1024 * 1024) { 
      return HtmlService.createHtmlOutput(`
        <div style="font-family:sans-serif;text-align:center;margin-top:50px;">
          <h2>Archivo Grande</h2>
          <a href="https://drive.google.com/uc?export=download&id=${file.getId()}" style="background:#0033A0;color:white;padding:15px;text-decoration:none;">Descargar</a>
        </div>
      `).setFaviconUrl('https://www.g4s.com/favicon.ico');
    }

    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>G4S - ${fileName}</title>
        <style>
          body { margin: 0; padding: 0; background-color: #f3f4f6; height: 100vh; display: flex; align-items: center; justify-content: center; font-family: sans-serif; }
          .card { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center; }
          .btn { background: #D32F2F; color: white; padding: 12px 20px; text-decoration: none; border-radius: 6px; font-weight: bold; cursor: pointer; border: none; }
          .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #D32F2F; border-radius: 50%; width: 24px; height: 24px; animation: spin 1s linear infinite; margin: 15px auto; }
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        </style>
      </head>
      <body>
        <div class="card">
          <div id="loader"><div class="spinner"></div><h3>Procesando...</h3></div>
          <div id="content" style="display:none;">
            <h3>Listo</h3>
            <p>${fileName}</p>
            <button id="dlBtn" class="btn">Guardar Archivo</button>
          </div>
        </div>
        <script>
          window.onload = function() {
            const rawBase64 = "${base64}";
            const byteCharacters = atob(rawBase64);
            const byteNumbers = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) { byteNumbers[i] = byteCharacters.charCodeAt(i); }
            const blob = new Blob([new Uint8Array(byteNumbers)], {type: "${mimeType}"});
            const url = URL.createObjectURL(blob);
            const btn = document.getElementById('dlBtn');
            btn.onclick = function() {
              const a = document.createElement('a');
              a.href = url; a.download = "${fileName}"; a.click();
            };
            document.getElementById('loader').style.display = 'none';
            document.getElementById('content').style.display = 'block';
            setTimeout(() => btn.click(), 800);
          };
        </script>
      </body>
      </html>
    `;
    return HtmlService.createHtmlOutput(html).setFaviconUrl('https://www.g4s.com/favicon.ico');
  } catch (e) {
    return HtmlService.createHtmlOutput(`<h3>Error de Sistema: ${e.message}</h3>`).setFaviconUrl('https://www.g4s.com/favicon.ico');
  }
}

function enviarAppSheetAPI(tableName, rowData) {
  const appId = "c0817cfb-b068-4a46-ae3b-228c0385a486";
  const accessKey = "V2-gaw9Q-LcMsx-wfJof-pFCgC-u6igd-FMxtR-23Zr1-V3O4K"; 
  const url = `https://api.appsheet.com/api/v1/apps/${appId}/tables/${tableName}/Action`;
  const payload = {
    "Action": "Add",
    "Properties": { "Locale": "es-CO", "Timezone": "SA Pacific Standard Time", "RunAsUserEmail": rowData["Usuario Actualización"] },
    "Rows": [ rowData ]
  };
  const options = {
    "method": "post", "contentType": "application/json",
    "headers": { "ApplicationAccessKey": accessKey },
    "payload": JSON.stringify(payload), "muteHttpExceptions": true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    console.error("Error en la API de AppSheet: " + e);
    return null;
  }
}

/**
 * ------------------------------------------------------------------
 * REFACTORIZACIÓN INTEGRAL: CONTROL DE ACTIVOS POR MEMORIA DE GOOGLE SHEETS
 * (Remplaza completamente a BigQuery emulando los conteos y joins en JS)
 * ------------------------------------------------------------------
 */
function getAssetsData(email, { action, payload = {} }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  try {
    switch (action) {
      case 'getClients': {
        const clients = getDataFromSheet('Clientes');
        const sedes = getDataFromSheet('Sedes');
        const pisos = getDataFromSheet('Pisos');
        const activos = getDataFromSheet('Activos');

        let filteredClients = clients;
        if (!context.isAdmin) {
          if (!context.assignedCustomerNames || context.assignedCustomerNames.length === 0) return [];
          const names = context.assignedCustomerNames.map(n => String(n).trim().toUpperCase());
          filteredClients = clients.filter(c => {
            const name = String(_getField(c, ['Nombre cliente', 'Nombre Cliente', 'Nombre', 'RazonSocial'])).trim().toUpperCase();
            return names.includes(name);
          });
        }

        return filteredClients.map(c => {
          const clientId = String(_getField(c, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim();
          const clientName = String(_getField(c, ['Nombre cliente', 'Nombre Cliente', 'Nombre', 'RazonSocial'])).trim();
          
          const connectedSedes = sedes.filter(s => String(_getField(s, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim() === clientId)
                                      .map(s => String(_getField(s, ['ID Sede', 'Id Sede', 'Sede', 'IDSede'])).trim());
          const connectedPisos = pisos.filter(p => connectedSedes.includes(String(_getField(p, ['ID Sede', 'Id Sede', 'Sede'])).trim()))
                                      .map(p => String(_getField(p, ['ID Piso', 'Id Piso', 'Piso'])).trim());
          const totalActivos = activos.filter(a => connectedPisos.includes(String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim())).length;

          return { id_cliente: clientId, nombre_cliente: clientName, total_activos: totalActivos };
        }).sort((a, b) => a.nombre_cliente.localeCompare(b.nombre_cliente));
      }
      
      case 'getSites': {
        if (!payload.clientId) throw new Error("clientId es requerido.");
        const sedes = getDataFromSheet('Sedes');
        const pisos = getDataFromSheet('Pisos');
        const activos = getDataFromSheet('Activos');

        const targetSedes = sedes.filter(s => String(_getField(s, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim() === String(payload.clientId).trim());

        return targetSedes.map(s => {
          const siteId = String(_getField(s, ['ID Sede', 'Id Sede', 'Sede', 'IDSede'])).trim();
          const siteName = String(_getField(s, ['Nombre', 'Nombre_Sede', 'Nombre sede', 'Nombre Sede', 'Sede', 'Label']) || siteId).trim();

          const connectedPisos = pisos.filter(p => String(_getField(p, ['ID Sede', 'Id Sede', 'Sede'])).trim() === siteId)
                                      .map(p => String(_getField(p, ['ID Piso', 'Id Piso', 'Piso'])).trim());
          const totalActivos = activos.filter(a => connectedPisos.includes(String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim())).length;

          return { id_sede: siteId, nombre_sede: siteName, total_activos: totalActivos };
        }).sort((a, b) => a.nombre_sede.localeCompare(b.nombre_sede));
      }
      
      case 'getFloors': {
        if (!payload.siteId) throw new Error("siteId es requerido.");
        const pisos = getDataFromSheet('Pisos');
        const activos = getDataFromSheet('Activos');

        const targetPisos = pisos.filter(p => String(_getField(p, ['ID Sede', 'Id Sede', 'Sede'])).trim() === String(payload.siteId).trim());

        return targetPisos.map(p => {
          const floorId = String(_getField(p, ['ID Piso', 'Id Piso', 'Piso'])).trim();
          const floorName = String(_getField(p, ['Nombre Piso', 'Nombre piso', 'Nombre'])).trim();
          const nivel = _getField(p, ['Nivel', 'nivel']);
          const planoUrl = _getField(p, ['Imagen Plano URL', 'imagen_plano_url', 'Plano', 'Imagen']);
          const totalActivos = activos.filter(a => String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim() === floorId).length;

          return { id_piso: floorId, nombre_piso: floorName, nivel: nivel, imagen_plano_url: planoUrl, total_activos: totalActivos };
        }).sort((a, b) => a.nombre_piso.localeCompare(b.nombre_piso));
      }
      
      case 'getAssets': {
        if (!payload.floorId) throw new Error("floorId es requerido.");
        const activos = getDataFromSheet('Activos');
        let dispositivos = [];
        try { dispositivos = getDataFromSheet('Dispositivos'); } catch(e) { console.warn("Hoja auxiliar de dispositivos no cargada."); }

        const targetActivos = activos.filter(a => String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim() === String(payload.floorId).trim());

        return targetActivos.map(a => {
          const idDispositivo = String(_getField(a, ['ID Dispositivo', 'Id Dispositivo', 'id_dispositivo'])).trim();
          let tipoDispositivo = _getField(a, ['Tipo Dispositivo', 'Tipo dispositivo', 'tipo_dispositivo', 'Tipo']);
          
          if (!tipoDispositivo && dispositivos.length > 0) {
            const dispInfo = dispositivos.find(d => String(_getField(d, ['ID Dispositivo', 'Id Dispositivo'])).trim() === idDispositivo);
            if (dispInfo) tipoDispositivo = _getField(dispInfo, ['Clasificación', 'Clasificacion', 'clasificacion']);
          }
          if (!tipoDispositivo) tipoDispositivo = idDispositivo || "General";

          return {
            id_activo: String(_getField(a, ['ID Activo', 'Id Activo', 'id_activo'])).trim(),
            nombre_activo: String(_getField(a, ['Nombre Activo', 'Nombre activo', 'nombre_activo'])).trim(),
            tipo_dispositivo: tipoDispositivo,
            estado_activo: String(_getField(a, ['Estado Activo', 'Estado activo', 'estado_activo'])).trim(),
            coord_x: _getField(a, ['Coord X', 'coord_x', 'X']),
            coord_y: _getField(a, ['Coord Y', 'coord_y', 'Y']),
            fecha_actualizacion: _getField(a, ['Fecha Actualización', 'Fecha Actualizacion', 'fecha_actualizacion', 'Fecha']),
            foto_1: _getField(a, ['Foto 1', 'foto_1', 'Foto']),
            foto_2: _getField(a, ['Foto 2', 'foto_2']),
            foto_3: _getField(a, ['Foto 3', 'foto_3']),
            specs: _getField(a, ['Specs', 'specs', 'Datos Tecnicos', 'datos_tecnicos_json']),
            protocol: _getField(a, ['Protocol', 'protocol', 'Ultimo Protocolo', 'ultimo_protocolo_json'])
          };
        });
      }
      default: return [];
    }
  } catch (e) { 
    console.error("Error en getAssetsData", e);
    throw new Error("Error procesando inventario desde Hojas de Cálculo: " + e.message); 
  }
}
