/**
 * ------------------------------------------------------------------
 * CONFIGURACIÓN Y MAPEO DE HOJAS DE CÁLCULO (GOOGLE SHEETS)
 * ------------------------------------------------------------------
 * Coloque aquí los ID o las URL completas (p. ej. 'https://docs.google.com/spreadsheets/d/ID/edit')
 * de cada una de las hojas de Google Sheets necesarias.
 */

// Hoja Principal: Contiene Solicitudes, Solicitudes anexos, Estados historico, Observaciones historico, Solicitudes activos
const MAIN_SPREADSHEET_ID        = '1MC76eZZt7qiso2M8LMz777_xJnzrl_ZpZptDZBnPlDo';

// Hoja de Permisos: Contiene las pestañas Permisos y Usuarios filtro
const PERMISSIONS_SPREADSHEET_ID = '1zcZZGe_93ytWXtCF1kmk_Y8zc5b5cL1xH34i7v1w01k'; 

// Hoja de Clientes: Contiene la pestaña Clientes
const CLIENTS_SPREADSHEET_ID     = '1hHWPJF9KSC0opplpCNgHRNkW6CLf7StXG2Y31m6yUpo';

// Hoja de Sedes: Contiene la pestaña Sedes
const SEDES_SPREADSHEET_ID       = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM';

// Hoja de Pisos: Contiene la pestaña Pisos (puede ser el mismo ID de Sedes si están en el mismo archivo)
const PISOS_SPREADSHEET_ID       = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM';

// Hoja de Activos: Contiene la pestaña Activos
const ACTIVOS_SPREADSHEET_ID     = '1JU8c1MidgV4DRFg6W-GxZ2tHkfKNqGt1_cR5VDTehC4';

// Hoja de Dispositivos (Opcional): Contiene catálogo de dispositivos
const DISPOSITIVOS_SPREADSHEET_ID = '1JU8c1MidgV4DRFg6W-GxZ2tHkfKNqGt1_cR5VDTehC4';

// Configuración para saber en qué Spreadsheet buscar cada tabla
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
  'Pisos': PISOS_SPREADSHEET_ID,
  'Solicitudes activos': MAIN_SPREADSHEET_ID,
  'Activos': ACTIVOS_SPREADSHEET_ID,
  'Dispositivos': DISPOSITIVOS_SPREADSHEET_ID,
};

/**
 * Extrae el ID del Spreadsheet a partir de un ID puro o de una URL completa de Google Sheets.
 * @param {string} idOrUrl ID o URL completa de Google Sheets.
 * @returns {string} ID formateado del Spreadsheet.
 */
function _extractSpreadsheetId(idOrUrl) {
  if (!idOrUrl) return '';
  const str = String(idOrUrl).trim();
  const match = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];
  return str;
}

/**
 * ------------------------------------------------------------------
 * OPTIMIZACIÓN: Reuso de Spreadsheet abiertos (memoria por ejecución)
 * ------------------------------------------------------------------
 */
const __SS_MEMO = {}; 
function _openSS(spreadsheetIdOrUrl) {
  const spreadsheetId = _extractSpreadsheetId(spreadsheetIdOrUrl);
  if (!spreadsheetId) throw new Error("ID o URL de Spreadsheet no válido.");
  if (!__SS_MEMO[spreadsheetId]) __SS_MEMO[spreadsheetId] = SpreadsheetApp.openById(spreadsheetId);
  return __SS_MEMO[spreadsheetId];
}

/**
 * ------------------------------------------------------------------
 * HELPER DE SEGURIDAD Y CONCURRENCIA
 * ------------------------------------------------------------------
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

/**
 * OPTIMIZACIÓN: memo de info de hoja por ejecución
 */
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

/**
 * ------------------------------------------------------------------
 * ROUTER INTELIGENTE
 * ------------------------------------------------------------------
 */
function doGet(e) {
  // ✅ ROUTER DE ARCHIVOS (MODO PROXY)
  if (e.parameter && e.parameter.v === 'archivo' && e.parameter.id) {
    return _renderFileView(e.parameter.id);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  // ✅ IMPORTANTE: Inyectamos la URL del script para el frontend
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

/**
 * Entry point para compatibilidad con código proporcionado por el usuario.
 * @param {string} action Acción a realizar.
 * @param {Object} payload Parámetros de la acción.
 * @returns {any} Resultado del apiHandler.
 */
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
      case 'getAnexoFileBase64': return getAnexoFileBase64(userEmail, payload);
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

// CACHE
const DETAIL_CACHE_VER = "v3"; 
function _detailCacheKey(email, id) {
  const e = String(email || "").toLowerCase().trim();
  const rid = String(id || "").trim();
  return `detail_${DETAIL_CACHE_VER}_${Utilities.base64Encode(e)}_${rid}`;
}
function _invalidateDetailCache(email, id) {
  try { CacheService.getScriptCache().remove(_detailCacheKey(email, id)); } catch (e) {}
}

// ------------------------------------------------------------------
// ✅ FUNCIÓN DE SANITIZACIÓN ROBUSTA (Sin espacios ni caracteres raros)
// ------------------------------------------------------------------
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

// ------------------------------------------------------------------
// ✅ FUNCIÓN UPLOAD V4: CARPETAS FIJAS POR TIPO (IDS ESPECÍFICOS)
// ------------------------------------------------------------------
function uploadAnexo(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  // --- IDs DE CARPETAS PROPORCIONADOS ---
  const IMAGES_FOLDER_ID = '1tzYk9jiQ7Lp_bSZylMn0vuzfWZt4xTHb'; // Carpeta para Fotos/Dibujos
  const DOCS_FOLDER_ID   = '1-CBsinL67dJUPfr8WXtKP1wM93B6zPDX'; // Carpeta para Documentos
  // ----------------------------------------

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

  // Limpieza de nombre
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
        // --- MODO FOTO: Guardar en carpeta específica IMAGES_FOLDER_ID ---
        
        try {
          const targetFolder = DriveApp.getFolderById(IMAGES_FOLDER_ID);
          file = targetFolder.createFile(blob);
          
          // OJO: Para que AppSheet vea la foto, escribimos: "NombreCarpeta/NombreArchivo"
          // Esto asume que la carpeta IMAGES_FOLDER_ID está en la ubicación correcta para AppSheet
          storedPath = `${targetFolder.getName()}/${finalName}`;
          
        } catch (e) {
          throw new Error("No se pudo acceder a la carpeta de Imágenes definida. Verifique el ID.");
        }
        
    } else {
        // --- MODO DOCUMENTO: Guardar en carpeta específica DOCS_FOLDER_ID ---
        
        try {
          const targetFolder = DriveApp.getFolderById(DOCS_FOLDER_ID);
          file = targetFolder.createFile(blob);
          
          // Para documentos usamos URL completa para descarga web
          storedPath = `https://drive.google.com/file/d/${file.getId()}/view`;
          
        } catch (e) {
          throw new Error("No se pudo acceder a la carpeta de Documentos definida. Verifique el ID.");
        }
    }

    // Permisos (Intento de hacerlos visibles para lectores)
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

    if (tipoAnexo === 'Foto') {
      row['Foto'] = storedPath;
    } else if (tipoAnexo === 'Dibujo') {
      row['Dibujo'] = storedPath;
    } else {
      row['Archivo'] = storedPath;
    }

    appendDataToSheet('Solicitudes anexos', row);
    _invalidateDetailCache(email, solicitudId);

    return { success: true, anexoId: anexoUuid, fileName: file.getName(), path: storedPath };
  });
}

// HELPERS
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
  const ss = SpreadsheetApp.openById(_extractSpreadsheetId(spreadsheetId));
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
  const file = DriveApp.getFileById(_extractSpreadsheetId(MAIN_SPREADSHEET_ID));
  const parents = file.getParents();
  if (parents.hasNext()) return parents.next();
  return DriveApp.getRootFolder();
}

function _resolveDriveFileFromAppSheetPath(pathValue) {
  const p = String(pathValue || "").trim();

  // Si ya es una URL de view, la devolvemos como tal
  if (/file\/d\/([^/]+)/.test(p)) {
     return { kind: "url", url: p };
  }
  
  // Soporte para URLs antiguas con id=
  if (/id=([^&]+)/.test(p)) {
     const id = p.match(/id=([^&]+)/)[1];
     return { kind: "url", url: `https://drive.google.com/file/d/${id}/view` };
  }

  // Lógica legacy para rutas relativas
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

  throw new Error(`Archivo no encontrado: ${filename}`);
}

function _findSolicitudHeaderFast(key) {
  return _findRowObjectByKey('Solicitudes', key, [
    'ID Solicitud', 'ID Solicitudes', 'Ticket G4S', 'Ticket Cliente', 'Ticket (Opcional)'
  ]);
}

// LOGICA NEGOCIO

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
      return parsed;
    } catch (e) {}
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

  const result = { header, services, history, documents };

  const json = JSON.stringify(result);
  if (json.length < 90000) cache.put(ck, json, 30);

  return result;
}

// ------------------------------------------------------------------
// ✅ CREATE REQUEST ACTUALIZADO CON API DE APPSHEET
// ------------------------------------------------------------------
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

    const ss = SpreadsheetApp.openById(_extractSpreadsheetId(MAIN_SPREADSHEET_ID));
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
      
      // ✅ MAPEADO CORRECTO SEGÚN SOLICITUD
      "Clasificación Solicitud": payload.clasificacion, 
      "Clasificación": payload.tipoServicio,
      "Técnicos Clientes": "Por disponibilidad", // Valor fijo solicitado

      "Prioridad Solicitud": payload.prioridad,
      "Solicitud": payload.solicitud,
      "Observación": payload.observacion,
      "Usuario Actualización": email
    };

    // --- INTEGRACIÓN DE LA API PARA ACTIVAR CORREOS ---
    // ✅ PRIORIDAD: Guardado directo para asegurar disponibilidad inmediata y evitar latencia de AppSheet
    appendDataToSheet('Solicitudes', newRow);
    SpreadsheetApp.flush(); // Aseguramos que los cambios se persistan antes de seguir

    // ✅ SECUNDARIO: Notificación a AppSheet para disparar automatizaciones (emails)
    try {
      enviarAppSheetAPI('Solicitudes', newRow);
    } catch (e) {
      console.warn("Notificación a AppSheet API falló o detectó duplicado, pero el ticket ya está en la hoja:", e);
    }
    // ------------------------------------------------

    try {
      const historyRow = {
        "ID Estado": Utilities.getUuid(),
        "ID Solicitudes": uuid,
        "Estado actual": "Creado",
        "Usuario Actualización": email,
        "Fecha Actualización": now
      };
      appendDataToSheet('Estados historico', historyRow);
    } catch (e) {
      console.warn("No se pudo guardar el historial inicial:", e);
    }

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
    const fileNameFromRow = _getField(row, ['Nombre']) || "Anexo";
    return { mode: "url", fileName: fileNameFromRow, url: resolved.url };
  }

  const file = resolved.file;
  return { mode: "url", url: `https://drive.google.com/file/d/${file.getId()}/view`, fileName: file.getName() };
}

function getAnexoFileBase64(email, { anexoId }) {
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
  const fileName = _getField(row, ['Nombre']) || "Archivo_G4S";
  let file = null;

  if (pathValue.includes("drive.google.com") || pathValue.includes("/d/")) {
    const idMatch = pathValue.match(/\/d\/([a-zA-Z0-9_-]+)/) || pathValue.match(/id=([a-zA-Z0-9_-]+)/);
    if (idMatch && idMatch[1]) {
      try { file = DriveApp.getFileById(idMatch[1]); } catch(e) {}
    }
  } else {
    try {
      const resolved = _resolveDriveFileFromAppSheetPath(pathValue);
      if (resolved && resolved.kind === "file") {
        file = resolved.file;
      } else if (resolved && resolved.kind === "url") {
        const idMatch = resolved.url.match(/\/d\/([a-zA-Z0-9_-]+)/) || resolved.url.match(/id=([a-zA-Z0-9_-]+)/);
        if (idMatch && idMatch[1]) {
          file = DriveApp.getFileById(idMatch[1]);
        }
      }
    } catch (e) {
      const parts = pathValue.split('/');
      const exactFileName = parts[parts.length - 1];
      if (exactFileName) {
        const filesIt = DriveApp.getFilesByName(exactFileName);
        if (filesIt.hasNext()) {
          file = filesIt.next();
        }
      }
    }
  }

  if (!file) {
    throw new Error("No se pudo encontrar el archivo físico en Google Drive.");
  }

  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());
  const contentType = blob.getContentType();
  const finalName = file.getName() || fileName;

  return { base64, contentType, fileName: finalName };
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
      
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {}

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
      idActivo: String(_getField(r, ['ID Activo', 'id_activo'])).trim(),
      nombreActivo: String(_getField(r, ['Nombre Activo', 'nombre_activo'])).trim(),
      qrSerial: String(_getField(r, ['QR Serial', 'qr_serial'])).trim(),
      nombreUbicacion: String(_getField(r, ['Nombre Ubicacion', 'nombre_ubicacion'])).trim(),
      estadoActivo: String(_getField(r, ['Estado Activo', 'estado_activo'])).trim(),
      funcionamiento: String(_getField(r, ['Funcionamiento', 'funcionamiento'])).trim()
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
  const found = rows.find(r => String(_getField(r, ['QR Serial', 'QR', 'Qr', 'Codigo QR', 'qr_serial'])).trim() === q);
  if (!found) return { found: false };
  return {
    found: true,
    activo: {
      idActivo: String(_getField(found, ['ID Activo', 'id_activo'])).trim(),
      nombreActivo: String(_getField(found, ['Nombre Activo', 'nombre_activo'])).trim(),
      qrSerial: q,
      nombreUbicacion: String(_getField(found, ['Nombre Ubicacion', 'nombre_ubicacion'])).trim(),
      estadoActivo: String(_getField(found, ['Estado Activo', 'estado_activo'])).trim(),
      funcionamiento: String(_getField(found, ['Funcionamiento', 'funcionamiento'])).trim()
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

// ------------------------------------------------------------------
// ✅ MODO PROXY V5: CORRECCIÓN DE RUTAS RELATIVAS (INTERFAZ ORIGINAL COMPLETA)
// ------------------------------------------------------------------
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
            if (filesIt.hasNext()) {
                file = filesIt.next();
            }
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
              const a = document.createElement('a'); a.href = url; a.download = "${fileName}"; a.click();
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

/**
 * ------------------------------------------------------------------
 * ✅ FUNCIÓN PUENTE: API DE APPSHEET
 * ------------------------------------------------------------------
 */
function enviarAppSheetAPI(tableName, rowData) {
  const appId = "c0817cfb-b068-4a46-ae3b-228c0385a486"; 
  const accessKey = "V2-gaw9Q-LcMsx-wfJof-pFCgC-u6igd-FMxtR-23Zr1-V3O4K"; 
  
  const url = `https://api.appsheet.com/api/v1/apps/${appId}/tables/${tableName}/Action`;
  
  const payload = {
    "Action": "Add",
    "Properties": { 
       "Locale": "es-CO", 
       "Timezone": "SA Pacific Standard Time",
       "RunAsUserEmail": rowData["Usuario Actualización"] 
    },
    "Rows": [ rowData ]
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": { "ApplicationAccessKey": accessKey },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const resText = response.getContentText();
    console.log("Respuesta AppSheet API: " + resText);
    return JSON.parse(resText);
  } catch (e) {
    console.error("Error en la API de AppSheet: " + e);
    return null;
  }
}

/**
 * ------------------------------------------------------------------
 * LÓGICA DE ACTIVOS (100% GOOGLE SHEETS)
 * ------------------------------------------------------------------
 */
/**
 * Manejador central para obtener datos de Activos desde Google Sheets.
 * Incluye validación de permisos y procesamiento de relaciones jerárquicas.
 * @param {string} email Email del usuario para validar contexto.
 * @param {Object} params Parámetros de la acción (action y payload).
 * @returns {Array} Resultados estructurados para la interfaz de activos.
 */
function getAssetsData(email, { action, payload = {} }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  try {
    switch (action) {
      case 'getClients': {
        const clientes = getDataFromSheet('Clientes');
        const sedes = getDataFromSheet('Sedes');
        const pisos = getDataFromSheet('Pisos');
        const activos = getDataFromSheet('Activos');

        let targetClientes = clientes;
        if (!context.isAdmin) {
          if (!context.assignedCustomerNames || context.assignedCustomerNames.length === 0) return [];
          const assignedUpper = context.assignedCustomerNames.map(n => String(n).trim().toUpperCase());
          targetClientes = clientes.filter(c => {
            const name = String(_getField(c, ['Nombre cliente', 'Nombre Cliente', 'Nombre', 'RazonSocial'])).trim().toUpperCase();
            return assignedUpper.includes(name);
          });
        }

        return targetClientes.map(c => {
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

          // Mapear la columna real del CSV "Número de piso" al parámetro "nivel" esperado en Index.html
          const nivel = _getField(p, ['Número de piso', 'Numero de piso', 'Nivel', 'nivel']);

          const planoUrl = _getField(p, ['Imagen Plano URL', 'imagen_plano_url', 'Plano', 'Imagen']);
          const totalActivos = activos.filter(a => String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim() === floorId).length;

          return { id_piso: floorId, nombre_piso: floorName, nivel: nivel, imagen_plano_url: planoUrl, total_activos: totalActivos };
        }).sort((a, b) => a.nombre_piso.localeCompare(b.nombre_piso));
      }
      
      case 'getAssets': {
        if (!payload.floorId) throw new Error("floorId es requerido.");
        const activos = getDataFromSheet('Activos');
        let dispositivos = [];
        try {
          dispositivos = getDataFromSheet('Dispositivos');
        } catch(e) {
          console.warn("Hoja auxiliar de dispositivos no cargada.");
        }

        const targetActivos = activos.filter(a => String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim() === String(payload.floorId).trim());
        return targetActivos.map(a => {
          const idDispositivo = String(_getField(a, ['ID Dispositivo', 'Id Dispositivo', 'id_dispositivo'])).trim();
          let tipoDispositivo = _getField(a, ['Tipo Dispositivo', 'Tipo dispositivo', 'tipo_dispositivo', 'Tipo']);

          if (!tipoDispositivo && dispositivos.length > 0) {
            const dispInfo = dispositivos.find(d => String(_getField(d, ['ID Dispositivo', 'Id Dispositivo'])).trim() === idDispositivo);
            if (dispInfo) tipoDispositivo = _getField(dispInfo, ['Clasificación', 'Clasificacion', 'clasificacion']);
          }
          if (!tipoDispositivo) tipoDispositivo = idDispositivo || "General";

          // Parseo seguro de la columna unificada "Ubicación plano" para obtener coord_x y coord_y individuales si viene como "45.5, 62.3"
          const ubicacionPlano = _getField(a, ['Ubicación plano', 'Ubicacion plano', 'ubicacion_plano']);
          let coordX = "";
          let coordY = "";

          if (ubicacionPlano) {
            const strCoords = String(ubicacionPlano).trim();
            if (strCoords.includes(',')) {
              const partesCoords = strCoords.split(',');
              if (partesCoords.length >= 2) {
                coordX = partesCoords[0].trim();
                coordY = partesCoords[1].trim();
              }
            }
          }

          return {
            id_activo: String(_getField(a, ['ID Activo', 'Id Activo', 'id_activo'])).trim(),
            nombre_activo: String(_getField(a, ['Nombre Activo', 'Nombre activo', 'nombre_activo'])).trim(),
            tipo_dispositivo: tipoDispositivo,
            estado_activo: String(_getField(a, ['Estado Activo', 'Estado activo', 'estado_activo'])).trim(),

            coord_x: coordX || _getField(a, ['Coord X', 'coord_x', 'X']),
            coord_y: coordY || _getField(a, ['Coord Y', 'coord_y', 'Y']),

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
