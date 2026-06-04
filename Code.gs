/**
 * ==================================================================
 * CONFIGURACIÓN DE TABLAS
 * ==================================================================
 */

const MAIN_SPREADSHEET_ID        = '1m1XnSzqlTqvKgvn9yLK8CKw_n1MmnUgCVKhrFrEJe4Q'; // 
const PERMISSIONS_SPREADSHEET_ID = '18Ur6__aI2xjDjpA_yl-oqMz7dNh1h8Ds1lhYFA7LB14'; // [cite: 48]
const CLIENTS_SPREADSHEET_ID     = '1GleT24bySpMAPd5lgRPvS0zQ4FsI8u04GGaTT05jLSo'; // [cite: 48]
const SEDES_SPREADSHEET_ID       = '1GEXwm13dOw8ExF_DxJ7BaL7nqdKP7rB35bH0dkEdktU'; // [cite: 49]
const ACTIVOS_SPREADSHEET_ID     = '1-n3WCFErmS7aOPR3l4VHmKdGh0EtvodyhEuahvWPUwg'; // [cite: 49]
// HOJAS MIGRADAS DE BIGQUERY A GOOGLE SHEETS
const PISOS_SPREADSHEET_ID       = '1WWA6xqMZtKHrTvacqP6KUlZLjfPoEVnZdfYpudR8DRw'; // [cite: 50]
const DISPOSITIVOS_SPREADSHEET_ID = '1r07LqMwpJyKQ2DvT7f9elteay4r80MBJl88NSiVvV2o'; // [cite: 51]

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
}; // [cite: 51]

/**
 * ------------------------------------------------------------------
 * OPTIMIZACIÓN: Apertura de Hojas de cálculo por ID o URL
 * ------------------------------------------------------------------
 */
const __SS_MEMO = {}; // [cite: 52]
function _openSS(idOrUrl) {
  if (!__SS_MEMO[idOrUrl]) {
    if (String(idOrUrl).includes("docs.google.com/spreadsheets")) {
      __SS_MEMO[idOrUrl] = SpreadsheetApp.openByUrl(idOrUrl); // [cite: 53]
    } else {
      __SS_MEMO[idOrUrl] = SpreadsheetApp.openById(idOrUrl); // [cite: 54]
    }
  }
  return __SS_MEMO[idOrUrl]; // [cite: 54]
}

/**
 * Helper concurrencia
 */
function _withLock(callback) {
  const lock = LockService.getScriptLock(); // [cite: 55]
  try {
    lock.waitLock(30000); // [cite: 55]
    const result = callback(); // [cite: 56]
    SpreadsheetApp.flush(); 
    return result; // [cite: 56]
  } catch (e) {
    console.error("Error de Lock/Concurrencia:", e); // [cite: 56]
    throw new Error("El servidor está ocupado. Intente de nuevo en unos segundos."); // [cite: 57]
  } finally {
    lock.releaseLock(); // [cite: 57]
  }
}

const __SHEET_INFO_MEMO = {}; // [cite: 58]
function _normHeader(x) { return String(x || "").trim().toLowerCase(); } // [cite: 58]

function _getSheetInfo(sheetName) {
  if (__SHEET_INFO_MEMO[sheetName]) return __SHEET_INFO_MEMO[sheetName]; // [cite: 58]
  const spreadsheetId = SHEET_CONFIG[sheetName]; // [cite: 59]
  if (!spreadsheetId) throw new Error(`Configuración no encontrada para la tabla: ${sheetName}`); // [cite: 59]

  const ss = _openSS(spreadsheetId); // [cite: 59]
  const sheet = ss.getSheetByName(sheetName); // [cite: 60]
  if (!sheet) {
    const info = { sheet: null, headers: [], headersNorm: [], indexByNorm: {}, lastRow: 0, lastCol: 0 }; // [cite: 60]
    __SHEET_INFO_MEMO[sheetName] = info; // [cite: 61]
    return info; // [cite: 61]
  }

  const lastRow = sheet.getLastRow(); // [cite: 61]
  const lastCol = sheet.getLastColumn(); // [cite: 61]
  if (lastRow < 1 || lastCol < 1) {
    const info = { sheet, headers: [], headersNorm: [], indexByNorm: {}, lastRow, lastCol }; // [cite: 62]
    __SHEET_INFO_MEMO[sheetName] = info; // [cite: 63]
    return info; // [cite: 63]
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || []; // [cite: 63]
  const headersNorm = headers.map(_normHeader); // [cite: 63]
  const indexByNorm = {}; // [cite: 64]
  headersNorm.forEach((h, i) => { if (h && indexByNorm[h] === undefined) indexByNorm[h] = i; }); // [cite: 64]
  const info = { sheet, headers, headersNorm, indexByNorm, lastRow, lastCol }; // [cite: 65]
  __SHEET_INFO_MEMO[sheetName] = info; // [cite: 65]
  return info; // [cite: 65]
}

function _findColIndex(headersNorm, candidateNames) {
  for (let i = 0; i < candidateNames.length; i++) { // [cite: 66]
    const idx = headersNorm.indexOf(_normHeader(candidateNames[i])); // [cite: 66]
    if (idx !== -1) return idx; // [cite: 67]
  }
  return -1; // [cite: 67]
}

function _rowToObject(headers, row) {
  const obj = {}; // [cite: 67]
  for (let i = 0; i < headers.length; i++) { // [cite: 68]
    let v = row[i]; // [cite: 68]
    if (v instanceof Date) v = v.toISOString(); // [cite: 69]
    obj[headers[i]] = v; // [cite: 69]
  }
  return obj; // [cite: 70]
}

function _mergeRowRuns(rowNumsSorted) {
  const runs = []; // [cite: 70]
  if (!rowNumsSorted.length) return runs; // [cite: 70]
  let s = rowNumsSorted[0], p = rowNumsSorted[0]; // [cite: 70]
  for (let i = 1; i < rowNumsSorted.length; i++) { // [cite: 71]
    const cur = rowNumsSorted[i]; // [cite: 71]
    if (cur === p + 1) { // [cite: 72]
      p = cur; // [cite: 72]
    } else { // [cite: 73]
      runs.push([s, p]); // [cite: 73]
      s = p = cur; // [cite: 73]
    }
  }
  runs.push([s, p]); // [cite: 74]
  return runs; // [cite: 74]
}

function _fetchRowRuns(sheet, runs, lastCol) {
  const rows = []; // [cite: 74]
  runs.forEach(([start, end]) => { // [cite: 75]
    const num = end - start + 1; // [cite: 75]
    const block = sheet.getRange(start, 1, num, lastCol).getValues(); // [cite: 75]
    for (let i = 0; i < block.length; i++) rows.push(block[i]); // [cite: 75]
  });
  return rows; // [cite: 76]
}

function _findRowObjectByKey(sheetName, keyValue, colCandidates) {
  const { sheet, headers, headersNorm, lastRow, lastCol } = _getSheetInfo(sheetName); // [cite: 76]
  if (!sheet || lastRow < 2) return null; // [cite: 77]

  const k = String(keyValue || "").trim(); // [cite: 77]
  if (!k) return null; // [cite: 77]
  for (let c = 0; c < colCandidates.length; c++) { // [cite: 78]
    const idx = headersNorm.indexOf(_normHeader(colCandidates[c])); // [cite: 78]
    if (idx === -1) continue; // [cite: 79]

    const colVals = sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues(); // [cite: 79]
    for (let i = 0; i < colVals.length; i++) { // [cite: 80]
      const v = String(colVals[i][0]).trim(); // [cite: 80]
      if (v === k) { // [cite: 81]
        const rowNum = i + 2; // [cite: 81]
        const row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0]; // [cite: 82]
        return { rowNum, obj: _rowToObject(headers, row) }; // [cite: 82]
      }
    }
  }
  return null; // [cite: 84]
}

function _getChildrenFast(sheetName, parentKeys) {
  const { sheet, headers, headersNorm, lastRow, lastCol } = _getSheetInfo(sheetName); // [cite: 84]
  if (!sheet || lastRow < 2) return []; // [cite: 85]

  const fkCandidates = ['ID Solicitudes', 'ID Solicitud', 'ID Solicitudes ']; // [cite: 85]
  const idxFk = _findColIndex(headersNorm, fkCandidates); // [cite: 86]
  if (idxFk === -1) return []; // [cite: 86]
  const idxDate = _findColIndex(headersNorm, ['Fecha Actualización', 'Fecha Actualizacion', 'Fecha', 'FechaCambio']); // [cite: 87]
  const set = new Set((parentKeys || []).map(x => String(x || "").trim()).filter(Boolean)); // [cite: 88]
  if (!set.size) return []; // [cite: 88]
  const fkVals = sheet.getRange(2, idxFk + 1, lastRow - 1, 1).getValues(); // [cite: 89]
  const rowNums = []; // [cite: 89]
  for (let i = 0; i < fkVals.length; i++) { // [cite: 90]
    const fk = String(fkVals[i][0]).trim(); // [cite: 90]
    if (fk && set.has(fk)) rowNums.push(i + 2); // [cite: 91]
  }
  if (!rowNums.length) return []; // [cite: 91]

  rowNums.sort((a, b) => a - b); // [cite: 91]
  const runs = _mergeRowRuns(rowNums); // [cite: 92]
  const rows = _fetchRowRuns(sheet, runs, lastCol); // [cite: 92]

  const out = []; // [cite: 92]
  for (let i = 0; i < rows.length; i++) { // [cite: 93]
    const r = rows[i]; // [cite: 93]
    const obj = _rowToObject(headers, r); // [cite: 94]
    let ts = 0; // [cite: 94]
    if (idxDate !== -1) { // [cite: 94]
      const dv = r[idxDate]; // [cite: 94]
      ts = (dv instanceof Date) ? dv.getTime() : (Date.parse(String(dv)) || 0); // [cite: 95]
    }
    obj.__ts = ts; // [cite: 95]
    out.push(obj); // [cite: 95]
  }
  out.sort((a, b) => (b.__ts || 0) - (a.__ts || 0)); // [cite: 96]
  out.forEach(o => delete o.__ts); // [cite: 96]
  return out; // [cite: 96]
}

function doGet(e) {
  if (e.parameter && e.parameter.v === 'archivo' && e.parameter.id) { // [cite: 97]
    return _renderFileView(e.parameter.id); // [cite: 97]
  }
  const template = HtmlService.createTemplateFromFile('Index'); // [cite: 98]
  template.scriptUrl = ScriptApp.getService().getUrl(); // [cite: 98]
  return template
    .evaluate()
    .setTitle('G4S Ticket Tracker') // [cite: 99]
    .setFaviconUrl('https://www.g4s.com/favicon.ico') // [cite: 99]
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // [cite: 99]
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // [cite: 99]
}

function _resolveCallerEmail(request) {
  const active = Session.getActiveUser().getEmail(); // [cite: 99]
  if (active) return String(active).toLowerCase().trim(); // [cite: 100]
  const p = request?.payload || {}; // [cite: 100]
  const fromClient = p.__clientEmail || p.clientEmail || request?.clientEmail || ""; // [cite: 100]
  const email = String(fromClient).toLowerCase().trim(); // [cite: 100]
  if (email && email.includes("@")) return email; // [cite: 101]
  return ""; // [cite: 101]
}

function getData(action, payload) {
  return apiHandler({ endpoint: 'getAssetsData', payload: { action: action, payload: payload } }); // [cite: 102]
}

function apiHandler(request) {
  const userEmail = _resolveCallerEmail(request); // [cite: 103]
  const { endpoint, payload } = request || {}; // [cite: 103]
  console.log(`🔒 [API CHECK] Endpoint: ${endpoint} | ActiveUser: ${Session.getActiveUser().getEmail()} | Resuelto: ${userEmail}`); // [cite: 104]
  try {
    if (!userEmail) throw new Error("No se pudo verificar la identidad del usuario."); // [cite: 105]
    switch (endpoint) { // [cite: 105]
      case 'getUserContext': return getUserContext(userEmail, payload?.ignoreCache); // [cite: 106]
      case 'getRequests': return getRequests(userEmail); // [cite: 106]
      case 'getRequestDetail': return getRequestDetail(userEmail, payload); // [cite: 106]
      case 'createRequest': return createRequest(userEmail, payload); // [cite: 107]
      case 'uploadAnexo': return uploadAnexo(userEmail, payload); // [cite: 107]
      case 'createSolicitudActivo': return createSolicitudActivo(userEmail, payload); // [cite: 107]
      case 'getAnexoDownload': return getAnexoDownload(userEmail, payload); // [cite: 108]
      case 'getSolicitudActivos': return getSolicitudActivos(userEmail, payload); // [cite: 108]
      case 'getActivosCatalog': return getActivosCatalog(userEmail); // [cite: 108]
      case 'getActivoByQr': return getActivoByQr(userEmail, payload); // [cite: 108]
      case 'getClassificationOptions': return getClassificationOptions(userEmail); // [cite: 109]
      case 'getBatchRequestDetails': return getBatchRequestDetails(userEmail, payload); // [cite: 109]
      case 'getAssetsData': return getAssetsData(userEmail, payload); // [cite: 109]
      default: throw new Error(`Endpoint desconocido: ${endpoint}`); // [cite: 109]
    }
  } catch (err) {
    console.error(`❌ ERROR DE SEGURIDAD/EJECUCIÓN: ${err.message}`, err); // [cite: 110]
    return { error: true, message: "Error procesando su solicitud. Contacte al administrador." }; // [cite: 111]
  }
}

const DETAIL_CACHE_VER = "v3"; // [cite: 111]
function _detailCacheKey(email, id) {
  const e = String(email || "").toLowerCase().trim(); // [cite: 112]
  const rid = String(id || "").trim(); // [cite: 112]
  return `detail_${DETAIL_CACHE_VER}_${Utilities.base64Encode(e)}_${rid}`; // [cite: 112]
}
function _invalidateDetailCache(email, id) {
  try { CacheService.getScriptCache().remove(_detailCacheKey(email, id)); } catch (e) {} // [cite: 113]
}

function _sanitizeFileName(name) {
  const n = String(name || 'anexo').trim(); // [cite: 113]
  return n.normalize("NFD").replace(/[\u0300-\u036f]/g, "") // [cite: 114]
          .replace(/[/\\]/g, '_')
          .replace(/[<>:"|?*]/g, '')
          .replace(/\s+/g, '_') 
          .slice(0, 120) || 'anexo';
}

function _ensureFolder(parent, name) {
  const it = parent.getFoldersByName(name); // [cite: 114]
  if (it.hasNext()) return it.next(); // [cite: 114]
  return parent.createFolder(name); // [cite: 114]
}

function _ensurePathFromRoot(root, parts) {
  let current = root; // [cite: 114]
  parts.forEach(p => { current = _ensureFolder(current, p); }); // [cite: 114]
  return current; // [cite: 114]
}

function uploadAnexo(email, payload) {
  const context = getUserContext(email); // [cite: 114]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 114, 115]

  const IMAGES_FOLDER_ID = '1tzYk9jiQ7Lp_bSZylMn0vuzfWZt4xTHb';  // [cite: 115]
  const DOCS_FOLDER_ID   = '1-CBsinL67dJUPfr8WXtKP1wM93B6zPDX'; // [cite: 115]

  const solicitudId = String(payload?.solicitudId || '').trim(); // [cite: 115]
  const tipoAnexo = String(payload?.tipoAnexo || 'Archivo').trim(); // [cite: 115]
  const fileNameInput = String(payload?.fileName || 'anexo'); // [cite: 115]
  const mimeType = String(payload?.mimeType || 'application/octet-stream').trim(); // [cite: 115, 116]
  const base64 = String(payload?.base64 || '').trim(); // [cite: 116]

  if (!solicitudId) throw new Error("solicitudId requerido"); // [cite: 116]
  if (!base64) throw new Error("base64 requerido"); // [cite: 116]
  const headerFound = _findRowObjectByKey('Solicitudes', solicitudId, ['ID Solicitud', 'ID Solicitudes']); // [cite: 117]
  if (!headerFound) throw new Error("Solicitud padre no encontrada."); // [cite: 117]
  const maxBytes = 10 * 1024 * 1024; // [cite: 118]
  const bytes = Utilities.base64Decode(base64); // [cite: 118]
  if (bytes.length > maxBytes) throw new Error("Archivo demasiado grande (máx 10MB)."); // [cite: 119]
  
  const safeFileName = _sanitizeFileName(fileNameInput).replace(/\s+/g, '_'); // [cite: 119]
  const shortId = solicitudId.replace(/-/g, '').slice(0, 8); // [cite: 120]
  const rand = Math.floor(Math.random() * 900000) + 100000; // [cite: 120]
  
  const extMatch = safeFileName.match(/\.([0-9a-z]+)$/i); // [cite: 120]
  const ext = extMatch ? extMatch[1] : (mimeType.includes('image') ? 'jpg' : 'pdf'); // [cite: 121]
  const baseName = safeFileName.replace(/\.[^/.]+$/, "").replace(/\./g, "_"); // [cite: 121]
  const finalName = `${shortId}_${tipoAnexo}_${rand}_${baseName}.${ext}`; // [cite: 122]
  const blob = Utilities.newBlob(bytes, mimeType, finalName); // [cite: 122]
  return _withLock(() => {
    let file;
    let storedPath;
    if (tipoAnexo === 'Foto' || tipoAnexo === 'Dibujo' || mimeType.startsWith('image/')) { // [cite: 123]
        try {
          const targetFolder = DriveApp.getFolderById(IMAGES_FOLDER_ID); // [cite: 123]
          file = targetFolder.createFile(blob); // [cite: 123]
          storedPath = `${targetFolder.getName()}/${finalName}`; // [cite: 123]
        } catch (e) {
          throw new Error("No se pudo acceder a la carpeta de Imágenes definida. Verifique el ID."); // [cite: 123, 124]
        }
    } else {
        try {
          const targetFolder = DriveApp.getFolderById(DOCS_FOLDER_ID); // [cite: 124]
          file = targetFolder.createFile(blob); // [cite: 124]
          storedPath = `https://drive.google.com/file/d/${file.getId()}/view`; // [cite: 124]
        } catch (e) {
          throw new Error("No se pudo acceder a la carpeta de Documentos definida. Verifique el ID."); // [cite: 124, 125]
        }
    }

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // [cite: 125]
    } catch(e) {
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); // [cite: 125]
      } catch(e2) {} // [cite: 126]
    }

    const anexoUuid = Utilities.getUuid(); // [cite: 126]
    const now = new Date(); // [cite: 126]
    const row = {
      "ID Solicitudes anexos": anexoUuid,
      "ID Solicitudes": solicitudId,
      "Tipo anexo": tipoAnexo,
      "Nombre": safeFileName,
      "Usuario Actualización": email,
      "Fecha Actualización": now
    }; // [cite: 127]
    if (tipoAnexo === 'Foto') { row['Foto'] = storedPath; }  // [cite: 128]
    else if (tipoAnexo === 'Dibujo') { row['Dibujo'] = storedPath; }  // [cite: 128, 129]
    else { row['Archivo'] = storedPath; } // [cite: 129]

    appendDataToSheet('Solicitudes anexos', row); // [cite: 129]
    _invalidateDetailCache(email, solicitudId); // [cite: 129]
    return { success: true, anexoId: anexoUuid, fileName: file.getName(), path: storedPath }; // [cite: 130]
  });
}

function getDataFromSheet(sheetName) {
  const { sheet, headers, lastRow, lastCol } = _getSheetInfo(sheetName); // [cite: 131]
  if (!sheet || lastRow < 2 || lastCol < 1) return []; // [cite: 132]
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues(); // [cite: 132]
  if (values.length < 2) return []; // [cite: 133]
  const data = values.slice(1); // [cite: 133]
  return data.map(row => _rowToObject(headers, row)); // [cite: 133]
}

function appendDataToSheet(sheetName, objectData) {
  const spreadsheetId = SHEET_CONFIG[sheetName]; // [cite: 134]
  const ss = _openSS(spreadsheetId);  // [cite: 134]
  const sheet = ss.getSheetByName(sheetName); // [cite: 134]
  if (!sheet) throw new Error(`Hoja ${sheetName} no encontrada.`); // [cite: 135]

  const lastCol = sheet.getLastColumn(); // [cite: 135]
  if (lastCol === 0) throw new Error(`La hoja ${sheetName} está vacía.`); // [cite: 136]

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]; // [cite: 136]
  const rowArray = headers.map(header => { // [cite: 137]
    let val = objectData[header]; // [cite: 137]
    if (val === undefined) { // [cite: 137]
       const cleanHeader = String(header).trim().toLowerCase(); // [cite: 137]
       const foundKey = Object.keys(objectData).find(k => String(k).trim().toLowerCase() === cleanHeader); // [cite: 137]
       if (foundKey) val = objectData[foundKey]; // [cite: 137]
    }
    return val === undefined || val === null ? "" : val; // [cite: 138]
  });
  sheet.appendRow(rowArray); // [cite: 138]
  
  const possibleTicketId = objectData["ID Solicitud"] || objectData["ID Solicitudes"] || ""; // [cite: 138]
  if (possibleTicketId) _invalidateDetailCache(String(objectData["Usuario Actualización"] || ""), possibleTicketId); // [cite: 138]
  return { success: true }; // [cite: 139]
}

function _getField(row, candidateNames) {
  if (!row) return ""; // [cite: 139]
  const keys = Object.keys(row); // [cite: 139]
  for (let i = 0; i < candidateNames.length; i++) { // [cite: 140]
    const c = candidateNames[i]; // [cite: 140]
    if (row[c] !== undefined && row[c] !== null && row[c] !== "") return row[c]; // [cite: 141]
    const k = keys.find(x => String(x).trim() === String(c).trim()); // [cite: 142]
    if (k && row[k] !== undefined && row[k] !== null && row[k] !== "") return row[k]; // [cite: 142]
  }
  return ""; // [cite: 143]
}

function _normalizePath(path) {
  if (!path) return ""; // [cite: 143]
  let p = String(path).trim(); // [cite: 143]
  if (/^https?:\/\//i.test(p)) return p; // [cite: 143]
  p = p.replace(/\\/g, '/'); // [cite: 144]
  p = p.replace(/^\/+/, ''); // [cite: 144]
  p = p.replace(/\/+/g, '/'); // [cite: 144]
  return p; // [cite: 144]
}

function _getRootFolderForFiles() {
  const file = DriveApp.getFileById(MAIN_SPREADSHEET_ID); // [cite: 145]
  const parents = file.getParents(); // [cite: 145]
  if (parents.hasNext()) return parents.next(); // [cite: 145]
  return DriveApp.getRootFolder(); // [cite: 145]
}

function _resolveDriveFileFromAppSheetPath(pathValue) {
  const p = String(pathValue || "").trim(); // [cite: 146]
  if (/file\/d\/([^/]+)/.test(p)) { return { kind: "url", url: p }; } // [cite: 146, 147]
  if (/id=([^&]+)/.test(p)) { // [cite: 147]
     const id = p.match(/id=([^&]+)/)[1]; // [cite: 147]
     return { kind: "url", url: `https://drive.google.com/file/d/${id}/view` }; // [cite: 147]
  }
  const root = _getRootFolderForFiles(); // [cite: 148]
  const parts = p.split('/').filter(Boolean); // [cite: 148]
  const filename = parts.pop(); // [cite: 148]
  try {
    let current = root; // [cite: 149]
    parts.forEach(folderName => { // [cite: 149]
      const it = current.getFoldersByName(folderName); // [cite: 149]
      if (!it.hasNext()) throw new Error(`Carpeta no encontrada: ${folderName}`); // [cite: 149]
      current = it.next(); // [cite: 149]
    });
    const files = current.getFilesByName(filename); // [cite: 150]
    if (files.hasNext()) return { kind: "file", file: files.next() }; // [cite: 150]
  } catch (e) {}

  const safeName = filename.replace(/"/g, '\\"'); // [cite: 151]
  const q = `name = "${safeName}" and trashed = false`; // [cite: 151]
  const it2 = DriveApp.searchFiles(q); // [cite: 152]
  if (it2.hasNext()) return { kind: "file", file: it2.next() }; // [cite: 152]
  throw new Error(`Archivo no encontrado: ${filename}`); // [cite: 152]
}

function _findSolicitudHeaderFast(key) {
  return _findRowObjectByKey('Solicitudes', key, [
    'ID Solicitud', 'ID Solicitudes', 'Ticket G4S', 'Ticket Cliente', 'Ticket (Opcional)'
  ]); // [cite: 153]
}

function getUserContext(email, ignoreCache = false) {
  const cache = CacheService.getScriptCache(); // [cite: 154]
  const cacheKey = `ctx_it_v6_${Utilities.base64Encode(email)}`; // [cite: 154]
  if (!ignoreCache) { // [cite: 155]
    const cachedData = cache.get(cacheKey); // [cite: 155]
    if (cachedData) return JSON.parse(cachedData); // [cite: 155]
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
    }; // [cite: 156]
    const allPermissions = getDataFromSheet('Permisos'); // [cite: 157]
    const userData = allPermissions.find(row => String(row['Correo']).toLowerCase() === email.toLowerCase()); // [cite: 157]
    if (userData) { // [cite: 158]
      context.isValidUser = true; // [cite: 158]
      const rol = (userData['Rol_Asignado'] || '').trim().toLowerCase(); // [cite: 158]
      if (rol === 'administrador') { // [cite: 159]
        context.role = 'Administrador'; // [cite: 159]
        context.isAdmin = true; // [cite: 159]
      }
    } else {
      console.warn(`⚠️ Usuario IT ${email} no encontrado en tabla Permisos.`); // [cite: 160]
    }

    if (!context.isValidUser) return context; // [cite: 161]

    const allRelations = getDataFromSheet('Usuarios filtro'); // [cite: 161]
    const myRelations = allRelations.filter(row => String(row['Usuario']).toLowerCase() === email.toLowerCase()); // [cite: 162]
    const assignedClientIds = []; // [cite: 162]
    myRelations.forEach(row => { // [cite: 163]
      const id = row['Cliente']; // [cite: 163]
      if (id) assignedClientIds.push(String(id)); // [cite: 163]
    });
    context.allowedCustomerIds = assignedClientIds; // [cite: 164]

    if (assignedClientIds.length > 0) {
      const allClientes = getDataFromSheet('Clientes'); // [cite: 164]
      const myClients = allClientes.filter(c => 
        assignedClientIds.includes(String(_getField(c, ['ID Cliente', 'Id Cliente', 'Cliente'])))
      ); // [cite: 165]
      myClients.forEach(c => { // [cite: 166]
        const clientName = _getField(c, ['Nombre cliente', 'Nombre Cliente', 'Nombre', 'RazonSocial']); // [cite: 166]
        if (clientName) context.assignedCustomerNames.push(String(clientName).trim()); // [cite: 166]
      });
      const allSedes = getDataFromSheet('Sedes'); // [cite: 167]
      const mySedes = allSedes.filter(sede => assignedClientIds.includes(String(_getField(sede, ['ID Cliente', 'Id Cliente', 'Cliente'])))); // [cite: 167]
      mySedes.forEach(sede => { // [cite: 168]
        const idSede = String(_getField(sede, ['ID Sede', 'Id Sede', 'Sede', 'IDSede'])).trim(); // [cite: 168]
        const nombreSede = _getField(sede, ['Nombre', 'Nombre_Sede', 'Nombre sede', 'Nombre Sede', 'Sede', 'Label']) || idSede; // [cite: 168]
        if (idSede) { // [cite: 168]
          context.allowedClientIds.push(idSede); // [cite: 168]
          context.clientNames[idSede] = String(nombreSede).trim() || idSede; // [cite: 168]
        }
      });
    }
    cache.put(cacheKey, JSON.stringify(context), 350); // [cite: 169]
    return context; // [cite: 169]
  } catch (e) {
    console.error("Error getUserContext", e); // [cite: 169]
    throw e; // [cite: 170]
  }
}

function getRequests(email) {
  const t0 = Date.now(); // [cite: 170]
  const context = getUserContext(email); // [cite: 170]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 170]
  try {
    let allRows = getDataFromSheet('Solicitudes'); // [cite: 171]
    let filteredRows = []; // [cite: 171]
    if (context.isAdmin) { // [cite: 172]
      filteredRows = allRows; // [cite: 172]
    } else {
      if (context.allowedClientIds.length === 0) return { data: [], total: 0 }; // [cite: 173]
      filteredRows = allRows.filter(row => context.allowedClientIds.includes(String(_getField(row, ['ID Sede'])))); // [cite: 174]
    }
    filteredRows.sort((a, b) => {
      const dateA = new Date(_getField(a, ['Fecha creación cliente', 'Fecha creacion cliente'])).getTime() || 0; // [cite: 174]
      const dateB = new Date(_getField(b, ['Fecha creación cliente', 'Fecha creacion cliente'])).getTime() || 0; // [cite: 174]
      return dateB - dateA; // [cite: 175]
    });
    console.log(`⚡ [PERF] getRequests: ${Date.now() - t0}ms | total=${filteredRows.length}`); // [cite: 175]
    return { data: filteredRows, total: filteredRows.length }; // [cite: 175]
  } catch (e) {
    console.error("Error getRequests", e); // [cite: 176]
    throw new Error("Error obteniendo datos."); // [cite: 176]
  }
}

function getRequestDetail(email, { id }) {
  const context = getUserContext(email); // [cite: 177]
  if (!context.isValidUser) throw new Error("Acceso Denegado"); // [cite: 177]
  if (!id) throw new Error("ID requerido"); // [cite: 178]

  const rid = String(id).trim(); // [cite: 178]
  const cache = CacheService.getScriptCache(); // [cite: 178]
  const ck = _detailCacheKey(email, rid); // [cite: 178]
  const cached = cache.get(ck); // [cite: 179]
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {} // [cite: 179, 180]
  }

  const headerFound = _findSolicitudHeaderFast(rid); // [cite: 180]
  if (!headerFound) throw new Error("Ticket no encontrado."); // [cite: 180]
  const header = headerFound.obj; // [cite: 181]

  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede'])).trim(); // [cite: 181]
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) { // [cite: 182]
      throw new Error("No tiene permisos para ver este ticket."); // [cite: 182]
    }
  }

  const parentKeys = [
    rid,
    String(_getField(header, ['Ticket G4S'])),
    String(_getField(header, ['Ticket Cliente', 'Ticket (Opcional)']))
  ].filter(x => x && x !== "undefined" && x !== "null").map(x => String(x).trim()); // [cite: 183]
  const services = _getChildrenFast('Observaciones historico', parentKeys); // [cite: 184]
  const history = _getChildrenFast('Estados historico', parentKeys); // [cite: 184]
  const documents = _getChildrenFast('Solicitudes anexos', parentKeys); // [cite: 184]
  const activos = _getChildrenFast('Solicitudes activos', parentKeys); // [cite: 185]
  const result = { header, services, history, documents, activos }; // [cite: 186]
  const json = JSON.stringify(result); // [cite: 186]
  if (json.length < 90000) cache.put(ck, json, 30); // [cite: 187]
  return result; // [cite: 187]
}

function createRequest(email, payload) {
  const context = getUserContext(email); // [cite: 187]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 188]
  if (!payload?.idSede || !payload?.solicitud || !payload?.observacion) { // [cite: 188]
    throw new Error("Faltan campos obligatorios."); // [cite: 188]
  }
  if (!context.isAdmin && !context.allowedClientIds.includes(String(payload.idSede))) { // [cite: 189]
    throw new Error("No tiene permisos para esta sede."); // [cite: 189]
  }

  return _withLock(() => {
    const now = new Date(); // [cite: 190]
    const uuid = Utilities.getUuid(); // [cite: 190]
    const allSedes = getDataFromSheet('Sedes'); // [cite: 190]
    const sedeInfo = allSedes.find(s => String(_getField(s, ['ID Sede', 'Id Sede', 'Sede'])).trim() === String(payload.idSede).trim()); // [cite: 190]
    const idCliente = sedeInfo ? _getField(sedeInfo, ['ID Cliente', 'Id Cliente', 'Cliente']) : null; // [cite: 190]

    let letraInicial = "X"; // [cite: 190]
    if (idCliente) { // [cite: 190]
      const allClientes = getDataFromSheet('Clientes'); // [cite: 190]
      const clienteInfo = allClientes.find(c => String(_getField(c, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim() === String(idCliente).trim()); // [cite: 190]
 
      if (clienteInfo) { // [cite: 191]
        const nombreCorto = _getField(clienteInfo, ['Nombre corto', 'Nombre_Corto', 'RazonSocial', 'Razón Social']) || "G"; // [cite: 191]
        letraInicial = String(nombreCorto).trim().charAt(0).toUpperCase(); // [cite: 191]
      }
    }

    const ss = _openSS(MAIN_SPREADSHEET_ID); // [cite: 191]
    const sheet = ss.getSheetByName('Solicitudes'); // [cite: 191]
    const nextRow = sheet.getLastRow() + 1; // [cite: 192]
    const rand = Math.floor(Math.random() * 90) + 10; // [cite: 192]
    const ticketG4S = `${letraInicial}${1000000 + nextRow}${rand}`; // [cite: 192]
    const newRow = {
      "ID Solicitud": uuid,
      "Ticket G4S": ticketG4S,
      "Fecha creación cliente": now,
      "Estado": "Creado",
      "ID Sede": String(payload.idSede).trim(),
      "Ticket Cliente": payload.ticketCliente || "", // [cite: 193, 194]
      "Clasificación Solicitud": payload.clasificacion,  // [cite: 194]
      "Clasificación": payload.tipoServicio, // [cite: 194]
      "Técnicos Clientes": "Por disponibilidad",  // [cite: 194]
      "Prioridad Solicitud": payload.prioridad, // [cite: 194]
      "Solicitud": payload.solicitud, // [cite: 194]
      "Observación": payload.observacion, // [cite: 194]
      "Usuario Actualización": email // [cite: 194]
    };
    appendDataToSheet('Solicitudes', newRow); // [cite: 195]
    SpreadsheetApp.flush(); // [cite: 195]

    try { enviarAppSheetAPI('Solicitudes', newRow); } catch (e) { console.warn("AppSheet Sync advertencia:", e); } // [cite: 195]

    try {
      const historyRow = {
        "ID Estado": Utilities.getUuid(),
        "ID Solicitudes": uuid,
        "Estado actual": "Creado",
        "Usuario Actualización": email,
        "Fecha Actualización": now
      }; // [cite: 196]
      appendDataToSheet('Estados historico', historyRow); // [cite: 197]
    } catch (e) { console.warn("Historial falló:", e); } // [cite: 197]

    _invalidateDetailCache(email, uuid); // [cite: 197]
    const returnRow = {
      ...newRow,
      "Fecha creación cliente": (newRow["Fecha creación cliente"] instanceof Date) ? newRow["Fecha creación cliente"].toISOString() : newRow["Fecha creación cliente"] // [cite: 198, 199]
    };
    return { success: true, solicitudId: uuid, ticketG4S: ticketG4S, GeneratedTicket: ticketG4S, Status: "Success", Rows: [returnRow], row: returnRow }; // [cite: 200]
  });
}

function getAnexoDownload(email, { anexoId }) {
  const context = getUserContext(email); // [cite: 201]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 201]
  if (!anexoId) throw new Error("anexoId requerido"); // [cite: 202]

  const found = _findRowObjectByKey('Solicitudes anexos', anexoId, [
    'ID Solicitudes anexos', 'ID Solicitud anexos', 'ID Anexo', 'ID', 'ID Solicitudes anexos '
  ]); // [cite: 202]
  if (!found) throw new Error("Anexo no encontrado."); // [cite: 203]
  const row = found.obj; // [cite: 203]
  const parentKey = _getField(row, ['ID Solicitudes', 'ID Solicitud']); // [cite: 203]
  const headerFound = _findSolicitudHeaderFast(parentKey); // [cite: 204]
  if (!headerFound) throw new Error("No se pudo validar la solicitud padre del anexo."); // [cite: 204]
  const header = headerFound.obj; // [cite: 205]

  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede'])).trim(); // [cite: 205]
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) { // [cite: 206]
      throw new Error("No tiene permisos para descargar este anexo."); // [cite: 206]
    }
  }

  const pathValue = _getField(row, ['Archivo', 'Archivo ', 'Foto', 'Dibujo', 'QR']) || ""; // [cite: 207]
  if (pathValue.includes("drive.google.com")) { // [cite: 208]
     return { mode: "url", url: pathValue, fileName: _getField(row, ['Nombre']) }; // [cite: 208]
  }
  const resolved = _resolveDriveFileFromAppSheetPath(pathValue); // [cite: 209]
  if (resolved.kind === "url") { // [cite: 209]
    return { mode: "url", fileName: _getField(row, ['Nombre']) || "Anexo", url: resolved.url }; // [cite: 209, 210]
  }
  const file = resolved.file; // [cite: 210]
  return { mode: "url", url: `https://drive.google.com/file/d/${file.getId()}/view`, fileName: file.getName() }; // [cite: 210]
}

function createSolicitudActivo(email, payload) {
  const context = getUserContext(email); // [cite: 211]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 211]
  const solicitudId = String(payload?.solicitudId || payload?.IDSolicitudes || payload?.idSolicitud || '').trim(); // [cite: 212]
  const qr = String(payload?.qrSerial || payload?.qr || payload?.QR || '').trim(); // [cite: 212]
  const idActivo = String(payload?.idActivo || payload?.activoId || payload?.IDActivo || '').trim(); // [cite: 213]
  const observaciones = String(payload?.observaciones || payload?.novedades || '').trim(); // [cite: 213]
  const dibujoBase64 = String(payload?.dibujoBase64 || '').trim(); // [cite: 214]

  if (!solicitudId) throw new Error("solicitudId requerido"); // [cite: 214]
  if (!qr) throw new Error("QR requerido"); // [cite: 214]
  if (!idActivo) throw new Error("ID Activo requerido"); // [cite: 215]

  const headerFound = _findSolicitudHeaderFast(solicitudId); // [cite: 215]
  if (!headerFound) throw new Error("Solicitud padre no encontrada."); // [cite: 215]
  const header = headerFound.obj; // [cite: 216]
  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede'])).trim(); // [cite: 216]
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) { // [cite: 217]
      throw new Error("No tiene permisos para asociar activos a este ticket."); // [cite: 217]
    }
  }

  return _withLock(() => {
    let dibujoPath = ""; // [cite: 218]
    if (dibujoBase64) { // [cite: 218]
      const bytes = Utilities.base64Decode(dibujoBase64); // [cite: 218]
      const root = _getRootFolderForFiles(); // [cite: 218]
      const folder = _ensurePathFromRoot(root, ['Info', 'Clientes', 'Activos']); // [cite: 218]
      const short = Utilities.getUuid().replace(/-/g, '').slice(0, 8); // [cite: 218]
      const rand = Math.floor(Math.random() * 900000) + 100000; // [cite: 218]
      const fileName = `${short}.Dibujo.${rand}.png`; // [cite: 218]
      const blob = Utilities.newBlob(bytes, 'image/png', fileName); // [cite: 218]
      
      const file = folder.createFile(blob); // [cite: 219]
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {} // [cite: 219]
      dibujoPath = `https://drive.google.com/uc?export=view&id=${file.getId()}`; // [cite: 219]
    }

    const now = new Date(); // [cite: 219]
    const rowId = Utilities.getUuid(); // [cite: 219]
    const row = {
      "ID Solicitudes activos": rowId,
      "ID Solicitudes": solicitudId,
      "QR": qr,
      "ID Activo": idActivo,
      "Observaciones": observaciones,
      "Dibujo": dibujoPath,
      "Usuario Actualización": email,
      "Fecha Actualización": now
    }; // [cite: 219, 220]
    appendDataToSheet('Solicitudes activos', row); // [cite: 220]
    _invalidateDetailCache(email, solicitudId); // [cite: 221]
    return { success: true, activoRowId: rowId, dibujoPath }; // [cite: 221]
  });
}

function getSolicitudActivos(email, { solicitudId }) {
  const context = getUserContext(email); // [cite: 222]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 222]
  const sid = String(solicitudId || '').trim(); // [cite: 223]
  if (!sid) throw new Error("solicitudId requerido"); // [cite: 223]
  const rows = _getChildrenFast('Solicitudes activos', [sid]); // [cite: 223]
  return { data: rows, total: rows.length }; // [cite: 224]
}

function getActivosCatalog(email) {
  const context = getUserContext(email); // [cite: 224]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 225]
  const cache = CacheService.getScriptCache(); // [cite: 225]
  const key = "activos_catalog_v2"; // [cite: 225]
  const cached = cache.get(key); // [cite: 225]
  if (cached) return JSON.parse(cached); // [cite: 226]

  const rows = getDataFromSheet('Activos'); // [cite: 226]
  const mapped = rows.map(r => { // [cite: 226]
    return {
      idActivo: String(_getField(r, ['ID Activo'])).trim(),
      nombreActivo: String(_getField(r, ['Nombre Activo'])).trim(),
      qrSerial: String(_getField(r, ['QR Serial'])).trim(),
      nombreUbicacion: String(_getField(r, ['Nombre Ubicacion'])).trim(),
      estadoActivo: String(_getField(r, ['Estado Activo'])).trim(),
      funcionamiento: String(_getField(r, ['Funcionamiento'])).trim()
    };
  }).filter(x => x.idActivo || x.qrSerial); // [cite: 226]
  const res = { data: mapped, total: mapped.length }; // [cite: 227]
  cache.put(key, JSON.stringify(res), 600); // [cite: 227]
  return res; // [cite: 227]
}

function getActivoByQr(email, payload) {
  const context = getUserContext(email); // [cite: 228]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 228]
  const q = String(payload?.qr || '').trim(); // [cite: 229]
  if (!q) throw new Error("qr requerido"); // [cite: 229]
  const rows = getDataFromSheet('Activos'); // [cite: 229]
  const found = rows.find(r => String(_getField(r, ['QR Serial', 'QR', 'Qr', 'Codigo QR'])).trim() === q); // [cite: 230]
  if (!found) return { found: false }; // [cite: 230]
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
  }; // [cite: 231]
}

function getBatchRequestDetails(email, { ids }) {
  const t0 = Date.now(); // [cite: 232]
  const context = getUserContext(email); // [cite: 232]
  if (!context.isValidUser) throw new Error("Acceso Denegado"); // [cite: 233]
  if (!ids || !Array.isArray(ids) || ids.length === 0) return {}; // [cite: 233]
  const targetIds = new Set(ids.map(x => String(x).trim())); // [cite: 234]
  const allServices = getDataFromSheet('Observaciones historico'); // [cite: 234]
  const allHistory = getDataFromSheet('Estados historico'); // [cite: 234]
  const allDocs = getDataFromSheet('Solicitudes anexos'); // [cite: 235]
  const allActivos = getDataFromSheet('Solicitudes activos'); // [cite: 235]

  const result = {}; // [cite: 235]
  targetIds.forEach(id => { result[id] = { services: [], history: [], documents: [], activos: [] }; }); // [cite: 236]
  const findParentIdInRow = (row) => {
    if (!row) return ""; // [cite: 237]
    const candidates = ['idsolicitud', 'idsolicitudes', 'ticketg4s', 'ticketcliente']; // [cite: 237]
    const keys = Object.keys(row); // [cite: 238]
    for (const key of keys) {
      const cleanKey = String(key).toLowerCase().replace(/[^a-z0-9]/g, ''); // [cite: 238]
      if (candidates.includes(cleanKey)) { // [cite: 239]
        const val = row[key]; // [cite: 239]
        if (val !== undefined && val !== null && val !== "") return String(val).trim(); // [cite: 240]
      }
    }
    return ""; // [cite: 241]
  };
  const groupByParentSmart = (rows, targetSet, targetKeyInResult) => {
    rows.forEach(row => {
      const parentId = findParentIdInRow(row); // [cite: 242]
      if (parentId && targetSet.has(parentId)) { // [cite: 242]
        if (!result[parentId][targetKeyInResult]) result[parentId][targetKeyInResult] = []; // [cite: 242]
        result[parentId][targetKeyInResult].push(row); // [cite: 242]
      }
    });
  };

  groupByParentSmart(allServices, targetIds, 'services'); // [cite: 243]
  groupByParentSmart(allHistory, targetIds, 'history'); // [cite: 243]
  groupByParentSmart(allDocs, targetIds, 'documents'); // [cite: 243]
  groupByParentSmart(allActivos, targetIds, 'activos'); // [cite: 243]
  console.log(`⚡ [BATCH SMART] Procesados ${ids.length} tickets. Tiempo: ${Date.now() - t0}ms`); // [cite: 244]
  return result; // [cite: 244]
}

function getClassificationOptions(email) {
  const context = getUserContext(email); // [cite: 244]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 245]
  return ["Visita técnica", "Visita comercial"]; // [cite: 245]
}

function _renderFileView(anexoId) {
  try {
    const found = _findRowObjectByKey('Solicitudes anexos', anexoId, [
      'ID Solicitudes anexos', 'ID Solicitud anexos', 'ID Anexo', 'ID'
    ]); // [cite: 246]
    if (!found) return HtmlService.createHtmlOutput("<h1>Archivo no encontrado en la base de datos.</h1>").setFaviconUrl('https://www.g4s.com/favicon.ico'); // [cite: 247]
    const row = found.obj; // [cite: 247]
    const pathValue = _getField(row, ['Archivo', 'Archivo ', 'Foto', 'Dibujo', 'QR']) || ""; // [cite: 248]
    const fileName = _getField(row, ['Nombre']) || "Archivo_G4S"; // [cite: 248]
    let file = null; // [cite: 248]

    if (pathValue.includes("drive.google.com") || pathValue.includes("/d/")) { // [cite: 249]
        const idMatch = pathValue.match(/\/d\/([a-zA-Z0-9_-]+)/) || pathValue.match(/id=([a-zA-Z0-9_-]+)/); // [cite: 249, 250]
        if (idMatch && idMatch[1]) { // [cite: 250]
            try { file = DriveApp.getFileById(idMatch[1]); } catch(e) {} // [cite: 250, 251]
        }
    } else {
        const parts = pathValue.split('/'); // [cite: 251]
        const exactFileName = parts[parts.length - 1];  // [cite: 251]
        if (exactFileName) { // [cite: 252]
            const filesIt = DriveApp.getFilesByName(exactFileName); // [cite: 252]
            if (filesIt.hasNext()) { file = filesIt.next(); } // [cite: 252]
        }
    }

    if (!file) {
       return HtmlService.createHtmlOutput(`
         <div style='font-family:sans-serif;text-align:center;padding:40px;'>
           <h1>Archivo no encontrado en Drive</h1>
           <p>No se pudo localizar el archivo físico: <b>${fileName}</b></p>
         </div>
       `).setFaviconUrl('https://www.g4s.com/favicon.ico'); // [cite: 253]
    }

    if (file.getSize() > 8 * 1024 * 1024) {  // [cite: 254]
      return HtmlService.createHtmlOutput(`
        <div style="font-family:sans-serif;text-align:center;margin-top:50px;">
          <h2>Archivo Grande</h2>
          <a href="https://drive.google.com/uc?export=download&id=${file.getId()}" style="background:#0033A0;color:white;padding:15px;text-decoration:none;">Descargar</a>
        </div>
      `).setFaviconUrl('https://www.g4s.com/favicon.ico'); // [cite: 254]
    }

    const blob = file.getBlob(); // [cite: 255]
    const base64 = Utilities.base64Encode(blob.getBytes()); // [cite: 255]
    const mimeType = blob.getContentType(); // [cite: 255]
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
    `; // [cite: 256, 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269]
    return HtmlService.createHtmlOutput(html).setFaviconUrl('https://www.g4s.com/favicon.ico'); // [cite: 270]
  } catch (e) {
    return HtmlService.createHtmlOutput(`<h3>Error de Sistema: ${e.message}</h3>`).setFaviconUrl('https://www.g4s.com/favicon.ico'); // [cite: 271]
  }
}

function enviarAppSheetAPI(tableName, rowData) {
  const appId = "c0817cfb-b068-4a46-ae3b-228c0385a486"; // [cite: 272]
  const accessKey = "V2-gaw9Q-LcMsx-wfJof-pFCgC-u6igd-FMxtR-23Zr1-V3O4K";  // [cite: 272]
  const url = `https://api.appsheet.com/api/v1/apps/${appId}/tables/${tableName}/Action`; // [cite: 272]
  const payload = {
    "Action": "Add",
    "Properties": { "Locale": "es-CO", "Timezone": "SA Pacific Standard Time", "RunAsUserEmail": rowData["Usuario Actualización"] }, // [cite: 273]
    "Rows": [ rowData ]
  }; // [cite: 273]
  const options = {
    "method": "post", "contentType": "application/json",
    "headers": { "ApplicationAccessKey": accessKey }, // [cite: 274]
    "payload": JSON.stringify(payload), "muteHttpExceptions": true // [cite: 274]
  }; // [cite: 274]
  try {
    const response = UrlFetchApp.fetch(url, options); // [cite: 275]
    return JSON.parse(response.getContentText()); // [cite: 275]
  } catch (e) {
    console.error("Error en la API de AppSheet: " + e); // [cite: 276]
    return null; // [cite: 276]
  }
}

/**
 * ------------------------------------------------------------------
 * REFACTORIZACIÓN INTEGRAL: CONTROL DE ACTIVOS POR MEMORIA DE GOOGLE SHEETS
 * (Remplaza completamente a BigQuery emulando los conteos y joins en JS)
 * ------------------------------------------------------------------
 */
function getAssetsData(email, { action, payload = {} }) {
  const context = getUserContext(email); // [cite: 277]
  if (!context.isValidUser) throw new Error("Acceso Denegado."); // [cite: 277]

  try {
    switch (action) {
      case 'getClients': {
        const clients = getDataFromSheet('Clientes'); // [cite: 278]
        const sedes = getDataFromSheet('Sedes'); // [cite: 278]
        const pisos = getDataFromSheet('Pisos'); // [cite: 279]
        const activos = getDataFromSheet('Activos'); // [cite: 279]

        let filteredClients = clients; // [cite: 279]
        if (!context.isAdmin) { // [cite: 280]
          if (!context.assignedCustomerNames || context.assignedCustomerNames.length === 0) return []; // [cite: 280]
          const names = context.assignedCustomerNames.map(n => String(n).trim().toUpperCase()); // [cite: 281]
          filteredClients = clients.filter(c => {
            const name = String(_getField(c, ['Nombre cliente', 'Nombre Cliente', 'Nombre', 'RazonSocial'])).trim().toUpperCase(); // [cite: 281]
            return names.includes(name); // [cite: 281]
          });
        }

        return filteredClients.map(c => {
          const clientId = String(_getField(c, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim(); // [cite: 282]
          const clientName = String(_getField(c, ['Nombre cliente', 'Nombre Cliente', 'Nombre', 'RazonSocial'])).trim(); // [cite: 282]
          
          const connectedSedes = sedes.filter(s => String(_getField(s, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim() === clientId)
                                      .map(s => String(_getField(s, ['ID Sede', 'Id Sede', 'Sede', 'IDSede'])).trim()); // [cite: 282, 283]
          const connectedPisos = pisos.filter(p => connectedSedes.includes(String(_getField(p, ['ID Sede', 'Id Sede', 'Sede'])).trim()))
                                      .map(p => String(_getField(p, ['ID Piso', 'Id Piso', 'Piso'])).trim()); // [cite: 283]
          const totalActivos = activos.filter(a => connectedPisos.includes(String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim())).length; // [cite: 284]

          return { id_cliente: clientId, nombre_cliente: clientName, total_activos: totalActivos }; // [cite: 284]
        }).sort((a, b) => a.nombre_cliente.localeCompare(b.nombre_cliente)); // [cite: 285]
      }
      
      case 'getSites': {
        if (!payload.clientId) throw new Error("clientId es requerido."); // [cite: 285]
        const sedes = getDataFromSheet('Sedes'); // [cite: 286]
        const pisos = getDataFromSheet('Pisos'); // [cite: 286]
        const activos = getDataFromSheet('Activos'); // [cite: 286]
        const targetSedes = sedes.filter(s => String(_getField(s, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim() === String(payload.clientId).trim()); // [cite: 287]
        return targetSedes.map(s => {
          const siteId = String(_getField(s, ['ID Sede', 'Id Sede', 'Sede', 'IDSede'])).trim(); // [cite: 288]
          const siteName = String(_getField(s, ['Nombre', 'Nombre_Sede', 'Nombre sede', 'Nombre Sede', 'Sede', 'Label']) || siteId).trim(); // [cite: 288]

          const connectedPisos = pisos.filter(p => String(_getField(p, ['ID Sede', 'Id Sede', 'Sede'])).trim() === siteId)
                                      .map(p => String(_getField(p, ['ID Piso', 'Id Piso', 'Piso'])).trim()); // [cite: 288, 289]
          const totalActivos = activos.filter(a => connectedPisos.includes(String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim())).length; // [cite: 289]

          return { id_sede: siteId, nombre_sede: siteName, total_activos: totalActivos }; // [cite: 289]
        }).sort((a, b) => a.nombre_sede.localeCompare(b.nombre_sede)); // [cite: 290]
      }
      
      case 'getFloors': {
        if (!payload.siteId) throw new Error("siteId es requerido."); // [cite: 290]
        const pisos = getDataFromSheet('Pisos'); // [cite: 291]
        const activos = getDataFromSheet('Activos'); // [cite: 291]

        const targetPisos = pisos.filter(p => String(_getField(p, ['ID Sede', 'Id Sede', 'Sede'])).trim() === String(payload.siteId).trim()); // [cite: 291]
        return targetPisos.map(p => {
          const floorId = String(_getField(p, ['ID Piso', 'Id Piso', 'Piso'])).trim(); // [cite: 292]
          const floorName = String(_getField(p, ['Nombre Piso', 'Nombre piso', 'Nombre'])).trim(); // [cite: 292]
          
          // AJUSTE 1: Mapear la columna real del CSV "Número de piso" al parámetro "nivel" esperado en Index.html
          const nivel = _getField(p, ['Número de piso', 'Numero de piso', 'Nivel', 'nivel']);
          
          const planoUrl = _getField(p, ['Imagen Plano URL', 'imagen_plano_url', 'Plano', 'Imagen']); // [cite: 292]
          const totalActivos = activos.filter(a => String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim() === floorId).length; // [cite: 292]

          return { id_piso: floorId, nombre_piso: floorName, nivel: nivel, imagen_plano_url: planoUrl, total_activos: totalActivos }; // [cite: 293]
        }).sort((a, b) => a.nombre_piso.localeCompare(b.nombre_piso)); // [cite: 293]
      }
      
      case 'getAssets': {
        if (!payload.floorId) throw new Error("floorId es requerido."); // [cite: 294]
        const activos = getDataFromSheet('Activos'); // [cite: 295]
        let dispositivos = []; // [cite: 295]
        try { dispositivos = getDataFromSheet('Dispositivos'); // [cite: 295]
        } catch(e) { console.warn("Hoja auxiliar de dispositivos no cargada."); } // [cite: 296]

        const targetActivos = activos.filter(a => String(_getField(a, ['ID Piso', 'Id Piso', 'Piso'])).trim() === String(payload.floorId).trim()); // [cite: 296]
        return targetActivos.map(a => {
          const idDispositivo = String(_getField(a, ['ID Dispositivo', 'Id Dispositivo', 'id_dispositivo'])).trim(); // [cite: 297]
          let tipoDispositivo = _getField(a, ['Tipo Dispositivo', 'Tipo dispositivo', 'tipo_dispositivo', 'Tipo']); // [cite: 297]
          
          if (!tipoDispositivo && dispositivos.length > 0) { // [cite: 297]
            const dispInfo = dispositivos.find(d => String(_getField(d, ['ID Dispositivo', 'Id Dispositivo'])).trim() === idDispositivo); // [cite: 297]
            if (dispInfo) tipoDispositivo = _getField(dispInfo, ['Clasificación', 'Clasificacion', 'clasificacion']); // [cite: 298]
          }
          if (!tipoDispositivo) tipoDispositivo = idDispositivo || "General"; // [cite: 298]

          // AJUSTE 2: Parseo seguro de la columna unificada "Ubicación plano" para obtener coord_x y coord_y individuales
          const ubicacionPlano = _getField(a, ['Ubicación plano', 'Ubicacion plano', 'ubicacion_plano']);
          let coordX = "";
          let coordY = "";
          
          if (ubicacionPlano) {
            const strCoords = String(ubicacionPlano).trim();
            if (strCoords.includes(',')) {
              const partesCoords = strCoords.split(',');
              if (partesCoords.length >= 2) {
                coordX = partesCoords[0].trim(); // Extrae la primera coordenada
                coordY = partesCoords[1].trim(); // Extrae la segunda coordenada
              }
            }
          }

          return {
            id_activo: String(_getField(a, ['ID Activo', 'Id Activo', 'id_activo'])).trim(), // [cite: 299]
            nombre_activo: String(_getField(a, ['Nombre Activo', 'Nombre activo', 'nombre_activo'])).trim(), // [cite: 299]
            tipo_dispositivo: tipoDispositivo, // [cite: 299]
            estado_activo: String(_getField(a, ['Estado Activo', 'Estado activo', 'estado_activo'])).trim(), // [cite: 299]
            
            // Inyección de coordenadas calculadas dinámicamente desde la columna unificada del CSV
            coord_x: coordX || _getField(a, ['Coord X', 'coord_x', 'X']), // [cite: 299]
            coord_y: coordY || _getField(a, ['Coord Y', 'coord_y', 'Y']), // [cite: 299]
            
            fecha_actualizacion: _getField(a, ['Fecha Actualización', 'Fecha Actualizacion', 'fecha_actualizacion', 'Fecha']), // [cite: 299]
            foto_1: _getField(a, ['Foto 1', 'foto_1', 'Foto']), // [cite: 299]
            foto_2: _getField(a, ['Foto 2', 'foto_2']), // [cite: 299]
            foto_3: _getField(a, ['Foto 3', 'foto_3']), // [cite: 300]
            specs: _getField(a, ['Specs', 'specs', 'Datos Tecnicos', 'datos_tecnicos_json']), // [cite: 300]
            protocol: _getField(a, ['Protocol', 'protocol', 'Ultimo Protocolo', 'ultimo_protocolo_json']) // [cite: 300]
          };
        });
      }
      default: return []; // [cite: 301]
    }
  } catch (e) { 
    console.error("Error en getAssetsData", e); // [cite: 302]
    throw new Error("Error procesando inventario desde Hojas de Cálculo: " + e.message); // [cite: 302]
  }
}
