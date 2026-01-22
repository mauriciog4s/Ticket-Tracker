/**
 * ------------------------------------------------------------------
 * CONFIGURACIÓN Y MAPEO DE HOJAS
 * ------------------------------------------------------------------
 */

/**
 * Obtiene la configuración de BigQuery desde ScriptProperties o valores por defecto.
 * Se recomienda almacenar la clave privada en ScriptProperties por seguridad (BQ_PRIVATE_KEY).
 * @returns {Object} Objeto con las credenciales de BigQuery.
 */
function _getBQConfig() {
  const props = PropertiesService.getScriptProperties().getProperties();
  return {
    private_key: props.BQ_PRIVATE_KEY || "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQC8pIXxJuXj2kgo\n3eJ0CuRG7QdZXazXTfTFt04VL6B9b2q+kHyR8UDefeg3LhaYq19jxazqoMvqFsQV\nU/tTCLzKYL+Yew4vy5NOkmENqXmtrjBHmxPqBoHk2R+7aR52RnGkGPcmNAmqcjv4\nCUjgCCq2ko4VDKIQav6/6Psrg1FoaQy4p1lJXZZVTJBU3vUfUIKzlLu7zOrmdZ/K\nUyK+nXSR5Sw6DeE5pf5ElG3QQ9KgjZ/FnzG9gRRWozdW8IgghY6Lw0efyE72ijjf\n873V0K8FqhqDYd1r9stCj05zl07UHqqLXx1ik8YcIxQQR63kcMkqLmysJZx/axds\naDRAvGQbAgMBAAECggEASgqdU/C7jLoxVnD4oCliPgBswQPGgl9jsnLnH+Oor3Ma\nx584daPmnS14Bqh9UAD7mNKOsyzXvJKg9eoXnBiy2RAuQ3AROmtB7zX/B/i7/JKA\n+qoAn/tb4nHiRZHV1gCCPDFcWE9Wd+MMbKdgRiaOdUiCofpqZd1JDhQo+YQ6YKsl\nFZdXPb9SxOFZOxDLjQY64/FV9gn3qFXBbMYf53yNmzd3l6aH/ERgWw3N9YZEgnFM\nBjKn6AN0JSsFLcacjtvgjNDE47U9jO5vdSKu0aFO+vDFBmlewou36i5S9AYLYyAQ\nNC5RDDQ/+WfPv3rGD7D4Nc9JGgK2PbEHRQypmDDNsQKBgQD3AzUE87voHefjH8xw\nWXlRd1H6+0OXt6BwohtSKJmoUvnmbtSjujSppT0EpQzdnrKJ5migbHu3xofaimcH\nmaeqTohuxtQq6E6Yly3kRGexlCAoj0LEW/Hmk1sIiuwI4ukqO6ZiCshMue8pErpE\n5Shb+2xR48nThaeHl8oqrfilCQKBgQDDgaGdKefwE/FCkMgM3ucB12SkHa37Whc9\nPn7YwrcqRC7ZkAvzkzRB/NcyW95fUcD8dYYvHyCHiksMxocwmJ38CbRJWibQlb3y\nMevWes60iAX4NjM1pWWIzADg8MUnncFlUZPWN/b7uxPzMrY1bijtxumGrme0jLwt\nfgOc6CkNAwKBgFhw3IXmYswsEP/APemoD4j8qOytFDl5NMe/MvsKsGGVPAamfhoV\nLI/lKuDD28Rp8tDvH1z5Gp7lRXUZAuS0vlR7A9xt8j9ep+14i6TkXSA2wgDjsmst\n5IHDFuALJZHU9Nj7PIp0A9184UWaf/j096tfbRww6+2BOEeTMH5xhcpJAoGAF9FD\nDxJ73xOO4L0ioe7F1cOXzyaOe4CONDfY3C9cgRmtW3PhANt+EkvrK4dln9cl25u1\nrSftnpWKbxQAhDsThBDqlcUV1XNooIjUYlyzseqgT4zK0E5GAFRaBw1N93WQifdW\nO1K2FBTGaWpUKE4zTkRdTrsQhz5d7mzbo9HkrmECgYAxR0UNwZeSLe+yZhdeK3cv\qz55RbNeHwrhE10PE49CwlUDDdTHk2qK7raAV+LMFEz2Lq8umXxx2OgJSEip3ty4\noVA5qOjr5M62v1wTbrDpmi2ItWXxuzH+oHVW3MBS4jnrbZzsoZ0ZF855xbgfAEwI\nDAygge9kB/HNsXs2OMufAw==\n-----END PRIVATE KEY-----\n",
    client_email: props.BQ_CLIENT_EMAIL || "tz1-bigquery@g4s-shared-tz1.iam.gserviceaccount.com",
    project_id: props.BQ_PROJECT_ID || "g4s-shared-tz1"
  };
}

const DATASET_ID = "ControlTower";

// IDs extraídos de las URLs proporcionadas
const MAIN_SPREADSHEET_ID = '1MC76eZZt7qiso2M8LMz777_xJnzrl_ZpZptDZBnPlDo'; 
const PERMISSIONS_SPREADSHEET_ID = '1zcZZGe_93ytWXtCF1kmk_Y8zc5b5cL1xH34i7v1w01k'; 
const CLIENTS_SPREADSHEET_ID = '1hHWPJF9KSC0opplpCNgHRNkW6CLf7StXG2Y31m6yUpo'; 
const SEDES_SPREADSHEET_ID = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM'; 
const ACTIVOS_SPREADSHEET_ID = '1JU8c1MidgV4DRFg6W-GxZ2tHkfKNqGt1_cR5VDTehC4'; 

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
  'Solicitudes activos': MAIN_SPREADSHEET_ID,
  'Activos': ACTIVOS_SPREADSHEET_ID,
};

/**
 * ------------------------------------------------------------------
 * OPTIMIZACIÓN: Reuso de Spreadsheet abiertos (memoria por ejecución)
 * ------------------------------------------------------------------
 */
const __SS_MEMO = {}; 
function _openSS(spreadsheetId) {
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
  const ss = SpreadsheetApp.openById(spreadsheetId); 
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

    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
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

// --- UTILIDADES DE CONEXIÓN BIGQUERY (OAuth2) ---
function _getBQService() {
  const config = _getBQConfig();
  return OAuth2.createService('BigQueryApp')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setPrivateKey(config.private_key)
    .setIssuer(config.client_email)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope('https://www.googleapis.com/auth/bigquery');
}

function _runBQQuery(query) {
  const config = _getBQConfig();
  const service = _getBQService();
  if (!service.hasAccess()) throw new Error('Error de Autenticación BigQuery: ' + service.getLastError());
  
  const url = `https://bigquery.googleapis.com/bigquery/v2/projects/${config.project_id}/queries`;
  const response = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + service.getAccessToken() },
    payload: JSON.stringify({ query: query, useLegacySql: false })
  });
  
  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);
  
  if (!json.rows) return [];
  const fields = json.schema.fields.map(f => f.name);
  return json.rows.map(row => {
    let obj = {};
    row.f.forEach((cell, i) => { obj[fields[i]] = cell.v; });
    return obj;
  });
}

/**
 * ------------------------------------------------------------------
 * LÓGICA DE ACTIVOS (BIGQUERY)
 * ------------------------------------------------------------------
 */
/**
 * Manejador central para obtener datos de Activos desde BigQuery.
 * Incluye validación de permisos y protección contra inyección SQL.
 * 
 * @param {string} email Email del usuario para validar contexto.
 * @param {Object} params Parámetros de la acción (action y payload).
 * @returns {Array} Resultados de la consulta a BigQuery.
 */
function getAssetsData(email, { action, payload = {} }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  const config = _getBQConfig();
  const projectId = config.project_id;
  
  // Helper para escapar comillas simples y prevenir inyección SQL básica
  const esc = (v) => String(v || '').replace(/'/g, "''");

  try {
    switch (action) {
      case 'getClients':
        let clientQuery = `SELECT DISTINCT id_cliente, nombre_cliente FROM \`${projectId}.${DATASET_ID}.DIM_CLIENTES\``;
        
        if (!context.isAdmin) {
          // Filtrado por nombre de cliente para mayor compatibilidad con Usuarios Filtro
          if (!context.assignedCustomerNames || context.assignedCustomerNames.length === 0) return [];
          const names = context.assignedCustomerNames.map(n => `'${esc(n).toUpperCase()}'`).join(',');
          clientQuery += ` WHERE UPPER(nombre_cliente) IN (${names})`;
        }
        
        clientQuery += ` ORDER BY nombre_cliente`;
        return _runBQQuery(clientQuery);
      
      case 'getSites':
        if (!payload.clientId) throw new Error("clientId es requerido.");
        return _runBQQuery(`
          SELECT id_sede, nombre_sede 
          FROM \`${projectId}.${DATASET_ID}.DIM_SEDES\` 
          WHERE id_cliente = '${esc(payload.clientId)}' 
          ORDER BY nombre_sede
        `);
      
      case 'getFloors':
        if (!payload.siteId) throw new Error("siteId es requerido.");
        return _runBQQuery(`
          SELECT id_piso, nombre_piso, nivel, imagen_plano_url 
          FROM \`${projectId}.${DATASET_ID}.DIM_PISOS\` 
          WHERE id_sede = '${esc(payload.siteId)}' 
          ORDER BY nombre_piso
        `);
      
      case 'getAssets':
        if (!payload.floorId) throw new Error("floorId es requerido.");
        return _runBQQuery(`
          SELECT 
            A.id_activo, 
            A.nombre_activo, 
            COALESCE(D.clasificacion, A.id_dispositivo) as tipo_dispositivo, 
            A.estado_activo, 
            A.coord_x, 
            A.coord_y, 
            A.fecha_actualizacion, 
            A.foto_1, 
            A.foto_2, 
            A.foto_3, 
            TO_JSON_STRING(A.datos_tecnicos_json) as specs,
            TO_JSON_STRING(A.ultimo_protocolo_json) as protocol
          FROM \`${projectId}.${DATASET_ID}.DIM_ACTIVOS\` A
          LEFT JOIN \`${projectId}.${DATASET_ID}.DIM_DISPOSITIVOS\` D 
            ON A.id_dispositivo = D.id_dispositivo
          WHERE A.id_piso = '${esc(payload.floorId)}'
          LIMIT 2000
        `);

      default: return [];
    }
  } catch (e) { 
    console.error("Error en getAssetsData", e);
    throw new Error("Error obteniendo datos de activos: " + e.message); 
  }
}
