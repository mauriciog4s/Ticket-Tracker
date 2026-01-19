/**
 * @fileoverview G4S TICKET TRACKER - BACKEND (Google Apps Script)
 * Lógica del servidor para gestionar tickets, anexos, activos y permisos,
 * utilizando Google Sheets como motor de persistencia masiva.
 *
 * @author Jules (Software Engineer)
 * @version 2.0
 */

/**
 * ------------------------------------------------------------------
 * CONFIGURACIÓN Y MAPEO DE HOJAS
 * ------------------------------------------------------------------
 */

/** @const {string} ID del Spreadsheet principal (Solicitudes, Estados, Observaciones) */
const MAIN_SPREADSHEET_ID = '1MC76eZZt7qiso2M8LMz777_xJnzrl_ZpZptDZBnPlDo'; 

/** @const {string} ID del Spreadsheet de permisos y usuarios */
const PERMISSIONS_SPREADSHEET_ID = '1zcZZGe_93ytWXtCF1kmk_Y8zc5b5cL1xH34i7v1w01k'; 

/** @const {string} ID del Spreadsheet de catálogo de clientes */
const CLIENTS_SPREADSHEET_ID = '1hHWPJF9KSC0opplpCNgHRNkW6CLf7StXG2Y31m6yUpo'; 

/** @const {string} ID del Spreadsheet de sedes operativas */
const SEDES_SPREADSHEET_ID = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM'; 

/** @const {string} ID del Spreadsheet de inventario de activos (Assets) */
const ACTIVOS_SPREADSHEET_ID = '1JU8c1MidgV4DRFg6W-GxZ2tHkfKNqGt1_cR5VDTehC4'; 

/**
 * @type {Object<string, string>}
 * Mapeo de nombres de tablas a sus respectivos contenedores (Spreadsheets).
 */
const SHEET_CONFIG = {
  'Solicitudes': MAIN_SPREADSHEET_ID,
  'Estados historico': MAIN_SPREADSHEET_ID,
  'Observaciones historico': MAIN_SPREADSHEET_ID,
  'Estados': MAIN_SPREADSHEET_ID,
  'Solicitudes anexos': MAIN_SPREADSHEET_ID,
  'Solicitudes activos': MAIN_SPREADSHEET_ID,
  'Permisos': PERMISSIONS_SPREADSHEET_ID,
  'Usuarios filtro': PERMISSIONS_SPREADSHEET_ID,
  'Clientes': CLIENTS_SPREADSHEET_ID,
  'Sedes': SEDES_SPREADSHEET_ID,
  'Activos': ACTIVOS_SPREADSHEET_ID,
};

/**
 * ------------------------------------------------------------------
 * GESTIÓN DE MEMORIA Y CACHÉ DE EJECUCIÓN
 * ------------------------------------------------------------------
 */

/** @type {Object<string, GoogleAppsScript.Spreadsheet.Spreadsheet>} Caché de instancias de SS */
const __SS_MEMO = {}; 

/**
 * Abre y cachea una instancia de Spreadsheet para optimizar lecturas múltiples.
 * @param {string} spreadsheetId
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function _openSS(spreadsheetId) {
  if (!__SS_MEMO[spreadsheetId]) __SS_MEMO[spreadsheetId] = SpreadsheetApp.openById(spreadsheetId);
  return __SS_MEMO[spreadsheetId];
}

/** @type {Object<string, Object>} Caché de metadatos de hojas */
const __SHEET_INFO_MEMO = {};

/**
 * Normaliza una cadena para comparaciones de encabezados.
 * @param {string} x
 * @return {string}
 */
function _normHeader(x) { return String(x || "").trim().toLowerCase(); }

/**
 * Obtiene y cachea información estructural de una hoja.
 * @param {string} sheetName Nombre de la hoja.
 * @return {Object} Metadatos de la hoja (headers, lastRow, lastCol, etc).
 */
function _getSheetInfo(sheetName) {
  if (__SHEET_INFO_MEMO[sheetName]) return __SHEET_INFO_MEMO[sheetName];

  const spreadsheetId = SHEET_CONFIG[sheetName];
  if (!spreadsheetId) throw new Error(`Configuración no encontrada para: ${sheetName}`);

  const ss = _openSS(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { sheet: null, headers: [], headersNorm: [], indexByNorm: {}, lastRow: 0, lastCol: 0 };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    return { sheet, headers: [], headersNorm: [], indexByNorm: {}, lastRow, lastCol };
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const headersNorm = headers.map(_normHeader);
  const indexByNorm = {};
  headersNorm.forEach((h, i) => { if (h && indexByNorm[h] === undefined) indexByNorm[h] = i; });

  const info = { sheet, headers, headersNorm, indexByNorm, lastRow, lastCol };
  __SHEET_INFO_MEMO[sheetName] = info;
  return info;
}

/**
 * ------------------------------------------------------------------
 * HELPERS DE DATOS Y BÚSQUEDA
 * ------------------------------------------------------------------
 */

/**
 * Busca el índice de una columna entre varios candidatos.
 * @param {string[]} headersNorm Array de encabezados normalizados.
 * @param {string[]} candidateNames Nombres posibles de la columna.
 * @return {number} Índice de la columna o -1 si no se encuentra.
 */
function _findColIndex(headersNorm, candidateNames) {
  for (let i = 0; i < candidateNames.length; i++) {
    const idx = headersNorm.indexOf(_normHeader(candidateNames[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Mapea una fila plana a un objeto JSON descriptivo.
 * @param {string[]} headers
 * @param {any[]} row
 * @return {Object}
 */
function _rowToObject(headers, row) {
  const obj = {};
  for (let i = 0; i < headers.length; i++) {
    let v = row[i];
    if (v instanceof Date) v = v.toISOString();
    obj[headers[i]] = v;
  }
  return obj;
}

/**
 * Identifica bloques continuos de filas para optimizar la lectura mediante rangos.
 * @param {number[]} rowNumsSorted Lista ordenada de números de fila.
 * @return {number[][]} Array de pares [inicio, fin].
 */
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

/**
 * Extrae físicamente los datos de los bloques de filas identificados.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number[][]} runs
 * @param {number} lastCol
 * @return {any[][]}
 */
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
 * Localiza un registro único mediante una búsqueda por columna/llave.
 * @param {string} sheetName
 * @param {string} keyValue
 * @param {string[]} colCandidates
 * @return {?{rowNum: number, obj: Object}}
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
      if (String(colVals[i][0]).trim() === k) {
        const rowNum = i + 2;
        const row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
        return { rowNum, obj: _rowToObject(headers, row) };
      }
    }
  }
  return null;
}

/**
 * Obtiene todos los registros relacionados (hijos) de forma masiva y optimizada.
 * @param {string} sheetName
 * @param {string[]} parentKeys IDs de referencia.
 * @return {Object[]} Lista de objetos hijos ordenada por fecha.
 */
function _getChildrenFast(sheetName, parentKeys) {
  const { sheet, headers, headersNorm, lastRow, lastCol } = _getSheetInfo(sheetName);
  if (!sheet || lastRow < 2) return [];

  const idxFk = _findColIndex(headersNorm, ['ID Solicitudes', 'ID Solicitud']);
  if (idxFk === -1) return [];

  const idxDate = _findColIndex(headersNorm, ['Fecha Actualización', 'Fecha', 'FechaCambio']);
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
 * SEGURIDAD Y CONCURRENCIA
 * ------------------------------------------------------------------
 */

/**
 * Garantiza la atomicidad de las operaciones de escritura mediante bloqueos.
 * @param {Function} callback
 * @return {any}
 */
function _withLock(callback) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const result = callback();
    SpreadsheetApp.flush();
    return result;
  } catch (e) {
    console.error("Lock Timeout:", e);
    throw new Error("El sistema está saturado. Reintente en un momento.");
  } finally {
    lock.releaseLock();
  }
}

/**
 * ------------------------------------------------------------------
 * ENTRADA PRINCIPAL (ROUTERS)
 * ------------------------------------------------------------------
 */

/**
 * Punto de entrada para la Web App. Maneja el visor de archivos y la carga inicial.
 * @param {GoogleAppsScript.Events.DoGet} e
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  if (e.parameter && e.parameter.v === 'archivo' && e.parameter.id) {
    return _renderFileView(e.parameter.id);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.scriptUrl = ScriptApp.getService().getUrl();
  
  return template
    .evaluate()
    .setTitle('G4S Ticket Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Resuelve la identidad del usuario de forma segura.
 * @param {Object} request
 * @return {string} Correo electrónico normalizado.
 */
function _resolveCallerEmail(request) {
  const active = Session.getActiveUser().getEmail();
  if (active) return String(active).toLowerCase().trim();

  const p = request?.payload || {};
  const email = String(p.__clientEmail || p.clientEmail || request?.clientEmail || "").toLowerCase().trim();
  return email && email.includes("@") ? email : "";
}

/**
 * Dispatcher central de la API. Enruta todas las peticiones desde el frontend.
 * @param {Object} request { endpoint: string, payload: Object }
 * @return {any} Respuesta JSON para el cliente.
 */
function apiHandler(request) {
  const userEmail = _resolveCallerEmail(request);
  const { endpoint, payload } = request || {};

  console.log(`🔒 [API] Call: ${endpoint} | User: ${userEmail}`);

  try {
    if (!userEmail) throw new Error("Sesión corporativa no válida.");

    switch (endpoint) {
      case 'getUserContext': return getUserContext(userEmail, payload?.ignoreCache);
      case 'getRequests': return getRequests(userEmail);
      case 'getRequestDetail': return getRequestDetail(userEmail, payload);
      case 'getBatchRequestDetails': return getBatchRequestDetails(userEmail, payload);
      case 'createRequest': return createRequest(userEmail, payload);
      case 'uploadAnexo': return uploadAnexo(userEmail, payload);
      case 'getAnexoDownload': return getAnexoDownload(userEmail, payload);
      case 'createSolicitudActivo': return createSolicitudActivo(userEmail, payload);
      case 'getSolicitudActivos': return getSolicitudActivos(userEmail, payload);
      case 'getActivosCatalog': return getActivosCatalog(userEmail);
      case 'getActivoByQr': return getActivoByQr(userEmail, payload);
      case 'getClassificationOptions': return getClassificationOptions(userEmail);

      default: throw new Error(`Endpoint '${endpoint}' no soportado.`);
    }
  } catch (err) {
    console.error(`❌ API Error [${endpoint}]:`, err);
    return { error: true, message: err.message };
  }
}

/**
 * ------------------------------------------------------------------
 * LÓGICA DE NEGOCIO Y DATOS
 * ------------------------------------------------------------------
 */

const DETAIL_CACHE_VER = "v3"; 

/** Genera una llave de caché única por usuario/ticket. */
function _detailCacheKey(email, id) {
  return `detail_${DETAIL_CACHE_VER}_${Utilities.base64Encode(email.toLowerCase().trim())}_${String(id).trim()}`;
}

/** Invalida la caché de un ticket tras una actualización. */
function _invalidateDetailCache(email, id) {
  try { CacheService.getScriptCache().remove(_detailCacheKey(email, id)); } catch (e) {}
}

/**
 * Resuelve los permisos y alcance del usuario (Sedes/Clientes permitidos).
 * @param {string} email
 * @param {boolean} [ignoreCache=false]
 * @return {Object} Contexto de usuario completo.
 */
function getUserContext(email, ignoreCache = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `ctx_it_v5_${Utilities.base64Encode(email)}`; 
  
  if (!ignoreCache) {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  try {
    let context = {
      email, role: 'Usuario', allowedClientIds: [], clientNames: {},
      assignedCustomerNames: [], isValidUser: false, isAdmin: false
    };

    const allPermissions = getDataFromSheet('Permisos');
    const userData = allPermissions.find(row => String(row['Correo']).toLowerCase() === email.toLowerCase());

    if (userData) {
      context.isValidUser = true;
      if ((userData['Rol_Asignado'] || '').toLowerCase().includes('administrador')) {
        context.role = 'Administrador'; context.isAdmin = true;
      }
    }

    if (!context.isValidUser) return context;

    const allRelations = getDataFromSheet('Usuarios filtro');
    const myRelations = allRelations.filter(row => String(row['Usuario']).toLowerCase() === email.toLowerCase());
    const assignedClientIds = myRelations.map(row => String(row['Cliente']).trim()).filter(Boolean);

    if (assignedClientIds.length > 0) {
      const allClientes = getDataFromSheet('Clientes');
      allClientes.filter(c => assignedClientIds.includes(String(_getField(c, ['ID Cliente', 'Cliente']))))
        .forEach(c => {
          const name = _getField(c, ['Nombre cliente', 'Nombre', 'RazonSocial']);
          if (name) context.assignedCustomerNames.push(String(name).trim());
        });

      const allSedes = getDataFromSheet('Sedes');
      allSedes.filter(sede => assignedClientIds.includes(String(_getField(sede, ['ID Cliente', 'Cliente']))))
        .forEach(sede => {
          const id = String(_getField(sede, ['ID Sede', 'IDSede'])).trim();
          const name = _getField(sede, ['Nombre', 'Nombre sede']) || id;
          if (id) {
            context.allowedClientIds.push(id);
            context.clientNames[id] = String(name).trim();
          }
        });
    }

    cache.put(cacheKey, JSON.stringify(context), 350);
    return context;
  } catch (e) {
    console.error("getUserContext error:", e);
    throw e;
  }
}

/**
 * Lista solicitudes filtradas por permisos.
 * @param {string} email
 * @return {Object} { data: Object[], total: number }
 */
function getRequests(email) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Usuario no autorizado.");

  try {
    let rows = getDataFromSheet('Solicitudes');
    if (!context.isAdmin) {
      rows = rows.filter(row => context.allowedClientIds.includes(String(_getField(row, ['ID Sede']))));
    }

    rows.sort((a, b) => {
      const dateA = new Date(_getField(a, ['Fecha creación cliente'])).getTime() || 0;
      const dateB = new Date(_getField(b, ['Fecha creación cliente'])).getTime() || 0;
      return dateB - dateA;
    });

    return { data: rows, total: rows.length };
  } catch (e) {
    throw new Error("Fallo al sincronizar solicitudes.");
  }
}

/**
 * Detalle extendido de un ticket específico.
 * @param {string} email
 * @param {Object} payload { id: string }
 * @return {Object}
 */
function getRequestDetail(email, { id }) {
  const context = getUserContext(email);
  if (!id) throw new Error("ID de ticket requerido.");

  const cache = CacheService.getScriptCache();
  const ck = _detailCacheKey(email, id);
  const cached = cache.get(ck);
  if (cached) return JSON.parse(cached);

  const headerFound = _findSolicitudHeaderFast(id);
  if (!headerFound) throw new Error("Ticket inexistente.");
  const header = headerFound.obj;

  if (!context.isAdmin) {
    const sedeId = String(_getField(header, ['ID Sede'])).trim();
    if (!context.allowedClientIds.includes(sedeId)) throw new Error("Privilegios insuficientes.");
  }

  const pKeys = [id, String(header['Ticket G4S']), String(header['Ticket Cliente'])].filter(x => x && x !== "null");

  const result = {
    header,
    services: _getChildrenFast('Observaciones historico', pKeys),
    history: _getChildrenFast('Estados historico', pKeys),
    documents: _getChildrenFast('Solicitudes anexos', pKeys),
    activos: _getChildrenFast('Solicitudes activos', pKeys)
  };

  const json = JSON.stringify(result);
  if (json.length < 90000) cache.put(ck, json, 60);
  return result;
}

/**
 * Registra una nueva solicitud.
 * @param {string} email
 * @param {Object} payload Datos del formulario.
 * @return {Object} Resultado de la creación.
 */
function createRequest(email, payload) {
  const context = getUserContext(email);
  if (!payload?.idSede || !payload?.solicitud) throw new Error("Faltan campos mandatorios.");

  return _withLock(() => {
    const uuid = Utilities.getUuid();
    const now = new Date();

    // Generación de Ticket G4S con prefijo dinámico
    const allSedes = getDataFromSheet('Sedes');
    const sede = allSedes.find(s => String(_getField(s, ['ID Sede'])).trim() === String(payload.idSede).trim());
    const idCli = sede ? _getField(sede, ['ID Cliente']) : null;
    let prefix = "G";
    if (idCli) {
      const cli = getDataFromSheet('Clientes').find(c => String(_getField(c, ['ID Cliente'])).trim() === String(idCli).trim());
      prefix = String(_getField(cli, ['Nombre corto']) || "G").charAt(0).toUpperCase();
    }

    const sheet = _openSS(MAIN_SPREADSHEET_ID).getSheetByName('Solicitudes');
    const ticketG4S = `${prefix}${1000000 + sheet.getLastRow() + 1}${Math.floor(Math.random() * 90) + 10}`;

    const newRow = {
      "ID Solicitud": uuid, "Ticket G4S": ticketG4S, "Fecha creación cliente": now,
      "Estado": "Creado", "ID Sede": payload.idSede, "Ticket Cliente": payload.ticketCliente || "",
      "Clasificación Solicitud": payload.clasificacion, "Clasificación": payload.tipoServicio,
      "Técnicos Clientes": "Por disponibilidad", "Prioridad Solicitud": payload.prioridad,
      "Solicitud": payload.solicitud, "Observación": payload.observacion, "Usuario Actualización": email
    };

    // Disparar automatización vía API y persistir
    const apiRes = enviarAppSheetAPI('Solicitudes', newRow);
    if (!apiRes || apiRes.error) appendDataToSheet('Solicitudes', newRow);

    appendDataToSheet('Estados historico', {
      "ID Estado": Utilities.getUuid(), "ID Solicitudes": uuid, "Estado actual": "Creado",
      "Usuario Actualización": email, "Fecha Actualización": now
    });

    _invalidateDetailCache(email, uuid);
    return { success: true, solicitudId: uuid, ticketG4S };
  });
}

/**
 * ------------------------------------------------------------------
 * GESTIÓN DE ARCHIVOS Y ANEXOS
 * ------------------------------------------------------------------
 */

/** Sanitiza nombres de archivo para Google Drive. */
function _sanitizeFileName(name) {
  return String(name || 'anexo').trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9._-]/gi, '_').slice(0, 100);
}

/**
 * Sube archivos a carpetas específicas (Fotos/Docs) según tipo.
 * @param {string} email
 * @param {Object} payload { solicitudId, tipoAnexo, base64, mimeType, fileName }
 * @return {Object}
 */
function uploadAnexo(email, payload) {
  const IMAGES_FOLDER_ID = '1tzYk9jiQ7Lp_bSZylMn0vuzfWZt4xTHb';
  const DOCS_FOLDER_ID   = '1-CBsinL67dJUPfr8WXtKP1wM93B6zPDX';

  if (!payload?.solicitudId || !payload?.base64) throw new Error("Datos insuficientes para carga.");

  const bytes = Utilities.base64Decode(payload.base64);
  const isImg = payload.tipoAnexo === 'Foto' || (payload.mimeType || "").startsWith('image/');
  const folder = DriveApp.getFolderById(isImg ? IMAGES_FOLDER_ID : DOCS_FOLDER_ID);
  
  const finalName = `${payload.solicitudId.slice(0,8)}_${payload.tipoAnexo}_${Date.now()}_${_sanitizeFileName(payload.fileName)}`;
  const file = folder.createFile(Utilities.newBlob(bytes, payload.mimeType, finalName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const row = {
    "ID Solicitudes anexos": Utilities.getUuid(), "ID Solicitudes": payload.solicitudId,
    "Tipo anexo": payload.tipoAnexo, "Nombre": payload.fileName, "Usuario Actualización": email, "Fecha Actualización": new Date()
  };
  const path = isImg ? `${folder.getName()}/${finalName}` : file.getUrl();
  row[payload.tipoAnexo === 'Foto' ? 'Foto' : payload.tipoAnexo === 'Dibujo' ? 'Dibujo' : 'Archivo'] = path;

  appendDataToSheet('Solicitudes anexos', row);
  _invalidateDetailCache(email, payload.solicitudId);
  return { success: true, path };
}

/** Resuelve un archivo físico en Drive desde un ID o ruta relativa. */
function _resolveDriveFile(pathValue) {
  const idMatch = String(pathValue).match(/\/d\/([a-zA-Z0-9_-]+)/) || String(pathValue).match(/id=([a-zA-Z0-9_-]+)/);
  if (idMatch) return DriveApp.getFileById(idMatch[1]);
  const parts = String(pathValue).split('/');
  const it = DriveApp.getFilesByName(parts[parts.length - 1]);
  return it.hasNext() ? it.next() : null;
}

/** Endpoint para obtener link de visualización de anexo. */
function getAnexoDownload(email, { anexoId }) {
  const row = _findRowObjectByKey('Solicitudes anexos', anexoId, ['ID Solicitudes anexos']).obj;
  const file = _resolveDriveFile(_getField(row, ['Archivo', 'Foto', 'Dibujo']));
  if (!file) throw new Error("Archivo no hallado en Drive.");
  return { mode: "url", url: file.getUrl(), fileName: file.getName() };
}

/** Sirve una página HTML de descarga/visualización (Proxy Mode). */
function _renderFileView(anexoId) {
  try {
    const row = _findRowObjectByKey('Solicitudes anexos', anexoId, ['ID Solicitudes anexos']).obj;
    const file = _resolveDriveFile(_getField(row, ['Archivo', 'Foto', 'Dibujo']));
    if (!file) return HtmlService.createHtmlOutput("Archivo no disponible.");

    const blob = file.getBlob();
    const html = `
      <body style="font-family:sans-serif; text-align:center; padding:50px; background:#f3f4f6;">
        <div style="background:white; padding:30px; border-radius:12px; box-shadow:0 4px 12px rgba(0,0,0,0.1); display:inline-block;">
          <h3>G4S - ${file.getName()}</h3>
          <p>Preparando descarga...</p>
          <button id="dl" style="background:#D32F2F; color:white; border:none; padding:10px 24px; border-radius:6px; cursor:pointer;">Guardar Archivo</button>
        </div>
        <script>
          const b = new Blob([new Uint8Array(atob("${Utilities.base64Encode(blob.getBytes())}").split("").map(c => c.charCodeAt(0)))], {type: "${blob.getContentType()}"});
          const u = URL.createObjectURL(b);
          document.getElementById('dl').onclick = () => { const a = document.createElement('a'); a.href = u; a.download = "${file.getName()}"; a.click(); };
          setTimeout(() => document.getElementById('dl').click(), 1000);
        </script>
      </body>`;
    return HtmlService.createHtmlOutput(html);
  } catch (e) { return HtmlService.createHtmlOutput("Error crítico del visor."); }
}

/**
 * ------------------------------------------------------------------
 * GESTIÓN DE ACTIVOS (QR)
 * ------------------------------------------------------------------
 */

/** Vincula un activo a una solicitud. */
function createSolicitudActivo(email, payload) {
  return _withLock(() => {
    const row = {
      "ID Solicitudes activos": Utilities.getUuid(), "ID Solicitudes": payload.solicitudId,
      "QR": payload.qrSerial, "ID Activo": payload.idActivo, "Observaciones": payload.observaciones,
      "Usuario Actualización": email, "Fecha Actualización": new Date()
    };
    appendDataToSheet('Solicitudes activos', row);
    _invalidateDetailCache(email, payload.solicitudId);
    return { success: true };
  });
}

/** Obtiene activos vinculados a un ticket. */
function getSolicitudActivos(email, { solicitudId }) {
  return { data: _getChildrenFast('Solicitudes activos', [solicitudId]), total: 0 };
}

/** Catálogo maestro de activos. */
function getActivosCatalog(email) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("activos_catalog_v2");
  if (cached) return JSON.parse(cached);

  const data = getDataFromSheet('Activos').map(r => ({
    idActivo: String(_getField(r, ['ID Activo'])), nombreActivo: String(_getField(r, ['Nombre Activo'])),
    qrSerial: String(_getField(r, ['QR Serial'])), nombreUbicacion: String(_getField(r, ['Nombre Ubicacion'])),
    estadoActivo: String(_getField(r, ['Estado Activo'])), funcionamiento: String(_getField(r, ['Funcionamiento']))
  }));
  const res = { data, total: data.length };
  cache.put("activos_catalog_v2", JSON.stringify(res), 600);
  return res;
}

/** Busca activo por QR. */
function getActivoByQr(email, payload) {
  const found = getDataFromSheet('Activos').find(r => String(_getField(r, ['QR Serial', 'QR'])).trim() === String(payload.qr).trim());
  return found ? { found: true, activo: { idActivo: _getField(found, ['ID Activo']), nombreActivo: _getField(found, ['Nombre Activo']), qrSerial: payload.qr } } : { found: false };
}

/**
 * Sincronización masiva para modo offline.
 * @param {string} email
 * @param {Object} payload { ids: string[] }
 * @return {Object} Mapa de detalles por ticket.
 */
function getBatchRequestDetails(email, { ids }) {
  const targetIds = new Set(ids);
  const result = {};
  targetIds.forEach(id => { result[id] = { services: [], history: [], documents: [], activos: [] }; });

  const tables = { 'Observaciones historico': 'services', 'Estados historico': 'history', 'Solicitudes anexos': 'documents', 'Solicitudes activos': 'activos' };
  Object.keys(tables).forEach(t => {
    getDataFromSheet(t).forEach(row => {
      const pId = String(_getField(row, ['ID Solicitudes', 'ID Solicitud'])).trim();
      if (targetIds.has(pId)) result[pId][tables[t]].push(row);
    });
  });
  return result;
}

/** Clasificaciones dinámicas para el formulario. */
function getClassificationOptions(email) {
  return ["Visita técnica", "Visita comercial", "Soporte Remoto", "Mantenimiento"];
}

/**
 * ------------------------------------------------------------------
 * HELPER DE ACCESO A DATOS (BAJO NIVEL)
 * ------------------------------------------------------------------
 */

/** Lee una tabla completa como array de objetos. */
function getDataFromSheet(sheetName) {
  const { sheet, headers, lastRow, lastCol } = _getSheetInfo(sheetName);
  return (!sheet || lastRow < 2) ? [] : sheet.getRange(2, 1, lastRow - 1, lastCol).getValues().map(r => _rowToObject(headers, r));
}

/** Inserta un objeto como nueva fila respetando los encabezados. */
function appendDataToSheet(sheetName, objectData) {
  const { sheet, headers } = _getSheetInfo(sheetName);
  if (!sheet) return;
  const row = headers.map(h => {
    const val = objectData[h] ?? objectData[Object.keys(objectData).find(k => _normHeader(k) === _normHeader(h))];
    return val ?? "";
  });
  sheet.appendRow(row);
  return { success: true };
}

/** Extrae un valor de un objeto probando varios nombres de propiedad. */
function _getField(row, candidates) {
  const keys = Object.keys(row || {});
  for (const c of candidates) {
    const k = keys.find(x => _normHeader(x) === _normHeader(c));
    if (k && row[k]) return row[k];
  }
  return "";
}

/** Shortcut para buscar el encabezado de solicitud. */
function _findSolicitudHeaderFast(key) {
  return _findRowObjectByKey('Solicitudes', key, ['ID Solicitud', 'Ticket G4S', 'Ticket Cliente']);
}

/**
 * Comunicación con AppSheet API para disparar eventos.
 * @param {string} tableName
 * @param {Object} rowData
 * @return {?Object}
 */
function enviarAppSheetAPI(tableName, rowData) {
  const appId = "c0817cfb-b068-4a46-ae3b-228c0385a486"; 
  const accessKey = "V2-gaw9Q-LcMsx-wfJof-pFCgC-u6igd-FMxtR-23Zr1-V3O4K"; 
  const url = `https://api.appsheet.com/api/v1/apps/${appId}/tables/${tableName}/Action`;
  
  const options = {
    "method": "post", "contentType": "application/json",
    "headers": { "ApplicationAccessKey": accessKey },
    "payload": JSON.stringify({ "Action": "Add", "Properties": { "Locale": "es-CO" }, "Rows": [ rowData ] }),
    "muteHttpExceptions": true
  };
  try {
    const res = UrlFetchApp.fetch(url, options);
    return JSON.parse(res.getContentText());
  } catch (e) { return null; }
}
