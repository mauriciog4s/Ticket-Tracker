/**
 * G4S TICKET TRACKER - BACKEND (Google Apps Script)
 *
 * Este archivo contiene la lógica del servidor para gestionar tickets,
 * anexos, activos y permisos, utilizando Google Sheets como base de datos.
 */

/**
 * ------------------------------------------------------------------
 * CONFIGURACIÓN Y MAPEO DE HOJAS
 * ------------------------------------------------------------------
 */

// IDs de los Spreadsheets que actúan como base de datos
const MAIN_SPREADSHEET_ID = '1MC76eZZt7qiso2M8LMz777_xJnzrl_ZpZptDZBnPlDo'; 
const PERMISSIONS_SPREADSHEET_ID = '1zcZZGe_93ytWXtCF1kmk_Y8zc5b5cL1xH34i7v1w01k'; 
const CLIENTS_SPREADSHEET_ID = '1hHWPJF9KSC0opplpCNgHRNkW6CLf7StXG2Y31m6yUpo'; 
const SEDES_SPREADSHEET_ID = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM'; 
const ACTIVOS_SPREADSHEET_ID = '1JU8c1MidgV4DRFg6W-GxZ2tHkfKNqGt1_cR5VDTehC4'; 

// Mapeo de tablas a sus respectivos Spreadsheets
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
 * OPTIMIZACIÓN: Gestión de Memoria y Caché de Ejecución
 * ------------------------------------------------------------------
 */

// Caché en memoria para evitar aperturas repetidas de Spreadsheets
const __SS_MEMO = {}; 
function _openSS(spreadsheetId) {
  if (!__SS_MEMO[spreadsheetId]) __SS_MEMO[spreadsheetId] = SpreadsheetApp.openById(spreadsheetId);
  return __SS_MEMO[spreadsheetId];
}

// Caché en memoria para metadatos de las hojas (encabezados, índices)
const __SHEET_INFO_MEMO = {};

/**
 * Normaliza nombres de encabezados para búsquedas insensibles a mayúsculas/espacios.
 */
function _normHeader(x) { return String(x || "").trim().toLowerCase(); }

/**
 * Obtiene información estructural de una hoja (hoja, encabezados, última fila/columna).
 */
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

/**
 * ------------------------------------------------------------------
 * HELPERS DE DATOS Y BÚSQUEDA
 * ------------------------------------------------------------------
 */

/**
 * Encuentra el índice de una columna basándose en nombres candidatos.
 */
function _findColIndex(headersNorm, candidateNames) {
  for (let i = 0; i < candidateNames.length; i++) {
    const idx = headersNorm.indexOf(_normHeader(candidateNames[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Convierte una fila de datos en un objeto JS usando los encabezados como llaves.
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
 * Agrupa números de fila consecutivos para realizar lecturas por bloques (optimización de red).
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
 * Lee bloques de filas desde la hoja.
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
 * Busca un objeto (fila) en una tabla por una llave y valor específicos.
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
 * Obtiene registros hijos (ej. observaciones o anexos) filtrados por IDs de solicitud.
 * Optimizado para lectura rápida por bloques.
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

  // Ordenar por fecha descendente
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
 * Ejecuta un callback protegiéndolo con un bloqueo de script para evitar condiciones de carrera.
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
    throw new Error("El servidor está ocupado. Por favor, intente de nuevo en unos segundos.");
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
 * Maneja las peticiones GET (Carga de la página y Visualización de archivos).
 */
function doGet(e) {
  // ROUTER DE ARCHIVOS (MODO PROXY para visualizar anexos sin problemas de CORS)
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
 * Resuelve el correo del llamante, ya sea por sesión activa o por payload (fallback).
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
 * Manejador central de peticiones desde el frontend (google.script.run).
 */
function apiHandler(request) {
  const userEmail = _resolveCallerEmail(request);
  const { endpoint, payload } = request || {};

  console.log(`🔒 [API] Endpoint: ${endpoint} | Usuario: ${userEmail}`);

  try {
    if (!userEmail) throw new Error("No se pudo verificar su identidad corporativa.");

    switch (endpoint) {
      // Contexto y Tickets
      case 'getUserContext': return getUserContext(userEmail, payload?.ignoreCache);
      case 'getRequests': return getRequests(userEmail);
      case 'getRequestDetail': return getRequestDetail(userEmail, payload);
      case 'getBatchRequestDetails': return getBatchRequestDetails(userEmail, payload);
      
      // Creación y Subida
      case 'createRequest': return createRequest(userEmail, payload);
      case 'uploadAnexo': return uploadAnexo(userEmail, payload);
      case 'getAnexoDownload': return getAnexoDownload(userEmail, payload);

      // Activos (QR)
      case 'createSolicitudActivo': return createSolicitudActivo(userEmail, payload);
      case 'getSolicitudActivos': return getSolicitudActivos(userEmail, payload);
      case 'getActivosCatalog': return getActivosCatalog(userEmail);
      case 'getActivoByQr': return getActivoByQr(userEmail, payload);
      
      // Utilidades
      case 'getClassificationOptions': return getClassificationOptions(userEmail);

      default: throw new Error(`El endpoint '${endpoint}' no está definido.`);
    }

  } catch (err) {
    console.error(`❌ ERROR API [${endpoint}]: ${err.message}`, err);
    return { error: true, message: err.message || "Error interno del servidor. Contacte a soporte." };
  }
}

/**
 * ------------------------------------------------------------------
 * LÓGICA DE NEGOCIO Y DATOS
 * ------------------------------------------------------------------
 */

// Gestión de Caché de Detalles
const DETAIL_CACHE_VER = "v3"; 
function _detailCacheKey(email, id) {
  const e = String(email || "").toLowerCase().trim();
  const rid = String(id || "").trim();
  return `detail_${DETAIL_CACHE_VER}_${Utilities.base64Encode(e)}_${rid}`;
}
function _invalidateDetailCache(email, id) {
  try { CacheService.getScriptCache().remove(_detailCacheKey(email, id)); } catch (e) {}
}

/**
 * Obtiene el contexto del usuario: rol, clientes y sedes permitidas.
 */
function getUserContext(email, ignoreCache = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `ctx_it_v5_${Utilities.base64Encode(email)}`; 
  
  if (!ignoreCache) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData);
  }

  try {
    let context = {
      email: email,
      role: 'Usuario',
      allowedClientIds: [],
      clientNames: {},
      assignedCustomerNames: [],
      isValidUser: false,
      isAdmin: false
    };

    // 1. Verificar permisos base
    const allPermissions = getDataFromSheet('Permisos');
    const userData = allPermissions.find(row => String(row['Correo']).toLowerCase() === email.toLowerCase());

    if (userData) {
      context.isValidUser = true;
      const rol = (userData['Rol_Asignado'] || '').trim().toLowerCase();
      if (rol === 'administrador') {
        context.role = 'Administrador';
        context.isAdmin = true;
      }
    }

    if (!context.isValidUser) return context;

    // 2. Resolver Clientes y Sedes (Filtros)
    const allRelations = getDataFromSheet('Usuarios filtro');
    const myRelations = allRelations.filter(row => String(row['Usuario']).toLowerCase() === email.toLowerCase());

    const assignedClientIds = [];
    myRelations.forEach(row => {
      const id = row['Cliente'];
      if (id) assignedClientIds.push(String(id));
    });

    if (assignedClientIds.length > 0) {
      // Nombres de Clientes
      const allClientes = getDataFromSheet('Clientes');
      const myClients = allClientes.filter(c => 
        assignedClientIds.includes(String(_getField(c, ['ID Cliente', 'Id Cliente', 'Cliente'])))
      );

      myClients.forEach(c => {
        const clientName = _getField(c, ['Nombre cliente', 'Nombre Cliente', 'Nombre', 'RazonSocial']);
        if (clientName) context.assignedCustomerNames.push(String(clientName).trim());
      });

      // Sedes (Hijas de Clientes)
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
    console.error("Error en getUserContext:", e);
    throw e;
  }
}

/**
 * Obtiene el listado de solicitudes filtrado por los permisos del usuario.
 */
function getRequests(email) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso denegado: Usuario no registrado.");

  try {
    let allRows = getDataFromSheet('Solicitudes');
    let filteredRows = [];

    if (context.isAdmin) {
      filteredRows = allRows;
    } else {
      if (context.allowedClientIds.length === 0) return { data: [], total: 0 };
      filteredRows = allRows.filter(row => context.allowedClientIds.includes(String(_getField(row, ['ID Sede']))));
    }

    // Ordenar por fecha de creación descendente
    filteredRows.sort((a, b) => {
      const dateA = new Date(_getField(a, ['Fecha creación cliente', 'Fecha creacion cliente'])).getTime() || 0;
      const dateB = new Date(_getField(b, ['Fecha creación cliente', 'Fecha creacion cliente'])).getTime() || 0;
      return dateB - dateA;
    });

    return { data: filteredRows, total: filteredRows.length };
  } catch (e) {
    console.error("Error en getRequests:", e);
    throw new Error("No se pudieron obtener las solicitudes.");
  }
}

/**
 * Obtiene el detalle de un ticket específico (header, observaciones, historial, anexos).
 */
function getRequestDetail(email, { id }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso denegado.");
  if (!id) throw new Error("Se requiere un ID de solicitud.");

  const rid = String(id).trim();
  const cache = CacheService.getScriptCache();
  const ck = _detailCacheKey(email, rid);

  const cached = cache.get(ck);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const headerFound = _findSolicitudHeaderFast(rid);
  if (!headerFound) throw new Error("El ticket solicitado no existe.");
  const header = headerFound.obj;

  // Validar permisos de sede
  if (!context.isAdmin) {
    const recordSedeId = String(_getField(header, ['ID Sede'])).trim();
    if (recordSedeId && !context.allowedClientIds.includes(recordSedeId)) {
      throw new Error("No tiene permisos para ver este ticket específico.");
    }
  }

  const parentKeys = [
    rid,
    String(_getField(header, ['Ticket G4S'])),
    String(_getField(header, ['Ticket Cliente', 'Ticket (Opcional)']))
  ].filter(x => x && x !== "undefined" && x !== "null").map(x => String(x).trim());

  const result = {
    header,
    services: _getChildrenFast('Observaciones historico', parentKeys),
    history: _getChildrenFast('Estados historico', parentKeys),
    documents: _getChildrenFast('Solicitudes anexos', parentKeys),
    activos: _getChildrenFast('Solicitudes activos', parentKeys)
  };

  const json = JSON.stringify(result);
  if (json.length < 90000) cache.put(ck, json, 30);

  return result;
}

/**
 * Crea una nueva solicitud de servicio.
 */
function createRequest(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso denegado.");

  if (!payload?.idSede || !payload?.solicitud || !payload?.observacion) {
    throw new Error("Por favor, complete todos los campos obligatorios.");
  }

  if (!context.isAdmin && !context.allowedClientIds.includes(String(payload.idSede))) {
    throw new Error("No tiene permisos para reportar en esta sede.");
  }

  return _withLock(() => {
    const now = new Date();
    const uuid = Utilities.getUuid();

    // Lógica de prefijo de Ticket G4S basado en el cliente
    const allSedes = getDataFromSheet('Sedes');
    const sedeInfo = allSedes.find(s => String(_getField(s, ['ID Sede', 'Id Sede', 'Sede'])).trim() === String(payload.idSede).trim());
    const idCliente = sedeInfo ? _getField(sedeInfo, ['ID Cliente', 'Id Cliente', 'Cliente']) : null;

    let letraInicial = "G"; // Default G4S
    if (idCliente) {
      const allClientes = getDataFromSheet('Clientes');
      const clienteInfo = allClientes.find(c => String(_getField(c, ['ID Cliente', 'Id Cliente', 'Cliente'])).trim() === String(idCliente).trim());
      if (clienteInfo) {
        const nombreCorto = _getField(clienteInfo, ['Nombre corto', 'Nombre_Corto', 'RazonSocial']) || "G";
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

    // Integrar con AppSheet API para disparar flujos automáticos (correos, etc.)
    const apiResult = enviarAppSheetAPI('Solicitudes', newRow);

    // Respaldo: si la API falla, guardamos directamente en la hoja
    if (!apiResult || (apiResult && apiResult.RestServerErrorMessage)) {
      console.warn("API de AppSheet falló o no disponible. Guardando directo.");
      appendDataToSheet('Solicitudes', newRow);
    }

    // Historial inicial
    try {
      appendDataToSheet('Estados historico', {
        "ID Estado": Utilities.getUuid(),
        "ID Solicitudes": uuid,
        "Estado actual": "Creado",
        "Usuario Actualización": email,
        "Fecha Actualización": now
      });
    } catch (e) { console.warn("Error guardando historial inicial:", e); }

    _invalidateDetailCache(email, uuid);

    const returnRow = { ...newRow };
    if (returnRow["Fecha creación cliente"] instanceof Date) returnRow["Fecha creación cliente"] = returnRow["Fecha creación cliente"].toISOString();

    return {
      success: true,
      solicitudId: uuid,
      ticketG4S: ticketG4S,
      row: returnRow
    };
  });
}

/**
 * ------------------------------------------------------------------
 * GESTIÓN DE ARCHIVOS Y ANEXOS
 * ------------------------------------------------------------------
 */

function _sanitizeFileName(name) {
  const n = String(name || 'anexo').trim();
  return n.normalize("NFD").replace(/[\u0300-\u036f]/g, "")
          .replace(/[/\\]/g, '_')
          .replace(/[<>:"|?*]/g, '')
          .replace(/\s+/g, '_')
          .slice(0, 120) || 'anexo';
}

/**
 * Sube un anexo (Foto, Dibujo o Archivo) a Drive y registra el enlace en Sheets.
 */
function uploadAnexo(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso denegado.");

  // Carpetas fijas de almacenamiento
  const IMAGES_FOLDER_ID = '1tzYk9jiQ7Lp_bSZylMn0vuzfWZt4xTHb';
  const DOCS_FOLDER_ID   = '1-CBsinL67dJUPfr8WXtKP1wM93B6zPDX';

  const solicitudId = String(payload?.solicitudId || '').trim();
  const tipoAnexo = String(payload?.tipoAnexo || 'Archivo').trim();
  const base64 = String(payload?.base64 || '').trim();

  if (!solicitudId || !base64) throw new Error("Faltan datos para la subida (solicitudId/base64).");

  const headerFound = _findRowObjectByKey('Solicitudes', solicitudId, ['ID Solicitud', 'ID Solicitudes']);
  if (!headerFound) throw new Error("La solicitud padre no existe.");
  
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > 10 * 1024 * 1024) throw new Error("El archivo excede el límite de 10MB.");

  const safeFileName = _sanitizeFileName(payload?.fileName || 'anexo');
  const shortId = solicitudId.replace(/-/g, '').slice(0, 8);
  const rand = Math.floor(Math.random() * 900000) + 100000;
  const ext = (payload?.mimeType || "").includes('image') ? 'jpg' : 'pdf';
  const finalName = `${shortId}_${tipoAnexo}_${rand}_${safeFileName.split('.')[0]}.${ext}`;

  const blob = Utilities.newBlob(bytes, payload?.mimeType || 'application/octet-stream', finalName);

  return _withLock(() => {
    let file;
    let storedPath;

    try {
      const isImg = tipoAnexo === 'Foto' || tipoAnexo === 'Dibujo' || (payload?.mimeType || "").startsWith('image/');
      const targetFolder = DriveApp.getFolderById(isImg ? IMAGES_FOLDER_ID : DOCS_FOLDER_ID);
      file = targetFolder.createFile(blob);

      // Formato de ruta compatible con la lógica del visor
      storedPath = isImg ? `${targetFolder.getName()}/${finalName}` : `https://drive.google.com/file/d/${file.getId()}/view`;

      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    } catch (e) {
      throw new Error(`Error al acceder a Drive: ${e.message}`);
    }

    const anexoUuid = Utilities.getUuid();
    const row = {
      "ID Solicitudes anexos": anexoUuid,
      "ID Solicitudes": solicitudId,
      "Tipo anexo": tipoAnexo,
      "Nombre": safeFileName,
      "Usuario Actualización": email,
      "Fecha Actualización": new Date()
    };

    if (tipoAnexo === 'Foto') row['Foto'] = storedPath;
    else if (tipoAnexo === 'Dibujo') row['Dibujo'] = storedPath;
    else row['Archivo'] = storedPath;

    appendDataToSheet('Solicitudes anexos', row);
    _invalidateDetailCache(email, solicitudId);

    return { success: true, anexoId: anexoUuid, path: storedPath };
  });
}

/**
 * Resuelve un archivo físico en Drive basándose en una ruta (AppSheet o directa).
 */
function _resolveDriveFile(pathValue) {
  const p = String(pathValue || "").trim();
  if (!p) return null;

  // 1. Caso ID de Drive directo en URL
  const idMatch = p.match(/\/d\/([a-zA-Z0-9_-]+)/) || p.match(/id=([a-zA-Z0-9_-]+)/);
  if (idMatch && idMatch[1]) {
    try { return DriveApp.getFileById(idMatch[1]); } catch(e) {}
  }

  // 2. Caso Nombre de Archivo (Rutas relativas de AppSheet)
  const parts = p.split('/');
  const fileName = parts[parts.length - 1];
  if (fileName) {
    const filesIt = DriveApp.getFilesByName(fileName);
    if (filesIt.hasNext()) return filesIt.next();
  }

  return null;
}

/**
 * Endpoint para obtener la URL de descarga o visualización de un anexo.
 */
function getAnexoDownload(email, { anexoId }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso denegado.");

  const found = _findRowObjectByKey('Solicitudes anexos', anexoId, ['ID Solicitudes anexos', 'ID Anexo', 'ID']);
  if (!found) throw new Error("Anexo no encontrado en la base de datos.");

  const pathValue = _getField(found.obj, ['Archivo', 'Foto', 'Dibujo']) || "";
  const file = _resolveDriveFile(pathValue);

  if (!file) throw new Error("No se pudo localizar el archivo físico en Google Drive.");

  return {
    mode: "url",
    url: `https://drive.google.com/file/d/${file.getId()}/view`,
    fileName: file.getName()
  };
}

/**
 * Sirve el visor de archivos (Modo Proxy) para evitar bloqueos del navegador.
 */
function _renderFileView(anexoId) {
  try {
    const found = _findRowObjectByKey('Solicitudes anexos', anexoId, ['ID Solicitudes anexos', 'ID Anexo', 'ID']);
    if (!found) return HtmlService.createHtmlOutput("<h1>Anexo no registrado.</h1>");

    const pathValue = _getField(found.obj, ['Archivo', 'Foto', 'Dibujo']) || "";
    const file = _resolveDriveFile(pathValue);

    if (!file) return HtmlService.createHtmlOutput("<h1>Archivo no disponible en Drive.</h1>");

    if (file.getSize() > 8 * 1024 * 1024) {
      return HtmlService.createHtmlOutput(`
        <div style="font-family:sans-serif;text-align:center;padding-top:50px;">
          <h2>Archivo demasiado grande para previsualizar</h2>
          <p>Tamaño: ${(file.getSize()/(1024*1024)).toFixed(2)}MB</p>
          <a href="${file.getDownloadUrl()}" style="background:#0033A0;color:white;padding:12px 24px;text-decoration:none;border-radius:4px;">Descargar Directamente</a>
        </div>
      `);
    }

    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    // Generar página minimalista de descarga automática
    const html = `
      <!DOCTYPE html>
      <html>
      <body style="font-family:sans-serif;text-align:center;padding:50px;background:#f3f4f6;">
        <div style="background:white;padding:30px;border-radius:8px;display:inline-block;box-shadow:0 2px 10px rgba(0,0,0,0.1);">
          <h3>Procesando: ${file.getName()}</h3>
          <p>Su descarga comenzará en breve...</p>
          <button id="btn" style="background:#D32F2F;color:white;border:none;padding:10px 20px;border-radius:4px;cursor:pointer;font-weight:bold;">Guardar ahora</button>
        </div>
        <script>
          const blob = new Blob([new Uint8Array(atob("${base64}").split("").map(c => c.charCodeAt(0)))], {type: "${blob.getContentType()}"});
          const url = URL.createObjectURL(blob);
          const save = () => { const a = document.createElement('a'); a.href = url; a.download = "${file.getName()}"; a.click(); };
          document.getElementById('btn').onclick = save;
          setTimeout(save, 1000);
        </script>
      </body>
      </html>
    `;
    return HtmlService.createHtmlOutput(html);
  } catch (e) {
    return HtmlService.createHtmlOutput(`<h3>Error crítico del visor: ${e.message}</h3>`);
  }
}

/**
 * ------------------------------------------------------------------
 * GESTIÓN DE ACTIVOS (QR)
 * ------------------------------------------------------------------
 */

function createSolicitudActivo(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso denegado.");

  const solicitudId = String(payload?.solicitudId || payload?.idSolicitud || '').trim();
  if (!solicitudId) throw new Error("ID de solicitud requerido.");

  return _withLock(() => {
    const rowId = Utilities.getUuid();
    const row = {
      "ID Solicitudes activos": rowId,
      "ID Solicitudes": solicitudId,
      "QR": String(payload?.qrSerial || '').trim(),
      "ID Activo": String(payload?.idActivo || '').trim(),
      "Observaciones": String(payload?.observaciones || '').trim(),
      "Usuario Actualización": email,
      "Fecha Actualización": new Date()
    };

    appendDataToSheet('Solicitudes activos', row);
    _invalidateDetailCache(email, solicitudId);
    return { success: true, rowId };
  });
}

function getSolicitudActivos(email, { solicitudId }) {
  if (!getUserContext(email).isValidUser) throw new Error("Acceso denegado.");
  const rows = _getChildrenFast('Solicitudes activos', [String(solicitudId).trim()]);
  return { data: rows, total: rows.length };
}

function getActivosCatalog(email) {
  if (!getUserContext(email).isValidUser) throw new Error("Acceso denegado.");

  const cache = CacheService.getScriptCache();
  const cached = cache.get("activos_catalog_v2");
  if (cached) return JSON.parse(cached);

  const rows = getDataFromSheet('Activos');
  const mapped = rows.map(r => ({
    idActivo: String(_getField(r, ['ID Activo'])).trim(),
    nombreActivo: String(_getField(r, ['Nombre Activo'])).trim(),
    qrSerial: String(_getField(r, ['QR Serial'])).trim(),
    nombreUbicacion: String(_getField(r, ['Nombre Ubicacion'])).trim(),
    estadoActivo: String(_getField(r, ['Estado Activo'])).trim(),
    funcionamiento: String(_getField(r, ['Funcionamiento'])).trim()
  })).filter(x => x.idActivo || x.qrSerial);

  const res = { data: mapped, total: mapped.length };
  cache.put("activos_catalog_v2", JSON.stringify(res), 600);
  return res;
}

function getActivoByQr(email, payload) {
  if (!getUserContext(email).isValidUser) throw new Error("Acceso denegado.");
  const q = String(payload?.qr || '').trim();
  const rows = getDataFromSheet('Activos');
  const found = rows.find(r => String(_getField(r, ['QR Serial', 'QR'])).trim() === q);

  if (!found) return { found: false };
  return {
    found: true,
    activo: {
      idActivo: String(_getField(found, ['ID Activo'])).trim(),
      nombreActivo: String(_getField(found, ['Nombre Activo'])).trim(),
      qrSerial: q,
      nombreUbicacion: String(_getField(found, ['Nombre Ubicacion'])).trim()
    }
  };
}

/**
 * ------------------------------------------------------------------
 * SINCRONIZACIÓN MASIVA (MODO OFFLINE)
 * ------------------------------------------------------------------
 */

function getBatchRequestDetails(email, { ids }) {
  if (!getUserContext(email).isValidUser) throw new Error("Acceso denegado.");
  if (!ids || !ids.length) return {};

  const targetIds = new Set(ids.map(x => String(x).trim()));
  const tables = ['Observaciones historico', 'Estados historico', 'Solicitudes anexos', 'Solicitudes activos'];
  const result = {};
  targetIds.forEach(id => { result[id] = { services: [], history: [], documents: [], activos: [] }; });

  const mapKey = { 'Observaciones historico': 'services', 'Estados historico': 'history', 'Solicitudes anexos': 'documents', 'Solicitudes activos': 'activos' };

  tables.forEach(table => {
    const rows = getDataFromSheet(table);
    rows.forEach(row => {
      const parentId = String(_getField(row, ['ID Solicitudes', 'ID Solicitud'])).trim();
      if (parentId && targetIds.has(parentId)) {
        result[parentId][mapKey[table]].push(row);
      }
    });
  });

  return result;
}

function getClassificationOptions(email) {
  if (!getUserContext(email).isValidUser) throw new Error("Acceso denegado.");
  return ["Visita técnica", "Visita comercial", "Soporte Remoto", "Mantenimiento"];
}

/**
 * ------------------------------------------------------------------
 * HELPER DE BAJO NIVEL: ACCESO A HOJAS
 * ------------------------------------------------------------------
 */

function getDataFromSheet(sheetName) {
  const { sheet, headers, lastRow, lastCol } = _getSheetInfo(sheetName);
  if (!sheet || lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.map(row => _rowToObject(headers, row));
}

function appendDataToSheet(sheetName, objectData) {
  const { sheet, headers, lastCol } = _getSheetInfo(sheetName);
  if (!sheet) throw new Error(`Hoja '${sheetName}' no encontrada.`);

  const rowArray = headers.map(header => {
    let val = objectData[header];
    if (val === undefined) {
      const norm = _normHeader(header);
      const key = Object.keys(objectData).find(k => _normHeader(k) === norm);
      val = key ? objectData[key] : "";
    }
    return val ?? "";
  });

  sheet.appendRow(rowArray);
  return { success: true };
}

function _getField(row, candidateNames) {
  if (!row) return "";
  const keys = Object.keys(row);
  for (const c of candidateNames) {
    const k = keys.find(x => _normHeader(x) === _normHeader(c));
    if (k && row[k] !== undefined && row[k] !== null && row[k] !== "") return row[k];
  }
  return "";
}

/**
 * ------------------------------------------------------------------
 * INTEGRACIÓN CON APPSHEET API
 * ------------------------------------------------------------------
 */
function enviarAppSheetAPI(tableName, rowData) {
  /**
   * 🔒 NOTA DE SEGURIDAD:
   * Se recomienda mover 'appId' y 'accessKey' a Propiedades del Script (ScriptProperties)
   * para evitar exponer llaves en el código fuente.
   * Ejemplo: PropertiesService.getScriptProperties().getProperty('APPSHEET_KEY');
   */
  const appId = "c0817cfb-b068-4a46-ae3b-228c0385a486"; 
  const accessKey = "V2-gaw9Q-LcMsx-wfJof-pFCgC-u6igd-FMxtR-23Zr1-V3O4K"; 
  const url = `https://api.appsheet.com/api/v1/apps/${appId}/tables/${tableName}/Action`;
  
  const payload = {
    "Action": "Add",
    "Properties": { "Locale": "es-CO", "Timezone": "SA Pacific Standard Time", "RunAsUserEmail": rowData["Usuario Actualización"] },
    "Rows": [ rowData ]
  };
  
  try {
    const response = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "headers": { "ApplicationAccessKey": accessKey },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    });
    return JSON.parse(response.getContentText());
  } catch (e) {
    console.error("Error AppSheet API:", e);
    return null;
  }
}
