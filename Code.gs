/**
 * ------------------------------------------------------------------
 * CONFIGURACIÓN Y MAPEO DE HOJAS
 * ------------------------------------------------------------------
 */

// ⚠️ REEMPLAZA ESTOS IDs CON LOS TUYOS:
const MAIN_SPREADSHEET_ID = '1MC76eZZt7qiso2M8LMz777_xJnzrl_ZpZptDZBnPlDo'; 
const PERMISSIONS_SPREADSHEET_ID = '1zcZZGe_93ytWXtCF1kmk_Y8zc5b5cL1xH34i7v1w01k'; 
const CLIENTS_SPREADSHEET_ID = '1hHWPJF9KSC0opplpCNgHRNkW6CLf7StXG2Y31m6yUpo'; 
const SEDES_SPREADSHEET_ID = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM'; 

const SHEET_CONFIG = {
  'Solicitudes': MAIN_SPREADSHEET_ID,
  'Estados historico': MAIN_SPREADSHEET_ID,
  'Observaciones historico': MAIN_SPREADSHEET_ID,
  'Solicitudes anexos': MAIN_SPREADSHEET_ID,
  'Permisos': PERMISSIONS_SPREADSHEET_ID,
  'Usuarios filtro': PERMISSIONS_SPREADSHEET_ID,
  'Clientes': CLIENTS_SPREADSHEET_ID,
  'Sedes': SEDES_SPREADSHEET_ID
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('G4S Ticket Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function apiHandler(request) {
  const userEmail = Session.getActiveUser().getEmail();
  const { endpoint, payload } = request;
  console.log(`🔒 [API] ${endpoint} | User: ${userEmail}`);
  try {
    if (!userEmail) throw new Error("Usuario no identificado.");
    switch (endpoint) {
      case 'getUserContext': return getUserContext(userEmail);
      case 'getRequests': return getRequests(userEmail);
      case 'getRequestDetail': return getRequestDetail(userEmail, payload);
      case 'createRequest': return createRequest(userEmail, payload);
      default: throw new Error(`Endpoint desconocido: ${endpoint}`);
    }
  } catch (err) {
    console.error(`❌ Error: ${err.message}`);
    return { error: true, message: err.message };
  }
}

// --- HELPERS ---

function getDataFromSheet(sheetName) {
  const spreadsheetId = SHEET_CONFIG[sheetName];
  if (!spreadsheetId) throw new Error(`Configuración faltante: ${sheetName}`);

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(h => h.toString().trim());
  const data = values.slice(1);

  return data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] === undefined ? "" : row[index];
    });
    return obj;
  });
}

function appendDataToSheet(sheetName, objectData) {
  const spreadsheetId = SHEET_CONFIG[sheetName];
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Hoja ${sheetName} no encontrada.`);

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowArray = headers.map(header => objectData[header] || "");

  sheet.appendRow(rowArray);
  return { success: true };
}

// --- LOGICA ---

function getUserContext(email) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `ctx_v6_${Utilities.base64Encode(email)}`; 
  const cachedData = cache.get(cacheKey);
  if (cachedData) return JSON.parse(cachedData);

  try {
    let context = { email: email, role: 'Usuario', allowedClientIds: [], clientNames: {}, isValidUser: false, isAdmin: false };
    const allPermissions = getDataFromSheet('Permisos');
    const userData = allPermissions.find(row => String(row['Correo']).trim().toLowerCase() === email.toLowerCase());

    if (userData) {
      context.isValidUser = true;
      if ((userData['Rol_Asignado'] || '').trim().toLowerCase() === 'administrador') { context.role = 'Administrador'; context.isAdmin = true; }
    }

    if (context.isValidUser) {
      const allRelations = getDataFromSheet('Usuarios filtro');
      const myRelations = allRelations.filter(row => String(row['Usuario']).trim().toLowerCase() === email.toLowerCase());
      const assignedClientIds = [];
      myRelations.forEach(row => { if (row['Cliente']) assignedClientIds.push(String(row['Cliente'])); });

      if (assignedClientIds.length > 0) {
        const allSedes = getDataFromSheet('Sedes');
        const mySedes = allSedes.filter(sede => assignedClientIds.includes(String(sede['ID Cliente'])));
        mySedes.forEach(sede => {
          const idSede = String(sede['ID Sede']);
          const nombreSede = sede['Nombre'] || sede['Nombre_Sede'] || sede['Sede'] || sede['Nombre Sede'] || idSede;
          if (idSede) { context.allowedClientIds.push(idSede); context.clientNames[idSede] = nombreSede; }
        });
      }
    }
    cache.put(cacheKey, JSON.stringify(context), 600);
    return context;
  } catch (e) { console.error("Error Context", e); throw e; }
}

function getRequests(email) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  try {
    let allRows = getDataFromSheet('Solicitudes');
    let filteredRows = [];

    if (context.isAdmin) {
      filteredRows = allRows;
    } else {
      if (context.allowedClientIds.length === 0) return { data: [], total: 0 };
      filteredRows = allRows.filter(row => context.allowedClientIds.includes(String(row['ID Sede'])));
    }
    filteredRows.sort((a, b) => {
      const dateA = new Date(a['Fecha creación cliente']).getTime() || 0;
      const dateB = new Date(b['Fecha creación cliente']).getTime() || 0;
      return dateB - dateA;
    });
    return { data: filteredRows, total: filteredRows.length };

  } catch (e) { console.error("Error getRequests", e); throw new Error("Error obteniendo datos."); }
}

function getRequestDetail(email, { id }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado");
  
  try {
    const allRequests = getDataFromSheet('Solicitudes');
    const header = allRequests.find(row => 
      String(row['ID Solicitud']).trim() === String(id).trim() || 
      String(row['Ticket G4S']).trim() === String(id).trim()
    );

    if (!header) throw new Error("Ticket no encontrado.");
    const realUuid = String(header['ID Solicitud']).trim();
    const ticketG4S = String(header['Ticket G4S']).trim();

    if (!context.isAdmin) {
      const recordClientId = String(header['ID Sede']);
      if (recordClientId && !context.allowedClientIds.includes(recordClientId)) throw new Error("No tiene permisos.");
    }

    // --- BÚSQUEDA INTELIGENTE DE RELACIONES (PLURAL O SINGULAR) ---
    const getChildren = (sheetName) => {
      const allChildren = getDataFromSheet(sheetName);
      if (allChildren.length === 0) return [];

      const keys = Object.keys(allChildren[0]);
      // Buscamos ID Solicitudes (Plural) O ID Solicitud (Singular)
      const foreignKey = keys.find(k => 
        k === 'ID Solicitudes' || 
        k === 'ID Solicitud' || 
        k === 'Id Solicitud' || 
        k === 'Solicitud'
      );
      
      if (!foreignKey) return [];

      return allChildren.filter(row => {
         const val = String(row[foreignKey]).trim();
         return val === realUuid || val === ticketG4S;
      });
    };

    return {
      header: header,
      services: getChildren('Observaciones historico'),
      history: getChildren('Estados historico'),
      documents: getChildren('Solicitudes anexos')
    };

  } catch (e) { console.error("Error Detalle", e); throw e; }
}

function createRequest(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");
  
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw new Error("Servidor ocupado."); }

  try {
    const now = new Date();
    const uuid = Utilities.getUuid();
    let letraInicial = "G"; 

    const allSedes = getDataFromSheet('Sedes');
    const sedeInfo = allSedes.find(s => String(s['ID Sede']) === String(payload.idSede));
    if (sedeInfo) {
        const idCliente = sedeInfo['ID Cliente'];
        const allClientes = getDataFromSheet('Clientes');
        const clienteInfo = allClientes.find(c => String(c['ID Cliente']) === String(idCliente));
        if (clienteInfo) {
            const nombre = clienteInfo['Nombre corto'] || clienteInfo['RazonSocial'] || "G";
            letraInicial = nombre.toString().trim().charAt(0).toUpperCase();
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
      "ID Sede": payload.idSede,
      "Ticket Cliente": payload.ticketCliente || "",
      "Clasificación": payload.clasificacion,
      "Prioridad Solicitud": payload.prioridad,
      "Solicitud": payload.solicitud,
      "Observación": payload.observacion,
      "Usuario Actualización": email
    };

    appendDataToSheet('Solicitudes', newRow);
    return { Rows: [newRow], Status: "Success", GeneratedTicket: ticketG4S }; 

  } catch (e) { console.error("Error createRequest", e); throw new Error("Error: " + e.message); } finally { lock.releaseLock(); }
}
