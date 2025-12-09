/**
 * ------------------------------------------------------------------
 * CONFIGURACIÓN Y MAPEO DE HOJAS
 * ------------------------------------------------------------------
 */

// IDs extraídos de las URLs proporcionadas
const MAIN_SPREADSHEET_ID = '1MC76eZZt7qiso2M8LMz777_xJnzrl_ZpZptDZBnPlDo'; // Solicitudes, Historicos, Anexos
const PERMISSIONS_SPREADSHEET_ID = '1zcZZGe_93ytWXtCF1kmk_Y8zc5b5cL1xH34i7v1w01k'; // Permisos, Usuarios filtro
const CLIENTS_SPREADSHEET_ID = '1hHWPJF9KSC0opplpCNgHRNkW6CLf7StXG2Y31m6yUpo'; // Clientes
const SEDES_SPREADSHEET_ID = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM'; // Sedes

// Configuración para saber en qué Spreadsheet buscar cada tabla
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

/**
 * Sirve el HTML principal (SPA)
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('G4S Ticket Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Router Principal de la API
 */
function apiHandler(request) {
  const userEmail = Session.getActiveUser().getEmail();
  const { endpoint, payload } = request;
  console.log(`🔒 [API CHECK] Endpoint: ${endpoint} | Usuario Real: ${userEmail}`);
  try {
    if (!userEmail) throw new Error("No se pudo verificar la identidad del usuario.");

    switch (endpoint) {
      case 'getUserContext': return getUserContext(userEmail);
      case 'getRequests': return getRequests(userEmail);
      case 'getRequestDetail': return getRequestDetail(userEmail, payload);
      case 'createRequest': return createRequest(userEmail, payload);
      default: throw new Error(`Endpoint desconocido: ${endpoint}`);
    }
  } catch (err) {
    console.error(`❌ ERROR DE SEGURIDAD/EJECUCIÓN: ${err.message}`);
    return { error: true, message: "Error procesando su solicitud. Contacte al administrador." };
  }
}

/**
 * ------------------------------------------------------------------
 * HELPERS GENÉRICOS
 * ------------------------------------------------------------------
 */

function getDataFromSheet(sheetName) {
  const spreadsheetId = SHEET_CONFIG[sheetName];
  if (!spreadsheetId) throw new Error(`Configuración no encontrada para la tabla: ${sheetName}`);

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length < 2) return [];

  const headers = values[0];
  const data = values.slice(1);

  return data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      let value = row[index];
      if (value instanceof Date) {
        value = value.toISOString();
      }
      obj[header] = value;
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
  if (lastCol === 0) throw new Error(`La hoja ${sheetName} está vacía (sin cabeceras).`);
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  const rowArray = headers.map(header => {
    const val = objectData[header];
    return val === undefined || val === null ? "" : val;
  });

  sheet.appendRow(rowArray);
  return { success: true };
}

/**
 * ------------------------------------------------------------------
 * LÓGICA DE NEGOCIO
 * ------------------------------------------------------------------
 */

function getUserContext(email) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `ctx_it_v3_${Utilities.base64Encode(email)}`; 
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

    if (context.isValidUser) {
      const allRelations = getDataFromSheet('Usuarios filtro');
      const myRelations = allRelations.filter(row => String(row['Usuario']).toLowerCase() === email.toLowerCase());
      
      const assignedClientIds = [];
      myRelations.forEach(row => {
        const id = row['Cliente']; 
        if (id) assignedClientIds.push(String(id));
      });

      if (assignedClientIds.length > 0) {
        const allSedes = getDataFromSheet('Sedes');
        const mySedes = allSedes.filter(sede => assignedClientIds.includes(String(sede['ID Cliente'])));

        mySedes.forEach(sede => {
          const idSede = String(sede['ID Sede']);
          const nombreSede = sede['Nombre'] || sede['Nombre_Sede'] || sede['Sede'] || sede['Nombre Sede'] || idSede;
          
          if (idSede) {
             context.allowedClientIds.push(idSede); 
             context.clientNames[idSede] = nombreSede; 
          }
        });
      }
    }

    cache.put(cacheKey, JSON.stringify(context), 600);
    return context;

  } catch (e) {
    console.error("Error getUserContext", e);
    throw e;
  }
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

  } catch (e) {
    console.error("Error getRequests", e);
    throw new Error("Error obteniendo datos.");
  }
}

function getRequestDetail(email, { id }) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado");
  if (!id) throw new Error("ID requerido");

  try {
    const allRequests = getDataFromSheet('Solicitudes');
    const header = allRequests.find(row => row['ID Solicitud'] == id);

    if (!header) throw new Error("Ticket no encontrado.");

    if (!context.isAdmin) {
      const recordClientId = String(header['ID Sede']);
      if (recordClientId && !context.allowedClientIds.includes(recordClientId)) {
        throw new Error("No tiene permisos para ver este ticket.");
      }
    }

    // RELACIÓN: Las tablas dependientes se filtran por 'ID Solicitud'
    const getChildren = (sheetName) => {
      const allChildren = getDataFromSheet(sheetName);
      return allChildren.filter(row => row['ID Solicitud'] == id);
    };

    return {
      header: header,
      services: getChildren('Observaciones historico'),
      history: getChildren('Estados historico'),
      documents: getChildren('Solicitudes anexos')
    };

  } catch (e) {
    console.error("Error Detalle", e);
    throw e;
  }
}

function createRequest(email, payload) {
  const context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  if (!payload.idSede || !payload.solicitud || !payload.observacion) {
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

    // --- LÓGICA DE GENERACIÓN DE TICKET G4S ---
    // 1. Obtener ID Cliente desde la Sede
    const allSedes = getDataFromSheet('Sedes');
    const sedeInfo = allSedes.find(s => String(s['ID Sede']) === String(payload.idSede));
    const idCliente = sedeInfo ? sedeInfo['ID Cliente'] : null;

    // 2. Obtener Nombre Corto desde Cliente
    let letraInicial = "X";
    if (idCliente) {
        const allClientes = getDataFromSheet('Clientes');
        const clienteInfo = allClientes.find(c => String(c['ID Cliente']) === String(idCliente));
        if (clienteInfo) {
            const nombreCorto = clienteInfo['Nombre corto'] || clienteInfo['Nombre_Corto'] || clienteInfo['RazonSocial'] || "G";
            letraInicial = nombreCorto.toString().trim().charAt(0).toUpperCase();
        }
    }

    // 3. Calcular RowNumber (Fila actual + 1)
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Solicitudes');
    const nextRow = sheet.getLastRow() + 1; // Equivalente a [_RowNumber] para la nueva fila

    // 4. Aleatorio (10-99)
    const rand = Math.floor(Math.random() * 90) + 10;

    // 5. Fórmula: UPPER(LEFT(Nombre,1)) & (1000000 + RowNum) & RAND
    const ticketG4S = `${letraInicial}${1000000 + nextRow}${rand}`;

    const newRow = {
      "ID Solicitud": uuid,
      "Ticket G4S": ticketG4S, // NUEVO CAMPO CALCULADO
      "Fecha creación cliente": now,
      "Estado": "Abierto",
      "ID Sede": payload.idSede,
      "Ticket Cliente": payload.ticketCliente || "",
      "Clasificación": payload.clasificacion,
      "Prioridad Solicitud": payload.prioridad,
      "Solicitud": payload.solicitud,
      "Observación": payload.observacion,
      "Usuario Actualización": email
    };

    appendDataToSheet('Solicitudes', newRow);
    
    // Retornamos el Ticket G4S generado para mostrarlo al usuario si es necesario
    return { Rows: [newRow], Status: "Success", GeneratedTicket: ticketG4S }; 

  } catch (e) {
    console.error("Error createRequest", e);
    throw new Error("Error guardando el ticket: " + e.message);
  } finally {
    lock.releaseLock();
  }
}
