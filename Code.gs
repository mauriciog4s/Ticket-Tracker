/**
 * ------------------------------------------------------------------
 * CONFIGURACIÓN Y MAPEO DE HOJAS
 * ------------------------------------------------------------------
 */

// IDs de los spreadsheets
const MAIN_SPREADSHEET_ID       = '1MC76eZZt7qiso2M8LMz777_xJnzrl_ZpZptDZBnPlDo';
const PERMISSIONS_SPREADSHEET_ID = '1zcZZGe_93ytWXtCF1kmk_Y8zc5b5cL1xH34i7v1w01k';
const CLIENTS_SPREADSHEET_ID    = '1hHWPJF9KSC0opplpCNgHRNkW6CLf7StXG2Y31m6yUpo';
const SEDES_SPREADSHEET_ID      = '1tbcmOM_LLwr62P6O1RjpYn3GirpzGyK98frYKVAqIsM';

// Datos de la app AppSheet (para construir la URL del archivo)
const APPSHEET_APP_NAME    = 'AppSolicitudes-5916254';
const APPSHEET_ATTACH_TABLE = 'Solicitudes anexos';

const SHEET_CONFIG = {
  'Solicitudes'            : MAIN_SPREADSHEET_ID,
  'Estados historico'      : MAIN_SPREADSHEET_ID,
  'Observaciones historico': MAIN_SPREADSHEET_ID,
  'Solicitudes anexos'     : MAIN_SPREADSHEET_ID,
  'Permisos'               : PERMISSIONS_SPREADSHEET_ID,
  'Usuarios filtro'        : PERMISSIONS_SPREADSHEET_ID,
  'Clientes'               : CLIENTS_SPREADSHEET_ID,
  'Sedes'                  : SEDES_SPREADSHEET_ID
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('G4S Ticket Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Handler genérico para llamadas desde el front.
 */
function apiHandler(request) {
  const userEmail = Session.getActiveUser().getEmail();
  const endpoint  = request && request.endpoint;
  const payload   = request && request.payload;

  console.log("🔒 [API]", endpoint, "| User:", userEmail);

  try {
    if (!userEmail) throw new Error("Usuario no identificado.");

    switch (endpoint) {
      case 'getUserContext':
        return getUserContext(userEmail);
      case 'getRequests':
        return getRequests(userEmail);
      case 'getRequestDetail':
        return getRequestDetail(userEmail, payload);
      case 'createRequest':
        return createRequest(userEmail, payload);
      default:
        throw new Error("Endpoint desconocido: " + endpoint);
    }
  } catch (err) {
    console.error("❌ Error API:", err);
    return { error: true, message: err.message };
  }
}

// ------------------------------------------------------------------
// HELPERS BÁSICOS DE HOJA
// ------------------------------------------------------------------

function getDataFromSheet(sheetName) {
  const spreadsheetId = SHEET_CONFIG[sheetName];
  if (!spreadsheetId) throw new Error("Configuración faltante: " + sheetName);

  const ss    = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const values  = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const data    = values.slice(1);

  return data.map(function (row) {
    const obj = {};
    headers.forEach(function (header, index) {
      obj[header] = (row[index] === undefined ? "" : row[index]);
    });
    return obj;
  });
}

function appendDataToSheet(sheetName, objectData) {
  const spreadsheetId = SHEET_CONFIG[sheetName];
  const ss    = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Hoja " + sheetName + " no encontrada.");

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowArray = headers.map(function (header) {
    return objectData[header] || "";
  });

  sheet.appendRow(rowArray);
  return { success: true };
}

// ------------------------------------------------------------------
// ARCHIVOS / ANEXOS
// ------------------------------------------------------------------

/**
 * A partir de un registro de "Solicitudes anexos" busca la columna
 * donde está la ruta del archivo y genera una URL lista para el front.
 *
 * Se buscan estas columnas (en orden):
 *  - Archivo
 *  - Foto
 *  - Dibujo
 *  - Image / Imagen / Picture / Photo (defensivo)
 */
function buildAttachmentUrlFromRecord(record) {
  if (!record) return "";

  var possibleKeys = [
    "Archivo",
    "Foto",
    "Dibujo",
    "Image",
    "Imagen",
    "Picture",
    "Photo"
  ];

  var rawPath = "";
  for (var i = 0; i < possibleKeys.length; i++) {
    var key = possibleKeys[i];
    if (record.hasOwnProperty(key) && record[key]) {
      rawPath = String(record[key]).trim();
      if (rawPath) break;
    }
  }

  if (!rawPath) return "";

  return getAttachmentUrlFromPath(rawPath);
}

/**
 * Convierte la ruta almacenada en la hoja en una URL usable.
 *
 * PRIORIDAD:
 * 1) Si la ruta parece ser la ruta relativa de AppSheet
 *    (ej: "/Info/Clientes/..." o "..._Images/..."), usamos
 *    https://www.appsheet.com/template/gettablefileurl
 * 2) En caso contrario, intentamos buscar el archivo en Drive por nombre.
 * 3) Si ya es una URL http(s), se devuelve tal cual.
 *
 * IMPORTANTE:
 * Para que funcione sin parámetro "signature", normalmente hay que
 * desactivar en AppSheet:
 *   Security → Options → "Require Image and File URL Signing".
 */
function getAttachmentUrlFromPath(path) {
  path = String(path || "").trim();
  if (!path) return "";

  // Si ya es URL completa
  if (/^https?:\/\//i.test(path)) return path;

  var cache    = CacheService.getScriptCache();
  var cacheKey = "att_v3_" + Utilities.base64Encode(path);
  var cached   = cache.get(cacheKey);
  if (cached) return cached;

  try {
    var normalizedPath = path;

    // CASO 1: Ruta AppSheet: Info/Clientes, o carpetas *_Images
    if (
      normalizedPath.indexOf("/Info/") === 0 ||
      normalizedPath.indexOf("Info/Clientes") !== -1 ||
      normalizedPath.indexOf("_Images/") !== -1
    ) {
      if (normalizedPath.charAt(0) !== "/") {
        normalizedPath = "/" + normalizedPath;
      }

      var base     = "https://www.appsheet.com/template/gettablefileurl";
      var appName  = encodeURIComponent(APPSHEET_APP_NAME);
      var table    = encodeURIComponent(APPSHEET_ATTACH_TABLE);
      var fileName = encodeURIComponent(normalizedPath);

      var url =
        base +
        "?appName=" +
        appName +
        "&tableName=" +
        table +
        "&fileName=" +
        fileName;

      cache.put(cacheKey, url, 21600); // 6 horas
      return url;
    }

    // CASO 2: Fallback → buscar en Drive por nombre de archivo
    var parts        = normalizedPath.split("/");
    var fileNameOnly = parts[parts.length - 1];
    if (fileNameOnly) {
      var files = DriveApp.getFilesByName(fileNameOnly);
      if (files.hasNext()) {
        var file    = files.next();
        var fileId  = file.getId();
        var driveUrl =
          "https://drive.google.com/file/d/" +
          fileId +
          "/view?usp=drivesdk";

        cache.put(cacheKey, driveUrl, 21600);
        return driveUrl;
      }
    }
  } catch (e) {
    console.error("Error resolviendo anexo:", path, e);
  }

  return "";
}

// ------------------------------------------------------------------
// LÓGICA DE NEGOCIO
// ------------------------------------------------------------------

function getUserContext(email) {
  var cache    = CacheService.getScriptCache();
  var cacheKey = "ctx_v6_" + Utilities.base64Encode(email);
  var cached   = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  try {
    var context = {
      email          : email,
      role           : "Usuario",
      allowedClientIds: [],
      clientNames    : {},
      isValidUser    : false,
      isAdmin        : false
    };

    var allPermissions = getDataFromSheet("Permisos");
    var userData = allPermissions.find(function (row) {
      return String(row["Correo"]).trim().toLowerCase() === email.toLowerCase();
    });

    if (userData) {
      context.isValidUser = true;
      if (String(userData["Rol_Asignado"] || "")
        .trim()
        .toLowerCase() === "administrador") {
        context.role  = "Administrador";
        context.isAdmin = true;
      }
    }

    if (context.isValidUser) {
      var allRelations = getDataFromSheet("Usuarios filtro");
      var myRelations = allRelations.filter(function (row) {
        return String(row["Usuario"]).trim().toLowerCase() === email.toLowerCase();
      });

      var assignedClientIds = [];
      myRelations.forEach(function (row) {
        if (row["Cliente"]) assignedClientIds.push(String(row["Cliente"]));
      });

      if (assignedClientIds.length > 0) {
        var allSedes = getDataFromSheet("Sedes");
        var mySedes = allSedes.filter(function (sede) {
          return assignedClientIds.indexOf(String(sede["ID Cliente"])) !== -1;
        });

        mySedes.forEach(function (sede) {
          var idSede = String(sede["ID Sede"]);
          var nombreSede =
            sede["Nombre"] ||
            sede["Nombre_Sede"] ||
            sede["Sede"] ||
            sede["Nombre Sede"] ||
            idSede;

          if (idSede) {
            context.allowedClientIds.push(idSede);
            context.clientNames[idSede] = nombreSede;
          }
        });
      }
    }

    cache.put(cacheKey, JSON.stringify(context), 600); // 10 min
    return context;
  } catch (e) {
    console.error("Error Context", e);
    throw e;
  }
}

function getRequests(email) {
  var context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  try {
    var allRows = getDataFromSheet("Solicitudes");
    var filteredRows = [];

    if (context.isAdmin) {
      filteredRows = allRows;
    } else {
      if (context.allowedClientIds.length === 0) {
        return { data: [], total: 0 };
      }
      filteredRows = allRows.filter(function (row) {
        return (
          context.allowedClientIds.indexOf(String(row["ID Sede"])) !== -1
        );
      });
    }

    filteredRows.sort(function (a, b) {
      var dateA = new Date(a["Fecha creación cliente"]).getTime() || 0;
      var dateB = new Date(b["Fecha creación cliente"]).getTime() || 0;
      return dateB - dateA;
    });

    return { data: filteredRows, total: filteredRows.length };
  } catch (e) {
    console.error("Error getRequests", e);
    throw new Error("Error obteniendo datos.");
  }
}

function getRequestDetail(email, params) {
  var id = params && params.id;
  var context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado");

  try {
    var allRequests = getDataFromSheet("Solicitudes");
    var header = allRequests.find(function (row) {
      return (
        String(row["ID Solicitud"]).trim() === String(id).trim() ||
        String(row["Ticket G4S"]).trim() === String(id).trim()
      );
    });

    if (!header) throw new Error("Ticket no encontrado.");

    var realUuid  = String(header["ID Solicitud"]).trim();
    var ticketG4S = String(header["Ticket G4S"]).trim();

    if (!context.isAdmin) {
      var recordClientId = String(header["ID Sede"]);
      if (
        recordClientId &&
        context.allowedClientIds.indexOf(recordClientId) === -1
      ) {
        throw new Error("No tiene permisos.");
      }
    }

    // Búsqueda de hijos (observaciones, historial, anexos)
    var getChildren = function (sheetName) {
      var allChildren = getDataFromSheet(sheetName);
      if (allChildren.length === 0) return [];

      var keys = Object.keys(allChildren[0]);
      var foreignKey = keys.find(function (k) {
        return (
          k === "ID Solicitudes" ||
          k === "ID Solicitud" ||
          k === "Id Solicitud" ||
          k === "Solicitud"
        );
      });

      if (!foreignKey) return [];

      return allChildren.filter(function (row) {
        var val = String(row[foreignKey]).trim();
        return val === realUuid || val === ticketG4S;
      });
    };

    var docsRaw = getChildren("Solicitudes anexos");
    var docsWithUrl = docsRaw.map(function (doc) {
      var url = buildAttachmentUrlFromRecord(doc);
      if (url) {
        // campo que usa el front (RequestDetail -> getDocUrl)
        doc.Url = url;
      }
      return doc;
    });

    return {
      header   : header,
      services : getChildren("Observaciones historico"),
      history  : getChildren("Estados historico"),
      documents: docsWithUrl
    };
  } catch (e) {
    console.error("Error Detalle", e);
    throw e;
  }
}

function createRequest(email, payload) {
  var context = getUserContext(email);
  if (!context.isValidUser) throw new Error("Acceso Denegado.");

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("Servidor ocupado.");
  }

  try {
    var now  = new Date();
    var uuid = Utilities.getUuid();
    var letraInicial = "G";

    var allSedes  = getDataFromSheet("Sedes");
    var sedeInfo = allSedes.find(function (s) {
      return String(s["ID Sede"]) === String(payload.idSede);
    });

    if (sedeInfo) {
      var idCliente   = sedeInfo["ID Cliente"];
      var allClientes = getDataFromSheet("Clientes");
      var clienteInfo = allClientes.find(function (c) {
        return String(c["ID Cliente"]) === String(idCliente);
      });

      if (clienteInfo) {
        var nombre =
          clienteInfo["Nombre corto"] ||
          clienteInfo["RazonSocial"] ||
          "G";
        letraInicial = String(nombre).trim().charAt(0).toUpperCase();
      }
    }

    var ss    = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Solicitudes");
    var nextRow = sheet.getLastRow() + 1;
    var rand    = Math.floor(Math.random() * 90) + 10;
    var ticketG4S = letraInicial + (1000000 + nextRow) + String(rand);

    var newRow = {
      "ID Solicitud"          : uuid,
      "Ticket G4S"            : ticketG4S,
      "Fecha creación cliente": now,
      "Estado"                : "Creado",
      "ID Sede"               : payload.idSede,
      "Ticket Cliente"        : payload.ticketCliente || "",
      "Clasificación"         : payload.clasificacion,
      "Prioridad Solicitud"   : payload.prioridad,
      "Solicitud"             : payload.solicitud,
      "Observación"           : payload.observacion,
      "Usuario Actualización" : email
    };

    appendDataToSheet("Solicitudes", newRow);

    return {
      Rows          : [newRow],
      Status        : "Success",
      GeneratedTicket: ticketG4S
    };
  } catch (e) {
    console.error("Error createRequest", e);
    throw new Error("Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}
