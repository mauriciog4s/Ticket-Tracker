
const BQ_CREDENTIALS = {
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQC8pIXxJuXj2kgo\n3eJ0CuRG7QdZXazXTfTFt04VL6B9b2q+kHyR8UDefeg3LhaYq19jxazqoMvqFsQV\nU/tTCLzKYL+Yew4vy5NOkmENqXmtrjBHmxPqBoHk2R+7aR52RnGkGPcmNAmqcjv4\nCUjgCCq2ko4VDKIQav6/6Psrg1FoaQy4p1lJXZZVTJBU3vUfUIKzlLu7zOrmdZ/K\nUyK+nXSR5Sw6DeE5pf5ElG3QQ9KgjZ/FnzG9gRRWozdW8IgghY6Lw0efyE72ijjf\n873V0K8FqhqDYd1r9stCj05zl07UHqqLXx1ik8YcIxQQR63kcMkqLmysJZx/axds\naDRAvGQbAgMBAAECggEASgqdU/C7jLoxVnD4oCliPgBswQPGgl9jsnLnH+Oor3Ma\nx584daPmnS14Bqh9UAD7mNKOsyzXvJKg9eoXnBiy2RAuQ3AROmtB7zX/B/i7/JKA\n+qoAn/tb4nHiRZHV1gCCPDFcWE9Wd+MMbKdgRiaOdUiCofpqZd1JDhQo+YQ6YKsl\nFZdXPb9SxOFZOxDLjQY64/FV9gn3qFXBbMYf53yNmzd3l6aH/ERgWw3N9YZEgnFM\nBjKn6AN0JSsFLcacjtvgjNDE47U9jO5vdSKu0aFO+vDFBmlewou36i5S9AYLYyAQ\nNC5RDDQ/+WfPv3rGD7D4Nc9JGgK2PbEHRQypmDDNsQKBgQD3AzUE87voHefjH8xw\nWXlRd1H6+0OXt6BwohtSKJmoUvnmbtSjujSppT0EpQzdnrKJ5migbHu3xofaimcH\nmaeqTohuxtQq6E6Yly3kRGexlCAoj0LEW/Hmk1sIiuwI4ukqO6ZiCshMue8pErpE\n5Shb+2xR48nThaeHl8oqrfilCQKBgQDDgaGdKefwE/FCkMgM3ucB12SkHa37Whc9\nPn7YwrcqRC7ZkAvzkzRB/NcyW95fUcD8dYYvHyCHiksMxocwmJ38CbRJWibQlb3y\nMevWes60iAX4NjM1pWWIzADg8MUnncFlUZPWN/b7uxPzMrY1bijtxumGrme0jLwt\nfgOc6CkNAwKBgFhw3IXmYswsEP/APemoD4j8qOytFDl5NMe/MvsKsGGVPAamfhoV\nLI/lKuDD28Rp8tDvH1z5Gp7lRXUZAuS0vlR7A9xt8j9ep+14i6TkXSA2wgDjsmst\n5IHDFuALJZHU9Nj7PIp0A9184UWaf/j096tfbRww6+2BOEeTMH5xhcpJAoGAF9FD\nDxJ73xOO4L0ioe7F1cOXzyaOe4CONDfY3C9cgRmtW3PhANt+EkvrK4dln9cl25u1\nrSftnpWKbxQAhDsThBDqlcUV1XNooIjUYlyzseqgT4zK0E5GAFRaBw1N93WQifdW\nO1K2FBTGaWpUKE4zTkRdTrsQhz5d7mzbo9HkrmECgYAxR0UNwZeSLe+yZhdeK3cv\nqz55RbNeHwrhE10PE49CwlUDDdTHk2qK7raAV+LMFEz2Lq8umXxx2OgJSEip3ty4\noVA5qOjr5M62v1wTbrDpmi2ItWXxuzH+oHVW3MBS4jnrbZzsoZ0ZF855xbgfAEwI\nDAygge9kB/HNsXs2OMufAw==\n-----END PRIVATE KEY-----\n",
  "client_email": "tz1-bigquery@g4s-shared-tz1.iam.gserviceaccount.com",
  "project_id": "g4s-shared-tz1"
};

const DATASET_ID = "ControlTower";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('G4S Visor de Activos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getData(action, payload) {
  const projectId = BQ_CREDENTIALS.project_id;
  
  try {
    switch (action) {
      // 1. Obtener Clientes
      case 'getClients':
        return runQuery(`
          SELECT DISTINCT id_cliente, nombre_cliente 
          FROM \`${projectId}.${DATASET_ID}.DIM_CLIENTES\` 
          ORDER BY nombre_cliente
        `);
      
      // 2. Obtener Sedes (Filtrado por Cliente)
      case 'getSites':
        return runQuery(`
          SELECT id_sede, nombre_sede 
          FROM \`${projectId}.${DATASET_ID}.DIM_SEDES\` 
          WHERE id_cliente = '${payload.clientId}' 
          ORDER BY nombre_sede
        `);
      
      // 3. Obtener Pisos (Filtrado por Sede) - Incluye URL del plano
      case 'getFloors':
        return runQuery(`
          SELECT id_piso, nombre_piso, imagen_plano_url 
          FROM \`${projectId}.${DATASET_ID}.DIM_PISOS\` 
          WHERE id_sede = '${payload.siteId}' 
          ORDER BY nombre_piso
        `);
      
      // 4. Obtener Activos (Filtrado por Piso)
      case 'getAssets':
        return runQuery(`
          SELECT 
            A.id_activo, 
            A.nombre_activo, 
            COALESCE(D.clasificacion, A.id_dispositivo) as tipo_dispositivo, 
            A.estado_activo, 
            A.coord_x, 
            A.coord_y, 
            
            -- AGREGADO: La fecha de actualización
            A.fecha_actualizacion, 

            A.foto_1, 
            A.foto_2, 
            A.foto_3, 
            
            TO_JSON_STRING(A.datos_tecnicos_json) as specs,
            TO_JSON_STRING(A.ultimo_protocolo_json) as protocol
          FROM \`${projectId}.${DATASET_ID}.DIM_ACTIVOS\` A
          LEFT JOIN \`${projectId}.${DATASET_ID}.DIM_DISPOSITIVOS\` D 
            ON A.id_dispositivo = D.id_dispositivo
          WHERE A.id_piso = '${payload.floorId}'
          LIMIT 2000
        `);

      default: return [];
    }
  } catch (e) { throw new Error("Error Backend: " + e.message); }
}

// --- UTILIDADES DE CONEXIÓN (OAuth2) ---
function getService() {
  return OAuth2.createService('BigQueryApp')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setPrivateKey(BQ_CREDENTIALS.private_key)
    .setIssuer(BQ_CREDENTIALS.client_email)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope('https://www.googleapis.com/auth/bigquery');
}

function runQuery(query) {
  const service = getService();
  if (!service.hasAccess()) throw new Error('Error de Autenticación: ' + service.getLastError());
  
  const url = `https://bigquery.googleapis.com/bigquery/v2/projects/${BQ_CREDENTIALS.project_id}/queries`;
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
