

````markdown
# G4S Ticket Tracker – Web App (Apps Script + React)

Versión web (Google Apps Script + React + Tailwind) del **G4S Ticket Tracker**, migrada desde una app original en AppSheet.

La aplicación permite que los usuarios corporativos gestionen **tickets de servicio**, vean **historial**, **observaciones** y descarguen **anexos** (PDF, imágenes, etc.) almacenados en Google Drive / AppSheet, todo con control de permisos basado en su correo.

---

## ✨ Features principales

- 🔐 **Autenticación corporativa**
  - Usa `Session.getActiveUser().getEmail()` para identificar al usuario.
  - Carga contexto de permisos desde una hoja de cálculo (`Permisos` + `Usuarios filtro`).

- 🧭 **Contexto por usuario / rol**
  - Roles: `Administrador` y `Usuario`.
  - Filtro automático de tickets por **Sedes** asignadas al usuario.
  - Mapeo de clientes y sedes desde hojas separadas (`Clientes`, `Sedes`).

- 🎫 **Gestión de tickets**
  - Listado de solicitudes con filtros: **Todo / Abierto / Cerrado**.
  - Vista de **detalle técnico** del ticket (similar al panel de AppSheet).
  - Creación de nuevos tickets con:
    - Selección de sede
    - Clasificación, tipo de servicio, prioridad
    - Campos personalizados (observación, ticket cliente)

- 🧮 **Backend en Google Sheets**
  - Todas las entidades se leen/escriben desde hojas de cálculo:
    - `Solicitudes`
    - `Estados historico`
    - `Observaciones historico`
    - `Solicitudes anexos`
    - `Permisos`
    - `Usuarios filtro`
    - `Clientes`
    - `Sedes`
  - Generación de IDs con `Utilities.getUuid()`.
  - Generación de Ticket G4S dinámico (prefijo por cliente + consecutivo + random).

- 📎 **Anexos integrados (AppSheet + Drive)**
  - Soporte para archivos gestionados originalmente por AppSheet.
  - Construcción de URLs públicas usando:
    - `gettablefileurl` de AppSheet cuando la columna `Archivo` trae rutas tipo:  
      `/Info/Clientes//Anexos/d9ae3f49.Archivo.153136.pdf`
    - Búsqueda de archivo en Drive por nombre como fallback.
  - Los anexos se muestran en el panel como lista descargable.

- 📊 **UI moderna**
  - React 18 (UMD) + Tailwind CDN.
  - Dashboard con tarjetas de métricas básicas (Total, Abiertos, Cerrados).
  - Sidebar con navegación: Inicio, Nueva solicitud, Tickets Activos, Historial, Configuración.
  - Modal de éxito al crear solicitudes.

---

## 🧱 Arquitectura general

### Backend – `Code.gs`

- Mapeo de IDs de spreadsheets:

  ```js
  const MAIN_SPREADSHEET_ID      = '...';
  const PERMISSIONS_SPREADSHEET_ID = '...';
  const CLIENTS_SPREADSHEET_ID   = '...';
  const SEDES_SPREADSHEET_ID     = '...';

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
````

* Endpoints expuestos vía `google.script.run`:

  ```js
  function apiHandler(request) {
    const userEmail = Session.getActiveUser().getEmail();
    const { endpoint, payload } = request;

    switch (endpoint) {
      case 'getUserContext':  return getUserContext(userEmail);
      case 'getRequests':     return getRequests(userEmail);
      case 'getRequestDetail':return getRequestDetail(userEmail, payload);
      case 'createRequest':   return createRequest(userEmail, payload);
      default: throw new Error('Endpoint desconocido');
    }
  }
  ```

* Helpers de acceso a Sheets:

  * `getDataFromSheet(sheetName)` → lee datos como arreglo de objetos.
  * `appendDataToSheet(sheetName, objectData)` → inserta nueva fila.

* Lógica de negocio:

  * `getUserContext(email)` → arma contexto (rol, sedes permitidas, nombres de sedes).
  * `getRequests(email)` → devuelve tickets filtrados por sede y ordenados por fecha.
  * `getRequestDetail(email, { id })` → devuelve:

    * `header` (detalle de la solicitud)
    * `services` (observaciones)
    * `history` (estados históricos)
    * `documents` (anexos ya enriquecidos con URL).

* Manejo de anexos:

  ```js
  const APPSHEET_APP_NAME   = 'AppSolicitudes-5916254';
  const APPSHEET_ATTACH_TABLE = 'Solicitudes anexos';

  function buildAttachmentUrlFromRecord(record) {
    const archivoKey = Object.keys(record).find(k =>
      String(k).trim().toLowerCase() === 'archivo'
    );
    if (!archivoKey) return '';
    return getAttachmentUrlFromPath(record[archivoKey]);
  }

  function getAttachmentUrlFromPath(path) {
    path = String(path || '').trim();
    if (!path) return '';

    // Si ya viene una URL
    if (/^https?:\/\//i.test(path)) return path;

    const cache = CacheService.getScriptCache();
    const cacheKey = 'att_v2_' + Utilities.base64Encode(path);
    const cached = cache.get(cacheKey);
    if (cached) return cached;

    // 1) Ruta AppSheet → gettablefileurl
    if (path.indexOf('/Info/') === 0 || path.indexOf('Info/Clientes') !== -1) {
      const base      = 'https://www.appsheet.com/template/gettablefileurl';
      const appName   = encodeURIComponent(APPSHEET_APP_NAME);
      const tableName = encodeURIComponent(APPSHEET_ATTACH_TABLE);
      const fileName  = encodeURIComponent(path);

      const url = `${base}?appName=${appName}&tableName=${tableName}&fileName=${fileName}`;
      cache.put(cacheKey, url, 21600);
      return url;
    }

    // 2) Fallback → buscar en Drive por nombre
    const parts    = path.split('/');
    const fileName = parts[parts.length - 1];
    if (fileName) {
      const files = DriveApp.getFilesByName(fileName);
      if (files.hasNext()) {
        const file  = files.next();
        const url   = `https://drive.google.com/file/d/${file.getId()}/view?usp=drivesdk`;
        cache.put(cacheKey, url, 21600);
        return url;
      }
    }

    return '';
  }
  ```

### Frontend – `Index.html`

* React 18 y ReactDOM 18 vía CDN.
* Babel standalone para usar JSX directamente en Apps Script.
* Tailwind configurado inline.

Componentes principales:

* `LoginView` – pantalla inicial y loading.
* `Sidebar` – navegación lateral.
* `HomeDashboard` – métricas generales y acciones rápidas.
* `CreateRequest` – formulario de creación de ticket.
* `RequestList` – listado filtrable/searchable de tickets.
* `RequestDetail` – vista de detalle técnico con:

  * Observaciones
  * Anexos (ícono + link de descarga)
  * Historial de estados
* `ConfigurationView` – datos del usuario y botón de resync.

Comunicación con backend:

```js
const runServer = (endpoint, payload = {}) =>
  new Promise((resolve, reject) => {
    if (typeof google === 'undefined' || !google.script) {
      reject(new Error('Ejecutar en GAS'));
      return;
    }
    google.script.run
      .withSuccessHandler(res => {
        if (res && res.error) reject(new Error(res.message));
        else resolve(res);
      })
      .withFailureHandler(err => reject(err))
      .apiHandler({ endpoint, payload });
  });
```

---

## 🚀 Cómo desplegar

1. Crear un proyecto en **Google Apps Script**.
2. Crear los archivos:

   * `Code.gs` → pegar el backend.
   * `Index.html` → pegar el frontend.
3. Configurar los IDs de los spreadsheets en `Code.gs`:

   * `MAIN_SPREADSHEET_ID`
   * `PERMISSIONS_SPREADSHEET_ID`
   * `CLIENTS_SPREADSHEET_ID`
   * `SEDES_SPREADSHEET_ID`
4. Revisar que las hojas tengan exactamente los nombres usados en `SHEET_CONFIG`.
5. Opcional (pero recomendado para anexos AppSheet):

   * En la app de AppSheet, desactivar la opción:
     **Security → Options → Require Image and File URL Signing**
     (o ajustar según la configuración usada).
6. Publicar como:

   * **Deploy → Test deployments / New deployment → Web app**
   * Ejecutar como: **Me (owner)**
   * Acceso: **Usuarios de la organización** (o lo que aplique).

---

## 🧰 Stack tecnológico

* **Frontend**

  * React 18 (UMD)
  * ReactDOM 18
  * Tailwind CSS (CDN)
* **Backend**

  * Google Apps Script (JavaScript)
  * Google Sheets (como base de datos ligera)
  * Google Drive (almacenamiento de archivos)
* **Integración externa**

  * AppSheet (rutas y almacenamiento original de anexos)
