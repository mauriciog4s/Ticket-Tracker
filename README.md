# G4S Ticket Tracker – Web App (Apps Script + React)

Versión web (Google Apps Script + React + Tailwind) del **G4S Ticket Tracker**, migrada desde una app original en AppSheet.

La aplicación permite que los usuarios corporativos gestionen **tickets de servicio**, vean **historial**, **observaciones**, vinculen **activos (Assets)** y descarguen **anexos** (PDF, imágenes, etc.) almacenados en Google Drive, todo con control de permisos basado en su correo institucional.

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
  - Búsqueda en tiempo real y paginación.
  - Vista de **detalle técnico** completo (Header, Observaciones, Historial, Anexos, Activos).
  - Creación de nuevos tickets con clasificación, prioridad y adjuntos.

- 🔍 **Gestión de Activos (QR)**
  - Vinculación de activos a tickets mediante escaneo o búsqueda de QR.
  - Catálogo de activos sincronizado desde Google Sheets.

- 📎 **Anexos y Archivos**
  - Subida de **Fotos, Dibujos y Documentos** directamente a carpetas específicas de Drive.
  - Sistema de visualización vía **Proxy** para evitar problemas de permisos de terceros.
  - Soporte para rutas originales de AppSheet y archivos directos de Drive.

- 📊 **UI moderna**
  - React 18 (UMD) + Tailwind CDN + Babel standalone.
  - Dashboard con métricas (Total, Abiertos, Cerrados).
  - Sidebar interactivo con información del cliente y estado de sincronización.
  - Almacenamiento local (Cache) para carga instantánea.

---

## 🧱 Arquitectura general

### Backend – `Code.gs`

El backend gestiona la persistencia en 5 Spreadsheets distintos y centraliza la lógica de negocio.

- **Mapeo de IDs de spreadsheets:**
  ```js
  const MAIN_SPREADSHEET_ID        = '...';
  const PERMISSIONS_SPREADSHEET_ID = '...';
  const CLIENTS_SPREADSHEET_ID     = '...';
  const SEDES_SPREADSHEET_ID       = '...';
  const ACTIVOS_SPREADSHEET_ID     = '...';

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
    'Activos': ACTIVOS_SPREADSHEET_ID
  };
  ```

- **Endpoints expuestos vía `apiHandler`:**
  - `getUserContext`: Obtiene rol y sedes permitidas.
  - `getRequests`: Lista de tickets filtrados.
  - `getRequestDetail`: Detalle completo de un ticket.
  - `createRequest`: Crea ticket e integra con **AppSheet API** para disparar correos.
  - `uploadAnexo`: Sube archivos a carpetas específicas de Drive.
  - `getAnexoDownload`: Resuelve la URL de descarga.
  - `createSolicitudActivo`: Vincula un activo QR al ticket.
  - `getSolicitudActivos`: Lista activos de un ticket.
  - `getActivosCatalog`: Catálogo completo de activos.
  - `getActivoByQr`: Busca activo específico por QR.
  - `getClassificationOptions`: Opciones dinámicas de clasificación.
  - `getBatchRequestDetails`: Sincronización masiva para modo offline/cache.

- **Sistema de Archivos (Proxy Mode):**
  Para garantizar que los usuarios puedan ver archivos sin errores de CORS o permisos de "Drive Viewer", la app implementa un router:
  `?v=archivo&id=...` → `_renderFileView(id)`
  Este método descarga el blob desde el servidor y lo sirve al cliente como una descarga directa o visualización.

### Frontend – `Index.html`

- **Tecnologías:** React 18, Tailwind CSS, Babel (JSX en el cliente).
- **Comunicación:** `runServer(endpoint, payload)` envuelve a `google.script.run`.
- **Persistencia:** Uso intensivo de `localStorage` para cachear tickets y detalles, permitiendo navegación instantánea.

---

## 🚀 Cómo desplegar

1.  Crear un proyecto en **Google Apps Script**.
2.  Subir los archivos `Code.gs` e `Index.html`.
3.  **Configurar IDs en `Code.gs`**:
    - Los 5 IDs de Spreadsheets mencionados arriba.
    - `IMAGES_FOLDER_ID`: Carpeta de Google Drive para fotos.
    - `DOCS_FOLDER_ID`: Carpeta de Google Drive para documentos/PDFs.
4.  **Permisos de Carpeta**: Las carpetas de Drive deben tener permisos de lectura para los usuarios (o estar dentro de una unidad compartida).
5.  **AppSheet API**: Si se desea el envío de correos original, configurar el `appId` y `accessKey` en la función `enviarAppSheetAPI`.
6.  **Publicar**:
    - Deploy → New deployment → Web app.
    - Execute as: **Me** (Owner).
    - Who has access: **Anyone within [Organization]**.

---

## 🧰 Stack tecnológico

-   **Frontend**: React 18, ReactDOM 18, Tailwind CSS (CDN), Babel.
-   **Backend**: Google Apps Script (V8 Engine).
-   **Base de Datos**: Google Sheets (Multi-spreadsheet).
-   **Almacenamiento**: Google Drive.
-   **Integraciones**: AppSheet API (disparadores de flujos).
