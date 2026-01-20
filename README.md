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

- 🔍 **Gestión de Activos (BigQuery Integration)**
  - Módulo avanzado de inventario conectado directamente a **Google BigQuery**.
  - Navegación jerárquica: **Cliente > Sede > Piso**.
  - **Vista Dual**: Alternancia entre vista de **Cuadrícula (Grid)** con fotos y **Mapa Interactivo (Plano)**.
  - **Ficha Técnica**: Visualización de datos técnicos y protocolos de mantenimiento en formato JSON estructurado.

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

## 🔄 Novedades y Cambios (vs Branch-origin)

Con respecto a la versión original (`Branch-origin`), se han implementado las siguientes mejoras y cambios estructurales:

### 1. Migración de Activos a BigQuery
Originalmente, los activos se gestionaban mediante Google Sheets. En la versión actual:
- Se integra la librería **OAuth2** para conexión segura con BigQuery.
- Consultas optimizadas a tablas de inventario (`DIM_CLIENTES`, `DIM_SEDES`, `DIM_PISOS`, `DIM_ACTIVOS`).
- Implementación de seguridad con escape de caracteres (`esc()`) para prevenir inyección SQL.

### 2. Nuevo Módulo de Visualización de Activos (`AssetsView`)
- **Navegación inteligente**: Auto-selección de Cliente/Sede/Piso cuando solo existe una opción disponible.
- **Interactividad en Planos**: Localización visual de activos sobre mapas de calor o planos de planta.
- **Paneles Laterales**: Detalles expandibles sin perder el contexto de la navegación.

### 3. Mejoras en el Contexto de Usuario
- Actualización del motor de caché a **v6**.
- Inclusión de `allowedCustomerIds` para un filtrado de seguridad más robusto a nivel de base de datos.

---

## 🧱 Arquitectura general

### Backend – `Code.gs`

El backend gestiona la persistencia en 5 Spreadsheets y una conexión a BigQuery.

- **Conectividad BigQuery:**
  - `_getBQConfig()`: Centraliza credenciales (Service Account). Se recomienda el uso de `ScriptProperties` para `BQ_PRIVATE_KEY`.
  - `_runBQQuery(query)`: Ejecuta SQL estándar y retorna objetos mapeados.

- **Endpoints principales:**
  - `getUserContext`: (v6) Obtiene rol, sedes y clientes permitidos.
  - `getAssetsData`: Manejador central para la lógica de inventario en BigQuery.
  - `apiHandler`: Router central que ahora incluye soporte para datos de activos.

### Frontend – `Index.html`

- **Componentes destacados:**
  - `AssetsView`: Componente principal del nuevo módulo de activos.
  - `MapViewer`: Renderiza el plano del piso y posiciona los activos dinámicamente (`coord_x`, `coord_y`).
  - `JsonBlock`: Formateador elegante para datos técnicos y protocolos.
  - `FileModal`: Visor integrado de imágenes y documentos PDF.

---

## 🚀 Cómo desplegar

1.  Crear un proyecto en **Google Apps Script**.
2.  Subir los archivos `Code.gs` e `Index.html`.
3.  **Configurar IDs en `Code.gs`** (Spreadsheets y carpetas de Drive).
4.  **Configurar BigQuery**:
    - Añadir la librería `OAuth2` (ID: `1B7_5B191Pn9ua_69CPv99Cof78Xh3XkBy9Wjy3YV59_t6Ksh9k5I8I54`).
    - Configurar `BQ_PRIVATE_KEY`, `BQ_CLIENT_EMAIL` y `BQ_PROJECT_ID` en las **Propiedades del Script**.
5.  **Publicar**:
    - Deploy → New deployment → Web app (Execute as: Me, Access: Anyone within Org).

---

## 🧰 Stack tecnológico

-   **Frontend**: React 18, Tailwind CSS, Font Awesome 6.4.0.
-   **Backend**: Google Apps Script, BigQuery (SQL).
-   **Seguridad**: OAuth2, Row-level filtering por email.
