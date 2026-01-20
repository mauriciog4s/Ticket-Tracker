# G4S Ticket Tracker – Registro de Cambios y Documentación

Este documento detalla los cambios realizados en el código original (rama `Branch-origin`) con respecto a la rama principal (`main`), enfocándose en la transición hacia una arquitectura basada en **BigQuery** y la implementación del módulo de **Activos**.

---

## 🚀 Resumen de Cambios Técnicos

### 1. Backend (`Code.gs`)

Se han añadido e integrado múltiples funcionalidades para soportar la conexión con Google BigQuery y mejorar la gestión de contexto de usuario.

#### **Nuevas Funciones y Constantes:**
- **`_getBQConfig()`**: Centraliza la obtención de credenciales de BigQuery (Project ID, Client Email y Private Key) desde `ScriptProperties`.
- **`DATASET_ID`**: Definición del dataset central `ControlTower`.
- **`_getBQService()`**: Configura el servicio de autenticación **OAuth2** para interactuar con la API de BigQuery.
- **`_runBQQuery(query)`**: Ejecuta consultas SQL estándar, maneja la autenticación y formatea los resultados en objetos JSON mapeados por columnas.
- **`getAssetsData(email, { action, payload })`**: Punto de entrada principal para la lógica de activos. Incluye:
    - `getClients`: Filtrado por permisos de usuario.
    - `getSites`: Obtención de sedes por cliente.
    - `getFloors`: Obtención de pisos por sede (incluye URLs de planos).
    - `getAssets`: Listado detallado de activos con soporte para datos técnicos (`specs`) y protocolos en formato JSON.
- **`getData(action, payload)`**: Función de conveniencia para invocar `getAssetsData` vía `apiHandler`.

#### **Modificaciones en Funciones Existentes:**
- **`apiHandler`**: Se añadió el caso `getAssetsData` para exponer las nuevas funcionalidades al frontend.
- **`getUserContext`**:
    - Actualización de la versión de caché a `v6`.
    - Implementación de `allowedCustomerIds` y `assignedCustomerNames` para filtrar datos de BigQuery basándose en la configuración de la hoja "Usuarios filtro".

---

### 2. Frontend (`Index.html`)

Se ha transformado la interfaz para incluir un sistema de gestión de activos robusto y visualmente interactivo.

#### **Nuevos Componentes React:**
- **`<AssetsView />`**:
    - Implementa la lógica de navegación jerárquica (Cliente > Sede > Piso).
    - Gestión de estados para carga (`loading`), activos y selección de vistas.
    - Regla de auto-selección: Si una lista tiene un solo elemento, se selecciona automáticamente.
- **`<MapViewer />`**:
    - Renderiza el plano del piso.
    - Posiciona dinámicamente los activos en el mapa utilizando coordenadas `coord_x` y `coord_y`.
    - Permite la interacción directa con los activos desde el plano.
- **`<JsonBlock />`**:
    - Formateador especializado para visualizar campos JSON complejos (Ficha técnica y Protocolos).
    - Detecta enlaces a archivos y permite previsualizarlos.
- **`<FileModal />`**:
    - Visor integrado para imágenes y documentos PDF sin salir de la aplicación.

#### **Mejoras en la UI/UX:**
- **Vista Dual**: Alternancia entre **Grid** (cuadrícula de fotos) y **Plano** (ubicación espacial).
- **Panel Lateral de Detalle**: Visualización limpia de fotos, estado operativo y especificaciones técnicas del activo seleccionado.
- **Integración en App**: Adición del acceso "Activos" en el Sidebar y manejo de rutas mediante el estado `view`.
- **Feedback Visual**: Implementación de un loader mejorado (`.loader-pro`) y transiciones de entrada para los paneles.

---

## 🛡️ Seguridad y Optimización

- **Protección SQL**: Se implementó la función `esc()` en el backend para escapar comillas simples y mitigar riesgos de inyección SQL en los parámetros de las consultas a BigQuery.
- **Caché v6**: Mejora en la persistencia del contexto del usuario para reducir llamadas redundantes a las hojas de configuración.
- **Batch Processing**: El sistema mantiene la capacidad de sincronización masiva para optimizar el rendimiento general.

---

## 🛠️ Requisitos de Configuración (Post-migración)

Para que estos cambios funcionen correctamente en un nuevo entorno:
1. **Librería OAuth2**: Debe estar vinculada al proyecto de Apps Script.
2. **Script Properties**: Es obligatorio configurar `BQ_PRIVATE_KEY`, `BQ_CLIENT_EMAIL` y `BQ_PROJECT_ID`.
3. **BigQuery Dataset**: Las tablas `DIM_CLIENTES`, `DIM_SEDES`, `DIM_PISOS`, `DIM_ACTIVOS` y `DIM_DISPOSITIVOS` deben existir en el dataset `ControlTower`.
