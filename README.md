# megudan-erp-v2
V2 del ERP Operativo de MEGUDAN CONSTRUCCIONES SOSTENIBLES SAS . Desarrollado con Google Apps Script y HTML Web Apps sobre Google Sheets , automatiza la gestión operativa (clientes, inventario, ventas, compras, obras)  y precalifica los datos para la integración contable con Siigo y Alegra

# V2 - ERP Operativo | MEGUDAN

Este repositorio contiene el código fuente de la **Segunda Versión (V2) del ERP Operativo** de **MEGUDAN CONSTRUCCIONES SOSTENIBLES SAS** [1, 2]. El sistema actúa como el núcleo de captura operacional de la compañía, donde se procesan y estructuran los datos antes de ser proyectados financieramente y enviados al Core de Integración Contable [4, 5].

## 🛠️ Tecnología y Arquitectura
*   **Plataforma Base:** Google Sheets (como base de datos tabular e interfaz operativa directa) [3, 4].
*   **Backend:** Google Apps Script (`.gs`) para el procesamiento lógico, validaciones y automatizaciones de flujos [3].
*   **Frontend:** Formularios e interfaces web interactivas (`.html` con CSS/JS) servidas mediante Web Apps (`doGet`) [3].

## 📦 Módulos Principales del ERP Operativo
1. **Configuración General (`CFG_`):** Información maestra de la empresa, sistema de consecutivos, impuestos locales/reglas tributarias y perfiles contables [2, 6-8].
2. **Seguridad y Usuarios (`USR_`):** Sistema de autenticación seguro, cifrado de contraseñas (Hash/Salt), sesiones activas y auditoría de cambios en tiempo real [3, 9, 10].
3. **Terceros (Clientes `CLI_` y Proveedores `PROV_`):** Maestros de datos validados bajo las reglas tributarias de la DIAN en Colombia (incluyendo régimen de IVA, responsabilidades fiscales y dígitos de verificación) [2, 11-13].
4. **Productos (`PROD_`):** Central de catálogo, unidades de medida y administración dinámica de listas de precios [14].
5. **Obras y Proyectos (`OBR_`):** Presupuestos, control de avances y asignación operativa de recursos (materiales, mano de obra y transporte) [15].
6. **Ventas y Compras (`VEN_` / `COM_`):** Registro de cabeceras, detalles e historiales con afectación automática sobre inventario, cartera y tesorería [16, 17].
7. **Inventario (`INV_`):** Control de existencias promedio ponderado, registro de movimientos (entradas/salidas), kardex automático, traslados y ajustes de stock [18, 19].
8. **Flujo de Caja, Tesorería y Auxiliares (`TES_` / `ING_` / `COS_` / `GAS_`):** Control y conciliación de cuentas bancarias y cajas menores frente a recaudos, costos y gastos operativos [4, 20-23].

## 🔌 Core de Integración Contable
Prepara, valida y encola de manera estructurada los documentos de ventas, compras y recaudos generados en este ERP para su posterior sincronización por API con **Siigo** y **Alegra** [5, 24].
