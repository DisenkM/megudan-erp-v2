/**************************************************************
* 000_NOTAS.gs
* ERP OPERATIVO MEGUDAN V2
* RESPONSABILIDAD:
* - Documentación técnica, conceptual y arquitectura del ERP.
* - Registro de flujos lógicos, dependencias e integraciones.
*
* FILOSOFÍA ESTRATÉGICA:
* - El ERP es dueño de la OPERACIÓN diaria de la compañía (Ventas, Compras,
*   Inventarios, Obras, Costos, Gastos, Cartera, CxP, Tesorería).
* - Siigo y Alegra actúan como la capa contable/fiscal externa conectada por API.
* - "Una sola fuente de verdad para cada dato": Se asocian IDs internos con IDs externos.
**************************************************************/

/*
 ARQUITECTURA DE DATOS:
 MAESTROS (Horizontal)       -> Entidades permanentes (Clientes, Proveedores, Productos, Obras, Cuentas).
 TRANSACCIONES (Horizontal)  -> Cabeceras y detalles de operaciones (Ventas, Compras, Movimientos).
 DERIVADAS (Horizontal)      -> Saldos, Kardex, Vencimientos generados de forma automatizada por lote.
 CONFIGURACIONES (Vertical)  -> Parámetros de comportamiento (Empresa, Sistema, Impuestos, Contabilidad).

 CONTROL DE ACCESO DUAL (Sheets vs Web):
 - Contexto Sheets: getUi() activo, bypass confiable concede nivel supremo (ADMINISTRADOR_LOCAL_SHEETS).
 - Contexto Web App: getUi() lanza error, exige token de sesión activo y reglas en USR_PERMISOS.
*/
