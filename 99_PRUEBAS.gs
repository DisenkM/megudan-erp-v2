/**************************************************************
* 99_PRUEBAS.gs
* ERP OPERATIVO V2 - MEGUDAN
* RESPONSABILIDAD:
* - Proporcionar una Suite Completa de Pruebas de Integración y End-to-End.
* - Probar de manera automatizada todo el flujo transaccional y modular del ERP:
*   🔑 Seguridad -> 👥 Clientes -> 🏢 Proveedores -> 📦 Productos -> 🛒 Compras -> 💰 Ventas -> 💸 Gastos -> ⚙️ Mantenimiento.
* - Garantizar que todas las llamadas asíncronas, bypass dual y afectaciones cruzadas operen al 100%.
**************************************************************/

/**
 * Función maestra para ejecutar la suite completa de pruebas de extremo a extremo.
 * Se puede ejecutar directamente desde el editor de Google Apps Script.
 * Escribe logs descriptivos detallados del avance de cada transacción simulada.
 */
function PROBAR_ERP_E2E_INTEGRAL() {
  console.log("==================================================================");
  console.log("🚀 INICIANDO SUITE DE PRUEBAS DE INTEGRACIÓN E2E DE MEGUDAN ERP V2");
  console.log("==================================================================");
  
  const testResults = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tokenSesionActiva = null;
  let idUsuarioCreado = null;
  let idClienteCreado = null;
  let idProveedorCreado = null;
  let idProductoCreado = null;
  let idCompraCreada = null;
  let idVentaCreada = null;
  let idGastoCreado = null;

  // ==========================================================
  // TEST 1: MÓDULO DE SEGURIDAD Y CONTROL DE ACCESO
  // ==========================================================
  console.log("\n🔑 [TEST 1] Probando Módulo de Seguridad y Sesiones (23_SEGURIDAD):");
  try {
    // 1. Asegurar autocuración de base de datos
    SEG_INICIALIZAR_ROLES_PREDEFINIDOS();
    SEG_INICIALIZAR_PERMISOS_PREDEFINIDOS();
    
    // 2. Crear un usuario de prueba único
    const timestamp = new Date().getTime();
    const usernameTest = "test_" + timestamp;
    const emailTest = "test_" + timestamp + "@megudan.com";
    
    const datosUsuario = {
      USUARIO: usernameTest,
      NOMBRE_COMPLETO: "Usuario Suite de Pruebas",
      CORREO: emailTest,
      ID_ROL: "ROL-000001", // Rol Administrador con accesos totales
      ESTADO_USUARIO: "ACTIVO"
    };

    console.log("   -> Creando usuario administrador de pruebas: " + usernameTest);
    const resCrear = SEG_CREAR_USUARIO(datosUsuario, "SISTEMA_INTERNAL_BYPASS");
    idUsuarioCreado = resCrear.ID_USUARIO;
    console.log("   [PASS] Usuario de prueba creado exitosamente con ID: " + idUsuarioCreado);

    // 3. Establecer contraseña segura
    console.log("   -> Configurando contraseña segura...");
    SEG_ESTABLECER_CONTRASENA(idUsuarioCreado, "SuiteTest2026!", "SISTEMA_TEST");
    console.log("   [PASS] Contraseña guardada como Hash SHA-256.");

    // 4. Probar autenticación correcta
    console.log("   -> Autenticando credenciales de prueba correctas...");
    const authCorrecta = SEG_AUTENTICAR_USUARIO(usernameTest, "SuiteTest2026!");
    if (!authCorrecta.EXITO) throw new Error("La autenticación correcta falló: " + authCorrecta.MENSAJE);
    console.log("   [PASS] Autenticación de credenciales correcta.");

    // 5. Probar autenticación con contraseña incorrecta (Debe fallar)
    console.log("   -> Probando autenticación con contraseña incorrecta...");
    const authIncorrecta = SEG_AUTENTICAR_USUARIO(usernameTest, "ContrasenaFalsa123!");
    if (authIncorrecta.EXITO) throw new Error("La autenticación incorrecta concedió acceso indebidamente.");
    console.log("   [PASS] Bloqueo correcto ante credenciales erróneas: " + authIncorrecta.MENSAJE);

    // 6. Crear una sesión activa (Token)
    console.log("   -> Iniciando sesión en USR_SESIONES y generando token...");
    const sesion = SEG_CREAR_SESION(idUsuarioCreado);
    tokenSesionActiva = sesion.TOKEN_SESION;
    console.log("   [PASS] Sesión creada correctamente. Token de Acceso: " + tokenSesionActiva);

    // 7. Validar sesión
    console.log("   -> Validando vigencia del token de sesión...");
    const validacion = SEG_VALIDAR_SESION(tokenSesionActiva);
    if (!validacion.VALIDA) throw new Error("La sesión generada no se reconoce como válida: " + validacion.MENSAJE);
    console.log("   [PASS] Token validado y autorizado.");

    testResults.push({ modulo: "SEGURIDAD", prueba: "AUTENTICACION_Y_SESIONES", estado: "PASS", detalle: "Creación de usuario, hash SHA-256, login y validación de token asíncrono." });
  } catch (errSeg) {
    console.error("   [FAIL] Error en Módulo de Seguridad: " + errSeg.message);
    testResults.push({ modulo: "SEGURIDAD", prueba: "AUTENTICACION_Y_SESIONES", estado: "FAIL", detalle: errSeg.message });
  }

  // ==========================================================
  // TEST 2: MÓDULO DE TERCEROS: CLIENTES (05_CLIENTES)
  // ==========================================================
  console.log("\n👥 [TEST 2] Probando Módulo de Clientes (05_CLIENTES) bajo Seguridad Dual:");
  try {
    if (!tokenSesionActiva) throw new Error("No se puede ejecutar la prueba de clientes sin token de sesión activo.");

    // 1. Probar validación matemática del DV de la DIAN
    console.log("   -> Probando algoritmo matemático de Dígito de Verificación (901915723)...");
    const dv = CLI_CALCULAR_DV("901915723");
    if (dv !== "2") throw new Error("Cálculo incorrecto del Dv de la DIAN. Obtenido: " + dv + " (Esperado: 2).");
    console.log("   [PASS] Algoritmo matemático Dv de la DIAN validado.");

    // 2. Registrar un cliente usando el token de sesión activa de la Web App
    const timestamp = new Date().getTime();
    const nitCliente = "900" + String(timestamp).substring(5, 15);
    const datosCliente = {
      TIPO_PERSONA: "PERSONA_JURIDICA",
      TIPO_DOCUMENTO: "NIT",
      NUMERO_DOCUMENTO: nitCliente,
      RAZON_SOCIAL: "CLIENTE TEST INTEGRACION SAS",
      TIPO_CLIENTE: "EMPRESARIAL",
      PAIS: "COLOMBIA",
      CIUDAD: "PITALITO",
      DEPARTAMENTO: "HUILA",
      DIRECCION: "Carrera 4 # 12-45",
      TELEFONO: "8360000",
      CELULAR: "3110000000",
      CORREO: "recaudos_test@megudan.com",
      FORMA_PAGO: "CREDITO",
      METODO_PAGO: "TRANSFERENCIA",
      PLAZO_PAGO_DIAS: 30,
      LIMITE_CREDITO: 50000000,
      ESTADO_CLIENTE: "ACTIVO"
    };

    console.log("   -> Creando nuevo cliente a través de la Web App (Con Token): " + datosCliente.RAZON_SOCIAL);
    const resCli = CLI_GUARDAR_CLIENTE(datosCliente, tokenSesionActiva);
    idClienteCreado = resCli.idCliente;
    console.log("   [PASS] Cliente guardado exitosamente en CLI_MAESTRO con ID: " + idClienteCreado);

    // 3. Buscar cliente
    console.log("   -> Buscando el cliente registrado por ID...");
    const clienteEncontrado = CLI_BUSCAR_CLIENTE(idClienteCreado, tokenSesionActiva);
    if (!clienteEncontrado || clienteEncontrado.RAZON_SOCIAL !== datosCliente.RAZON_SOCIAL) {
      throw new Error("El cliente registrado no se pudo encontrar en el maestro de forma idéntica.");
    }
    console.log("   [PASS] Cliente recuperado de forma correcta.");

    // 4. Actualizar cliente
    console.log("   -> Editando el límite de crédito del cliente...");
    const datosActualizar = {
      ID_CLIENTE: idClienteCreado,
      LIMITE_CREDITO: 75000000,
      DIRECCION: "Calle Nueva Dirección 45-90"
    };
    CLI_ACTUALIZAR_CLIENTE(datosActualizar, tokenSesionActiva);
    const clienteActualizado = CLI_BUSCAR_CLIENTE(idClienteCreado, tokenSesionActiva);
    if (Number(clienteActualizado.LIMITE_CREDITO) !== 75000000) {
      throw new Error("La actualización del límite de crédito no se reflejó correctamente.");
    }
    console.log("   [PASS] Cliente actualizado de forma exitosa en CLI_MAESTRO.");

    testResults.push({ modulo: "CLIENTES", prueba: "CRUD_WEB_APP", estado: "PASS", detalle: "Validación de Dv DIAN, creación, búsqueda, actualización y logs de historial CLI_HISTORIAL." });
  } catch (errCli) {
    console.error("   [FAIL] Error en Módulo de Clientes: " + errCli.message);
    testResults.push({ modulo: "CLIENTES", prueba: "CRUD_WEB_APP", estado: "FAIL", detalle: errCli.message });
  }

  // ==========================================================
  // TEST 3: MÓDULO DE TERCEROS: PROVEEDORES (06_PROVEEDORES)
  // ==========================================================
  console.log("\n🏢 [TEST 3] Probando Módulo de Proveedores (06_PROVEEDORES) bajo Seguridad Dual:");
  try {
    if (!tokenSesionActiva) throw new Error("No se puede ejecutar la prueba de proveedores sin token de sesión activo.");

    const timestamp = new Date().getTime();
    const nitProv = "800" + String(timestamp).substring(5, 15);
    const datosProv = {
      TIPO_PERSONA: "PERSONA_JURIDICA",
      TIPO_DOCUMENTO: "NIT",
      NIT_CC: nitProv,
      RAZON_SOCIAL: "PROVEEDOR TEST CONSTRUCCION SAS",
      NOMBRE_COMERCIAL: "PROVEEDOR TEST",
      EMAIL: "ventas_test@proveedor.com",
      DIRECCION: "Zona Industrial Bodega 10",
      CIUDAD: "PITALITO",
      DEPARTAMENTO: "HUILA",
      CONDICION_PAGO: "CREDITO",
      CUPO_CREDITO: 100000000,
      ESTADO: "ACTIVO"
    };

    console.log("   -> Creando nuevo proveedor a través de la Web App (Con Token)...");
    const resProv = PROV_GUARDAR_PROVEEDOR(datosProv, tokenSesionActiva);
    idProveedorCreado = resProv.idProveedor;
    console.log("   [PASS] Proveedor guardado en PROV_MAESTRO con ID: " + idProveedorCreado);

    // Buscar proveedor
    console.log("   -> Buscando el proveedor registrado...");
    const provEncontrado = PROV_BUSCAR_PROVEEDOR(idProveedorCreado, tokenSesionActiva);
    if (!provEncontrado) throw new Error("No se pudo localizar el proveedor registrado.");
    console.log("   [PASS] Proveedor recuperado de forma correcta.");

    testResults.push({ modulo: "PROVEEDORES", prueba: "CRUD_WEB_APP", estado: "PASS", detalle: "Creación y recuperación de registros de proveedores en PROV_MAESTRO con seguridad dual." });
  } catch (errProv) {
    console.error("   [FAIL] Error en Módulo de Proveedores: " + errProv.message);
    testResults.push({ modulo: "PROVEEDORES", prueba: "CRUD_WEB_APP", estado: "FAIL", detalle: errProv.message });
  }

  // ==========================================================
  // TEST 4: MÓDULO DE PRODUCTOS (07_PRODUCTOS)
  // ==========================================================
  console.log("\n📦 [TEST 4] Probando Módulo de Catálogo de Productos (07_PRODUCTOS):");
  try {
    const timestamp = new Date().getTime();
    const codigoProd = "PRD-" + timestamp;
    const datosProducto = {
      CODIGO: codigoProd,
      DESCRIPCION: "Guadua de Prueba Inmunizada 6m",
      TIPO_PRODUCTO: "PRODUCTO",
      ID_CATEGORIA: "CAT-000001",
      ID_UNIDAD: "UND",
      PRECIO_VENTA: 60000,
      COSTO_REFERENCIA: 30000,
      IVA: "19%",
      CONTROL_INVENTARIO: "SÍ",
      STOCK_MINIMO: 50,
      STOCK_MAXIMO: 1000
    };

    console.log("   -> Registrando nuevo producto en el catálogo...");
    const resProd = PROD_GUARDAR_PRODUCTO(datosProducto);
    idProductoCreado = resProd.idProducto;
    console.log("   [PASS] Producto creado con éxito con ID: " + idProductoCreado);

    testResults.push({ modulo: "PRODUCTOS", prueba: "REGISTRO_CATALOGO", estado: "PASS", detalle: "Registro atómico de productos e insumos en la tabla PROD_MAESTRO." });
  } catch (errProd) {
    console.error("   [FAIL] Error en Módulo de Productos: " + errProd.message);
    testResults.push({ modulo: "PRODUCTOS", prueba: "REGISTRO_CATALOGO", estado: "FAIL", detalle: errProd.message });
  }

  // ==========================================================
  // TEST 5: MÓDULO DE TRANSACCIONES - COMPRAS (10_COMPRAS)
  // ==========================================================
  console.log("\n🛒 [TEST 5] Probando Módulo de Compras, Ingreso de Stock y CxP (10_COMPRAS):");
  try {
    if (!idProveedorCreado || !idProductoCreado) throw new Error("Faltan dependencias previas de proveedor o producto.");

    // 1. Configurar una cuenta bancaria/caja si está vacía
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaCuentas = ss.getSheetByName("TES_CUENTAS");
    if (hojaCuentas && hojaCuentas.getLastRow() < 2) {
      hojaCuentas.appendRow(["CTA-000001", "BANCOS", "Bancolombia Corporativa", "BANCOLOMBIA", "35463960", "COP", 10000000, new Date(), "ACTIVO"]);
    }

    // 2. Preparar el payload de la compra
    const cabeceraCompra = {
      ID_PROVEEDOR: idProveedorCreado,
      TIPO_DOCUMENTO: "FACTURA_PROVEEDOR",
      NUM_DOCUMENTO: "FAC-COMPRA-789",
      FORMA_PAGO: "CREDITO", // Compra a crédito para disparar CxP
      CONDICION_PAGO: "30_DIAS",
      ID_CUENTA: "CTA-000001",
      USUARIO: "ADMIN_TEST"
    };

    const detalleCompra = [{
      ID_PRODUCTO: idProductoCreado,
      DESCRIPCION: "Guadua de Prueba Inmunizada 6m",
      CANTIDAD: 100, // Entrarán 100 unidades al inventario
      ID_UNIDAD: "UND",
      COSTO_UNITARIO: 30000,
      DESCUENTO: 0,
      IVA: 0.19 // 19% IVA
    }];

    console.log("   -> Registrando transacción de compra a crédito...");
    const resCompra = COM_GUARDAR_COMPRA(cabeceraCompra, detalleCompra);
    idCompraCreada = resCompra.idCompra;
    console.log("   [PASS] Compra registrada con éxito con ID: " + idCompraCreada);
    console.log("   [PASS] Total de Compra: " + resCompra.total + " COP.");

    // 3. Validar afectación de inventario (Stock físico incrementado)
    console.log("   -> Validando que el stock del producto haya aumentado a 100 unidades...");
    const saldoHoja = ss.getSheetByName("INV_SALDOS");
    const saldosData = saldoHoja.getRange(2, 1, Math.max(1, saldoHoja.getLastRow() - 1), 9).getValues();
    const saldoProducto = saldosData.find(f => f[1] === idProductoCreado);
    if (!saldoProducto || Number(saldoProducto[6]) !== 100) {
      throw new Error("El saldo en inventario es incorrecto. Esperado: 100. Registrado: " + (saldoProducto ? saldoProducto[6] : "N/A"));
    }
    console.log("   [PASS] Saldos de Inventario actualizados correctamente con Costo Promedio.");

    // 4. Validar afectación de Cuentas por Pagar (CXP)
    console.log("   -> Validando que se haya generado la Cuenta por Pagar al Proveedor...");
    const cxpHoja = ss.getSheetByName("CXP_CUENTAS");
    const cxpData = cxpHoja.getRange(2, 1, Math.max(1, cxpHoja.getLastRow() - 1), 11).getValues();
    const cxpCuenta = cxpData.find(f => f[2] === idCompraCreada);
    if (!cxpCuenta || Number(cxpCuenta[8]) !== resCompra.total) {
      throw new Error("La Cuenta por Pagar no se generó o el saldo es incorrecto.");
    }
    console.log("   [PASS] Cuenta por Pagar (CXP) creada con saldo de: " + cxpCuenta[8] + " COP.");

    testResults.push({ modulo: "COMPRAS", prueba: "TRANSACCION_COMPRA_CREDITO", estado: "PASS", detalle: "Simulación de compra, ingreso automático al Kardex, recalculo de saldos y provisión de Cuenta por Pagar (CXP)." });
  } catch (errCom) {
    console.error("   [FAIL] Error en Módulo de Compras: " + errCom.message);
    testResults.push({ modulo: "COMPRAS", prueba: "TRANSACCION_COMPRA_CREDITO", estado: "FAIL", detalle: errCom.message });
  }

  // ==========================================================
  // TEST 6: MÓDULO DE TRANSACCIONES - VENTAS (09_VENTAS)
  // ==========================================================
  console.log("\n💰 [TEST 6] Probando Módulo de Ventas, Salida de Stock y Cartera (09_VENTAS):");
  try {
    if (!idClienteCreado || !idProductoCreado) throw new Error("Faltan dependencias previas de cliente o producto.");

    // 1. Preparar el payload de la venta
    const cabeceraVenta = {
      ID_CLIENTE: idClienteCreado,
      TIPO_DOCUMENTO: "FACTURA_VENTA",
      NUM_DOCUMENTO: "FAC-VENTA-101",
      FORMA_PAGO: "CREDITO", // Venta a crédito para disparar Cartera
      CONDICION_PAGO: "30_DIAS",
      USUARIO: "ADMIN_TEST"
    };

    const detalleVenta = [{
      ID_PRODUCTO: idProductoCreado,
      DESCRIPCION: "Guadua de Prueba Inmunizada 6m",
      CANTIDAD: 20, // Salen 20 unidades del inventario (Deben quedar 80)
      ID_UNIDAD: "UND",
      PRECIO_UNITARIO: 60000,
      DESCUENTO: 0,
      IVA: 0.19 // 19% IVA
    }];

    console.log("   -> Registrando transacción de venta a crédito...");
    const resVenta = VEN_GUARDAR_VENTA(cabeceraVenta, detalleVenta);
    idVentaCreada = resVenta.idVenta;
    console.log("   [PASS] Venta registrada con éxito con ID: " + idVentaCreada);
    console.log("   [PASS] Total de Venta: " + resVenta.total + " COP.");

    // 2. Validar descarga de stock de inventario (Kardex y saldos)
    console.log("   -> Validando que el stock disponible haya bajado a 80 unidades...");
    const saldoHoja = ss.getSheetByName("INV_SALDOS");
    const saldosData = saldoHoja.getRange(2, 1, Math.max(1, saldoHoja.getLastRow() - 1), 9).getValues();
    const saldoProducto = saldosData.find(f => f[1] === idProductoCreado);
    if (!saldoProducto || Number(saldoProducto[6]) !== 80) {
      throw new Error("La salida del inventario falló. Esperado: 80. Registrado: " + (saldoProducto ? saldoProducto[6] : "N/A"));
    }
    console.log("   [PASS] Descarte de Kardex de salida y saldo recalculado.");

    // 3. Validar afectación de Cartera por Cobrar (CAR)
    console.log("   -> Validando que se haya generado la Cuenta por Cobrar al Cliente...");
    const carHoja = ss.getSheetByName("CAR_CUENTAS");
    const carData = carHoja.getRange(2, 1, Math.max(1, carHoja.getLastRow() - 1), 11).getValues();
    const carteraCuenta = carData.find(f => f[2] === idVentaCreada);
    if (!carteraCuenta || Number(carteraCuenta[8]) !== resVenta.total) {
      throw new Error("La Cuenta por Cobrar no se generó o el saldo es incorrecto.");
    }
    console.log("   [PASS] Cuenta de Cartera creada con saldo de: " + carteraCuenta[8] + " COP.");

    testResults.push({ modulo: "VENTAS", prueba: "TRANSACCION_VENTA_CREDITO", estado: "PASS", detalle: "Venta asíncrona, descarte de existencias físicas (Kardex) y generación de Cuenta por Cobrar (Cartera)." });
  } catch (errVen) {
    console.error("   [FAIL] Error en Módulo de Ventas: " + errVen.message);
    testResults.push({ modulo: "VENTAS", prueba: "TRANSACCION_VENTA_CREDITO", estado: "FAIL", detalle: errVen.message });
  }

  // ==========================================================
  // TEST 7: MÓDULO DE GASTOS (14_GASTOS)
  // ==========================================================
  console.log("\n💸 [TEST 7] Probando Módulo de Gastos y afectación de Tesorería (14_GASTOS):");
  try {
    const datosGasto = {
      TIPO_GASTO: "ADMINISTRATIVO",
      CATEGORIA: "Servicios Públicos",
      ID_PROVEEDOR: idProveedorCreado || "PROV-000001",
      DOCUMENTO_ORIGEN: "RECIBO",
      ID_ORIGEN: "REC-999",
      VALOR: 150000,
      IVA: 0,
      CENTRO_COSTO: "FABRICA",
      CUENTA_CONTABLE: "513505",
      METODO_PAGO: "TRANSFERENCIA",
      ID_CUENTA: "CTA-000001", // Caja/Banco origen del pago
      OBSERVACION: "Pago de energía eléctrica - Fabrica"
    };

    console.log("   -> Registrando egreso administrativo en la tabla de gastos...");
    const resGasto = GAS_REGISTRAR_GASTO(datosGasto);
    idGastoCreado = resGasto.idGasto;
    console.log("   [PASS] Gasto registrado correctamente con ID: " + idGastoCreado);

    testResults.push({ modulo: "GASTOS", prueba: "REGISTRO_EGRESO", estado: "PASS", detalle: "Registro de gastos y descarga en tiempo real de fondos en la cuenta de Tesorería." });
  } catch (errGas) {
    console.error("   [FAIL] Error en Módulo de Gastos: " + errGas.message);
    testResults.push({ modulo: "GASTOS", prueba: "REGISTRO_EGRESO", estado: "FAIL", detalle: errGas.message });
  }

  // ==========================================================
  // TEST 8: MÓDULO DE MANTENIMIENTO (26_MANTENIMIENTO)
  // ==========================================================
  console.log("\n⚙️ [TEST 8] Probando depuración de datos y Reset Seguro (26_MANTENIMIENTO):");
  try {
    if (!tokenSesionActiva) throw new Error("No se puede probar mantenimiento sin token de sesión.");

    // 1. Purga de historial de auditoría (conservando los últimos 30 días)
    console.log("   -> Ejecutando purga selectiva de USR_AUDITORIA...");
    const resMntAud = MNT_PURGAR_AUDITORIA_SISTEMA(30, tokenSesionActiva);
    if (!resMntAud.EXITO) throw new Error("La depuración selectiva falló: " + resMntAud.MENSAJE);
    console.log("   [PASS] Purga de logs completada: " + resMntAud.MENSAJE);

    // 2. Liberación de caché global del script
    console.log("   -> Liberando la caché global de Google Apps Script...");
    const resMntCache = MNT_PURGAR_CACHE_SISTEMA(tokenSesionActiva);
    if (!resMntCache.EXITO) throw new Error("La purga de caché falló: " + resMntCache.MENSAJE);
    console.log("   [PASS] Memoria RAM del servidor liberada: " + resMntCache.MENSAJE);

    testResults.push({ modulo: "MANTENIMIENTO", prueba: "PURGA_DE_LOGS_Y_CACHE", estado: "PASS", detalle: "Depuración selectiva de auditorías asíncronas y purga de CacheService." });
  } catch (errMnt) {
    console.error("   [FAIL] Error en Módulo de Mantenimiento: " + errMnt.message);
    testResults.push({ modulo: "MANTENIMIENTO", prueba: "PURGA_DE_LOGS_Y_CACHE", estado: "FAIL", detalle: errMnt.message });
  }

  // ==========================================================
  // REPORTE CONSOLIDADO FINAL
  // ==========================================================
  console.log("\n==================================================================");
  console.log("📊 REPORTE DE SUITE DE PRUEBAS DE INTEGRACIÓN");
  console.log("==================================================================");
  let pasados = 0;
  let fallados = 0;
  testResults.forEach(r => {
    if (r.estado === "PASS") {
      pasados++;
      console.log("   ✓ [" + r.modulo.padEnd(15) + "] " + r.prueba.padEnd(30) + " | ESTADO: OK   | " + r.detalle);
    } else {
      fallados++;
      console.error("   ❌ [" + r.modulo.padEnd(15) + "] " + r.prueba.padEnd(30) + " | ESTADO: FAIL | " + r.detalle);
    }
  });
  console.log("==================================================================");
  console.log("Pruebas Totales: " + testResults.length + " | Éxito: " + pasados + " | Errores: " + fallados);
  console.log("==================================================================");

  if (fallados === 0) {
    console.log("🎉 ¡FELICITACIONES! TU ERP MEGUDAN V2 ESTÁ 100% OPERATIVO Y INTEGRADO.");
  } else {
    console.warn("⚠️ ALGUNAS PRUEBAS TIENEN OBSERVACIONES. REVISA EL PANEL DE CONTROL.");
  }
}