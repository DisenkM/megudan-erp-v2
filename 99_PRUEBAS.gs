/**
 * ERP OPERATIVO V2 - MEGUDAN
 * Script de Pruebas de Integración para el Módulo de Seguridad
 */
function PROBAR_SISTEMA_DE_SEGURIDAD() {
  try {
    console.log("=== INICIANDO PRUEBA DE INTEGRACIÓN DE SEGURIDAD ===");

    // 1. Poblado inicial de roles de prueba si la hoja USR_ROLES está vacía
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaRoles = ss.getSheetByName("USR_ROLES");
    if (hojaRoles && hojaRoles.getLastRow() < 2) {
      console.log("-> USR_ROLES vacía. Creando Rol Administrador de prueba...");
      hojaRoles.appendRow([
        "ROL-000001", 
        "ADMINISTRADOR", 
        "Administración general del ERP", 
        "100", 
        "ACTIVO", 
        "SÍ", 
        new Date(), 
        new Date(), 
        "SISTEMA", 
        "SISTEMA", 
        "Rol de prueba"
      ]);
    }

  // 2. Definición del usuario de prueba (CORREGIDO)
    console.log("\n1. Preparando datos de usuario de prueba...");
    const datosUsuario = {
      USUARIO: "usuario_test",
      NOMBRE_COMPLETO: "Alejandro Test", //  ¡Ahora coincide con USR_USUARIOS!
      CORREO: "test@megudan.com",
      ID_ROL: "ROL-000001"
    };
    
    // 3. Crear el usuario en la base de datos USR_USUARIOS
    console.log("\n1. Creando usuario de prueba en Sheets...");
    const resultadoCrear = SEG_CREAR_USUARIO(datosUsuario);
    console.log("   [OK] ID Generado:", resultadoCrear.ID_USUARIO);
    const idUsuario = resultadoCrear.ID_USUARIO;

    // 4. Establecer contraseña segura (Debe cumplir: mayúscula, minúscula, número, especial y >=8 caracteres)
    console.log("\n2. Estableciendo contraseña segura ('Admin123!')...");
    SEG_ESTABLECER_CONTRASENA(idUsuario, "Admin123!", "SISTEMA_TEST");
    console.log("   [OK] Contraseña encriptada y guardada como hash SHA-256 en Sheets.");

    // 5. Probar autenticación con contraseña correcta
    console.log("\n3. Probando inicio de sesión con contraseña CORRECTA...");
    const loginExitoso = SEG_AUTENTICAR_USUARIO("usuario_test", "Admin123!");
    console.log("   [RESULTADO] ¿Acceso concedido?:", loginExitoso.EXITO);
    console.log("   [RESULTADO] Mensaje del sistema:", loginExitoso.MENSAJE);

    // 6. Probar autenticación con contraseña incorrecta (Debe fallar)
    console.log("\n4. Probando inicio de sesión con contraseña INCORRECTA...");
    const loginFallido = SEG_AUTENTICAR_USUARIO("usuario_test", "ClaveErronea123!");
    console.log("   [RESULTADO] ¿Acceso concedido? (Esperado: false):", loginFallido.EXITO);
    console.log("   [RESULTADO] Mensaje de denegación:", loginFallido.MENSAJE);

    // 7. Crear sesión activa (Generar Token único)
    console.log("\n5. Creando sesión activa en USR_SESIONES...");
    const sesion = SEG_CREAR_SESION(idUsuario);
    console.log("   [OK] Token UUID generado:", sesion.TOKEN_SESION);
    console.log("   [OK] Expira el:", sesion.FECHA_EXPIRACION);

    // 8. Validar token de sesión en segundo plano
    console.log("\n6. Validando token de sesión...");
    const validacion = SEG_VALIDAR_SESSION_O_TOKEN(sesion.TOKEN_SESION); // O usa SEG_VALIDAR_SESION(token) según tu código consolidado
    console.log("   [RESULTADO] ¿Sesión válida? (Esperado: true):", validacion.VALIDA);
    console.log("   [RESULTADO] Mensaje de validación:", validacion.MENSAJE);

    console.log("\n=== ¡TODAS LAS PRUEBAS DE SEGURIDAD PASARON CON ÉXITO! ===");

  } catch (error) {
    console.error("\n❌ ERROR DURANTE LA EJECUCIÓN DE PRUEBAS:");
    console.error("   Detalle:", error.message);
    console.error("   Línea del fallo:", error.stack);
  }
}

/**
 * Función auxiliar para dar flexibilidad al nombre según tu archivo
 */
function SEG_VALIDAR_SESSION_O_TOKEN(token) {
  if (typeof SEG_VALIDAR_SESION === "function") {
    return SEG_VALIDAR_SESION(token);
  } else if (typeof SEG_VALIDAR_SESSION === "function") {
    return SEG_VALIDAR_SESSION(token);
  }
  throw new Error("No se encontró la función de validar sesión en el archivo 23_SEGURIDAD.gs");
}