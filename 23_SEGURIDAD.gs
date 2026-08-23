// ============================================================
// 23. SEGURIDAD Y CONTROL DE ACCESO
// ARCHIVO: 23_SEGURIDAD.gs
// ============================================================


  // ============================================================
  // 01. CONFIGURACIÓN DEL MÓDULO DE SEGURIDAD
  // ============================================================

      const SEG_CONFIG = {

      // ----------------------------------------------------------
      // 01.01 HOJAS DEL MÓDULO
      // ----------------------------------------------------------

      HOJA_USUARIOS: "USR_USUARIOS",

      HOJA_ROLES: "USR_ROLES",

      HOJA_PERMISOS: "USR_PERMISOS",

      HOJA_SESIONES: "USR_SESIONES",

      HOJA_AUDITORIA: "USR_AUDITORIA",


      // ----------------------------------------------------------
      // 01.02 PREFIJOS DE IDENTIFICADORES
      // ----------------------------------------------------------

      PREFIJO_USUARIO: "USR",

      PREFIJO_ROL: "ROL",

      PREFIJO_PERMISO: "PER",

      PREFIJO_SESION: "SES",

      PREFIJO_AUDITORIA: "AUD",


      // ----------------------------------------------------------
      // 01.03 CONFIGURACIÓN DE IDENTIFICADORES
      // ----------------------------------------------------------

      DIGITOS_ID: 6,


      // ----------------------------------------------------------
      // 01.04 ESTADOS DE USUARIO
      // ----------------------------------------------------------

      ESTADO_USUARIO_ACTIVO: "ACTIVO",

      ESTADO_USUARIO_INACTIVO: "INACTIVO",

      ESTADO_USUARIO_BLOQUEADO: "BLOQUEADO",


      // ----------------------------------------------------------
      // 01.05 ESTADOS DE ROL
      // ----------------------------------------------------------

      ESTADO_ROL_ACTIVO: "ACTIVO",

      ESTADO_ROL_INACTIVO: "INACTIVO",


      // ----------------------------------------------------------
      // 01.06 ESTADOS DE PERMISOS
      // ----------------------------------------------------------

      ESTADO_PERMISO_ACTIVO: "ACTIVO",

      ESTADO_PERMISO_INACTIVO: "INACTIVO",


      // ----------------------------------------------------------
      // 01.07 ESTADOS DE SESIÓN
      // ----------------------------------------------------------

      ESTADO_SESION_ACTIVA: "ACTIVA",

      ESTADO_SESION_CERRADA: "CERRADA",

      ESTADO_SESION_EXPIRADA: "EXPIRADA",

      ESTADO_SESION_BLOQUEADA: "BLOQUEADA",

      ESTADO_SESION_REVOCADA: "REVOCADA",


      // ----------------------------------------------------------
      // 01.08 DURACIÓN DE SESIONES
      // ----------------------------------------------------------

      DURACION_SESION_HORAS: 8,

      TIEMPO_INACTIVIDAD_MINUTOS: 30,


      // ----------------------------------------------------------
      // 01.09 SEGURIDAD DE ACCESO
      // ----------------------------------------------------------

      MAXIMO_INTENTOS_LOGIN: 5,

      BLOQUEO_USUARIO_MINUTOS: 30,


      // ----------------------------------------------------------
      // 01.10 ORÍGENES DE ACCESO
      // ----------------------------------------------------------

      ORIGEN_GOOGLE_SHEETS: "GOOGLE_SHEETS",

      ORIGEN_APLICACION_WEB: "APLICACION_WEB",

      ORIGEN_SISTEMA: "SISTEMA",


      // ----------------------------------------------------------
      // 01.11 TIPOS DE ACCESO
      // ----------------------------------------------------------

      TIPO_ACCESO_WEB: "WEB",

      TIPO_ACCESO_GOOGLE_SHEETS: "GOOGLE_SHEETS",


      // ----------------------------------------------------------
      // 01.12 RESULTADOS DE OPERACIONES
      // ----------------------------------------------------------

      RESULTADO_EXITOSO: "EXITOSO",

      RESULTADO_DENEGADO: "DENEGADO",

      RESULTADO_ERROR: "ERROR",


      // ----------------------------------------------------------
      // 01.13 ACCIONES DEL SISTEMA
      // ----------------------------------------------------------

      ACCION_ACCEDER: "ACCEDER",

      ACCION_VER: "VER",

      ACCION_CREAR: "CREAR",

      ACCION_EDITAR: "EDITAR",

      ACCION_ELIMINAR: "ELIMINAR",

      ACCION_ANULAR: "ANULAR",

      ACCION_APROBAR: "APROBAR",

      ACCION_EXPORTAR: "EXPORTAR",

      ACCION_IMPORTAR: "IMPORTAR",

      ACCION_CONFIGURAR: "CONFIGURAR",

      ACCION_ADMINISTRAR: "ADMINISTRAR",

      ACCION_INICIAR_SESION: "INICIAR_SESION",

      ACCION_CERRAR_SESION: "CERRAR_SESION",


      // ----------------------------------------------------------
      // 01.14 CONFIGURACIÓN DEL ADMINISTRADOR INICIAL
      // ----------------------------------------------------------

      ROL_ADMINISTRADOR: "ADMINISTRADOR",


      // ----------------------------------------------------------
      // 01.15 CONFIGURACIÓN DE AUDITORÍA
      // ----------------------------------------------------------

      REGISTRAR_AUDITORIA: true,

      REGISTRAR_LOGIN: true,

      REGISTRAR_ERRORES_SEGURIDAD: true

    };

  // ============================================================
  // 02. UTILIDADES DE USUARIOS
  // Funciones de apoyo específicas para la gestión de usuarios.
  // ============================================================


    // ============================================================
    // 02.01 OBTENER HOJA DE USUARIOS
    // Obtiene la hoja USR_USUARIOS.
    // ============================================================

    function SEG_OBTENER_HOJA_USUARIOS() {

      const ss = SpreadsheetApp.getActiveSpreadsheet();

      const hoja = ss.getSheetByName(
        SEG_CONFIG.HOJA_USUARIOS
      );

      if (!hoja) {

        throw new Error(
          "No existe la hoja " +
          SEG_CONFIG.HOJA_USUARIOS + "."
        );

      }

      return hoja;

    }


    // ============================================================
    // 02.02 OBTENER ENCABEZADOS DE USUARIOS
    // Obtiene y normaliza los encabezados de USR_USUARIOS.
    // ============================================================

    function SEG_OBTENER_ENCABEZADOS_USUARIOS() {

      const hoja =
        SEG_OBTENER_HOJA_USUARIOS();

      const ultimaColumna =
        hoja.getLastColumn();

      if (ultimaColumna === 0) {

        throw new Error(
          "La hoja " +
          SEG_CONFIG.HOJA_USUARIOS +
          " no contiene encabezados."
        );

      }

      return hoja
        .getRange(
          1,
          1,
          1,
          ultimaColumna
        )
        .getDisplayValues()[0]
        .map(function(encabezado) {

          return String(encabezado || "")
            .trim()
            .toUpperCase();

        });

    }


    // ============================================================
    // 02.03 NORMALIZAR USUARIO
    // Elimina espacios innecesarios y convierte el usuario
    // a mayúsculas para realizar validaciones consistentes.
    // ============================================================

    function SEG_NORMALIZAR_USUARIO(usuario) {

      return String(usuario || "")
        .trim()
        .toUpperCase()
        .replace(/\s+/g, "");

    }


    // ============================================================
    // 02.04 NORMALIZAR CORREO
    // Elimina espacios innecesarios y convierte el correo
    // a minúsculas para realizar validaciones consistentes.
    // ============================================================

    function SEG_NORMALIZAR_CORREO(correo) {

      return String(correo || "")
        .trim()
        .toLowerCase();

    }


    // ============================================================
    // 02.05 GENERAR ID DE USUARIO
    // Genera el siguiente identificador disponible.
    // Ejemplo:
    // USR-000001
    // USR-000002
    // ============================================================

    function SEG_GENERAR_ID_USUARIO() {

      const hoja =
        SEG_OBTENER_HOJA_USUARIOS();

      const ultimaFila =
        hoja.getLastRow();


      // ----------------------------------------------------------
      // Si no existen registros
      // ----------------------------------------------------------

      if (ultimaFila < 2) {

        return SEG_CONFIG.PREFIJO_USUARIO +
          "-" +
          String(1).padStart(
            SEG_CONFIG.DIGITOS_ID,
            "0"
          );

      }


      // ----------------------------------------------------------
      // Obtener encabezados
      // ----------------------------------------------------------

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_USUARIOS();

      const indiceID =
        encabezados.indexOf("ID_USUARIO");

      if (indiceID === -1) {

        throw new Error(
          "No se encontró la columna ID_USUARIO."
        );

      }


      // ----------------------------------------------------------
      // Obtener IDs existentes
      // ----------------------------------------------------------

      const ids = hoja
        .getRange(
          2,
          indiceID + 1,
          ultimaFila - 1,
          1
        )
        .getDisplayValues()
        .flat();


      // ----------------------------------------------------------
      // Buscar el número más alto
      // ----------------------------------------------------------

      let numeroMayor = 0;

      ids.forEach(function(id) {

        const textoID =
          String(id || "").trim();

        if (
          textoID.startsWith(
            SEG_CONFIG.PREFIJO_USUARIO + "-"
          )
        ) {

          const numero = parseInt(
            textoID.replace(
              SEG_CONFIG.PREFIJO_USUARIO + "-",
              ""
            ),
            10
          );

          if (
            !isNaN(numero) &&
            numero > numeroMayor
          ) {

            numeroMayor = numero;

          }

        }

      });


      // ----------------------------------------------------------
      // Generar siguiente ID
      // ----------------------------------------------------------

      const siguienteNumero =
        numeroMayor + 1;

      return SEG_CONFIG.PREFIJO_USUARIO +
        "-" +
        String(siguienteNumero).padStart(
          SEG_CONFIG.DIGITOS_ID,
          "0"
        );

    }


    // ============================================================
    // 02.06 OBTENER REGISTROS DE USUARIOS
    // Obtiene todos los registros existentes en USR_USUARIOS.
    // ============================================================

    function SEG_OBTENER_USUARIOS() {

      const hoja =
        SEG_OBTENER_HOJA_USUARIOS();

      const ultimaFila =
        hoja.getLastRow();

      const ultimaColumna =
        hoja.getLastColumn();


      // ----------------------------------------------------------
      // Validar existencia de registros
      // ----------------------------------------------------------

      if (ultimaFila < 2) {

        return [];

      }


      // ----------------------------------------------------------
      // Obtener registros
      // ----------------------------------------------------------

      return hoja
        .getRange(
          2,
          1,
          ultimaFila - 1,
          ultimaColumna
        )
        .getDisplayValues();

    }


    // ============================================================
    // 02.07 BUSCAR USUARIO POR CAMPO
    // Busca un usuario por coincidencia exacta.
    // Ejemplos:
    // ID_USUARIO
    // USUARIO
    // CORREO
    // ============================================================

    function SEG_BUSCAR_USUARIO(
      campoBusqueda,
      valorBusqueda
    ) {

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_USUARIOS();

      const usuarios =
        SEG_OBTENER_USUARIOS();

      const campo =
        String(campoBusqueda || "")
          .trim()
          .toUpperCase();

      const indiceCampo =
        encabezados.indexOf(campo);

      if (indiceCampo === -1) {

        throw new Error(
          "No se encontró el campo " +
          campo +
          " en " +
          SEG_CONFIG.HOJA_USUARIOS + "."
        );

      }


      // ----------------------------------------------------------
      // Normalizar valor de búsqueda según el campo
      // ----------------------------------------------------------

      let valorNormalizado;

      if (campo === "CORREO") {

        valorNormalizado =
          SEG_NORMALIZAR_CORREO(valorBusqueda);

      } else {

        valorNormalizado =
          SEG_NORMALIZAR_USUARIO(valorBusqueda);

      }


      // ----------------------------------------------------------
      // Buscar coincidencia
      // ----------------------------------------------------------

      const filaEncontrada =
        usuarios.find(function(fila) {

          let valorRegistro;

          if (campo === "CORREO") {

            valorRegistro =
              SEG_NORMALIZAR_CORREO(
                fila[indiceCampo]
              );

          } else {

            valorRegistro =
              SEG_NORMALIZAR_USUARIO(
                fila[indiceCampo]
              );

          }

          return (
            valorRegistro === valorNormalizado
          );

        });


      // ----------------------------------------------------------
      // Usuario no encontrado
      // ----------------------------------------------------------

      if (!filaEncontrada) {

        return null;

      }


      // ----------------------------------------------------------
      // Convertir fila en objeto
      // ----------------------------------------------------------

      const usuario = {};

      encabezados.forEach(
        function(campoEncabezado, indice) {

          usuario[campoEncabezado] =
            filaEncontrada[indice] || "";

        }
      );


      return usuario;

    }


    // ============================================================
    // 02.08 BUSCAR FILA DE USUARIO
    // Busca la posición real de un usuario en Google Sheets.
    // Devuelve el número de fila o null.
    // ============================================================

    function SEG_BUSCAR_FILA_USUARIO(
      campoBusqueda,
      valorBusqueda
    ) {

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_USUARIOS();

      const usuarios =
        SEG_OBTENER_USUARIOS();

      const campo =
        String(campoBusqueda || "")
          .trim()
          .toUpperCase();

      const indiceCampo =
        encabezados.indexOf(campo);

      if (indiceCampo === -1) {

        throw new Error(
          "No se encontró el campo " +
          campo +
          " en " +
          SEG_CONFIG.HOJA_USUARIOS + "."
        );

      }


      // ----------------------------------------------------------
      // Normalizar valor de búsqueda
      // ----------------------------------------------------------

      let valorNormalizado;

      if (campo === "CORREO") {

        valorNormalizado =
          SEG_NORMALIZAR_CORREO(valorBusqueda);

      } else {

        valorNormalizado =
          SEG_NORMALIZAR_USUARIO(valorBusqueda);

      }


      // ----------------------------------------------------------
      // Buscar posición
      // ----------------------------------------------------------

      for (
        let i = 0;
        i < usuarios.length;
        i++
      ) {

        let valorRegistro;

        if (campo === "CORREO") {

          valorRegistro =
            SEG_NORMALIZAR_CORREO(
              usuarios[i][indiceCampo]
            );

        } else {

          valorRegistro =
            SEG_NORMALIZAR_USUARIO(
              usuarios[i][indiceCampo]
            );

        }


        if (
          valorRegistro === valorNormalizado
        ) {

          // La fila 1 contiene los encabezados.
          return i + 2;

        }

      }


      return null;

    }


    // ============================================================
    // 02.09 VALIDAR USUARIO DUPLICADO
    // Verifica si el nombre de usuario ya existe.
    // Permite excluir un ID_USUARIO al momento de editar.
    // ============================================================

    function SEG_VALIDAR_USUARIO_DUPLICADO(
      nombreUsuario,
      idUsuarioExcluir
    ) {

      const usuario =
        SEG_BUSCAR_USUARIO(
          "USUARIO",
          nombreUsuario
        );

      if (!usuario) {

        return false;

      }


      // ----------------------------------------------------------
      // Si estamos editando el mismo usuario,
      // no se considera duplicado.
      // ----------------------------------------------------------

      if (
        idUsuarioExcluir &&
        usuario.ID_USUARIO === idUsuarioExcluir
      ) {

        return false;

      }


      return true;

    }


    // ============================================================
    // 02.10 VALIDAR CORREO DUPLICADO
    // Verifica si el correo ya está registrado.
    // Permite excluir un ID_USUARIO al momento de editar.
    // ============================================================

    function SEG_VALIDAR_CORREO_DUPLICADO(
      correo,
      idUsuarioExcluir
    ) {

      const usuario =
        SEG_BUSCAR_USUARIO(
          "CORREO",
          correo
        );

      if (!usuario) {

        return false;

      }


      // ----------------------------------------------------------
      // Si estamos editando el mismo usuario,
      // no se considera duplicado.
      // ----------------------------------------------------------

      if (
        idUsuarioExcluir &&
        usuario.ID_USUARIO === idUsuarioExcluir
      ) {

        return false;

      }


      return true;

    }


    // ============================================================
    // 02.11 OBTENER FECHA Y HORA ACTUAL
    // Centraliza la fecha utilizada en los registros de usuarios.
    // ============================================================

    function SEG_AHORA() {

      return new Date();

    }
    
  // ============================================================
  // 03. CRUD DE USUARIOS
  // Creación, consulta, actualización y control de usuarios.
  // ============================================================

    // ============================================================
    // 03.01 CREAR USUARIO
    // Crea un nuevo usuario en USR_USUARIOS.
    // ============================================================

    function SEG_CREAR_USUARIO(datos) {

      // ----------------------------------------------------------
      // 03.01.01 VALIDAR DATOS RECIBIDOS
      // ----------------------------------------------------------

      if (!datos || typeof datos !== "object") {

        throw new Error(
          "Debe proporcionar la información del usuario."
        );

      }


      // ----------------------------------------------------------
      // 03.01.02 OBTENER HOJA Y ENCABEZADOS
      // ----------------------------------------------------------

      const hoja =
        SEG_OBTENER_HOJA_USUARIOS();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_USUARIOS();


      // ----------------------------------------------------------
      // 03.01.03 VALIDAR CAMPOS OBLIGATORIOS
      // ----------------------------------------------------------
      // IMPORTANTE:
      // Esta lista debe coincidir con los campos obligatorios
      // definidos en USR_USUARIOS.
      // ----------------------------------------------------------

      const camposObligatorios = [
        "USUARIO",
        "NOMBRE",
        "CORREO",
        "ID_ROL"
      ];

      camposObligatorios.forEach(function(campo) {

        if (
          datos[campo] === undefined ||
          datos[campo] === null ||
          String(datos[campo]).trim() === ""
        ) {

          throw new Error(
            "El campo " + campo + " es obligatorio."
          );

        }

      });


      // ----------------------------------------------------------
      // 03.01.04 NORMALIZAR DATOS PRINCIPALES
      // ----------------------------------------------------------

      datos.USUARIO =
        SEG_NORMALIZAR_USUARIO(
          datos.USUARIO
        );

      datos.CORREO =
        SEG_NORMALIZAR_CORREO(
          datos.CORREO
        );

      datos.NOMBRE =
        String(datos.NOMBRE || "")
          .trim();


      // ----------------------------------------------------------
      // 03.01.05 VALIDAR USUARIO DUPLICADO
      // ----------------------------------------------------------

      if (
        SEG_VALIDAR_USUARIO_DUPLICADO(
          datos.USUARIO
        )
      ) {

        throw new Error(
          "El usuario " +
          datos.USUARIO +
          " ya se encuentra registrado."
        );

      }


      // ----------------------------------------------------------
      // 03.01.06 VALIDAR CORREO DUPLICADO
      // ----------------------------------------------------------

      if (
        SEG_VALIDAR_CORREO_DUPLICADO(
          datos.CORREO
        )
      ) {

        throw new Error(
          "El correo " +
          datos.CORREO +
          " ya se encuentra registrado."
        );

      }


      // ----------------------------------------------------------
      // 03.01.07 GENERAR DATOS DEL SISTEMA
      // ----------------------------------------------------------

      const ahora =
        SEG_AHORA();

      const idUsuario =
        SEG_GENERAR_ID_USUARIO();


      // ----------------------------------------------------------
      // 03.01.08 CONSTRUIR OBJETO DEL USUARIO
      // ----------------------------------------------------------
      // Se parte de los datos recibidos y se agregan
      // automáticamente los campos del sistema.
      // ----------------------------------------------------------

      const usuario = {};

      encabezados.forEach(function(campo) {

        usuario[campo] =
          datos[campo] !== undefined
            ? datos[campo]
            : "";

      });


      // ----------------------------------------------------------
      // 03.01.09 ASIGNAR VALORES AUTOMÁTICOS
      // ----------------------------------------------------------

      usuario.ID_USUARIO =
        idUsuario;

      usuario.USUARIO =
        datos.USUARIO;

      usuario.NOMBRE =
        datos.NOMBRE;

      usuario.CORREO =
        datos.CORREO;

      usuario.ESTADO_USUARIO =
        datos.ESTADO_USUARIO ||
        SEG_CONFIG.ESTADO_USUARIO_ACTIVO;

      usuario.FECHA_CREACION =
        ahora;

      usuario.FECHA_ACTUALIZACION =
        ahora;


      // ----------------------------------------------------------
      // 03.01.10 ASIGNAR USUARIO DE AUDITORÍA
      // ----------------------------------------------------------

      usuario.USUARIO_CREACION =
        datos.USUARIO_CREACION ||
        "SISTEMA";

      usuario.USUARIO_ACTUALIZACION =
        datos.USUARIO_ACTUALIZACION ||
        usuario.USUARIO_CREACION;


      // ----------------------------------------------------------
      // 03.01.11 CONVERTIR OBJETO EN FILA
      // ----------------------------------------------------------

      const nuevaFila =
        encabezados.map(function(campo) {

          return usuario[campo] !== undefined
            ? usuario[campo]
            : "";

        });


      // ----------------------------------------------------------
      // 03.01.12 GUARDAR USUARIO
      // ----------------------------------------------------------

      hoja.appendRow(nuevaFila);


      // ----------------------------------------------------------
      // 03.01.13 DEVOLVER RESULTADO
      // ----------------------------------------------------------

      return {

        EXITO: true,

        MENSAJE:
          "Usuario creado correctamente.",

        ID_USUARIO:
          idUsuario,

        USUARIO:
          usuario

      };

    }


    // ============================================================
    // 03.02 CONSULTAR USUARIO
    // Busca un usuario por ID_USUARIO.
    // ============================================================

    function SEG_CONSULTAR_USUARIO(idUsuario) {

      if (
        !idUsuario ||
        String(idUsuario).trim() === ""
      ) {

        throw new Error(
          "Debe indicar el ID_USUARIO."
        );

      }

      return SEG_BUSCAR_USUARIO(
        "ID_USUARIO",
        idUsuario
      );

    }


    // ============================================================
    // 03.03 ACTUALIZAR USUARIO
    // Actualiza la información permitida de un usuario.
    // ============================================================

    function SEG_ACTUALIZAR_USUARIO(
      idUsuario,
      datos
    ) {

      // ----------------------------------------------------------
      // 03.03.01 VALIDAR DATOS
      // ----------------------------------------------------------

      if (
        !idUsuario ||
        String(idUsuario).trim() === ""
      ) {

        throw new Error(
          "Debe indicar el ID_USUARIO."
        );

      }

      if (!datos || typeof datos !== "object") {

        throw new Error(
          "Debe proporcionar los datos a actualizar."
        );

      }


      // ----------------------------------------------------------
      // 03.03.02 BUSCAR USUARIO EXISTENTE
      // ----------------------------------------------------------

      const usuarioActual =
        SEG_CONSULTAR_USUARIO(idUsuario);

      if (!usuarioActual) {

        throw new Error(
          "No se encontró el usuario " +
          idUsuario + "."
        );

      }


      // ----------------------------------------------------------
      // 03.03.03 OBTENER HOJA Y ENCABEZADOS
      // ----------------------------------------------------------

      const hoja =
        SEG_OBTENER_HOJA_USUARIOS();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_USUARIOS();

      const filaUsuario =
        SEG_BUSCAR_FILA_USUARIO(
          "ID_USUARIO",
          idUsuario
        );


      // ----------------------------------------------------------
      // 03.03.04 PROTEGER CAMPOS DEL SISTEMA
      // ----------------------------------------------------------

      const camposProtegidos = [
        "ID_USUARIO",
        "FECHA_CREACION",
        "USUARIO_CREACION"
      ];


      // ----------------------------------------------------------
      // 03.03.05 CONSTRUIR NUEVA INFORMACIÓN
      // ----------------------------------------------------------

      const usuarioActualizado =
        Object.assign(
          {},
          usuarioActual
        );

      Object.keys(datos).forEach(function(campo) {

        const campoNormalizado =
          String(campo)
            .trim()
            .toUpperCase();

        if (
          !camposProtegidos.includes(
            campoNormalizado
          )
        ) {

          usuarioActualizado[campoNormalizado] =
            datos[campo];

        }

      });


      // ----------------------------------------------------------
      // 03.03.06 NORMALIZAR USUARIO
      // ----------------------------------------------------------

      if (datos.USUARIO !== undefined) {

        usuarioActualizado.USUARIO =
          SEG_NORMALIZAR_USUARIO(
            datos.USUARIO
          );

        if (
          SEG_VALIDAR_USUARIO_DUPLICADO(
            usuarioActualizado.USUARIO,
            idUsuario
          )
        ) {

          throw new Error(
            "El nombre de usuario ya está registrado."
          );

        }

      }


      // ----------------------------------------------------------
      // 03.03.07 NORMALIZAR CORREO
      // ----------------------------------------------------------

      if (datos.CORREO !== undefined) {

        usuarioActualizado.CORREO =
          SEG_NORMALIZAR_CORREO(
            datos.CORREO
          );

        if (
          SEG_VALIDAR_CORREO_DUPLICADO(
            usuarioActualizado.CORREO,
            idUsuario
          )
        ) {

          throw new Error(
            "El correo ya está registrado."
          );

        }

      }


      // ----------------------------------------------------------
      // 03.03.08 ACTUALIZAR AUDITORÍA
      // ----------------------------------------------------------

      usuarioActualizado.FECHA_ACTUALIZACION =
        SEG_AHORA();

      usuarioActualizado.USUARIO_ACTUALIZACION =
        datos.USUARIO_ACTUALIZACION ||
        "SISTEMA";


      // ----------------------------------------------------------
      // 03.03.09 CONVERTIR OBJETO EN FILA
      // ----------------------------------------------------------

      const filaActualizada =
        encabezados.map(function(campo) {

          return usuarioActualizado[campo] !== undefined
            ? usuarioActualizado[campo]
            : "";

        });


      // ----------------------------------------------------------
      // 03.03.10 GUARDAR ACTUALIZACIÓN
      // ----------------------------------------------------------

      hoja
        .getRange(
          filaUsuario,
          1,
          1,
          encabezados.length
        )
        .setValues([
          filaActualizada
        ]);


      // ----------------------------------------------------------
      // 03.03.11 DEVOLVER RESULTADO
      // ----------------------------------------------------------

      return {

        EXITO: true,

        MENSAJE:
          "Usuario actualizado correctamente.",

        ID_USUARIO:
          idUsuario,

        USUARIO:
          usuarioActualizado

      };

    }


    // ============================================================
    // 03.04 CAMBIAR ESTADO DE USUARIO
    // Cambia el estado de un usuario.
    // Estados permitidos:
    // ACTIVO / INACTIVO / BLOQUEADO
    // ============================================================

    function SEG_CAMBIAR_ESTADO_USUARIO(
      idUsuario,
      nuevoEstado,
      usuarioActualizacion
    ) {

      const estado =
        String(nuevoEstado || "")
          .trim()
          .toUpperCase();


      // ----------------------------------------------------------
      // 03.04.01 VALIDAR ESTADO
      // ----------------------------------------------------------

      const estadosPermitidos = [
        SEG_CONFIG.ESTADO_USUARIO_ACTIVO,
        SEG_CONFIG.ESTADO_USUARIO_INACTIVO,
        SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO
      ];

      if (
        !estadosPermitidos.includes(estado)
      ) {

        throw new Error(
          "Estado de usuario no válido: " +
          nuevoEstado
        );

      }


      // ----------------------------------------------------------
      // 03.04.02 ACTUALIZAR USUARIO
      // ----------------------------------------------------------

      return SEG_ACTUALIZAR_USUARIO(
        idUsuario,
        {

          ESTADO_USUARIO: estado,

          USUARIO_ACTUALIZACION:
            usuarioActualizacion || "SISTEMA"

        }
      );

    }


    // ============================================================
    // 03.05 BLOQUEAR USUARIO
    // Cambia el estado del usuario a BLOQUEADO.
    // ============================================================

    function SEG_BLOQUEAR_USUARIO(
      idUsuario,
      usuarioActualizacion
    ) {

      return SEG_CAMBIAR_ESTADO_USUARIO(
        idUsuario,
        SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO,
        usuarioActualizacion
      );

    }


    // ============================================================
    // 03.06 DESBLOQUEAR USUARIO
    // Cambia el estado del usuario nuevamente a ACTIVO.
    // ============================================================

    function SEG_DESBLOQUEAR_USUARIO(
      idUsuario,
      usuarioActualizacion
    ) {

      return SEG_CAMBIAR_ESTADO_USUARIO(
        idUsuario,
        SEG_CONFIG.ESTADO_USUARIO_ACTIVO,
        usuarioActualizacion
      );

    }


  // ============================================================
  // 04. CRUD DE ROLES
  // Creación, consulta, actualización y control de roles.
  // ============================================================


    // ============================================================
    // 04.01 OBTENER HOJA DE ROLES
    // Obtiene la hoja USR_ROLES.
    // ============================================================

    function SEG_OBTENER_HOJA_ROLES() {

      const ss = SpreadsheetApp.getActiveSpreadsheet();

      const hoja = ss.getSheetByName(
        SEG_CONFIG.HOJA_ROLES
      );

      if (!hoja) {

        throw new Error(
          "No existe la hoja " +
          SEG_CONFIG.HOJA_ROLES + "."
        );

      }

      return hoja;

    }


    // ============================================================
    // 04.02 OBTENER ENCABEZADOS DE ROLES
    // Obtiene y normaliza los encabezados de USR_ROLES.
    // ============================================================

    function SEG_OBTENER_ENCABEZADOS_ROLES() {

      const hoja =
        SEG_OBTENER_HOJA_ROLES();

      const ultimaColumna =
        hoja.getLastColumn();

      if (ultimaColumna === 0) {

        throw new Error(
          "La hoja " +
          SEG_CONFIG.HOJA_ROLES +
          " no contiene encabezados."
        );

      }

      return hoja
        .getRange(
          1,
          1,
          1,
          ultimaColumna
        )
        .getDisplayValues()[0]
        .map(function(encabezado) {

          return String(encabezado || "")
            .trim()
            .toUpperCase();

        });

    }


    // ============================================================
    // 04.03 NORMALIZAR NOMBRE DE ROL
    // Elimina espacios innecesarios y convierte el nombre
    // del rol a mayúsculas para realizar validaciones.
    // ============================================================

    function SEG_NORMALIZAR_ROL(nombreRol) {

      return String(nombreRol || "")
        .trim()
        .toUpperCase()
        .replace(/\s+/g, " ");

    }


    // ============================================================
    // 04.04 OBTENER REGISTROS DE ROLES
    // Obtiene todos los registros existentes en USR_ROLES.
    // ============================================================

    function SEG_OBTENER_ROLES() {

      const hoja =
        SEG_OBTENER_HOJA_ROLES();

      const ultimaFila =
        hoja.getLastRow();

      const ultimaColumna =
        hoja.getLastColumn();

      if (ultimaFila < 2) {

        return [];

      }

      return hoja
        .getRange(
          2,
          1,
          ultimaFila - 1,
          ultimaColumna
        )
        .getDisplayValues();

    }


    // ============================================================
    // 04.05 GENERAR ID DE ROL
    // Genera el siguiente identificador disponible.
    // Ejemplo:
    // ROL-000001
    // ROL-000002
    // ============================================================

    function SEG_GENERAR_ID_ROL() {

      const hoja =
        SEG_OBTENER_HOJA_ROLES();

      const ultimaFila =
        hoja.getLastRow();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_ROLES();

      const indiceID =
        encabezados.indexOf("ID_ROL");

      if (indiceID === -1) {

        throw new Error(
          "No se encontró la columna ID_ROL."
        );

      }

      if (ultimaFila < 2) {

        return SEG_CONFIG.PREFIJO_ROL +
          "-" +
          String(1).padStart(
            SEG_CONFIG.DIGITOS_ID,
            "0"
          );

      }


      // ----------------------------------------------------------
      // Obtener IDs existentes
      // ----------------------------------------------------------

      const ids = hoja
        .getRange(
          2,
          indiceID + 1,
          ultimaFila - 1,
          1
        )
        .getDisplayValues()
        .flat();

      let numeroMayor = 0;

      ids.forEach(function(id) {

        const textoID =
          String(id || "").trim();

        if (
          textoID.startsWith(
            SEG_CONFIG.PREFIJO_ROL + "-"
          )
        ) {

          const numero = parseInt(
            textoID.replace(
              SEG_CONFIG.PREFIJO_ROL + "-",
              ""
            ),
            10
          );

          if (
            !isNaN(numero) &&
            numero > numeroMayor
          ) {

            numeroMayor = numero;

          }

        }

      });


      // ----------------------------------------------------------
      // Generar siguiente ID
      // ----------------------------------------------------------

      return SEG_CONFIG.PREFIJO_ROL +
        "-" +
        String(numeroMayor + 1).padStart(
          SEG_CONFIG.DIGITOS_ID,
          "0"
        );

    }


    // ============================================================
    // 04.06 BUSCAR ROL
    // Busca un rol por coincidencia exacta.
    // Puede buscar por:
    // - ID_ROL
    // - NOMBRE_ROL
    // ============================================================

    function SEG_BUSCAR_ROL(
      campoBusqueda,
      valorBusqueda
    ) {

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_ROLES();

      const roles =
        SEG_OBTENER_ROLES();

      const campo =
        String(campoBusqueda || "")
          .trim()
          .toUpperCase();

      const indiceCampo =
        encabezados.indexOf(campo);

      if (indiceCampo === -1) {

        throw new Error(
          "No se encontró el campo " +
          campo +
          " en " +
          SEG_CONFIG.HOJA_ROLES + "."
        );

      }

      const valorNormalizado =
        SEG_NORMALIZAR_ROL(valorBusqueda);


      // ----------------------------------------------------------
      // Buscar coincidencia
      // ----------------------------------------------------------

      const filaEncontrada =
        roles.find(function(fila) {

          const valorRegistro =
            SEG_NORMALIZAR_ROL(
              fila[indiceCampo]
            );

          return (
            valorRegistro === valorNormalizado
          );

        });


      if (!filaEncontrada) {

        return null;

      }


      // ----------------------------------------------------------
      // Convertir fila en objeto
      // ----------------------------------------------------------

      const rol = {};

      encabezados.forEach(function(campo, indice) {

        rol[campo] =
          filaEncontrada[indice] || "";

      });

      return rol;

    }


    // ============================================================
    // 04.07 BUSCAR FILA DE ROL
    // Busca la fila física del rol en USR_ROLES.
    // Devuelve el número de fila o null.
    // ============================================================

    function SEG_BUSCAR_FILA_ROL(
      campoBusqueda,
      valorBusqueda
    ) {

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_ROLES();

      const roles =
        SEG_OBTENER_ROLES();

      const campo =
        String(campoBusqueda || "")
          .trim()
          .toUpperCase();

      const indiceCampo =
        encabezados.indexOf(campo);

      if (indiceCampo === -1) {

        throw new Error(
          "No se encontró el campo " +
          campo +
          " en " +
          SEG_CONFIG.HOJA_ROLES + "."
        );

      }

      const valorNormalizado =
        SEG_NORMALIZAR_ROL(valorBusqueda);


      // ----------------------------------------------------------
      // Buscar posición
      // ----------------------------------------------------------

      for (
        let i = 0;
        i < roles.length;
        i++
      ) {

        const valorRegistro =
          SEG_NORMALIZAR_ROL(
            roles[i][indiceCampo]
          );

        if (
          valorRegistro === valorNormalizado
        ) {

          return i + 2;

        }

      }

      return null;

    }


    // ============================================================
    // 04.08 VALIDAR ROL DUPLICADO
    // Verifica si el nombre del rol ya existe.
    // Permite excluir un ID_ROL al momento de editar.
    // ============================================================

    function SEG_VALIDAR_ROL_DUPLICADO(
      nombreRol,
      idRolExcluir
    ) {

      const rol =
        SEG_BUSCAR_ROL(
          "NOMBRE_ROL",
          nombreRol
        );

      if (!rol) {

        return false;

      }

      if (
        idRolExcluir &&
        rol.ID_ROL === idRolExcluir
      ) {

        return false;

      }

      return true;

    }


    // ============================================================
    // 04.09 CREAR ROL
    // Crea un nuevo rol en USR_ROLES.
    // ============================================================

    function SEG_CREAR_ROL(datos) {

      // ----------------------------------------------------------
      // VALIDAR DATOS
      // ----------------------------------------------------------

      if (!datos || typeof datos !== "object") {

        throw new Error(
          "Debe proporcionar la información del rol."
        );

      }

      if (
        !datos.NOMBRE_ROL ||
        String(datos.NOMBRE_ROL).trim() === ""
      ) {

        throw new Error(
          "El campo NOMBRE_ROL es obligatorio."
        );

      }


      // ----------------------------------------------------------
      // OBTENER HOJA Y ENCABEZADOS
      // ----------------------------------------------------------

      const hoja =
        SEG_OBTENER_HOJA_ROLES();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_ROLES();


      // ----------------------------------------------------------
      // NORMALIZAR NOMBRE
      // ----------------------------------------------------------

      const nombreRol =
        SEG_NORMALIZAR_ROL(
          datos.NOMBRE_ROL
        );


      // ----------------------------------------------------------
      // VALIDAR DUPLICADO
      // ----------------------------------------------------------

      if (
        SEG_VALIDAR_ROL_DUPLICADO(nombreRol)
      ) {

        throw new Error(
          "El rol " +
          nombreRol +
          " ya se encuentra registrado."
        );

      }


      // ----------------------------------------------------------
      // GENERAR DATOS AUTOMÁTICOS
      // ----------------------------------------------------------

      const ahora =
        SEG_AHORA();

      const idRol =
        SEG_GENERAR_ID_ROL();


      // ----------------------------------------------------------
      // CONSTRUIR OBJETO
      // ----------------------------------------------------------

      const rol = {};

      encabezados.forEach(function(campo) {

        rol[campo] =
          datos[campo] !== undefined
            ? datos[campo]
            : "";

      });


      // ----------------------------------------------------------
      // ASIGNAR DATOS DEL SISTEMA
      // ----------------------------------------------------------

      rol.ID_ROL =
        idRol;

      rol.NOMBRE_ROL =
        nombreRol;

      rol.ESTADO_ROL =
        datos.ESTADO_ROL ||
        "ACTIVO";

      rol.FECHA_CREACION =
        ahora;

      rol.FECHA_ACTUALIZACION =
        ahora;

      rol.USUARIO_CREACION =
        datos.USUARIO_CREACION ||
        "SISTEMA";

      rol.USUARIO_ACTUALIZACION =
        datos.USUARIO_ACTUALIZACION ||
        rol.USUARIO_CREACION;


      // ----------------------------------------------------------
      // CONVERTIR A FILA
      // ----------------------------------------------------------

      const nuevaFila =
        encabezados.map(function(campo) {

          return rol[campo] !== undefined
            ? rol[campo]
            : "";

        });


      // ----------------------------------------------------------
      // GUARDAR ROL
      // ----------------------------------------------------------

      hoja.appendRow(nuevaFila);


      return {

        EXITO: true,

        MENSAJE:
          "Rol creado correctamente.",

        ID_ROL:
          idRol,

        ROL:
          rol

      };

    }


    // ============================================================
    // 04.10 CONSULTAR ROL
    // Busca un rol por ID_ROL.
    // ============================================================

    function SEG_CONSULTAR_ROL(idRol) {

      if (
        !idRol ||
        String(idRol).trim() === ""
      ) {

        throw new Error(
          "Debe indicar el ID_ROL."
        );

      }

      return SEG_BUSCAR_ROL(
        "ID_ROL",
        idRol
      );

    }


    // ============================================================
    // 04.11 ACTUALIZAR ROL
    // Actualiza la información permitida de un rol.
    // ============================================================

    function SEG_ACTUALIZAR_ROL(
      idRol,
      datos
    ) {

      const rolActual =
        SEG_CONSULTAR_ROL(idRol);

      if (!rolActual) {

        throw new Error(
          "No se encontró el rol " +
          idRol + "."
        );

      }

      if (!datos || typeof datos !== "object") {

        throw new Error(
          "Debe proporcionar los datos a actualizar."
        );

      }


      // ----------------------------------------------------------
      // OBTENER INFORMACIÓN
      // ----------------------------------------------------------

      const hoja =
        SEG_OBTENER_HOJA_ROLES();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_ROLES();

      const filaRol =
        SEG_BUSCAR_FILA_ROL(
          "ID_ROL",
          idRol
        );


      // ----------------------------------------------------------
      // PROTEGER CAMPOS DEL SISTEMA
      // ----------------------------------------------------------

      const camposProtegidos = [
        "ID_ROL",
        "FECHA_CREACION",
        "USUARIO_CREACION"
      ];


      // ----------------------------------------------------------
      // CONSTRUIR ROL ACTUALIZADO
      // ----------------------------------------------------------

      const rolActualizado =
        Object.assign({}, rolActual);

      Object.keys(datos).forEach(function(campo) {

        const campoNormalizado =
          String(campo)
            .trim()
            .toUpperCase();

        if (
          !camposProtegidos.includes(
            campoNormalizado
          )
        ) {

          rolActualizado[campoNormalizado] =
            datos[campo];

        }

      });


      // ----------------------------------------------------------
      // VALIDAR NOMBRE DEL ROL
      // ----------------------------------------------------------

      if (datos.NOMBRE_ROL !== undefined) {

        rolActualizado.NOMBRE_ROL =
          SEG_NORMALIZAR_ROL(
            datos.NOMBRE_ROL
          );

        if (
          SEG_VALIDAR_ROL_DUPLICADO(
            rolActualizado.NOMBRE_ROL,
            idRol
          )
        ) {

          throw new Error(
            "El nombre del rol ya está registrado."
          );

        }

      }


      // ----------------------------------------------------------
      // ACTUALIZAR AUDITORÍA
      // ----------------------------------------------------------

      rolActualizado.FECHA_ACTUALIZACION =
        SEG_AHORA();

      rolActualizado.USUARIO_ACTUALIZACION =
        datos.USUARIO_ACTUALIZACION ||
        "SISTEMA";


      // ----------------------------------------------------------
      // GUARDAR ACTUALIZACIÓN
      // ----------------------------------------------------------

      const filaActualizada =
        encabezados.map(function(campo) {

          return rolActualizado[campo] !== undefined
            ? rolActualizado[campo]
            : "";

        });

      hoja
        .getRange(
          filaRol,
          1,
          1,
          encabezados.length
        )
        .setValues([
          filaActualizada
        ]);


      return {

        EXITO: true,

        MENSAJE:
          "Rol actualizado correctamente.",

        ID_ROL:
          idRol,

        ROL:
          rolActualizado

      };

    }


    // ============================================================
    // 04.12 CAMBIAR ESTADO DE ROL
    // Cambia el estado entre ACTIVO e INACTIVO.
    // ============================================================

    function SEG_CAMBIAR_ESTADO_ROL(
      idRol,
      nuevoEstado,
      usuarioActualizacion
    ) {

      const estado =
        String(nuevoEstado || "")
          .trim()
          .toUpperCase();

      const estadosPermitidos = [
        "ACTIVO",
        "INACTIVO"
      ];

      if (
        !estadosPermitidos.includes(estado)
      ) {

        throw new Error(
          "Estado de rol no válido: " +
          nuevoEstado
        );

      }

      return SEG_ACTUALIZAR_ROL(
        idRol,
        {

          ESTADO_ROL: estado,

          USUARIO_ACTUALIZACION:
            usuarioActualizacion || "SISTEMA"

        }
      );

    }


    // ============================================================
    // 04.13 ACTIVAR ROL
    // Cambia el estado del rol a ACTIVO.
    // ============================================================

    function SEG_ACTIVAR_ROL(
      idRol,
      usuarioActualizacion
    ) {

      return SEG_CAMBIAR_ESTADO_ROL(
        idRol,
        "ACTIVO",
        usuarioActualizacion
      );

    }


    // ============================================================
    // 04.14 INACTIVAR ROL
    // Cambia el estado del rol a INACTIVO.
    // ============================================================

    function SEG_INACTIVAR_ROL(
      idRol,
      usuarioActualizacion
    ) {

      return SEG_CAMBIAR_ESTADO_ROL(
        idRol,
        "INACTIVO",
        usuarioActualizacion
      );

    }

  // ============================================================
  // 05. GESTIÓN Y VALIDACIÓN DE PERMISOS
  // Administración, consulta y validación de permisos
  // asignados a los roles del ERP.
  // ============================================================


    // ============================================================
    // 05.01 OBTENER HOJA DE PERMISOS
    // Obtiene la hoja USR_PERMISOS.
    // ============================================================

    function SEG_OBTENER_HOJA_PERMISOS() {

      const ss = SpreadsheetApp.getActiveSpreadsheet();

      const hoja = ss.getSheetByName(
        SEG_CONFIG.HOJA_PERMISOS
      );

      if (!hoja) {

        throw new Error(
          "No existe la hoja " +
          SEG_CONFIG.HOJA_PERMISOS + "."
        );

      }

      return hoja;

    }


    // ============================================================
    // 05.02 OBTENER ENCABEZADOS DE PERMISOS
    // Obtiene y normaliza los encabezados de USR_PERMISOS.
    // ============================================================

    function SEG_OBTENER_ENCABEZADOS_PERMISOS() {

      const hoja =
        SEG_OBTENER_HOJA_PERMISOS();

      const ultimaColumna =
        hoja.getLastColumn();

      if (ultimaColumna === 0) {

        throw new Error(
          "La hoja " +
          SEG_CONFIG.HOJA_PERMISOS +
          " no contiene encabezados."
        );

      }

      return hoja
        .getRange(
          1,
          1,
          1,
          ultimaColumna
        )
        .getDisplayValues()[0]
        .map(function(encabezado) {

          return String(encabezado || "")
            .trim()
            .toUpperCase();

        });

    }


    // ============================================================
    // 05.03 OBTENER REGISTROS DE PERMISOS
    // Obtiene todos los registros existentes.
    // ============================================================

    function SEG_OBTENER_PERMISOS() {

      const hoja =
        SEG_OBTENER_HOJA_PERMISOS();

      const ultimaFila =
        hoja.getLastRow();

      const ultimaColumna =
        hoja.getLastColumn();

      if (ultimaFila < 2) {

        return [];

      }

      return hoja
        .getRange(
          2,
          1,
          ultimaFila - 1,
          ultimaColumna
        )
        .getDisplayValues();

    }


    // ============================================================
    // 05.04 NORMALIZAR VALOR DE PERMISO
    // Estandariza valores para comparaciones.
    // ============================================================

    function SEG_NORMALIZAR_PERMISO(valor) {

      return String(valor || "")
        .trim()
        .toUpperCase()
        .replace(/\s+/g, "_");

    }


    // ============================================================
    // 05.05 BUSCAR PERMISOS POR ROL
    // Devuelve todos los permisos asignados a un ID_ROL.
    // ============================================================

    function SEG_OBTENER_PERMISOS_ROL(idRol) {

      if (
        !idRol ||
        String(idRol).trim() === ""
      ) {

        throw new Error(
          "Debe indicar el ID_ROL."
        );

      }

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_PERMISOS();

      const registros =
        SEG_OBTENER_PERMISOS();

      const idxIdRol =
        encabezados.indexOf("ID_ROL");

      if (idxIdRol === -1) {

        throw new Error(
          "No se encontró la columna ID_ROL en " +
          SEG_CONFIG.HOJA_PERMISOS + "."
        );

      }


      // ----------------------------------------------------------
      // Buscar permisos pertenecientes al rol
      // ----------------------------------------------------------

      const permisos =
        registros.filter(function(fila) {

          return (
            String(fila[idxIdRol] || "")
              .trim()
              .toUpperCase() ===
            String(idRol)
              .trim()
              .toUpperCase()
          );

        });


      // ----------------------------------------------------------
      // Convertir filas en objetos
      // ----------------------------------------------------------

      return permisos.map(function(fila) {

        const permiso = {};

        encabezados.forEach(function(campo, indice) {

          permiso[campo] =
            fila[indice] || "";

        });

        return permiso;

      });

    }


    // ============================================================
    // 05.06 BUSCAR PERMISO ESPECÍFICO
    // Busca un permiso de un rol según módulo y acción.
    // ============================================================

    function SEG_BUSCAR_PERMISO(
      idRol,
      modulo,
      accion
    ) {

      const permisos =
        SEG_OBTENER_PERMISOS_ROL(idRol);

      const moduloNormalizado =
        SEG_NORMALIZAR_PERMISO(modulo);

      const accionNormalizada =
        SEG_NORMALIZAR_PERMISO(accion);


      return permisos.find(function(permiso) {

        const moduloRegistro =
          SEG_NORMALIZAR_PERMISO(
            permiso.MODULO
          );

        const accionRegistro =
          SEG_NORMALIZAR_PERMISO(
            permiso.ACCION
          );

        return (
          moduloRegistro === moduloNormalizado &&
          accionRegistro === accionNormalizada
        );

      }) || null;

    }


    // ============================================================
    // 05.07 VALIDAR PERMISO DE ROL
    // Determina si un rol tiene autorización para ejecutar
    // una acción dentro de un módulo.
    // ============================================================

    function SEG_VALIDAR_PERMISO_ROL(
      idRol,
      modulo,
      accion
    ) {

      // ----------------------------------------------------------
      // VALIDAR ROL
      // ----------------------------------------------------------

      const rol =
        SEG_CONSULTAR_ROL(idRol);

      if (!rol) {

        return false;

      }

      if (
        String(rol.ESTADO_ROL || "")
          .trim()
          .toUpperCase() !== "ACTIVO"
      ) {

        return false;

      }


      // ----------------------------------------------------------
      // BUSCAR PERMISO
      // ----------------------------------------------------------

      const permiso =
        SEG_BUSCAR_PERMISO(
          idRol,
          modulo,
          accion
        );

      if (!permiso) {

        return false;

      }


      // ----------------------------------------------------------
      // VALIDAR ESTADO DEL PERMISO
      // ----------------------------------------------------------
      // Se admite ESTADO_PERMISO como:
      // ACTIVO / INACTIVO
      // ----------------------------------------------------------

      if (
        String(
          permiso.ESTADO_PERMISO || ""
        )
          .trim()
          .toUpperCase() !== "ACTIVO"
      ) {

        return false;

      }


      // ----------------------------------------------------------
      // VALIDAR AUTORIZACIÓN
      // ----------------------------------------------------------
      // Se admite PERMITIDO como:
      // SI / SÍ / TRUE / VERDADERO / 1
      // ----------------------------------------------------------

      const permitido =
        String(permiso.PERMITIDO || "")
          .trim()
          .toUpperCase();

      const valoresPermitidos = [
        "SI",
        "SÍ",
        "TRUE",
        "VERDADERO",
        "1"
      ];

      return valoresPermitidos.includes(
        permitido
      );

    }


    // ============================================================
    // 05.08 VALIDAR PERMISO DE USUARIO
    // Obtiene el rol del usuario y valida si dicho rol tiene
    // permiso para ejecutar una acción.
    // ============================================================

    function SEG_VALIDAR_PERMISO_USUARIO(
      idUsuario,
      modulo,
      accion
    ) {

      // ----------------------------------------------------------
      // CONSULTAR USUARIO
      // ----------------------------------------------------------

      const usuario =
        SEG_CONSULTAR_USUARIO(idUsuario);

      if (!usuario) {

        return false;

      }


      // ----------------------------------------------------------
      // VALIDAR ESTADO DEL USUARIO
      // ----------------------------------------------------------

      if (
        String(usuario.ESTADO_USUARIO || "")
          .trim()
          .toUpperCase() !==
        SEG_CONFIG.ESTADO_USUARIO_ACTIVO
      ) {

        return false;

      }


      // ----------------------------------------------------------
      // VALIDAR QUE TENGA ROL
      // ----------------------------------------------------------

      if (
        !usuario.ID_ROL ||
        String(usuario.ID_ROL).trim() === ""
      ) {

        return false;

      }


      // ----------------------------------------------------------
      // VALIDAR PERMISO DEL ROL
      // ----------------------------------------------------------

      return SEG_VALIDAR_PERMISO_ROL(
        usuario.ID_ROL,
        modulo,
        accion
      );

    }


    // ============================================================
    // 05.09 VERIFICAR ACCESO
    // Igual que validar permiso, pero genera error cuando
    // el usuario no tiene autorización.
    // Útil antes de ejecutar procesos críticos.
    // ============================================================

    function SEG_VERIFICAR_ACCESO(
      idUsuario,
      modulo,
      accion
    ) {

      const autorizado =
        SEG_VALIDAR_PERMISO_USUARIO(
          idUsuario,
          modulo,
          accion
        );

      if (!autorizado) {

        throw new Error(
          "ACCESO DENEGADO. " +
          "El usuario no tiene permiso para " +
          accion +
          " en el módulo " +
          modulo + "."
        );

      }

      return true;

    }


    // ============================================================
    // 05.10 LISTAR PERMISOS DE USUARIO
    // Devuelve todos los permisos correspondientes al rol
    // asignado al usuario.
    // ============================================================

    function SEG_LISTAR_PERMISOS_USUARIO(idUsuario) {

      const usuario =
        SEG_CONSULTAR_USUARIO(idUsuario);

      if (!usuario) {

        throw new Error(
          "No se encontró el usuario " +
          idUsuario + "."
        );

      }

      if (
        !usuario.ID_ROL ||
        String(usuario.ID_ROL).trim() === ""
      ) {

        return [];

      }

      return SEG_OBTENER_PERMISOS_ROL(
        usuario.ID_ROL
      );

    }

  // ============================================================
  // 06. AUTENTICACIÓN Y CONTRASEÑAS
  // Gestión de credenciales, hash, validación y cambio
  // de contraseñas de los usuarios del ERP.
  // ============================================================


    // ============================================================
    // 06.01 GENERAR HASH DE CONTRASEÑA
    // Convierte una contraseña en un hash SHA-256.
    // La contraseña original no debe almacenarse.
    // ============================================================

    function SEG_GENERAR_HASH_CONTRASENA(contrasena) {

      if (
        contrasena === undefined ||
        contrasena === null ||
        String(contrasena) === ""
      ) {

        throw new Error(
          "Debe proporcionar una contraseña."
        );

      }


      // ----------------------------------------------------------
      // Convertir contraseña a bytes
      // ----------------------------------------------------------

      const bytes =
        Utilities.newBlob(
          String(contrasena)
        ).getBytes();


      // ----------------------------------------------------------
      // Generar hash SHA-256
      // ----------------------------------------------------------

      const hashBytes =
        Utilities.computeDigest(
          Utilities.DigestAlgorithm.SHA_256,
          bytes
        );


      // ----------------------------------------------------------
      // Convertir hash a hexadecimal
      // ----------------------------------------------------------

      return hashBytes
        .map(function(byte) {

          const valor =
            byte < 0
              ? byte + 256
              : byte;

          return (
            "0" +
            valor.toString(16)
          ).slice(-2);

        })
        .join("");

    }


    // ============================================================
    // 06.02 VALIDAR SEGURIDAD DE CONTRASEÑA
    // Valida requisitos mínimos de seguridad.
    // ============================================================

    function SEG_VALIDAR_SEGURIDAD_CONTRASENA(contrasena) {

      const password =
        String(contrasena || "");


      // ----------------------------------------------------------
      // RESULTADO DE VALIDACIÓN
      // ----------------------------------------------------------

      const resultado = {

        VALIDA: false,

        MENSAJES: []

      };


      // ----------------------------------------------------------
      // LONGITUD MÍNIMA
      // ----------------------------------------------------------

      if (password.length < 8) {

        resultado.MENSAJES.push(
          "La contraseña debe tener mínimo 8 caracteres."
        );

      }


      // ----------------------------------------------------------
      // MAYÚSCULA
      // ----------------------------------------------------------

      if (!/[A-Z]/.test(password)) {

        resultado.MENSAJES.push(
          "La contraseña debe incluir una letra mayúscula."
        );

      }


      // ----------------------------------------------------------
      // MINÚSCULA
      // ----------------------------------------------------------

      if (!/[a-z]/.test(password)) {

        resultado.MENSAJES.push(
          "La contraseña debe incluir una letra minúscula."
        );

      }


      // ----------------------------------------------------------
      // NÚMERO
      // ----------------------------------------------------------

      if (!/[0-9]/.test(password)) {

        resultado.MENSAJES.push(
          "La contraseña debe incluir un número."
        );

      }


      // ----------------------------------------------------------
      // CARÁCTER ESPECIAL
      // ----------------------------------------------------------

      if (
        !/[!@#$%^&*()_+\-=\[\]{};':"\\|,.<>/?]/.test(
          password
        )
      ) {

        resultado.MENSAJES.push(
          "La contraseña debe incluir un carácter especial."
        );

      }


      // ----------------------------------------------------------
      // RESULTADO FINAL
      // ----------------------------------------------------------

      resultado.VALIDA =
        resultado.MENSAJES.length === 0;


      return resultado;

    }


    // ============================================================
    // 06.03 BUSCAR USUARIO PARA AUTENTICACIÓN
    // Busca un usuario por:
    // - USUARIO
    // - CORREO
    // ============================================================

    function SEG_BUSCAR_USUARIO_LOGIN(credencial) {

      if (
        !credencial ||
        String(credencial).trim() === ""
      ) {

        return null;

      }


      // ----------------------------------------------------------
      // OBTENER USUARIOS
      // ----------------------------------------------------------

      const hoja =
        SEG_OBTENER_HOJA_USUARIOS();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_USUARIOS();

      const ultimaFila =
        hoja.getLastRow();

      const ultimaColumna =
        hoja.getLastColumn();


      if (ultimaFila < 2) {

        return null;

      }


      // ----------------------------------------------------------
      // UBICAR COLUMNAS
      // ----------------------------------------------------------

      const idxUsuario =
        encabezados.indexOf("USUARIO");

      const idxCorreo =
        encabezados.indexOf("CORREO");


      if (idxUsuario === -1) {

        throw new Error(
          "No existe la columna USUARIO."
        );

      }

      if (idxCorreo === -1) {

        throw new Error(
          "No existe la columna CORREO."
        );

      }


      // ----------------------------------------------------------
      // NORMALIZAR CREDENCIAL
      // ----------------------------------------------------------

      const criterio =
        String(credencial)
          .trim()
          .toUpperCase();


      // ----------------------------------------------------------
      // OBTENER REGISTROS
      // ----------------------------------------------------------

      const registros =
        hoja.getRange(
          2,
          1,
          ultimaFila - 1,
          ultimaColumna
        ).getDisplayValues();


      // ----------------------------------------------------------
      // BUSCAR USUARIO
      // ----------------------------------------------------------

      const filaEncontrada =
        registros.find(function(fila) {

          const usuario =
            String(fila[idxUsuario] || "")
              .trim()
              .toUpperCase();

          const correo =
            String(fila[idxCorreo] || "")
              .trim()
              .toUpperCase();

          return (
            usuario === criterio ||
            correo === criterio
          );

        });


      if (!filaEncontrada) {

        return null;

      }


      // ----------------------------------------------------------
      // CONVERTIR FILA EN OBJETO
      // ----------------------------------------------------------

      const usuario = {};

      encabezados.forEach(function(campo, indice) {

        usuario[campo] =
          filaEncontrada[indice] || "";

      });


      return usuario;

    }


    // ============================================================
    // 06.04 ESTABLECER CONTRASEÑA
    // Crea o reemplaza la contraseña de un usuario.
    // La contraseña se almacena únicamente como HASH.
    // ============================================================

    function SEG_ESTABLECER_CONTRASENA(
      idUsuario,
      nuevaContrasena,
      usuarioActualizacion
    ) {

      // ----------------------------------------------------------
      // VALIDAR USUARIO
      // ----------------------------------------------------------

      const usuario =
        SEG_CONSULTAR_USUARIO(idUsuario);

      if (!usuario) {

        throw new Error(
          "No se encontró el usuario indicado."
        );

      }


      // ----------------------------------------------------------
      // VALIDAR SEGURIDAD
      // ----------------------------------------------------------

      const validacion =
        SEG_VALIDAR_SEGURIDAD_CONTRASENA(
          nuevaContrasena
        );

      if (!validacion.VALIDA) {

        throw new Error(
          validacion.MENSAJES.join(" ")
        );

      }


      // ----------------------------------------------------------
      // GENERAR HASH
      // ----------------------------------------------------------

      const hash =
        SEG_GENERAR_HASH_CONTRASENA(
          nuevaContrasena
        );


      // ----------------------------------------------------------
      // ACTUALIZAR USUARIO
      // ----------------------------------------------------------

      return SEG_ACTUALIZAR_USUARIO(
        idUsuario,
        {

          HASH_CONTRASENA: hash,

          FECHA_CAMBIO_CONTRASENA:
            SEG_AHORA(),

          DEBE_CAMBIAR_CONTRASENA:
            "NO",

          USUARIO_ACTUALIZACION:
            usuarioActualizacion || "SISTEMA"

        }
      );

    }


    // ============================================================
    // 06.05 VALIDAR CONTRASEÑA
    // Compara una contraseña ingresada contra el hash
    // almacenado para el usuario.
    // ============================================================

    function SEG_VALIDAR_CONTRASENA(
      usuario,
      contrasena
    ) {

      if (!usuario) {

        return false;

      }

      if (
        !usuario.HASH_CONTRASENA ||
        String(usuario.HASH_CONTRASENA).trim() === ""
      ) {

        return false;

      }


      // ----------------------------------------------------------
      // GENERAR HASH DE CONTRASEÑA INGRESADA
      // ----------------------------------------------------------

      const hashIngresado =
        SEG_GENERAR_HASH_CONTRASENA(
          contrasena
        );


      // ----------------------------------------------------------
      // COMPARAR HASHES
      // ----------------------------------------------------------

      return (
        String(hashIngresado) ===
        String(usuario.HASH_CONTRASENA)
      );

    }


    // ============================================================
    // 06.06 AUTENTICAR USUARIO
    // Valida:
    // 1. Usuario o correo.
    // 2. Existencia del usuario.
    // 3. Estado.
    // 4. Contraseña.
    // ============================================================

    function SEG_AUTENTICAR_USUARIO(
      credencial,
      contrasena
    ) {

      // ----------------------------------------------------------
      // BUSCAR USUARIO
      // ----------------------------------------------------------

      const usuario =
        SEG_BUSCAR_USUARIO_LOGIN(
          credencial
        );


      if (!usuario) {

        return {

          EXITO: false,

          CODIGO:
            "USUARIO_NO_ENCONTRADO",

          MENSAJE:
            "Usuario o correo incorrecto."

        };

      }


      // ----------------------------------------------------------
      // VALIDAR ESTADO
      // ----------------------------------------------------------

      const estado =
        String(usuario.ESTADO_USUARIO || "")
          .trim()
          .toUpperCase();


      if (
        estado ===
        SEG_CONFIG.ESTADO_USUARIO_BLOQUEADO
      ) {

        return {

          EXITO: false,

          CODIGO:
            "USUARIO_BLOQUEADO",

          MENSAJE:
            "El usuario se encuentra bloqueado."

        };

      }


      if (
        estado !==
        SEG_CONFIG.ESTADO_USUARIO_ACTIVO
      ) {

        return {

          EXITO: false,

          CODIGO:
            "USUARIO_INACTIVO",

          MENSAJE:
            "El usuario no se encuentra activo."

        };

      }


      // ----------------------------------------------------------
      // VALIDAR CONTRASEÑA
      // ----------------------------------------------------------

      const contrasenaValida =
        SEG_VALIDAR_CONTRASENA(
          usuario,
          contrasena
        );


      if (!contrasenaValida) {

        return {

          EXITO: false,

          CODIGO:
            "CONTRASENA_INCORRECTA",

          MENSAJE:
            "Credenciales incorrectas."

        };

      }


      // ----------------------------------------------------------
      // AUTENTICACIÓN CORRECTA
      // ----------------------------------------------------------

      return {

        EXITO: true,

        CODIGO:
          "AUTENTICACION_CORRECTA",

        MENSAJE:
          "Autenticación correcta.",

        USUARIO: {

          ID_USUARIO:
            usuario.ID_USUARIO,

          USUARIO:
            usuario.USUARIO,

          NOMBRE:
            usuario.NOMBRE,

          CORREO:
            usuario.CORREO,

          ID_ROL:
            usuario.ID_ROL

        },

        DEBE_CAMBIAR_CONTRASENA:
          String(
            usuario.DEBE_CAMBIAR_CONTRASENA || ""
          )
            .trim()
            .toUpperCase() === "SI" ||
          String(
            usuario.DEBE_CAMBIAR_CONTRASENA || ""
          )
            .trim()
            .toUpperCase() === "SÍ"

      };

    }


    // ============================================================
    // 06.07 CAMBIAR CONTRASEÑA
    // Valida la contraseña actual antes de establecer una nueva.
    // ============================================================

    function SEG_CAMBIAR_CONTRASENA(
      idUsuario,
      contrasenaActual,
      nuevaContrasena,
      usuarioActualizacion
    ) {

      // ----------------------------------------------------------
      // CONSULTAR USUARIO
      // ----------------------------------------------------------

      const usuario =
        SEG_CONSULTAR_USUARIO(idUsuario);

      if (!usuario) {

        throw new Error(
          "No se encontró el usuario."
        );

      }


      // ----------------------------------------------------------
      // VALIDAR CONTRASEÑA ACTUAL
      // ----------------------------------------------------------

      const contrasenaActualValida =
        SEG_VALIDAR_CONTRASENA(
          usuario,
          contrasenaActual
        );


      if (!contrasenaActualValida) {

        throw new Error(
          "La contraseña actual es incorrecta."
        );

      }


      // ----------------------------------------------------------
      // VALIDAR QUE LA NUEVA SEA DIFERENTE
      // ----------------------------------------------------------

      if (
        String(contrasenaActual) ===
        String(nuevaContrasena)
      ) {

        throw new Error(
          "La nueva contraseña debe ser diferente."
        );

      }


      // ----------------------------------------------------------
      // ESTABLECER NUEVA CONTRASEÑA
      // ----------------------------------------------------------

      return SEG_ESTABLECER_CONTRASENA(
        idUsuario,
        nuevaContrasena,
        usuarioActualizacion ||
          usuario.USUARIO
      );

    }


  // ============================================================
  // 07. GESTIÓN DE SESIONES
  // Creación, validación, renovación y cierre de sesiones
  // de usuarios del ERP.
  // ============================================================


    // ============================================================
    // 07.01 CONFIGURACIÓN DE SESIONES
    // Define la duración y parámetros generales.
    // ============================================================

    const SEG_SESIONES_CONFIG = {

      DURACION_HORAS: 8,

      ESTADO_ACTIVA: "ACTIVA",

      ESTADO_CERRADA: "CERRADA",

      ESTADO_EXPIRADA: "EXPIRADA"

    };


    // ============================================================
    // 07.02 GENERAR TOKEN DE SESIÓN
    // Genera un identificador único para cada sesión.
    // ============================================================

    function SEG_GENERAR_TOKEN_SESION() {

      return Utilities.getUuid() +
        "-" +
        Utilities.getUuid();

    }


    // ============================================================
    // 07.03 GENERAR ID DE SESIÓN
    // Genera un identificador interno.
    // Ejemplo: SES-000001
    // ============================================================

    function SEG_GENERAR_ID_SESION() {

      const ss =
        SpreadsheetApp.getActiveSpreadsheet();

      const hoja =
        ss.getSheetByName("USR_SESIONES");

      if (!hoja) {

        throw new Error(
          "No existe la hoja USR_SESIONES."
        );

      }

      const ultimaFila =
        hoja.getLastRow();

      const ultimaColumna =
        hoja.getLastColumn();


      if (ultimaFila < 2) {

        return "SES-000001";

      }


      // ----------------------------------------------------------
      // OBTENER ENCABEZADOS
      // ----------------------------------------------------------

      const encabezados =
        hoja
          .getRange(
            1,
            1,
            1,
            ultimaColumna
          )
          .getDisplayValues()[0]
          .map(function(encabezado) {

            return String(encabezado || "")
              .trim()
              .toUpperCase();

          });


      const idxIdSesion =
        encabezados.indexOf("ID_SESION");

      if (idxIdSesion === -1) {

        throw new Error(
          "No se encontró la columna ID_SESION en USR_SESIONES."
        );

      }


      // ----------------------------------------------------------
      // OBTENER IDS EXISTENTES
      // ----------------------------------------------------------

      const ids =
        hoja
          .getRange(
            2,
            idxIdSesion + 1,
            ultimaFila - 1,
            1
          )
          .getDisplayValues()
          .flat();


      let numeroMayor = 0;


      ids.forEach(function(id) {

        const texto =
          String(id || "")
            .trim()
            .toUpperCase();

        if (
          texto.startsWith("SES-")
        ) {

          const numero =
            parseInt(
              texto.replace("SES-", ""),
              10
            );

          if (
            !isNaN(numero) &&
            numero > numeroMayor
          ) {

            numeroMayor = numero;

          }

        }

      });


      // ----------------------------------------------------------
      // GENERAR SIGUIENTE ID
      // ----------------------------------------------------------

      return "SES-" +
        String(numeroMayor + 1)
          .padStart(6, "0");

    }


    // ============================================================
    // 07.04 OBTENER HOJA DE SESIONES
    // ============================================================

    function SEG_OBTENER_HOJA_SESIONES() {

      const ss =
        SpreadsheetApp.getActiveSpreadsheet();

      const hoja =
        ss.getSheetByName("USR_SESIONES");

      if (!hoja) {

        throw new Error(
          "No existe la hoja USR_SESIONES."
        );

      }

      return hoja;

    }


    // ============================================================
    // 07.05 OBTENER ENCABEZADOS DE SESIONES
    // ============================================================

    function SEG_OBTENER_ENCABEZADOS_SESIONES() {

      const hoja =
        SEG_OBTENER_HOJA_SESIONES();

      const ultimaColumna =
        hoja.getLastColumn();

      if (ultimaColumna === 0) {

        throw new Error(
          "La hoja USR_SESIONES no contiene encabezados."
        );

      }

      return hoja
        .getRange(
          1,
          1,
          1,
          ultimaColumna
        )
        .getDisplayValues()[0]
        .map(function(encabezado) {

          return String(encabezado || "")
            .trim()
            .toUpperCase();

        });

    }


    // ============================================================
    // 07.06 CREAR SESIÓN
    // Crea una nueva sesión para un usuario autenticado.
    // ============================================================

    function SEG_CREAR_SESION(idUsuario) {

      // ----------------------------------------------------------
      // VALIDAR USUARIO
      // ----------------------------------------------------------

      const usuario =
        SEG_CONSULTAR_USUARIO(idUsuario);

      if (!usuario) {

        throw new Error(
          "No se encontró el usuario para crear la sesión."
        );

      }


      // ----------------------------------------------------------
      // VALIDAR ESTADO
      // ----------------------------------------------------------

      if (
        String(usuario.ESTADO_USUARIO || "")
          .trim()
          .toUpperCase() !==
        SEG_CONFIG.ESTADO_USUARIO_ACTIVO
      ) {

        throw new Error(
          "No se puede crear una sesión para un usuario inactivo."
        );

      }


      // ----------------------------------------------------------
      // GENERAR DATOS DE SESIÓN
      // ----------------------------------------------------------

      const ahora =
        new Date();

      const fechaExpiracion =
        new Date(
          ahora.getTime() +
          (
            SEG_SESIONES_CONFIG.DURACION_HORAS *
            60 *
            60 *
            1000
          )
        );

      const idSesion =
        SEG_GENERAR_ID_SESION();

      const tokenSesion =
        SEG_GENERAR_TOKEN_SESION();


      // ----------------------------------------------------------
      // OBTENER HOJA Y ENCABEZADOS
      // ----------------------------------------------------------

      const hoja =
        SEG_OBTENER_HOJA_SESIONES();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_SESIONES();


      // ----------------------------------------------------------
      // CREAR OBJETO DE SESIÓN
      // ----------------------------------------------------------

      const sesion = {};

      encabezados.forEach(function(campo) {

        sesion[campo] = "";

      });


      sesion.ID_SESION =
        idSesion;

      sesion.ID_USUARIO =
        usuario.ID_USUARIO;

      sesion.USUARIO =
        usuario.USUARIO;

      sesion.TOKEN_SESION =
        tokenSesion;

      sesion.ESTADO_SESION =
        SEG_SESIONES_CONFIG.ESTADO_ACTIVA;

      sesion.FECHA_INICIO =
        ahora;

      sesion.FECHA_EXPIRACION =
        fechaExpiracion;

      sesion.FECHA_CIERRE =
        "";

      sesion.ULTIMA_ACTIVIDAD =
        ahora;


      // ----------------------------------------------------------
      // CONVERTIR OBJETO EN FILA
      // ----------------------------------------------------------

      const nuevaFila =
        encabezados.map(function(campo) {

          return sesion[campo] !== undefined
            ? sesion[campo]
            : "";

        });


      // ----------------------------------------------------------
      // GUARDAR SESIÓN
      // ----------------------------------------------------------

      hoja.appendRow(nuevaFila);


      // ----------------------------------------------------------
      // DEVOLVER SESIÓN
      // ----------------------------------------------------------

      return {

        EXITO: true,

        ID_SESION:
          idSesion,

        TOKEN_SESION:
          tokenSesion,

        ID_USUARIO:
          usuario.ID_USUARIO,

        FECHA_EXPIRACION:
          fechaExpiracion

      };

    }


    // ============================================================
    // 07.07 BUSCAR SESIÓN POR TOKEN
    // ============================================================

    function SEG_BUSCAR_SESION(tokenSesion) {

      if (
        !tokenSesion ||
        String(tokenSesion).trim() === ""
      ) {

        return null;

      }


      const hoja =
        SEG_OBTENER_HOJA_SESIONES();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_SESIONES();

      const ultimaFila =
        hoja.getLastRow();

      const ultimaColumna =
        hoja.getLastColumn();

      if (ultimaFila < 2) {

        return null;

      }


      const idxToken =
        encabezados.indexOf("TOKEN_SESION");

      if (idxToken === -1) {

        throw new Error(
          "No se encontró la columna TOKEN_SESION."
        );

      }


      const registros =
        hoja
          .getRange(
            2,
            1,
            ultimaFila - 1,
            ultimaColumna
          )
          .getValues();


      // ----------------------------------------------------------
      // BUSCAR TOKEN
      // ----------------------------------------------------------

      for (
        let i = 0;
        i < registros.length;
        i++
      ) {

        if (
          String(registros[i][idxToken]) ===
          String(tokenSesion)
        ) {

          const sesion = {};

          encabezados.forEach(
            function(campo, indice) {

              sesion[campo] =
                registros[i][indice];

            }
          );


          // Guardamos internamente la fila física.
          sesion._FILA =
            i + 2;

          return sesion;

        }

      }

      return null;

    }


    // ============================================================
    // 07.08 VALIDAR SESIÓN
    // Verifica existencia, estado y fecha de expiración.
    // ============================================================

    function SEG_VALIDAR_SESION(tokenSesion) {

      const sesion =
        SEG_BUSCAR_SESION(tokenSesion);

      if (!sesion) {

        return {

          VALIDA: false,

          CODIGO:
            "SESION_NO_ENCONTRADA",

          MENSAJE:
            "La sesión no existe o no es válida."

        };

      }


      // ----------------------------------------------------------
      // VALIDAR ESTADO
      // ----------------------------------------------------------

      const estado =
        String(sesion.ESTADO_SESION || "")
          .trim()
          .toUpperCase();

      if (
        estado !==
        SEG_SESIONES_CONFIG.ESTADO_ACTIVA
      ) {

        return {

          VALIDA: false,

          CODIGO:
            "SESION_NO_ACTIVA",

          MENSAJE:
            "La sesión no se encuentra activa."

        };

      }


      // ----------------------------------------------------------
      // VALIDAR EXPIRACIÓN
      // ----------------------------------------------------------

      const ahora =
        new Date();

      const fechaExpiracion =
        new Date(
          sesion.FECHA_EXPIRACION
        );

      if (
        ahora.getTime() >=
        fechaExpiracion.getTime()
      ) {

        // Cambiar automáticamente a EXPIRADA
        SEG_CAMBIAR_ESTADO_SESION(
          tokenSesion,
          SEG_SESIONES_CONFIG.ESTADO_EXPIRADA
        );

        return {

          VALIDA: false,

          CODIGO:
            "SESION_EXPIRADA",

          MENSAJE:
            "La sesión ha expirado."

        };

      }


      // ----------------------------------------------------------
      // ACTUALIZAR ÚLTIMA ACTIVIDAD
      // ----------------------------------------------------------

      SEG_ACTUALIZAR_ACTIVIDAD_SESION(
        tokenSesion
      );


      return {

        VALIDA: true,

        CODIGO:
          "SESION_VALIDA",

        MENSAJE:
          "Sesión válida.",

        SESION:
          sesion

      };

    }


    // ============================================================
    // 07.09 ACTUALIZAR ACTIVIDAD DE SESIÓN
    // ============================================================

    function SEG_ACTUALIZAR_ACTIVIDAD_SESION(
      tokenSesion
    ) {

      const sesion =
        SEG_BUSCAR_SESION(tokenSesion);

      if (!sesion) {

        return false;

      }


      const hoja =
        SEG_OBTENER_HOJA_SESIONES();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_SESIONES();

      const idxActividad =
        encabezados.indexOf(
          "ULTIMA_ACTIVIDAD"
        );

      if (idxActividad === -1) {

        return false;

      }

      hoja
        .getRange(
          sesion._FILA,
          idxActividad + 1
        )
        .setValue(
          new Date()
        );

      return true;

    }


    // ============================================================
    // 07.10 CAMBIAR ESTADO DE SESIÓN
    // ============================================================

    function SEG_CAMBIAR_ESTADO_SESION(
      tokenSesion,
      nuevoEstado
    ) {

      const sesion =
        SEG_BUSCAR_SESION(tokenSesion);

      if (!sesion) {

        return false;

      }


      const estadosPermitidos = [
        SEG_SESIONES_CONFIG.ESTADO_ACTIVA,
        SEG_SESIONES_CONFIG.ESTADO_CERRADA,
        SEG_SESIONES_CONFIG.ESTADO_EXPIRADA
      ];

      const estado =
        String(nuevoEstado || "")
          .trim()
          .toUpperCase();

      if (
        !estadosPermitidos.includes(estado)
      ) {

        throw new Error(
          "Estado de sesión no válido."
        );

      }


      const hoja =
        SEG_OBTENER_HOJA_SESIONES();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_SESIONES();

      const idxEstado =
        encabezados.indexOf(
          "ESTADO_SESION"
        );

      if (idxEstado === -1) {

        throw new Error(
          "No existe la columna ESTADO_SESION."
        );

      }


      hoja
        .getRange(
          sesion._FILA,
          idxEstado + 1
        )
        .setValue(estado);


      return true;

    }


    // ============================================================
    // 07.11 CERRAR SESIÓN
    // Cierra manualmente una sesión activa.
    // ============================================================

    function SEG_CERRAR_SESION(tokenSesion) {

      const sesion =
        SEG_BUSCAR_SESION(tokenSesion);

      if (!sesion) {

        return {

          EXITO: false,

          MENSAJE:
            "No se encontró la sesión."

        };

      }


      // ----------------------------------------------------------
      // CAMBIAR ESTADO
      // ----------------------------------------------------------

      SEG_CAMBIAR_ESTADO_SESION(
        tokenSesion,
        SEG_SESIONES_CONFIG.ESTADO_CERRADA
      );


      // ----------------------------------------------------------
      // REGISTRAR FECHA DE CIERRE
      // ----------------------------------------------------------

      const hoja =
        SEG_OBTENER_HOJA_SESIONES();

      const encabezados =
        SEG_OBTENER_ENCABEZADOS_SESIONES();

      const idxFechaCierre =
        encabezados.indexOf(
          "FECHA_CIERRE"
        );

      if (idxFechaCierre !== -1) {

        hoja
          .getRange(
            sesion._FILA,
            idxFechaCierre + 1
          )
          .setValue(
            new Date()
          );

      }


      return {

        EXITO: true,

        MENSAJE:
          "Sesión cerrada correctamente."

      };

    }

  // ============================================================
  // 08. VALIDACIÓN DE ACCESO
  // Integra sesión, usuario, rol y permisos para autorizar
  // o denegar acciones dentro del ERP.
  // ============================================================


    // ============================================================
    // 08.01 OBTENER USUARIO DESDE SESIÓN
    // Valida el token de sesión y obtiene el usuario asociado.
    // ============================================================

    function SEG_OBTENER_USUARIO_SESION(tokenSesion) {

      // ----------------------------------------------------------
      // VALIDAR SESIÓN
      // ----------------------------------------------------------

      const validacion =
        SEG_VALIDAR_SESION(tokenSesion);

      if (!validacion.VALIDA) {

        return null;

      }


      // ----------------------------------------------------------
      // OBTENER ID DEL USUARIO
      // ----------------------------------------------------------

      const idUsuario =
        validacion.SESION.ID_USUARIO;

      if (
        !idUsuario ||
        String(idUsuario).trim() === ""
      ) {

        return null;

      }


      // ----------------------------------------------------------
      // CONSULTAR USUARIO
      // ----------------------------------------------------------

      return SEG_CONSULTAR_USUARIO(
        idUsuario
      );

    }


    // ============================================================
    // 08.02 VALIDAR USUARIO DESDE SESIÓN
    // Verifica que exista y permanezca activo.
    // ============================================================

    function SEG_VALIDAR_USUARIO_SESION(tokenSesion) {

      const usuario =
        SEG_OBTENER_USUARIO_SESION(
          tokenSesion
        );

      if (!usuario) {

        return {

          VALIDO: false,

          CODIGO:
            "USUARIO_SESION_NO_ENCONTRADO",

          MENSAJE:
            "No se encontró un usuario válido para la sesión."

        };

      }


      // ----------------------------------------------------------
      // VALIDAR ESTADO DEL USUARIO
      // ----------------------------------------------------------

      const estado =
        String(usuario.ESTADO_USUARIO || "")
          .trim()
          .toUpperCase();

      if (
        estado !==
        SEG_CONFIG.ESTADO_USUARIO_ACTIVO
      ) {

        return {

          VALIDO: false,

          CODIGO:
            "USUARIO_NO_ACTIVO",

          MENSAJE:
            "El usuario no se encuentra activo."

        };

      }


      return {

        VALIDO: true,

        CODIGO:
          "USUARIO_VALIDO",

        MENSAJE:
          "Usuario válido.",

        USUARIO:
          usuario

      };

    }


    // ============================================================
    // 08.03 VALIDAR ROL DESDE SESIÓN
    // Obtiene el rol del usuario y verifica que esté activo.
    // ============================================================

    function SEG_VALIDAR_ROL_SESION(tokenSesion) {

      // ----------------------------------------------------------
      // VALIDAR USUARIO
      // ----------------------------------------------------------

      const validacionUsuario =
        SEG_VALIDAR_USUARIO_SESION(
          tokenSesion
        );

      if (!validacionUsuario.VALIDO) {

        return {

          VALIDO: false,

          CODIGO:
            validacionUsuario.CODIGO,

          MENSAJE:
            validacionUsuario.MENSAJE

        };

      }


      const usuario =
        validacionUsuario.USUARIO;


      // ----------------------------------------------------------
      // VALIDAR ID DEL ROL
      // ----------------------------------------------------------

      if (
        !usuario.ID_ROL ||
        String(usuario.ID_ROL).trim() === ""
      ) {

        return {

          VALIDO: false,

          CODIGO:
            "USUARIO_SIN_ROL",

          MENSAJE:
            "El usuario no tiene un rol asignado."

        };

      }


      // ----------------------------------------------------------
      // CONSULTAR ROL
      // ----------------------------------------------------------

      const rol =
        SEG_CONSULTAR_ROL(
          usuario.ID_ROL
        );

      if (!rol) {

        return {

          VALIDO: false,

          CODIGO:
            "ROL_NO_ENCONTRADO",

          MENSAJE:
            "No se encontró el rol asignado al usuario."

        };

      }


      // ----------------------------------------------------------
      // VALIDAR ESTADO DEL ROL
      // ----------------------------------------------------------

      const estadoRol =
        String(rol.ESTADO_ROL || "")
          .trim()
          .toUpperCase();

      if (estadoRol !== "ACTIVO") {

        return {

          VALIDO: false,

          CODIGO:
            "ROL_NO_ACTIVO",

          MENSAJE:
            "El rol asignado al usuario no se encuentra activo."

        };

      }


      return {

        VALIDO: true,

        CODIGO:
          "ROL_VALIDO",

        MENSAJE:
          "Rol válido.",

        USUARIO:
          usuario,

        ROL:
          rol

      };

    }


    // ============================================================
    // 08.04 VALIDAR ACCESO COMPLETO
    // Valida:
    //
    // 1. Token de sesión.
    // 2. Estado de la sesión.
    // 3. Usuario.
    // 4. Estado del usuario.
    // 5. Rol.
    // 6. Estado del rol.
    // 7. Permiso del módulo.
    // 8. Acción solicitada.
    // ============================================================

    function SEG_VALIDAR_ACCESO(
      tokenSesion,
      modulo,
      accion
    ) {

      // ----------------------------------------------------------
      // VALIDAR DATOS DE ACCESO
      // ----------------------------------------------------------

      if (
        !tokenSesion ||
        String(tokenSesion).trim() === ""
      ) {

        return {

          AUTORIZADO: false,

          CODIGO:
            "TOKEN_NO_PROPORCIONADO",

          MENSAJE:
            "Debe proporcionar un token de sesión."

        };

      }


      if (
        !modulo ||
        String(modulo).trim() === ""
      ) {

        return {

          AUTORIZADO: false,

          CODIGO:
            "MODULO_NO_PROPORCIONADO",

          MENSAJE:
            "Debe indicar el módulo."

        };

      }


      if (
        !accion ||
        String(accion).trim() === ""
      ) {

        return {

          AUTORIZADO: false,

          CODIGO:
            "ACCION_NO_PROPORCIONADA",

          MENSAJE:
            "Debe indicar la acción."

        };

      }


      // ----------------------------------------------------------
      // VALIDAR SESIÓN
      // ----------------------------------------------------------

      const validacionSesion =
        SEG_VALIDAR_SESION(tokenSesion);

      if (!validacionSesion.VALIDA) {

        return {

          AUTORIZADO: false,

          CODIGO:
            validacionSesion.CODIGO,

          MENSAJE:
            validacionSesion.MENSAJE

        };

      }


      // ----------------------------------------------------------
      // VALIDAR USUARIO Y ROL
      // ----------------------------------------------------------

      const validacionRol =
        SEG_VALIDAR_ROL_SESION(
          tokenSesion
        );

      if (!validacionRol.VALIDO) {

        return {

          AUTORIZADO: false,

          CODIGO:
            validacionRol.CODIGO,

          MENSAJE:
            validacionRol.MENSAJE

        };

      }


      // ----------------------------------------------------------
      // OBTENER DATOS
      // ----------------------------------------------------------

      const usuario =
        validacionRol.USUARIO;

      const rol =
        validacionRol.ROL;


      // ----------------------------------------------------------
      // VALIDAR PERMISO
      // ----------------------------------------------------------

      const autorizado =
        SEG_VALIDAR_PERMISO_ROL(
          rol.ID_ROL,
          modulo,
          accion
        );

      if (!autorizado) {

        return {

          AUTORIZADO: false,

          CODIGO:
            "PERMISO_DENEGADO",

          MENSAJE:
            "No tiene permiso para ejecutar esta acción.",

          ID_USUARIO:
            usuario.ID_USUARIO,

          ID_ROL:
            rol.ID_ROL,

          MODULO:
            SEG_NORMALIZAR_PERMISO(modulo),

          ACCION:
            SEG_NORMALIZAR_PERMISO(accion)

        };

      }


      // ----------------------------------------------------------
      // ACCESO AUTORIZADO
      // ----------------------------------------------------------

      return {

        AUTORIZADO: true,

        CODIGO:
          "ACCESO_AUTORIZADO",

        MENSAJE:
          "Acceso autorizado.",

        ID_USUARIO:
          usuario.ID_USUARIO,

        USUARIO:
          usuario.USUARIO,

        ID_ROL:
          rol.ID_ROL,

        ROL:
          rol.NOMBRE_ROL,

        MODULO:
          SEG_NORMALIZAR_PERMISO(modulo),

        ACCION:
          SEG_NORMALIZAR_PERMISO(accion),

        SESION:
          validacionSesion.SESION

      };

    }


    // ============================================================
    // 08.05 VERIFICAR ACCESO OBLIGATORIO
    // Valida el acceso y detiene el proceso mediante un error
    // cuando el usuario no está autorizado.
    // Esta función debe utilizarse antes de ejecutar procesos
    // críticos del ERP.
    // ============================================================

    function SEG_VERIFICAR_ACCESO_SESION(
      tokenSesion,
      modulo,
      accion
    ) {

      const resultado =
        SEG_VALIDAR_ACCESO(
          tokenSesion,
          modulo,
          accion
        );

      if (!resultado.AUTORIZADO) {

        throw new Error(
          "ACCESO DENEGADO [" +
          resultado.CODIGO +
          "]: " +
          resultado.MENSAJE
        );

      }

      return resultado;

    }


    // ============================================================
    // 08.06 OBTENER CONTEXTO DE SEGURIDAD
    // Devuelve toda la información de seguridad asociada
    // a una sesión válida.
    //
    // Incluye:
    // - Sesión.
    // - Usuario.
    // - Rol.
    // - Estado de validación.
    // ============================================================

    function SEG_OBTENER_CONTEXTO_SEGURIDAD(
      tokenSesion
    ) {

      // ----------------------------------------------------------
      // VALIDAR SESIÓN
      // ----------------------------------------------------------

      const validacionSesion =
        SEG_VALIDAR_SESION(tokenSesion);

      if (!validacionSesion.VALIDA) {

        return {

          VALIDO: false,

          CODIGO:
            validacionSesion.CODIGO,

          MENSAJE:
            validacionSesion.MENSAJE

        };

      }


      // ----------------------------------------------------------
      // VALIDAR ROL
      // ----------------------------------------------------------

      const validacionRol =
        SEG_VALIDAR_ROL_SESION(
          tokenSesion
        );

      if (!validacionRol.VALIDO) {

        return {

          VALIDO: false,

          CODIGO:
            validacionRol.CODIGO,

          MENSAJE:
            validacionRol.MENSAJE

        };

      }


      // ----------------------------------------------------------
      // DEVOLVER CONTEXTO COMPLETO
      // ----------------------------------------------------------

      return {

        VALIDO: true,

        CODIGO:
          "CONTEXTO_VALIDO",

        MENSAJE:
          "Contexto de seguridad válido.",

        SESION:
          validacionSesion.SESION,

        USUARIO:
          validacionRol.USUARIO,

        ROL:
          validacionRol.ROL

      };

    }


    // ============================================================
    // 08.07 VALIDAR ACCESO A VARIAS ACCIONES
    // Permite consultar múltiples permisos en una sola llamada.
    //
    // Ejemplo:
    // [
    //   { MODULO: "CLIENTES", ACCION: "CONSULTAR" },
    //   { MODULO: "CLIENTES", ACCION: "CREAR" }
    // ]
    // ============================================================

    function SEG_VALIDAR_MULTIPLES_ACCESOS(
      tokenSesion,
      accesos
    ) {

      if (!Array.isArray(accesos)) {

        throw new Error(
          "La lista de accesos debe ser un arreglo."
        );

      }


      const resultados = [];

      accesos.forEach(function(acceso) {

        const modulo =
          acceso.MODULO ||
          acceso.modulo;

        const accion =
          acceso.ACCION ||
          acceso.accion;

        const resultado =
          SEG_VALIDAR_ACCESO(
            tokenSesion,
            modulo,
            accion
          );

        resultados.push({
          MODULO: modulo,
          ACCION: accion,
          AUTORIZADO:
            resultado.AUTORIZADO,
          CODIGO:
            resultado.CODIGO,
          MENSAJE:
            resultado.MENSAJE
        });

      });


      return resultados;

    }