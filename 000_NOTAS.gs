/**************************************************************
* 000_NOTAS.gs
* RESPONSABILIDAD:
* - Documentación técnica conceptual y arquitectura del ERP Megudan V2.
* - Registro de flujos lógicos, dependencias e integraciones contables.
**************************************************************/

/*
  MEGUDAN ERP V2 — ARQUITECTURA GENERAL
  =====================================
  1. Propósito: Capa de control administrativo-operativo sobre Google Sheets + Web Apps.
  2. Integraciones: Alegra y Siigo actúan como capa contable/fiscal externa por API.
  3. Tres Libros Independientes:
     - Libro 1: ERP Operativo (Este proyecto)
     - Libro 2: Financiero Gerencial
     - Libro 3: Core Contable e Integración
*/

/*
USUARIO ABRE EL LINK
        ↓
       doGet(e)
        ↓
¿QUÉ RUTA SOLICITA?
        ↓
 ┌───────────────┬────────────────┐
 ↓               ↓                ↓
LOGIN        DASHBOARD          ERROR
 ↓               ↓
F2_LOGIN      Validar TOKEN
                    ↓
              F3_DASHBOARD






*/