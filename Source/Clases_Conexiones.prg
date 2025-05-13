/*
 * Proyecto: Descargate
 * Fichero: Clases_Conexiones.prg
 * Descripción:
 * Autor:
 * Fecha: 28/11/2023
 */

#include "Xailer.ch"
#include "Digio.ch"

//------------------------------------------------------------------------------

CLASS tConexion FROM tComponent

   PROPERTY oOrigen           INIT nil
   PROPERTY oAlertas          INIT nil

   PROPERTY cQuery            INIT ""
   PROPERTY cAccion           INIT ""

   PROPERTY aDatos            INIT {}

   PROPERTY lConectado        INIT .f.

   METHOD Create()
   METHOD Conecta()
   METHOD Ejecuta()

ENDCLASS

//------------------------------------------------------------------------------

METHOD Create CLASS tConexion

   WITH OBJECT ::oalertas := tAlerta():New()
      :Create()
   END WITH

   ::Super:Create()

RETURN self

//--------------------------------------------------------------------------

METHOD Conecta() CLASS tConexion

   LOCAL hpaso:= {=>}, aerrores:= {}, herror:= {}
   LOCAL carma:= ""

   ::oOrigen:lConnected:=.t.

   aerrores:= ::oOrigen:GetErrors()

   ::oOrigen:DelErrors()

   if Len(aerrores) > 0

      *============================================================================================================================================================================
      * MENSAJES PARA GRABAR EN ARCHIVO LOG
      *============================================================================================================================================================================

      ::oAlertas:cPrefijo:= "!! "

      ::oAlertas:aLogs:={}

      AAdd(::oAlertas:aLogs , "----------------------------------------------------------------------")
      AAdd(::oAlertas:aLogs , "ERROR en acceso a conexion a servidor " + ::oOrigen:cHost)
      AAdd(::oAlertas:aLogs , "  Accion: " + ::cAccion)

      for each herror in aerrores

         AAdd(::oAlertas:aLogs , herror[01] + " - " + herror[02])
         AAdd(::oAlertas:aLogs , tostring(herror[03]) + " - " + herror[04])

      next

      AAdd(::oAlertas:aLogs , "fin error conexion a servidor " + ::oOrigen:cHost)
      AAdd(::oAlertas:aLogs , "----------------------------------------------------------------------")
      AAdd(::oAlertas:aLogs , "")

      ::oAlertas:ejecutaAlerta()

      Application:Terminate()
      ::End()
      quit

   endif

   ::lConectado := .t.

RETURN NIL

//--------------------------------------------------------------------------

METHOD Ejecuta(leserror) CLASS tConexion

   LOCAL aerrores:= {}, herror:= {=>}, carma:= ""

   ::oOrigen:execute(::cQuery)

   aerrores:= ::oOrigen:GetErrors()

   ::oOrigen:DelErrors()

   if Len(aerrores) > 0

      *=============================================================================================================================================================================
      * MENSAJES PARA GRABAR EN ARCHIVO LOG
      *=============================================================================================================================================================================

      ::oAlertas:cPrefijo:= "// "

      ::oAlertas:aLogs:={}

      AAdd(::oAlertas:aLogs , "----------------------------------------------------------------------")
      AAdd(::oAlertas:aLogs , "ERROR en acceso a conexion a servidor " + ::oOrigen:cHost)
      AAdd(::oAlertas:aLogs , "  Accion: " + ::cAccion)

      for each herror in aerrores

         AAdd(::oAlertas:aLogs , herror[01] + " - " + herror[02])
         AAdd(::oAlertas:aLogs , tostring(herror[03]) + " - " + herror[04])

      next

      AAdd(::oAlertas:aLogs , "fin error conexion a servidor " + ::oOrigen:cHost)
      AAdd(::oAlertas:aLogs , "----------------------------------------------------------------------")
      AAdd(::oAlertas:aLogs , "")

      ::oAlertas:lAlertaLocal :=.f.
      ::oAlertas:lAlertaNube  :=.f.

      if appdata:oConexion_local:lConectado

         carma:= "insert into descargate.alertas (mensaje) "                             ;
                +"  values ('Error en base de datos accion: " + ::cAccion + "') "

         ::oAlertas:lAlertaLocal :=.t.
         ::oAlertas:cQueryLocal  := carma
         ::oAlertas:cAccionLocal := "grabando referencia en local de intento de conexion fallida a servidor central (nube)"

      endif

      ::oAlertas:ejecutaAlerta(leserror)

      Application:Terminate()
      ::End()
      quit

   endif

RETURN NIL

//--------------------------------------------------------------------------
