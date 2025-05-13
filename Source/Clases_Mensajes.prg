/*
 * Proyecto: Descargate
 * Fichero: Clases_Mensajes.prg
 * Descripción:
 * Autor:
 * Fecha: 28/11/2023
 */

#include "Xailer.ch"
#include "Digio.ch"

//------------------------------------------------------------------------------

CLASS tDgsLogs FROM tComponent

   PROPERTY cFile             INIT dgs_file_log

   PROPERTY lExiste           INIT .f.
   PROPERTY lSeconds          INIT .t.

   PROPERTY nHandle           INIT nil

   METHOD Create()
   METHOD CreaArchivoLog()
   METHOD GrabaUnMensajeEnLog(cTexto)
   METHOD GrabaMensajesEnLog(aLogs)

ENDCLASS

//------------------------------------------------------------------------------

METHOD Create CLASS tDgsLogs

   ::Super:Create()

RETURN self

//--------------------------------------------------------------------------

METHOD CreaArchivoLog() CLASS tDgsLogs

   if !File(::cFile)

      ::nHandle:= FCreate(::cFile)

      LogFile(DToC(Date())+" "+Time()+" !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ENTRO EN CREACION DE ARCHIVO LOG !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")

      if ::nHandle == -1

         LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
         LogFile(DToC(Date())+" "+Time()+" !!!!! error al intentar crear archivo " + ::cfile +" para guardar logs")
         LogFile(DToC(Date())+" "+Time()+" !!!!! error = " + ToString(FError()))
         LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")

         return(nil)

      endif

      IF !FClose(::nHandle)
         LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
         LogFile(DToC(Date())+" "+Time()+" !!!!! error al intentar cerrar archivo log " + ::cfile )
         LogFile(DToC(Date())+" "+Time()+" !!!!! error = " + ToString(FError()))
         LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
      endif

   endif

   ::lExiste := .t.

RETURN NIL

//--------------------------------------------------------------------------

METHOD GrabaUnMensajeEnLog(cTexto) CLASS tDgsLogs

   ::nHandle := FOpen(::cfile, 2)

   if ::nHandle == -1

      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
      LogFile(DToC(Date())+" "+Time()+" !!!!! error al intentar abrir archivo log " + ::cfile )
      LogFile(DToC(Date())+" "+Time()+" !!!!! error = " + ToString(FError()))
      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")

      return(nil)

   endif

   FSeek(::nHandle, 0, 2)

   cTexto := DToC(Date())+" "+Time()+"  "+cTexto

   if ::lSeconds
      cTexto += "  |  " + Str(seconds(),12,4)
   endif

   cTexto += crlf

   if FWrite(::nHandle, cTexto) < Len(cTexto)

      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
      LogFile(DToC(Date())+" "+Time()+" !!!!! error al intentar grabar texto en archivo log " + ::cfile )
      LogFile(DToC(Date())+" "+Time()+" !!!!! error = " + ToString(FError()))
      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")

   endif

   IF !FClose(::nHandle)
      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
      LogFile(DToC(Date())+" "+Time()+" !!!!! error al intentar cerrar archivo log " + ::cfile )
      LogFile(DToC(Date())+" "+Time()+" !!!!! error = " + ToString(FError()))
      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
   endif



RETURN NIL

//--------------------------------------------------------------------------

METHOD GrabaMensajesEnLog(aLogs) CLASS tDgsLogs

   LOCAL cTexto

   ::nHandle := FOpen(::cfile, 2)

   if ::nHandle == -1

      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
      LogFile(DToC(Date())+" "+Time()+" !!!!! error al intentar abrir archivo log " + ::cfile )
      LogFile(DToC(Date())+" "+Time()+" !!!!! error = " + ToString(FError()))
      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")

      return(nil)

   endif

   FSeek(::nHandle, 0, 2)

   for each cTexto in aLogs

      cTexto := DToC(Date())+" "+Time()+"  "+cTexto

      if ::lSeconds
         cTexto += "  |  " + Str(seconds(),12,4)
      endif

      cTexto += crlf

      if FWrite(::nHandle, cTexto) < Len(cTexto)

         LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
         LogFile(DToC(Date())+" "+Time()+" !!!!! error al intentar grabar texto en archivo log " + ::cfile )
         LogFile(DToC(Date())+" "+Time()+" !!!!! error = " + ToString(FError()))
         LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")

      endif

   next

   IF !FClose(::nHandle)
      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
      LogFile(DToC(Date())+" "+Time()+" !!!!! error al intentar cerrar archivo log " + ::cfile )
      LogFile(DToC(Date())+" "+Time()+" !!!!! error = " + ToString(FError()))
      LogFile(DToC(Date())+" "+Time()+" --------------------------------------------------------------------------------------------------------")
   endif

RETURN NIL

//------------------------------------------------------------------------------

CLASS tAlerta FROM tComponent

   PROPERTY aLogs             INIT nil
   PROPERTY cPrefijo          INIT "---- "

   PROPERTY lAlertaNube       INIT .f.
   PROPERTY cQueryNube        INIT ""
   PROPERTY cAccionNube       INIT ""

   PROPERTY oFileLog          INIT AppData:oFileLog

   METHOD Create()
   METHOD EjecutaAlerta()

ENDCLASS

//------------------------------------------------------------------------------

METHOD Create CLASS tAlerta

   LOCAL cHtml
   LOCAL hitem:={=>}

   ::Super:Create()

RETURN self

//--------------------------------------------------------------------------

METHOD EjecutaAlerta(leserror) CLASS tAlerta

   LOCAL cTexto := ""

   if leserror = nil
      leserror:=.f.
   endif

   if Len(::aLogs) > 0

      AppData:oFileLog:GrabaMensajesEnLog(::aLogs)

   endif

   if ::lAlertaNube .and. appdata:oConexion_nube:lConectado .and. !leserror

      appdata:oConexion_nube:cquery  := ::cQueryNube
      appdata:oConexion_nube:cAccion := ::cAccionNube
      appdata:oConexion_nube:ejecuta(.t.)

   endif

RETURN NIL

//--------------------------------------------------------------------------
