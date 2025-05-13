/*
 * Proyecto: Api-Dygy
 * Fichero: Principal.prg
 * Descripción:
 * Autor:
 * Fecha: 17/01/2023
 */

#include "Xailer.ch"

CLASS principal FROM TForm

   METHOD CreateForm()
   METHOD FormInitialize( oSender )
   METHOD FormShow( oSender )

ENDCLASS

#include "Principal.xfm"

//------------------------------------------------------------------------------

METHOD FormInitialize( oSender ) CLASS principal

   *===============================================================================================================================================================================
   * INICIAMOS UNA NUEVA INSTANCIA DE ARCHIVO LOG
   *===============================================================================================================================================================================

   with object AppData:oFileLog := tDgsLogs():new
      :Create()
   END

   AppData:oFileLog:CreaArchivoLog()

   *===============================================================================================================================================================================
   * INICIAMOS UNA NUEVA INSTANCIA DEL API, ESTA INSTANCIA DESPUES DE CUMPLIR LA ORDEN CON LA QUE LO INICIARON SE CERRARA.
   *===============================================================================================================================================================================

   AppData:oFileLog:GrabaMensajesEnLog({"=======================================================================================================================================" , ;
                                        "Entre a API-DYGY.EXE"})

   *===============================================================================================================================================================================
   * INSTANCIAMOS UN OBJECTO CONEXION PARA EL SERVIDOR LOCAL
   *===============================================================================================================================================================================

   WITH OBJECT appdata:oConexion_nube := tConexion():New( Self )
      :Create()
   END WITH

   *===============================================================================================================================================================================
   * APERTURAMOS LA BASE DE DATOS LOCAL
   *===============================================================================================================================================================================

   WITH OBJECT appdata:oConexion_nube:oOrigen := TMariaDBDataSource():New( Self )
      :cHost                 := "167.114.103.204"
      :cUser                 := "root"
      :nPort                 := 3308
      :cPassword             := "PasaleMaria9!"
      :lDisplayErrors        := .f.
      :lAbortOnErrors        := .f.
      :Create()
   END WITH

   *===============================================================================================================================================================================
   * INTENTAMOS ESTABLECER LA CONEXION, EN CASO DE ERROR EL METODO CONECTA REVISARA QUE TODO ESTE BIEN, EN CASO CONTRARIO ABORTARA LA APLICACION
   *===============================================================================================================================================================================

   appdata:oConexion_nube:cAccion           := "intento de conexion a BD Local"
   appdata:oConexion_nube:conecta()








   *===============================================================================================================================================================================
   * EN CASO DE REQUERIR PROBAR LA GENERACION DE UN ARCHIVO, INTERCEPTAMOS LA VARIABLE APPDATA:CFILE_JSON PARA SIMULAR LA PETICION DE RADAR
   *===============================================================================================================================================================================

   if appdata:cfile_json = nil .or. len(appdata:cfile_json) = 0

      appdata:cfile_json:= "c:\apache24\htdocs\dygy\proceso\con-car-cli-10612202300190203.json"

   endif

RETURN Nil

//------------------------------------------------------------------------------

METHOD FormShow( oSender ) CLASS principal

   LOCAL hpaso:= {=>}, hestado:= {=>}, cdatos:= "", ckey := "", oError
   LOCAL neste:= 0, alogs:= {}

   AppData:oFileLog:GrabaUnMensajeEnLog("en onShow de API-DYGY:principal appdata:cfile_json="+appdata:cfile_json)

   *===============================================================================================================================================================================
   * ANTES DE INICIAR VALIDAMOS QUE EXISTA EL ARCHIVO JSON, DE NO ENCONTRARLO GRABAREMOS EL MENSAJE EN ARCHIVO LOG Y NO SERA POSIBLE ENVIAR MENSAJE A USUARIO, NO HABRA MANERA DE
   * COMUNICAR A USUARIO.
   *===============================================================================================================================================================================

   if !file( appdata:cfile_json )

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "NO se encontro archivo json " + appdata:cfile_json
      hestado["cTituloTextoUsuario"]       := "El sistema ha detectado un problema con su peticion"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] := HB_JsonEncode( hestado )

      AppData:oFileLog:GrabaMensajesEnLog({"------------------------------------------------------------------------------------------------" , ;
                                           "ocurrio un error en API-DYGY al procesar peticion de cgi "+appdata:cfile_json                     , ;
                                           hestado["cDescripcion"], hestado["cTituloTextoUsuario"], hestado["cTextoUsuario"]                  , ;
                                           "------------------------------------------------------------------------------------------------"})

      *============================================================================================================================================================================
      * NOTA: INTENTAMOS ENVIAR MENSAJE A USUARIO PARTIENDO SOLO DEL NOMBRE DEL ARCHIVO JSON, NO TENEMOS MAS DATOS
      *============================================================================================================================================================================

      hpaso["carchivo_error"] := strtran( appdata:cfile_json, ".json", ".err")

      hb_memowrit( hpaso["carchivo_error"], hpaso["cgraba"] )

      Application:Terminate()
      return(nil)

   endif

   *===============================================================================================================================================================================
   * LEEMOS LOS PARAMETROS DE LA PETICION DESDE EL ARCHIVO JSON, SU UBICACION Y NOMBRE LA TENEMOS EN APPDATA:CFILE_JSON
   *===============================================================================================================================================================================

   cdatos := hb_memoread( appdata:cfile_json )

   *===============================================================================================================================================================================
   * VALIDAMOS QUE EL CONTENIDO DEL ARCHIVO JSON NO ESTE VACIO.
   *===============================================================================================================================================================================

   if Len(cdatos) = 0

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "El archivo JSON viene VACIO " + appdata:cfile_json
      hestado["cTituloTextoUsuario"]       := "El sistema ha detectado un problema con su peticion"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] := HB_JsonEncode( hestado )

      AppData:oFileLog:GrabaMensajesEnLog({"------------------------------------------------------------------------------------------------" , ;
                                           "ocurrio un error en API-DYGY al procesar peticion de cgi "+appdata:cfile_json                     , ;
                                           hestado["cDescripcion"], hestado["cTituloTextoUsuario"], hestado["cTextoUsuario"]                  , ;
                                           "------------------------------------------------------------------------------------------------"})

      *============================================================================================================================================================================
      * NOTA: INTENTAMOS ENVIAR MENSAJE A USUARIO PARTIENDO SOLO DEL NOMBRE DEL ARCHIVO JSON, NO TENEMOS MAS DATOS
      *============================================================================================================================================================================

      hpaso["carchivo_error"] := strtran( appdata:cfile_json, ".json", ".err")

      hb_memowrit( hpaso["carchivo_error"], hpaso["cgraba"] )

      Application:Terminate()
      return(nil)

   endif

   *===============================================================================================================================================================================
   * DECODIFICAMOS LOS DATOS EN UNA VARIBLE HASH QUE UBICAMOS EN APPDATA:HADATOS, CON ESTO ASEGURAMOS QUE EN CUALQUIER PARTE DE LA INSTANCIA TENDREMOS A LA MANO ESTOS DATOS
   *===============================================================================================================================================================================

   HB_JsonDecode( cdatos , @appdata:hreporte)

   *===============================================================================================================================================================================
   * VALIDACION DE LOS CAMPOS MINIMOS NECESARIOS, VALIDAMOS LA PARTE COMUN DEL HASH HREPORTE HPASO["LBASICOS"] NO DIRA SI ALGUN PARAMETRO FALTO
   *===============================================================================================================================================================================

   hpaso["lbasicos"] := .t.

   hpaso["hreporte_paso"] :={=>}

   inicializa_estructura_reportes( @hpaso["hreporte_paso"])

   hpaso["alista_keys"] := hb_HKeys( hpaso["hreporte_paso"])

   hpaso["cgraba"]:= ""

   for each ckey in hpaso["alista_keys"]

      if !hb_HHasKey( appdata:hreporte, ckey )

         inicializa_estructura_estado_peticion_api( @hestado )

         hestado["cEstado"]                   := "Cancelado"
         hestado["cDescripcion"]              := "Falta parametro " + ckey + " en archivo recibido con parametros"
         hestado["cTituloTextoUsuario"]       := "El sistema ha detectado un problema con su peticion"
         hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

         hpaso["cgraba"] += HB_JsonEncode( hestado )

         *=========================================================================================================================================================================
         * INTENTAMOS PROTEGER EL ENVIO SI NO EXISTE LA LLAVE "carchivo_error" FORMANDO EL NOMBRE DESDE EL ARCVHIVO JSON QUE RECIBIMOS
         *=========================================================================================================================================================================

         if !hb_HHasKey( appdata:hreporte, "carchivo_error" )

            hpaso["carchivo_error"] := strtran( appdata:cfile_json, ".json", ".err")

            hb_memowrit( hpaso["carchivo_error"], hpaso["cgraba"] )

            else

            hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

         endif

         hpaso["lbasicos"] := .f.

         AppData:oFileLog:GrabaMensajesEnLog({"------------------------------------------------------------------------------------------------" , ;
                                              "ocurrio un error en API-DYGY en revision de parametros BASICOS de "+appdata:cfile_json            , ;
                                              hestado["cDescripcion"]                                                                            , ;
                                              "------------------------------------------------------------------------------------------------"})

      endif

   next

   if !hpaso["lbasicos"]
      ::End()
      Application:Terminate()
      quit
   endif




























   *===============================================================================================================================================================================
   * AGREGAMOS UNA VARIABLE A APPDATA:LDATENDIDA PARA DETECTAR SI LA PETICION TIENE UNA FUNCION DE RESPUESTA, EN CASO DE NO SER ASI ENVIAREMOS UN MENSAJE DE ERROR A LA
   * APLICACION CGI
   *===============================================================================================================================================================================

   appdata:hreporte["latendida"] := .f.

   *%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   *% INICIAMOS A RUTEAR DEPENDIENDO DE LA SOLICTUD RECIBIDA
   *%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

   TRY

      if appdata:hreporte["cfuncion"]="consulta_cartera_clientes_excel"
         consulta_cartera_clientes_excel()
      endif

      if appdata:hreporte["cfuncion"]="consulta_cartera_clientes_pdf"
         consulta_cartera_clientes_pdf()
      endif

      if appdata:hreporte["cfuncion"]="consulta_cartera_grupos_pdf"
         consulta_cartera_grupos_pdf()
      endif

      if appdata:hreporte["cfuncion"]="consulta_cartera_grupos_excel"
         consulta_cartera_grupos_excel()
      endif

   CATCH oError

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "Ocurrio un error irrecuperable en llamado a API en funcion " + appdata:hreporte["cfuncion"]
      hestado["cTituloTextoUsuario"]       := "El sistema ha detectado un problema con su peticion"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] += HB_JsonEncode( hestado )

      hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

      alogs:= {"------------------------------------------------------------------------------------------------------------------------" , ;
               "  !! el sistema API-DYGY cayo en un error detectado por catch general  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!"                     , ;
               "  !! al intentar ejecutar funcion " + appdata:hreporte["cfuncion"]                                                        , ;
               "  !! " + oError:subsystem +": "+ oError:Description +" "+ oError:Operation}

      for neste:= 2 to Len(oError:cargo)

         AAdd(alogs,"  !! " + oError:cargo[neste,01] + "(" + ToString(oError:cargo[neste,02]) + ")")

      next neste

      AAdd(alogs,"  !!    ---- --------------------------")
      AAdd(alogs,"  !!    codigo error " + ToString(oError:subcode))
      AAdd(alogs,"  !!    modulo " + ProcName(1))
      AAdd(alogs,"  !!    ---- --------------------------")
      AAdd(alogs,"------------------------------------------------------------------------------------------------------------------------")

      AppData:oFileLog:GrabaMensajesEnLog(alogs)

      Application:Terminate()

      quit

   END

   *===============================================================================================================================================================================
   * VALIDAMOS QUE LA PETICION FUE ATENDIDA POR UNA FUNCION
   *===============================================================================================================================================================================

   IF !appdata:hreporte["latendida"]

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "La Peticion no fue atendida por plataforma API"
      hestado["cTituloTextoUsuario"]       := "El sistema ha detectado un problema con su peticion"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] += HB_JsonEncode( hestado )

      hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

      AppData:oFileLog:GrabaMensajesEnLog({"------------------------------------------------------------------------------------------------" , ;
                                           "## el sistema API-DYGY detecto que la funcion " + appdata:hreporte["cfuncion"]                    , ;
                                           "## no fue atendida por no identificarla en el ruteador del api"                                   , ;
                                           "------------------------------------------------------------------------------------------------"})


   ENDIF

   AppData:oFileLog:GrabaMensajesEnLog({"finalizo aplicacion API-DYGY desde modulo principal"                                                                                       , ;
                                        "======================================================================================================================================="})

   if !appdata:lno_cierres
      Application:Terminate()
   endif

RETURN Nil

//------------------------------------------------------------------------------
