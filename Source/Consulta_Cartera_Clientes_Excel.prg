/*
 * Proyecto: Api-Dygy
 * Fichero: Consulta_Cartera_Clientes_Excel.prg
 * Descripción:
 * Autor:
 * Fecha: 21/01/2023
 */

#include "Xailer.ch"
#include "Digio.ch"

//------------------------------------------------------------------------------

FUNCTION consulta_cartera_clientes_excel()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, hpunto:= {=>}, ohoja, nrenglon, nid_factura

   LogFile(DToC(Date())+" "+Time()+"   ******** en API-DYGY Inicializando funcion consulta_cartera_clientes_excel   "+Str(seconds(),12,4))

   *=================================================================================================================================================
   * INDICAMOS QUE LA PETICION SI ESTA ATENDIDA POR UNA FUNCION Y ASI EVITAMOS ENVIAR EL MENSAJE DE QUE LA PETICION NO TIENE FUNCION
   *=================================================================================================================================================

   appdata:hreporte["latendida"] := .t.

   *=================================================================================================================================================
   * LOS PARAMETROS BASICOS DE HREPORTE YA FUERON VALIDADOS, AHORA VAMOS POR LOS ESPECIFICOS DE LA FUNCION
   *=================================================================================================================================================

   hpaso["cgraba"] := ""

   hpaso["akeys"] := {"cfecha_inicial", "cfecha_final", "ctipo_tablero", "cidcliente", "cnombre_cliente", "cidempresa"}

   if !valida_keys(@hpaso)
      return(nil)
   endif

   *=================================================================================================================================================
   * AGREGAMOS LA VARIABLE LABORTAPROCESO A EL OBJETO HREPORTE PARA SABER CUANDO UN SUBPROCESO GENERO UN ERROR QUE MERECE QUE SE ABORTE LA PETICION
   *=================================================================================================================================================

   appdata:hreporte["laborta_proceso"] := .f.

   *=================================================================================================================================================
   * CONEXIONES A BASES DE DATOS
   *=================================================================================================================================================

   if !conecta_bd()
      return(nil)
   endif

   LogFile(DToC(Date())+" "+Time()+"   ******** en APi_DYGY ya conectado a servidor de bd   "+Str(seconds(),12,4))

   *=================================================================================================================================================
   * LA PETICION QUE RECIBIMOS ABARCA 6 TIPOS DE FORMATOS EXCEL, VAMOS A DEFINIR UNA FUNCION PARA CADA TIPO
   *=================================================================================================================================================

   if appdata:hreporte["ctipo_tablero"]="1" .or. appdata:hreporte["ctipo_tablero"]="4"

      genera_nuevo_excel_cartera_clientes_analiticos_un_cliente()

      return(nil)

   endif

   if appdata:hreporte["ctipo_tablero"]="2" .or. appdata:hreporte["ctipo_tablero"]="5"

      genera_nuevo_excel_cartera_clientes_auxiliar_un_cliente()
      *genera_excel_cartera_clientes_auxiliar_un_cliente()

      return(nil)

   endif

   if appdata:hreporte["ctipo_tablero"]="3" .or. appdata:hreporte["ctipo_tablero"]="6"

      *genera_excel_cartera_clientes_antiguedad_un_cliente()

      genera_nuevo_excel_cartera_clientes_antiguedad_un_cliente()

      return(nil)

   endif

   *=================================================================================================================================================
   * SI LLEGAMOS A ESTA LINEAS ASUMIMOS QUE NO DETECTAMOS EL FORMATO SOLICITADO POR EL CGI, ABORTAMOS Y MANDAMOS MENSAJE
   *=================================================================================================================================================

   inicializa_estructura_estado_peticion_api( @appdata:hestado )

   appdata:hestado["cEstado"]                   := "Proceso"
   appdata:hestado["cDescripcion"]              := "seguimiento normal"
   appdata:hestado["cTituloTextoUsuario"]       := "El sistema no encontro funcion solicitada..."
   appdata:hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

   hpaso["cgraba"] := HB_JsonEncode( appdata:hestado )

   hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

   Application:Terminate()
   quit

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION genera_nuevo_excel_cartera_clientes_antiguedad_un_cliente()

   LOCAL hpaso:= {=>}, hadoerror:= {=>}, hregistro:= {=>}
   LOCAL oExcel := nil, oHoja:= nil

   *=================================================================================================================================================
   * PREPARAMOS LA VARIABLE HASH PARA REPORTAR ERRORES
   *=================================================================================================================================================

   inicializa_estructura_hadoerror( @hadoerror )

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Solicitando informacion a servidor..")

   *=================================================================================================================================================
   * OBTENEMOS LOS CONJUNTOS DE DATOS NECESARIOS
   *=================================================================================================================================================

   hpaso["cidempresa"]      := appdata:hreporte["cidempresa"]
   hpaso["cidgrupo"]        := "0"
   hpaso["cidcliente"]      := appdata:hreporte["cidcliente"]
   hpaso["chasta"]          := "'" + fechaammdd(appdata:hreporte["cfecha_final"], .t.) +"'"
   hpaso["ctipo_respuesta"] := "1"
   hpaso["ctipo_respuesta"] := IIF(appdata:hreporte["ctipo_tablero"]="6", "3", "1")

   hpaso["carma"]:=" call digio.sp_consulta_antiguedad_de_clientes_web("       ;
                         +hpaso["cidempresa"] +", "                            ;
                         +hpaso["cidgrupo"]   +", "                            ;
                         +hpaso["cidcliente"] +", "                            ;
                         +hpaso["chasta"]     +", "                            ;
                         +hpaso["ctipo_respuesta"]+")"

   AppData:oFileLog:GrabaUnMensajeEnLog("en genera_nuevo_excel_cartera_clientes_antiguedad_un_cliente carma="+ hpaso["carma"])

   hpaso["adant"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   AppData:oFileLog:GrabaUnMensajeEnLog("** despues de query")

   hadoerror["cquery"]          := hpaso["carma"]
   hadoerror["ctitulo_mensaje"] := "Oops...   El sistema no puede obtener datos de antiguedad de cliente"

   if error_ado(hadoerror)
      quit
   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Inicializando documento Excel...")

   * ==============================================================================================================================================================================
   * OBJETO PRINCIPAL
   * ==============================================================================================================================================================================

   with object oExcel := tExcelFromXlm():new
      :cNombreDocumento   := appdata:hreporte["cfile"]
      :cDisco             := "c"
      :cRuta              := appdata:hreporte["cruta_entregable"]
      :lBorraXmls         := .f.
      :lRastrea           := .f.

      :Create()

   END

   * ==============================================================================================================================================================================
   * PRIMERA HOJA
   * ==============================================================================================================================================================================

   with object oHoja := tExcelHoja(oExcel):new
      :cNombre             := "Antiguedad"
      :lInmoviliza         := .t.
      :lAutoFiltro         := .t.
      :lPautado            := .t.
      :cColorFondoAlterno1 := "FFFFFFFF"
      :cColorFondoAlterno2 := "FFDFDFDF"

      with object :oEstiloEncabezados := tEstiloExcel():New
         :lNegrita         := .t.
         :lWrap            := .t.
         :cColorFondo      := "FF800000"
         :cColorFuente     := "FFFFFFFF"

         :Create()

      END

      :Create(oExcel)

   END


   AppData:oFileLog:GrabaUnMensajeEnLog("En excel_por_xlm iniciando definicion  de hoja 1")

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE TITULOS Y SUBTITULOS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   *oHoja:AgregaTitulo({"cTexto"          => appdata:hReporte["cnombre_empresa" ,             ;

   oHoja:AgregaTitulo({"cTexto"          => "DIGIOSOFT ASESORES" ,                           ;
                       "lBrincaRenglon"  => .t.,                                             ;
                       "cColorFondo"     => "FF800000" ,                                     ;
                       "cColorFuente"    => "FFFFFFFF" ,                                     ;
                       "cColorFondoAbajo"=> "FFFFFFFF"})

   if appdata:hreporte["ctipo_tablero"] = "3"
      oHoja:AgregaTitulo({"cTexto" => "ANTIGUEDAD DE SALDOS DE " + appdata:hreporte["cnombre_cliente"]  , "lBrincaRenglon" => .t. , "nFuente" => 16})
   endif

   if appdata:hreporte["ctipo_tablero"] = "6"
      oHoja:AgregaTitulo({"cTexto" => "ANTIGUEDAD DE SALDOS (TODOS LOS CLIENTES)"                       , "lBrincaRenglon" => .t. , "nFuente" => 16})
   endif

   oHoja:AgregaTitulo({"cTexto" => "AL " + fectal(appdata:hreporte["cfecha_final"])                     , "lBrincaRenglon" => .t. , "nFuente" => 16, "lNegrita" => .t.})

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE COLUMNAS DE PRIMERA HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oHoja:AgregaColumna({"cTitulo" => "FACTURA"                              , "cNombreHash" => "factura"                , "nAncho" => 14, "cFormato" => "000000" , "cAliHorizontal" => "center" })
   oHoja:AgregaColumna({"cTitulo" => "FECHA"                                , "cNombreHash" => "fecha"                  , "nAncho" => 14, "cFormato" => "dd/mmm/yyyy" })
   oHoja:AgregaColumna({"cTitulo" => "CONCEPTO"                             , "cNombreHash" => "descripcion"            , "nAncho" => 70, "lWrap" => .t.})
   oHoja:AgregaColumna({"cTitulo" => "IMPORTE"                              , "cNombreHash" => "importe"                , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   oHoja:AgregaColumna({"cTitulo" => "PAGOS"                                , "cNombreHash" => "saldo"                  , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   oHoja:AgregaColumna({"cTitulo" => "SALDO"                                , "cNombreHash" => "pagos"                  , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   oHoja:AgregaColumna({"cTitulo" => "NO VENCIDO"                           , "cNombreHash" => "no_vencido"             , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   oHoja:AgregaColumna({"cTitulo" => "VENCIDO 1 - 30 DIAS"                  , "cNombreHash" => "vencido_01_30_dias"     , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   oHoja:AgregaColumna({"cTitulo" => "VENCIDO 31 - 60 DIAS"                 , "cNombreHash" => "vencido_31_60_dias"     , "nAncho" => 20 ,"cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   oHoja:AgregaColumna({"cTitulo" => "VENCIDO 61-90 DIAS "                  , "cNombreHash" => "vencido_61_90_dias"     , "nAncho" => 20 ,"cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   oHoja:AgregaColumna({"cTitulo" => "VENCIDO MAS DE 90 DIAS "              , "cNombreHash" => "vencido_mas_de_90_dias" , "nAncho" => 20 ,"cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })


   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINIMOS ESTILOS ESPECIALES PARA CIERTOS RENGLONES DE LOS DATOS - TOTALES POR CLIENTE
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oHoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF833C0C" , "nFuente" => 14 ,  "lNegrita" => .t.})
   oHoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF833C0C" , "nFuente" => 12 ,  "lNegrita" => .t.})
   oHoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FF800000" , "cColorFuente"    => "FFFFFFFF" , "nFuente" => 14 ,  "lNegrita" => .t.})

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINIMOS ESTILOS LIBRES A UTILIZARSE EN CUALQUIER PARTE DEL LIBRO
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oExcel:AgregaEstiloLibre({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF000000" , "nFuente" => 14 })

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * HAREMOS UN RECORRIDO AGREGANDO A LOS RENGLONES ESPECIALES SU ID DE ESTILOS PERSONALIZADOS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   for each hregistro in hpaso["adant"]

      hregistro["descripcion"]    := AnsiToUTF8(hregistro["descripcion"])
      hregistro["nombre_cliente"] := AnsiToUTF8(hregistro["nombre_cliente"])

      if hregistro["clase"]="04" .and. appdata:hreporte["ctipo_tablero"] = "6"
         hregistro["id_estilo_personalizado"]:= 1
         hregistro["descripcion"]            := hregistro["nombre_cliente"]
         hregistro["factura"]                := ""
         hregistro["fecha"]                  := ""
         hregistro["importe"]                := ""
         hregistro["saldo"]                  := ""
         hregistro["pagos"]                  := ""
         hregistro["no_vencido"]             := ""
         hregistro["vencido_01_30_dias"]     := ""
         hregistro["vencido_31_60_dias"]     := ""
         hregistro["vencido_61_90_dias"]     := ""
         hregistro["vencido_mas_de_90_dias"] := ""
      endif

      if hregistro["clase"]="12" .and. appdata:hreporte["ctipo_tablero"] = "3"
         hregistro["id_estilo_personalizado"]   := 3
         hregistro["descripcion"]               := "TOTALES"
         hregistro["factura"]                   := ""
         hregistro["fecha"]                     := ""
      endif

      if hregistro["clase"]="12" .and. appdata:hreporte["ctipo_tablero"] = "6"
         hregistro["id_estilo_personalizado"]   := 2
         hregistro["descripcion"]               := ""
         hregistro["factura"]                   := ""
         hregistro["fecha"]                     := ""
         hregistro["lbrinca_renglon_en_blanco"] := .t.
         hregistro["id_estilo_libre_brinco"]    := 1
      endif

      if hregistro["clase"]="16"
         hregistro["id_estilo_personalizado"]:= 3
         hregistro["descripcion"]               := "TOTAL GENERAL"
         hregistro["factura"]                   := ""
         hregistro["fecha"]                     := ""
      endif

   next

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE DATOS DE HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oHoja:aDatos := hpaso["adant"]

   AAdd(oExcel:aHojas , oHoja)

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * INICIAMOS PROCESO DE GENERACION DE XMLS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oExcel:GeneraExcel()

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   mensaje_avance("Gestionando Entrega de Archivo Solicitado")

   hpaso["cgraba"] := HB_JsonEncode(appdata:hestado )

   hb_memowrit( appdata:hreporte["carchivo_finalizacion"], hpaso["cgraba"] )

   *borra_temporales()

   AppData:oFileLog:GrabaUnMensajeEnLog("Finalizando y saliendo de funcion genera_nuevo_excel_cartera_clientes_antiguedad_un_cliente")

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION genera_nuevo_excel_cartera_clientes_auxiliar_un_cliente()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, hregistro:= {=>}, oexcel, ohoja, nrenglon

   *=================================================================================================================================================
   * PREPARAMOS LA VARIABLE HASH PARA REPORTAR ERRORES
   *=================================================================================================================================================

   inicializa_estructura_hadoerror( @hadoerror )

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Solicitando informacion a servidor..")

   *=================================================================================================================================================
   * OBTENEMOS LOS CONJUNTOS DE DATOS NECESARIOS
   *=================================================================================================================================================

   hpaso["cidempresa"]      := appdata:hreporte["cidempresa"]
   hpaso["cidgrupo"]        := "0"
   hpaso["cidcliente"]      := appdata:hreporte["cidcliente"]
   hpaso["cdesde"]          := "'" + fechaammdd(appdata:hreporte["cfecha_inicial"], .t.) +"'"
   hpaso["chasta"]          := "'" + fechaammdd(appdata:hreporte["cfecha_final"], .t.) +"'"
   hpaso["ctipo_respuesta"] := IIF(appdata:hreporte["ctipo_tablero"]="5", "11", "1")

   hpaso["carma"]:=" call digio.sp_consulta_cartera_grupos_de_clientes_web("   ;
                         +hpaso["cidempresa"] +", "                            ;
                         +hpaso["cidgrupo"]   +", "                            ;
                         +hpaso["cidcliente"] +", "                            ;
                         +hpaso["cdesde"]     +", "                            ;
                         +hpaso["chasta"]     +", "                            ;
                         +hpaso["ctipo_respuesta"]+")"

   AppData:oFileLog:GrabaUnMensajeEnLog("en genera_nuevo_excel_cartera_clientes_antiguedad_un_cliente carma="+ hpaso["carma"])

   hpaso["adaux"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   AppData:oFileLog:GrabaUnMensajeEnLog("despues de query")

   hadoerror["cquery"]          := hpaso["carma"]
   hadoerror["ctitulo_mensaje"] := "Oops...   El sistema no puede obtener datos de antiguedad de cliente"

   if error_ado(hadoerror)
      quit
   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Inicializando documento Excel...")

   * ==============================================================================================================================================================================
   * OBJETO PRINCIPAL
   * ==============================================================================================================================================================================

   with object oexcel := tExcelFromXlm():new
      :cNombreDocumento   := appdata:hreporte["cfile"]
      :cDisco             := "c"
      :cRuta              := appdata:hreporte["cruta_entregable"]
      :lBorraXmls         := .f.
      :lRastrea           := .f.

      :Create()

   END

   * ==============================================================================================================================================================================
   * PRIMERA HOJA
   * ==============================================================================================================================================================================

   with object ohoja := tExcelHoja(oexcel):new
      :cNombre             := "Auxiliar"
      :lInmoviliza         := .t.
      :lAutoFiltro         := .t.
      :lPautado            := .t.
      :cColorFondoAlterno1 := "FFFFFFFF"
      :cColorFondoAlterno2 := "FFDFDFDF"

      with object :oEstiloEncabezados := tEstiloExcel():New
         :lNegrita         := .t.
         :lWrap            := .t.
         :cColorFondo      := "FF800000"
         :cColorFuente     := "FFFFFFFF"

         :Create()

      END

      :Create(oexcel)

   END


   AppData:oFileLog:GrabaUnMensajeEnLog("En excel_por_xlm iniciando definicion  de hoja 1")

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE TITULOS Y SUBTITULOS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   *oHoja:AgregaTitulo({"cTexto"          => appdata:hReporte["cnombre_empresa" ,             ;

   ohoja:AgregaTitulo({"cTexto"          => "DIGIOSOFT ASESORES" ,                           ;
                       "lBrincaRenglon"  => .t.,                                             ;
                       "cColorFondo"     => "FF800000" ,                                     ;
                       "cColorFuente"    => "FFFFFFFF" ,                                     ;
                       "cColorFondoAbajo"=> "FFFFFFFF"})

   if appdata:hreporte["ctipo_tablero"] = "5"
      ohoja:AgregaTitulo({"cTexto" => "AUXILIAR DE CLIENTES "                                       ,  "lBrincaRenglon" => .t. , "nFuente" => 16})
   endif

   if appdata:hreporte["ctipo_tablero"] = "2"
      ohoja:AgregaTitulo({"cTexto" => "AUXILIAR DE CLIENTE " + appdata:hreporte["cnombre_cliente"]  , "lBrincaRenglon" => .t. , "nFuente" => 16})
   endif

   ohoja:AgregaTitulo({"cTexto" => "DEL: "+fectal(appdata:hreporte["cfecha_inicial"])+"   AL: "+fectal(appdata:hreporte["cfecha_final"])  , "lBrincaRenglon" => .t. , "nFuente" => 16, "lNegrita" => .t.})

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE COLUMNAS DE PRIMERA HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:AgregaColumna({"cTitulo" => "FACTURA"                              , "cNombreHash" => "esfactura"              , "nAncho" => 14, "cFormato" => "000000" , "cAliHorizontal" => "center" })
   ohoja:AgregaColumna({"cTitulo" => "FECHA"                                , "cNombreHash" => "fecha"                  , "nAncho" => 14, "cFormato" => "dd/mmm/yyyy" , "cAliHorizontal" => "center" })
   ohoja:AgregaColumna({"cTitulo" => "CONCEPTO"                             , "cNombreHash" => "descripcion"            , "nAncho" => 90, "lWrap" => .t.})
   ohoja:AgregaColumna({"cTitulo" => "SALDO ANTERIOR"                       , "cNombreHash" => "anterior"               , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "CARGOS"                               , "cNombreHash" => "cargo"                  , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "ABONOS"                               , "cNombreHash" => "abono"                  , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "SALDO ACTUAL"                         , "cNombreHash" => "actual"                 , "nAncho" => 20 ,"cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINIMOS ESTILOS ESPECIALES PARA CIERTOS RENGLONES DE LOS DATOS - TOTALES POR CLIENTE
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF833C0C" , "nFuente" => 14 ,  "lNegrita" => .t.})
   ohoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF833C0C" , "nFuente" => 12 ,  "lNegrita" => .t.})
   ohoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FF800000" , "cColorFuente"    => "FFFFFFFF" , "nFuente" => 14 ,  "lNegrita" => .t.})
   ohoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FFFFE699" , "cColorFuente"    => "FF000000" })



   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINIMOS ESTILOS LIBRES A UTILIZARSE EN CUALQUIER PARTE DEL LIBRO
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oexcel:AgregaEstiloLibre({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF000000" , "nFuente" => 14 })

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * HAREMOS UN RECORRIDO AGREGANDO A LOS RENGLONES ESPECIALES SU ID DE ESTILOS PERSONALIZADOS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   for each hregistro in hpaso["adaux"]

      hregistro["anterior"]:= ""

      hregistro["nombre_cliente"] := AnsiToUTF8(hregistro["nombre_cliente"])
      hregistro["descripcion"]    := AnsiToUTF8(hregistro["descripcion"])

      * --------------------- 10 - TOTAL CLIENTE (QUE PUEDE SER AHORRADO POR NO MOVIMIENTOS) ----------

      if hregistro["clase"]="10"

         hregistro["lomite_renglon"]    := 1

         if hregistro:__enumindex > 1
            hpaso["adaux"][hregistro:__enumindex-1, "anterior"]                  := hregistro["actual"]
            hpaso["adaux"][hregistro:__enumindex-1, "actual"]                    := hregistro["actual"]
            hpaso["adaux"][hregistro:__enumindex-1, "cargo"]                     := 0
            hpaso["adaux"][hregistro:__enumindex-1, "abono"]                     := 0
            hpaso["adaux"][hregistro:__enumindex-1, "lbrinca_renglon_en_blanco"] := .t.
            hpaso["adaux"][hregistro:__enumindex-1, "id_estilo_libre_brinco"]    := 1

         endif

         loop

      endif
      * --------------------- 03 - INICIO DE CLIENTE --------------------------------------------------

      if hregistro["clase"]="03" .and. appdata:hreporte["ctipo_tablero"] = "2"
         hregistro["descripcion"]             := "SALDO ANTERIOR"
         hregistro["fecha"]                     := ""
         hregistro["anterior"]                := hregistro["actual"]
         hregistro["cargo"]                   := ""
         hregistro["abono"]                   := ""
         hregistro["actual"]                  := ""
      endif

      if hregistro["clase"]="03" .and. appdata:hreporte["ctipo_tablero"] = "5"

         hregistro["descripcion"]             := hregistro["cliente"] +"-"+ hregistro["nombre_cliente"]
         hregistro["fecha"]                     := ""
         hregistro["anterior"]                := hregistro["actual"]
         hregistro["cargo"]                   := ""
         hregistro["abono"]                   := ""
         hregistro["actual"]                  := ""
         hregistro["id_estilo_personalizado"] := 1

      endif

      * --------------------- 07 - REGISTRO DE PAGO ---------------------------------------------------

      if hregistro["clase"]="07"
         hregistro["id_estilo_personalizado"]   := 4
      endif


      * --------------------- 09 - TOTAL POR CLIENTE --------------------------------------------------

      if hregistro["clase"]="09" .and. appdata:hreporte["ctipo_tablero"] = "2"
         hregistro["fecha"]                     := ""
         hregistro["descripcion"]               := "TOTALES"
         hregistro["anterior"]                  := hregistro["actual"] + hregistro["abono"] - hregistro["cargo"]
         hregistro["id_estilo_personalizado"]   := 3
      endif

      if hregistro["clase"]="09" .and. appdata:hreporte["ctipo_tablero"] = "5"
         hregistro["fecha"]                     := ""
         hregistro["descripcion"]               := "TOTAL CLIENTE"
         hregistro["anterior"]                  := hregistro["actual"] + hregistro["abono"] - hregistro["cargo"]
         hregistro["id_estilo_personalizado"]   := 1
         hregistro["lbrinca_renglon_en_blanco"] := .t.
         hregistro["id_estilo_libre_brinco"]    := 1
      endif


      * --------------------- 14 - TOTAL GENERAL (TODOS LOS CLIENTES)----------------------------------

      if hregistro["clase"]="14" .and. appdata:hreporte["ctipo_tablero"] = "5"
         hregistro["fecha"]                     := ""
         hregistro["descripcion"]               := "TOTAL CLIENTE"
         hregistro["anterior"]                  := hregistro["actual"] + hregistro["abono"] - hregistro["cargo"]
         hregistro["id_estilo_personalizado"]   := 3
      endif

   next

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE DATOS DE HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:aDatos := hpaso["adaux"]

   AAdd(oexcel:aHojas , ohoja)

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * INICIAMOS PROCESO DE GENERACION DE XMLS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oexcel:GeneraExcel()

   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------

   mensaje_avance("Gestionando Entrega de Archivo Solicitado")

   hpaso["cgraba"] := HB_JsonEncode(appdata:hestado )

   hb_memowrit( appdata:hreporte["carchivo_finalizacion"], hpaso["cgraba"] )

   *borra_temporales()

   AppData:oFileLog:GrabaUnMensajeEnLog("Finalizando y saliendo de funcion genera_nuevo_excel_cartera_clientes_antiguedad_un_cliente")

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION genera_nuevo_excel_cartera_clientes_analiticos_un_cliente()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, hregistro:= {=>}, oExcel, ohoja, nrenglon

   *=================================================================================================================================================
   * PREPARAMOS LA VARIABLE HASH PARA REPORTAR ERRORES
   *=================================================================================================================================================

   inicializa_estructura_hadoerror( @hadoerror )

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Solicitando informacion a servidor..")

   *=================================================================================================================================================
   * OBTENEMOS LOS CONJUNTOS DE DATOS NECESARIOS
   *=================================================================================================================================================

   hpaso["cidempresa"]      := appdata:hreporte["cidempresa"]
   hpaso["cidgrupo"]        := "0"
   hpaso["cidcliente"]      := appdata:hreporte["cidcliente"]
   hpaso["cdesde"]          := "'" + fechaammdd(appdata:hreporte["cfecha_inicial"], .t.) +"'"
   hpaso["chasta"]          := "'" + fechaammdd(appdata:hreporte["cfecha_final"], .t.) +"'"
   hpaso["ctipo_respuesta"] := IIF(appdata:hreporte["ctipo_tablero"]="4", "12", "2")

   hpaso["carma"]:=" call digio.sp_consulta_cartera_grupos_de_clientes_web("   ;
                         +hpaso["cidempresa"] +", "                            ;
                         +hpaso["cidgrupo"]   +", "                            ;
                         +hpaso["cidcliente"] +", "                            ;
                         +hpaso["cdesde"]     +", "                            ;
                         +hpaso["chasta"]     +", "                            ;
                         +hpaso["ctipo_respuesta"]+")"

   AppData:oFileLog:GrabaUnMensajeEnLog("en genera_nuevo_excel_cartera_clientes_analiticos_cliente carma="+ hpaso["carma"])

   hpaso["adana"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   AppData:oFileLog:GrabaUnMensajeEnLog("despues de query")

   hadoerror["cquery"]          := hpaso["carma"]
   hadoerror["ctitulo_mensaje"] := "Oops...   El sistema no puede obtener datos de antiguedad de cliente"

   if error_ado(hadoerror)
      quit
   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Inicializando documento Excel...")

   * ==============================================================================================================================================================================
   * OBJETO PRINCIPAL
   * ==============================================================================================================================================================================

   with object oExcel := tExcelFromXlm():new
      :cNombreDocumento   := appdata:hreporte["cfile"]
      :cDisco             := "c"
      :cRuta              := appdata:hreporte["cruta_entregable"]
      :lBorraXmls         := .f.
      :lRastrea           := .f.

      :Create()

   END

   * ==============================================================================================================================================================================
   * PRIMERA HOJA
   * ==============================================================================================================================================================================

   with object ohoja := tExcelHoja(oExcel):new
      :cNombre             := "Auxiliar"
      :lInmoviliza         := .t.
      :lAutoFiltro         := .t.
      :lPautado            := .t.
      :cColorFondoAlterno1 := "FFFFFFFF"
      :cColorFondoAlterno2 := "FFDFDFDF"

      with object :oEstiloEncabezados := tEstiloExcel():New
         :lNegrita         := .t.
         :lWrap            := .t.
         :cColorFondo      := "FF800000"
         :cColorFuente     := "FFFFFFFF"

         :Create()

      END

      :Create(oExcel)

   END


   AppData:oFileLog:GrabaUnMensajeEnLog("En excel_por_xlm iniciando definicion de documento")

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE TITULOS Y SUBTITULOS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   *oHoja:AgregaTitulo({"cTexto"          => appdata:hReporte["cnombre_empresa" ,             ;

   ohoja:AgregaTitulo({"cTexto"          => "DIGIOSOFT ASESORES" ,                           ;
                       "lBrincaRenglon"  => .t.,                                             ;
                       "cColorFondo"     => "FF800000" ,                                     ;
                       "cColorFuente"    => "FFFFFFFF" ,                                     ;
                       "cColorFondoAbajo"=> "FFFFFFFF"})

   if appdata:hreporte["ctipo_tablero"] = "5"
      ohoja:AgregaTitulo({"cTexto" => "RELACIONES ANALITICAS DE CARTERA"                            ,  "lBrincaRenglon" => .t. , "nFuente" => 16})
   endif

   if appdata:hreporte["ctipo_tablero"] = "2"
      ohoja:AgregaTitulo({"cTexto" => "RELACIONES ANALITICAS DE CARTERA      CLIENTE: " + appdata:hreporte["cnombre_cliente"]  , "lBrincaRenglon" => .t. , "nFuente" => 16})
   endif

   ohoja:AgregaTitulo({"cTexto" => "DEL: "+fectal(appdata:hreporte["cfecha_inicial"])+"   AL: "+fectal(appdata:hreporte["cfecha_final"])  , "lBrincaRenglon" => .t. , "nFuente" => 16, "lNegrita" => .t.})

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE COLUMNAS DE PRIMERA HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:AgregaColumna({"cTitulo" => "CLIENTE"                              , "cNombreHash" => "cliente"                , "nAncho" => 14, "cFormato" => "000000" , "cAliHorizontal" => "center" })
   ohoja:AgregaColumna({"cTitulo" => "NOMBRE"                               , "cNombreHash" => "nombre"                 , "nAncho" => 70, "lWrap" => .t.})
   ohoja:AgregaColumna({"cTitulo" => "SALDO ANTERIOR"                       , "cNombreHash" => "saldo_anterior"         , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "CARGOS"                               , "cNombreHash" => "total_cargos"           , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "ABONOS"                               , "cNombreHash" => "total_abonos"           , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "SALDO ACTUAL"                         , "cNombreHash" => "saldo_actual"           , "nAncho" => 20 ,"cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINIMOS ESTILOS ESPECIALES PARA CIERTOS RENGLONES DE LOS DATOS - TOTALES POR CLIENTE
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FF800000" , "cColorFuente"    => "FFFFFFFF" , "nFuente" => 14 ,  "lNegrita" => .t.})



   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINIMOS ESTILOS LIBRES A UTILIZARSE EN CUALQUIER PARTE DEL LIBRO
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oExcel:AgregaEstiloLibre({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF000000" , "nFuente" => 14 })

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * HAREMOS UN RECORRIDO AGREGANDO A LOS RENGLONES ESPECIALES SU ID DE ESTILOS PERSONALIZADOS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   for each hregistro in hpaso["adana"]

      hregistro["nombre"] := AnsiToUTF8(hregistro["nombre"])

      * --------------------- 14 - TOTAL GENERAL (TODOS LOS CLIENTES)----------------------------------

      if hregistro["clase"]="14" .and. appdata:hreporte["ctipo_tablero"] = "4"
         hregistro["nombre"]                    := "TOTALES"
         hregistro["id_estilo_personalizado"]   := 1
      endif

   next

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE DATOS DE HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:aDatos := hpaso["adana"]

   AAdd(oExcel:aHojas , ohoja)

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * INICIAMOS PROCESO DE GENERACION DE XMLS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oExcel:GeneraExcel()

   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------

   mensaje_avance("Gestionando Entrega de Archivo Solicitado")

   hpaso["cgraba"] := HB_JsonEncode(appdata:hestado )

   hb_memowrit( appdata:hreporte["carchivo_finalizacion"], hpaso["cgraba"] )

   *borra_temporales()

   AppData:oFileLog:GrabaUnMensajeEnLog("Finalizando y saliendo de funcion genera_nuevo_excel_cartera_clientes_analiticos_un_cliente")

RETURN NIL


//------------------------------------------------------------------------------
