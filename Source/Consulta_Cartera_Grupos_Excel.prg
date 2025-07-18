/*
 * Proyecto: Api-Dygy
 * Fichero: Consulta_Cartera_Grupos_Excel.prg
 * Descripci�n:
 * Autor:
 * Fecha: 23/01/2023
 */

#include "Xailer.ch"
#include "Digio.ch"

//------------------------------------------------------------------------------

FUNCTION consulta_cartera_grupos_excel()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, hpunto:= {=>}, ohoja, nrenglon, nid_factura

   LogFile(DToC(Date())+" "+Time()+"   ******** en API-DYGY Inicializando funcion consulta_cartera_grupos_excel   "+Str(seconds(),12,4))

   *=================================================================================================================================================
   * INDICAMOS QUE LA PETICION SI ESTA ATENDIDA POR UNA FUNCION Y ASI EVITAMOS ENVIAR EL MENSAJE DE QUE LA PETICION NO TIENE FUNCION
   *=================================================================================================================================================

   appdata:hreporte["latendida"] := .t.

   *=================================================================================================================================================
   * LOS PARAMETROS BASICOS DE HREPORTE YA FUERON VALIDADOS, AHORA VAMOS POR LOS ESPECIFICOS DE LA FUNCION
   *=================================================================================================================================================

   hpaso["cgraba"] := ""

   hpaso["akeys"] := {"cfecha_inicial", "cfecha_final", "ctipo_tablero", "cidgrupo", "cnombre_grupo", "cidempresa"}

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

   if appdata:hreporte["ctipo_tablero"]="1"

      genera_excel_cartera_grupos_analiticos_un_grupo()

      return(nil)

   endif

   if appdata:hreporte["ctipo_tablero"]="2"

      genera_excel_cartera_grupos_auxiliar_un_grupo()

      return(nil)

   endif

   if appdata:hreporte["ctipo_tablero"]="3"

      genera_excel_cartera_grupos_antiguedad_un_grupo()

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

FUNCTION genera_excel_cartera_grupos_antiguedad_un_grupo()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, oExcel, hregistro:= {=>}, ohoja, nrenglon

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
   hpaso["cidgrupo"]        := appdata:hreporte["cidgrupo"]
   hpaso["cidcliente"]      := "0"
   hpaso["chasta"]          := "'" + fechaammdd(appdata:hreporte["cfecha_final"], .t.) +"'"
   hpaso["ctipo_respuesta"] := "2"

   hpaso["carma"]:=" call digio.sp_consulta_antiguedad_de_clientes_web("       ;
                         +hpaso["cidempresa"] +", "                            ;
                         +hpaso["cidgrupo"]   +", "                            ;
                         +hpaso["cidcliente"] +", "                            ;
                         +hpaso["chasta"]     +", "                            ;
                         +hpaso["ctipo_respuesta"]+")"

   AppData:oFileLog:GrabaUnMensajeEnLog("en genera_excel_cartera_grupos_antiguedad_un_grupo carma="+ hpaso["carma"])

   hpaso["adant"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   AppData:oFileLog:GrabaUnMensajeEnLog("despues de query="+ hpaso["carma"])

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
      :lBorraXmls         := .t.
      :lRastrea           := .f.

      :Create()

   END

   * ==============================================================================================================================================================================
   * PRIMERA HOJA
   * ==============================================================================================================================================================================

   with object ohoja := tExcelHoja(oExcel):new
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

   ohoja:AgregaTitulo({"cTexto"          => "DIGIOSOFT ASESORES" ,                           ;
                       "lBrincaRenglon"  => .t.,                                             ;
                       "cColorFondo"     => "FF800000" ,                                     ;
                       "cColorFuente"    => "FFFFFFFF" ,                                     ;
                       "cColorFondoAbajo"=> "FFFFFFFF"})

   ohoja:AgregaTitulo({"cTexto" => "ANTIGUEDAD DE SALDOS DE GRUPO: " + appdata:hreporte["cnombre_grupo"] , "lBrincaRenglon" => .t. , "nFuente" => 16})

   ohoja:AgregaTitulo({"cTexto" => "AL " + fectal(appdata:hreporte["cfecha_final"])                       , "lBrincaRenglon" => .t. , "nFuente" => 16, "lNegrita" => .t.})

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE COLUMNAS DE PRIMERA HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:AgregaColumna({"cTitulo" => "FACTURA"                              , "cNombreHash" => "factura"                , "nAncho" => 14, "cFormato" => "000000" , "cAliHorizontal" => "center" })
   ohoja:AgregaColumna({"cTitulo" => "FECHA"                                , "cNombreHash" => "fecha"                  , "nAncho" => 14, "cFormato" => "dd/mmm/yyyy" })
   ohoja:AgregaColumna({"cTitulo" => "CONCEPTO"                             , "cNombreHash" => "descripcion"            , "nAncho" => 70, "lWrap" => .t.})
   ohoja:AgregaColumna({"cTitulo" => "IMPORTE"                              , "cNombreHash" => "importe"                , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "PAGOS"                                , "cNombreHash" => "saldo"                  , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "SALDO"                                , "cNombreHash" => "pagos"                  , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "NO VENCIDO"                           , "cNombreHash" => "no_vencido"             , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "VENCIDO 1 - 30 DIAS"                  , "cNombreHash" => "vencido_01_30_dias"     , "nAncho" => 20, "cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "VENCIDO 31 - 60 DIAS"                 , "cNombreHash" => "vencido_31_60_dias"     , "nAncho" => 20 ,"cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "VENCIDO 61-90 DIAS "                  , "cNombreHash" => "vencido_61_90_dias"     , "nAncho" => 20 ,"cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })
   ohoja:AgregaColumna({"cTitulo" => "VENCIDO MAS DE 90 DIAS "              , "cNombreHash" => "vencido_mas_de_90_dias" , "nAncho" => 20 ,"cFormato" => "###,###,##0.00" , "cAliHorizontal" => "right" })

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINIMOS ESTILOS ESPECIALES PARA CIERTOS RENGLONES DE LOS DATOS - TOTALES POR CLIENTE
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF833C0C" , "nFuente" => 14 ,  "lNegrita" => .t.})
   ohoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF833C0C" , "nFuente" => 12 ,  "lNegrita" => .t.})
   ohoja:AgregaEstiloRenglonPorColumnas({"cColorFondo"     => "FF800000" , "cColorFuente"    => "FFFFFFFF" , "nFuente" => 14 ,  "lNegrita" => .t.})

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINIMOS ESTILOS LIBRES A UTILIZARSE EN CUALQUIER PARTE DEL LIBRO
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   oExcel:AgregaEstiloLibre({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF000000" , "nFuente" => 14 })

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * HAREMOS UN RECORRIDO AGREGANDO A LOS RENGLONES ESPECIALES SU ID DE ESTILOS PERSONALIZADOS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   for each hregistro in hpaso["adant"]

      hregistro["nombre_cliente"] := AnsiToUTF8(hregistro["nombre_cliente"])
      hregistro["descripcion"]    := AnsiToUTF8(hregistro["descripcion"])

      * --------------------- 02 - ENCABEZADO POR GRUPO -----------------------------------------------

      if hregistro["clase"]="02"
         hregistro["lomite_renglon"]  := .t.
         loop
      endif

      * --------------------- 04 - ENCABEZADO POR CLIENTE ---------------------------------------------

      if hregistro["clase"]="04"
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

      * --------------------- 12 - TOTAL POR CLIENTE --------------------------------------------------

      if hregistro["clase"]="12"
         hregistro["id_estilo_personalizado"]   := 2
         hregistro["descripcion"]               := ""
         hregistro["factura"]                   := ""
         hregistro["fecha"]                     := ""
         hregistro["lbrinca_renglon_en_blanco"] := .t.
         hregistro["id_estilo_libre_brinco"]    := 1
      endif

      * --------------------- 14 - TOTAL POR GRUPO ----------------------------------------------------

      if hregistro["clase"]="14"
         hregistro["id_estilo_personalizado"]   := 3
         hregistro["descripcion"]               := "TOTALES"
         hregistro["factura"]                   := ""
         hregistro["fecha"]                     := ""
      endif

      * --------------------- 16 - TOTAL GENERAL ------------------------------------------------------

      if hregistro["clase"]="16"
         hregistro["id_estilo_personalizado"]:= 3
         hregistro["descripcion"]               := "TOTAL GENERAL"
         hregistro["factura"]                   := ""
         hregistro["fecha"]                     := ""
      endif

   next
/*
   -- ====================================================================================================================
   -- CODIGOS DE CAMPO CLASE
   -- ====================================================================================================================
   --   02 - ENCABEZADO DE GRUPO
   --   04 - ENCABEZADO DE CLIENTE
   --   08 - FACTURA
   --   12 - TOTAL POR CLIENTE
   --   14 - TOTAL GRUPO
   --   16 - TOTAL GENERAL
*/

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE DATOS DE HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:aDatos := hpaso["adant"]

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

   AppData:oFileLog:GrabaUnMensajeEnLog("Finalizando y saliendo de funcion genera_excel_cartera_grupos_antiguedad_un_grupo")

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION genera_excel_cartera_grupos_auxiliar_un_grupo()

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
   hpaso["cidgrupo"]        := appdata:hreporte["cidgrupo"]
   hpaso["cidcliente"]      := "0"
   hpaso["cdesde"]          := "'" + fechaammdd(appdata:hreporte["cfecha_inicial"], .t.) +"'"
   hpaso["chasta"]          := "'" + fechaammdd(appdata:hreporte["cfecha_final"], .t.) +"'"
   hpaso["ctipo_respuesta"] := "4"

   hpaso["carma"]:=" call digio.sp_consulta_cartera_grupos_de_clientes_web("   ;
                         +hpaso["cidempresa"] +", "                            ;
                         +hpaso["cidgrupo"]   +", "                            ;
                         +hpaso["cidcliente"] +", "                            ;
                         +hpaso["cdesde"]     +", "                            ;
                         +hpaso["chasta"]     +", "                            ;
                         +hpaso["ctipo_respuesta"]+")"


   AppData:oFileLog:GrabaUnMensajeEnLog("en genera_excel_cartera_grupos_auxiliar_un_grupo   carma="+ hpaso["carma"])

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

   with object oExcel := tExcelFromXlm():new
      :cNombreDocumento   := appdata:hreporte["cfile"]
      :cDisco             := "c"
      :cRuta              := appdata:hreporte["cruta_entregable"]
      :lBorraXmls         := .t.
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

   ohoja:AgregaTitulo({"cTexto" => "AUXILIAR DE GRUPO: " + appdata:hreporte["cnombre_grupo"]        ,  "lBrincaRenglon" => .t. , "nFuente" => 16})

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

   oExcel:AgregaEstiloLibre({"cColorFondo"     => "FFFFFFFF" , "cColorFuente"    => "FF000000" , "nFuente" => 14 })

   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * HAREMOS UN RECORRIDO AGREGANDO A LOS RENGLONES ESPECIALES SU ID DE ESTILOS PERSONALIZADOS
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   for each hregistro in hpaso["adaux"]

      hregistro["anterior"]:= ""

      hregistro["nombre_cliente"] := AnsiToUTF8(hregistro["nombre_cliente"])
      hregistro["descripcion"]    := AnsiToUTF8(hregistro["descripcion"])

      * --------------------- 01 - INICIO DE GRUPO (NO SE NECESITA) -----------------------------------

      if hregistro["clase"]="01"

         hregistro["lomite_renglon"]    := 1
         loop

      endif

      * --------------------- 12 - TOTAL GRUPO (NO SE NECESITA) ---------------------------------------

      if hregistro["clase"]="12"

         hregistro["lomite_renglon"]    := 1
         loop

      endif

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

      if hregistro["clase"]="03"

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

      if hregistro["clase"]="09"
         hregistro["fecha"]                     := ""
         hregistro["descripcion"]               := "TOTAL CLIENTE"
         hregistro["anterior"]                  := hregistro["actual"] + hregistro["abono"] - hregistro["cargo"]
         hregistro["id_estilo_personalizado"]   := 1
         hregistro["lbrinca_renglon_en_blanco"] := .t.
         hregistro["id_estilo_libre_brinco"]    := 1
      endif


      * --------------------- 14 - TOTAL GENERAL (TODOS LOS CLIENTES)----------------------------------

      if hregistro["clase"]="14"
         hregistro["fecha"]                     := ""
         hregistro["descripcion"]               := "TOTAL GENERAL"
         hregistro["anterior"]                  := hregistro["actual"] + hregistro["abono"] - hregistro["cargo"]
         hregistro["id_estilo_personalizado"]   := 3
      endif

   next
/*
      -- ====================================================================================================================
      -- CAMPO CLASE
      -- ====================================================================================================================
      -- 01 - ENCABEZADO DE GRUPO
      -- 03 - ENCABEZADO DE CLIENTE (SALDO INICIAL)
      -- 05 - FACTURAS
      -- 07 - PAGOS
      -- 08 - APLICACION DE NOTA DE CREDITO
      -- 09 - TOTAL CLIENTE
      -- 10 - TOTAL CLIENTE INNECESARIO (SUGERIDO RENGLON EN BLANCO)
      -- 12 - TOTAL GRUPO
      -- 14 - TOTAL GENERAL
      -- ====================================================================================================================
*/
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE DATOS DE HOJA
   * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

   ohoja:aDatos := hpaso["adaux"]

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

   AppData:oFileLog:GrabaUnMensajeEnLog("Finalizando y saliendo de funcion genera_excel_cartera_grupos_auxiliar_un_grupo")

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION genera_excel_cartera_grupos_analiticos_un_grupo()

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
   hpaso["cidgrupo"]        := appdata:hreporte["cidgrupo"]
   hpaso["cidcliente"]      := "0"
   hpaso["cdesde"]          := "'" + fechaammdd(appdata:hreporte["cfecha_inicial"], .t.) +"'"
   hpaso["chasta"]          := "'" + fechaammdd(appdata:hreporte["cfecha_final"], .t.) +"'"
   hpaso["ctipo_respuesta"] := "5"

   hpaso["carma"]:=" call digio.sp_consulta_cartera_grupos_de_clientes_web("   ;
                         +hpaso["cidempresa"] +", "                            ;
                         +hpaso["cidgrupo"]   +", "                            ;
                         +hpaso["cidcliente"] +", "                            ;
                         +hpaso["cdesde"]     +", "                            ;
                         +hpaso["chasta"]     +", "                            ;
                         +hpaso["ctipo_respuesta"]+")"

   AppData:oFileLog:GrabaUnMensajeEnLog("en genera_excel_cartera_grupos_analiticos_un_grupo    carma="+ hpaso["carma"])

   hpaso["adana"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   AppData:oFileLog:GrabaUnMensajeEnLog("despues de query")

   hadoerror["cquery"]          := hpaso["carma"]
   hadoerror["ctitulo_mensaje"] := "Oops...   El sistema no puede obtener datos de analiticos de grupo"

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

   ohoja:AgregaTitulo({"cTexto" => "SALDOS ANALITICOS DE GRUPO: " + appdata:hreporte["cnombre_grupo"]   ,  "lBrincaRenglon" => .t. , "nFuente" => 16})

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

      * --------------------- 12 - TOTAL GRUPO (NO SE NECESITA) ---------------------------------------

      if hregistro["clase"]="12"
         hregistro["lomite_renglon"]            := .t.
      endif

      * --------------------- 14 - TOTAL GENERAL (TODOS LOS CLIENTES)----------------------------------

      if hregistro["clase"]="14"
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



/*
   LogFile(DToC(Date())+" "+Time()+"   ******** en genera_excel_cartera_grupos_analiticos_un_grupo carma="+ hpaso["carma"]+"     "+Str(seconds(),12,4))

   hpaso["adana"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   LogFile(DToC(Date())+" "+Time()+"   ******** despues de query      "+Str(seconds(),12,4))

   hadoerror["cquery"]          := hpaso["carma"]
   hadoerror["ctitulo_mensaje"] := "Oops...   El sistema no puede obtener datos de antiguedad de cliente"

   if error_ado(hadoerror)
      quit
   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Inicializando documento Excel...")

   *=================================================================================================================================================
   * PREPARAMOS LAS VARIABLES QUE UTILIZAREMOS PARA LA GENERACION DEL ARCHIVO
   *=================================================================================================================================================

   appdata:hexcel        := {=>}

   *=================================================================================================================================================
   * ARMAMOS LA PESTAÑA 1
   *=================================================================================================================================================

   * PESTA1
   *     1        2         3          4        5       6
   *|         |        | SALDO    |        |        | SALDO  |
   *| CLIENTE | NOMBRE | ANTERIOR | CARGOS | ABONOS | ACTUAL |

   appdata:hexcel["hpesta1"] := {=>}

   *hpaso["cfecha_inicial"]:= CToD(appdata:hreporte["cfecha_inicial"])
   hpaso["cfecha_inicial"] := fectac(appdata:hreporte["cfecha_inicial"])
   hpaso["cfecha_final"]   := fectac(appdata:hreporte["cfecha_final"])

   appdata:hexcel["hpesta1"]["objeto"]      := nil
   appdata:hexcel["hpesta1"]["nombre"]      := "general"
   appdata:hexcel["hpesta1"]["acolus"]      := {}
   appdata:hexcel["hpesta1"]["asubtitulos"] := {}
   appdata:hexcel["hpesta1"]["cdatos"]      := ""
   appdata:hexcel["hpesta1"]["nrenglones"]  := 0
   appdata:hexcel["hpesta1"]["lfiltro"]     := .t.

   AAdd(appdata:hexcel["hpesta1"]["acolus"], {"titulo" => "CLIENTE"                          , "ancho" => 14.0  , "alineacion" => "cen"       , "formato" => ""                , "wrap" => .f., "nfuente" => 12})
   AAdd(appdata:hexcel["hpesta1"]["acolus"], {"titulo" => "NOMBRE"                           , "ancho" => 80.0  , "alineacion" => "izq"       , "formato" => ""                , "wrap" => .t., "nfuente" => 12})
   AAdd(appdata:hexcel["hpesta1"]["acolus"], {"titulo" => "SALDO ANTERIOR"                   , "ancho" => 18.0  , "alineacion" => "der"       , "formato" => " ###,###,##0.00" , "wrap" => .f., "nfuente" => 12})
   AAdd(appdata:hexcel["hpesta1"]["acolus"], {"titulo" => "CARGOS"                           , "ancho" => 18.0  , "alineacion" => "der"       , "formato" => " ###,###,##0.00" , "wrap" => .f., "nfuente" => 12})
   AAdd(appdata:hexcel["hpesta1"]["acolus"], {"titulo" => "ABONOS"                           , "ancho" => 18.0  , "alineacion" => "der"       , "formato" => " ###,###,##0.00" , "wrap" => .f., "nfuente" => 12})
   AAdd(appdata:hexcel["hpesta1"]["acolus"], {"titulo" => "SALDO ACTUAL"                     , "ancho" => 18.0  , "alineacion" => "der"       , "formato" => " ###,###,##0.00" , "wrap" => .f., "nfuente" => 12})

   appdata:hexcel["hpesta1"]["titulo"]:= appdata:hreporte["cnombre_empresa"]

   AAdd(appdata:hexcel["hpesta1"]["asubtitulos"], {"nrenglon" => 0 , "ctexto" => "SALDOS ANALITICOS DE GRUPO " + appdata:hreporte["cnombre_grupo"]})
   AAdd(appdata:hexcel["hpesta1"]["asubtitulos"], {"nrenglon" => 0 , "ctexto" => "DEL: "+fectal(appdata:hreporte["cfecha_inicial"])+"   AL: "+fectal(appdata:hreporte["cfecha_final"])})

   LogFile(DToC(Date())+" "+Time()+"   ******** armando pestaña 1 paso 1   "+Str(seconds(),12,4))

   *=================================================================================================================================================
   * ARMAMOS TITULOS Y ENCABEZADOS DE PESTAÑA 1
   *=================================================================================================================================================

   appdata:hexcel["hpesta1"]["cdatos"]          += appdata:hexcel["hpesta1"]["titulo"]+crlf+crlf
   appdata:hexcel["hpesta1"]["nrenglones"]      += 1
   appdata:hexcel["hpesta1"]["nrenglon_titulo"] := appdata:hexcel["hpesta1"]["nrenglones"]
   appdata:hexcel["hpesta1"]["nrenglones"]      += 1

   for each hrenglon in appdata:hexcel["hpesta1"]["asubtitulos"]

      appdata:hexcel["hpesta1"]["cdatos"]          += hrenglon["ctexto"]+crlf
      appdata:hexcel["hpesta1"]["nrenglones"]      += 1

      hrenglon["nrenglon"] := appdata:hexcel["hpesta1"]["nrenglones"]

   next

   appdata:hexcel["hpesta1"]["cdatos"]          += crlf
   appdata:hexcel["hpesta1"]["nrenglones"]      += 1

   for each hrenglon in appdata:hexcel["hpesta1"]["acolus"]

      appdata:hexcel["hpesta1"]["cdatos"]          += hrenglon["titulo"]+ctab

   next

   appdata:hexcel["hpesta1"]["cdatos"]          += crlf

   appdata:hexcel["hpesta1"]["nrenglones"]          += 1
   appdata:hexcel["hpesta1"]["nrenglon_encabezado"]:= appdata:hexcel["hpesta1"]["nrenglones"]

   LogFile(DToC(Date())+" "+Time()+"   ******** armando pestaña 1  paso 2   "+Str(seconds(),12,4))

   *=================================================================================================================================================
   * MENSAJE DE AVANCE PARA USUARIO
   *=================================================================================================================================================

   mensaje_avance("Armando seccion Principal...")

   *=================================================================================================================================================
   * EN CASO DE SER NECESARIOS CREAMOS VARIABLES PARA GUARDAR LOS RENGLONES CON CARACTERISTICAS ESPECIALES PARA FORMATEOS ESPECIALES, Y HACEMOS EL
   * RECORRIDO EN EL CONJUNTO DE DATOS PARA ARMAR EL STRING QUE PASAREMOS A EXCEL
   *=================================================================================================================================================

   for each hrenglon in hpaso["adana"]

      if hrenglon["clase"]="01" .or. hrenglon["clase"]="12"
         loop
      endif

      hpaso["ccliente"]                    := iif(hrenglon["cliente"]=nil              , "", hrenglon["cliente"])
      hpaso["cnombre"]                     := iif(hrenglon["nombre"]=nil               , "", hrenglon["nombre"])
      hpaso["canterior"]                   := LTrim(Str(hrenglon["saldo_anterior"]     , 14,2))
      hpaso["ccargo"]                      := LTrim(Str(hrenglon["total_cargos"]       , 14,2))
      hpaso["cabono"]                      := LTrim(Str(hrenglon["total_abonos"]       , 14,2))
      hpaso["csaldo"]                      := LTrim(Str(hrenglon["saldo_actual"]       , 14,2))

      *-------------------------------------------------------------------------
      * TOTALES GENERALES
      *-------------------------------------------------------------------------

      if hrenglon["clase"]="14"

         hpaso["cnombre"]                  := "TOTAL GENERAL"

      endif

      * PESTA1
      *     1        2        3          4        5       6
      *|         |       |          |        |        |       |
      *| FACTURA | FECHA | CONCEPTO | CARGOS | ABONOS | SALDO |

      *-------------------------------------------------------------------------
      * DATOS NORMALES
      *-------------------------------------------------------------------------

      appdata:hexcel["hpesta1"]["cdatos"] += hpaso["ccliente"] + ctab           ;
                                     + hpaso["cnombre"] + ctab                  ;
                                     + hpaso["canterior"] + ctab                ;
                                     + hpaso["ccargo"] + ctab                   ;
                                     + hpaso["cabono"] + ctab                   ;
                                     + hpaso["csaldo"] + crlf

      appdata:hexcel["hpesta1"]["nrenglones"] += 1

   next

   appdata:hexcel["hpesta1"]["nrenglon_tope"] := appdata:hexcel["hpesta1"]["nrenglones"]

   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------

   * mensaje_avance("Inicializando Seccion de fechas de captura validas...")

   *----------------------------------------------------------------------------
   * ARMAMOS LA PESTAÑA 2
   *----------------------------------------------------------------------------

   * PESTA2
   *    1         2
   *|       |           |
   *| FECHA | PRODUCTOS |
   /*
   appdata:hexcel["hpesta2"] := {=>}

   appdata:hexcel["hpesta2"]["objeto"]      := nil
   appdata:hexcel["hpesta2"]["nombre"]      := "fechas"
   appdata:hexcel["hpesta2"]["acolus"]      := {}
   appdata:hexcel["hpesta2"]["asubtitulos"] := {}
   appdata:hexcel["hpesta2"]["cdatos"]      := ""
   appdata:hexcel["hpesta2"]["nrenglones"]  := 0
   appdata:hexcel["hpesta2"]["lfiltro"]     := .t.

   AAdd(appdata:hexcel["hpesta2"]["acolus"], {"titulo" => "FECHA"                                 , "ancho" => 20  , "alineacion" => "cen"       , "formato" => "dd-mmm-aaaa"     , "wrap" => .f.})
   AAdd(appdata:hexcel["hpesta2"]["acolus"], {"titulo" => "PRODUCTOS"                             , "ancho" => 30  , "alineacion" => "der"       , "formato" => "###,###,##0"     , "wrap" => .f.})

   appdata:hexcel["hpesta2"]["titulo"]:= appdata:hreporte["cnombre_empresa"]

   AAdd(appdata:hexcel["hpesta2"]["asubtitulos"], {"nrenglon" => 0 , "ctexto" => "PUNTO DE VENTA: " + hpunto["descripcion"]})
   AAdd(appdata:hexcel["hpesta2"]["asubtitulos"], {"nrenglon" => 0 , "ctexto" => "FECHAS CON CAPTURA DE INVENTARIO FISICO"})
   AAdd(appdata:hexcel["hpesta2"]["asubtitulos"], {"nrenglon" => 0 , "ctexto" => "DEL: "+fectal(appdata:hreporte["cfecha_inicial"])+" AL: "+fectal(appdata:hreporte["cfecha_final"])})

   LogFile(DToC(Date())+" "+Time()+"   ******** armando pestaña 1  paso 1   "+Str(seconds(),12,4))

   *----------------------------------------------------------------------------
   * ARMAMOS TITULOS Y ENCABEZADOS DE PESTAÑA 2
   *----------------------------------------------------------------------------

   appdata:hexcel["hpesta2"]["cdatos"]          += appdata:hexcel["hpesta2"]["titulo"]+crlf+crlf
   appdata:hexcel["hpesta2"]["nrenglones"]      += 1
   appdata:hexcel["hpesta2"]["nrenglon_titulo"] := appdata:hexcel["hpesta2"]["nrenglones"]
   appdata:hexcel["hpesta2"]["nrenglones"]      += 1

   for each hrenglon in appdata:hexcel["hpesta2"]["asubtitulos"]

      appdata:hexcel["hpesta2"]["cdatos"]          += hrenglon["ctexto"]+crlf
      appdata:hexcel["hpesta2"]["nrenglones"]      += 1

      hrenglon["nrenglon"] := appdata:hexcel["hpesta2"]["nrenglones"]

   next

   appdata:hexcel["hpesta2"]["cdatos"]          += crlf
   appdata:hexcel["hpesta2"]["nrenglones"]      += 1

   for each hrenglon in appdata:hexcel["hpesta2"]["acolus"]

      appdata:hexcel["hpesta2"]["cdatos"]          += hrenglon["titulo"]+ctab

   next

   appdata:hexcel["hpesta2"]["cdatos"]          += crlf

   appdata:hexcel["hpesta2"]["nrenglones"]          += 1
   appdata:hexcel["hpesta2"]["nrenglon_encabezado"]:= appdata:hexcel["hpesta2"]["nrenglones"]

   LogFile(DToC(Date())+" "+Time()+"   ******** armando pestaña 2  paso 2   "+Str(seconds(),12,4))

   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------

   mensaje_avance("Solicitando Informacion a Servidor...")

   *----------------------------------------------------------------------------
   * ARMAMOS CONJUNTO DE DATOS DE PESTAÑA 2
   *----------------------------------------------------------------------------

   hpaso["carma"]:=" select fecha, count(idproducto) as productos "                 ;
                  +" from     restaurantes_local.inventario_fisico "                ;
                  +" where    idpunto_venta = "+appdata:hreporte["cidpunto"]        ;
                  +" group by fecha "                                               ;
                  +" order by fecha desc"

   LogFile(DToC(Date())+" "+Time()+"   "+hpaso["carma"])

   hpaso["adfec"] := hpunto["origen"]:QueryArrayHash(hpaso["carma"])

   hadoerror["cquery"]          := hpaso["carma"]

   if error_ado(hadoerror)
      return(nil)
   endif

   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------

   mensaje_avance("Armando Seccion de fechas con captura validas...")

   *----------------------------------------------------------------------------

   for each hrenglon in hpaso["adfec"]

      hpaso["cfecha"]       := fectac(hrenglon["fecha"])
      hpaso["productos"]    := LTrim(Str(hrenglon["productos"]          , 14,0))

      appdata:hexcel["hpesta2"]["cdatos"] += hpaso["cfecha"] + ctab                    ;
                                   + hpaso["productos"] + crlf

      appdata:hexcel["hpesta2"]["nrenglones"] += 1

   next

   appdata:hexcel["hpesta2"]["cdatos"] += crlf

   appdata:hexcel["hpesta2"]["nrenglones"] += 1

   appdata:hexcel["hpesta2"]["nrenglon_tope"] := appdata:hexcel["hpesta2"]["nrenglones"]

   * PESTA2
   *    1         2
   *|       |           |
   *| FECHA | PRODUCTOS |

   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------

   mensaje_avance("Creando Documento Excel...")

   *----------------------------------------------------------------------------
   * CREAMOS EL OBJETO EXCEL QUE CONTENDRA LAS PESTAÑAS
   *----------------------------------------------------------------------------

   appdata:oexcel := TOleAuto():New( "Excel.Application" )

   appdata:oexcel:WorkBooks:Add()

   *----------------------------------------------------------------------------
   * RESOLVEMOS LAS PESTAÑAS
   *----------------------------------------------------------------------------

 //  resuelve_pestana("hpesta2")
   resuelve_pestana("hpesta1")

   *----------------------------------------------------------------------------
   * FORMATOS ESPECIALES
   *----------------------------------------------------------------------------

   appdata:oexcel:sheets("general"):select()

   ohoja := appdata:oexcel:ActiveSheet

   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------
/*
   mensaje_avance("Configurando titulos de grupos...")

   *----------------------------------------------------------------------------
   * RENGLONES DE PAGOS EN AMARILLO
   *----------------------------------------------------------------------------

   for each nrenglon in hpaso["arenglones_pagos"]

      hpaso["crango"]:="A" + ToString(nrenglon) + ":" + appdata:hexcel["hpesta1"]["cletra_tope"] + ToString(nrenglon)

      ohoja:Range(hpaso["crango"]):Interior:Color         := dgs_excel_color_amarillo_bajo

   next
/*
   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------

   mensaje_avance("Configurando totales de grupos...")

   *----------------------------------------------------------------------------
   * RENGLONES TOTALES POR GRUPOS
   *----------------------------------------------------------------------------

   for each nrenglon in hpaso["atotales_grupos"]

      hpaso["crango"]:="A" + ToString(nrenglon) + ":" + appdata:hexcel["hpesta1"]["cletra_tope"] + ToString(nrenglon)

      ohoja:Range(hpaso["crango"]):Font:color       := dgs_excel_color_interior_titulos
      ohoja:Range(hpaso["crango"]):Font:Bold        := .t.
      ohoja:Range(hpaso["crango"]):Font:size        := 16
      ohoja:Range(hpaso["crango"]):VerticalAlignment:=-4108

   next

   *----------------------------------------------------------------------------
   * GRABAMOS EL ARCHIVO A ENTREGAR A CLIENTE
   *----------------------------------------------------------------------------

   appdata:hreporte["cnombre_completo_entregable"] := StrTran(appdata:hreporte["cnombre_completo_entregable"], "/", "\")

   appdata:hexcel["active"] := appdata:oexcel:ActiveWorkbook

   appdata:hexcel["active"]:saveas( appdata:hreporte["cnombre_completo_entregable"], 51, nil, nil, .f.)  // parametro 51 es para formato XLSX,  56 es para XLS

   appdata:oexcel:quit()
*/
   *----------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *----------------------------------------------------------------------------

   mensaje_avance("Gestionando Entrega de Archivo Solicitado")

   hpaso["cgraba"] := HB_JsonEncode(appdata:hestado )

   hb_memowrit( appdata:hreporte["carchivo_finalizacion"], hpaso["cgraba"] )

   *borra_temporales()

   LogFile(DToC(Date())+" "+Time()+"   ******** Finalizando y saliendo de funcion genera_excel_cartera_clientes_antiguedad_un_cliente   "+Str(seconds(),12,4))


RETURN NIL

//------------------------------------------------------------------------------

