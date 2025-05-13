/*
/*
 * Proyecto: Api-Dygy
 * Fichero: Consulta_Cartera_Clientes_Pdf.prg
 * Descripci�n:
 * Autor:
 * Fecha: 24/01/2023
 */

#include "Xailer.ch"
#include "Digio.ch"

//------------------------------------------------------------------------------

FUNCTION consulta_cartera_clientes_pdf()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, hpunto:= {=>}, ohoja, nrenglon, nid_factura

   LogFile(DToC(Date())+" "+Time()+"   ******** en API-DYGY Inicializando funcion consulta_cartera_clientes_pdf   "+Str(seconds(),12,4))

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
/*
         AAdd(:aitems, {"cid" => "1", "ctexto" => "Relaciones Analiticas"})
         AAdd(:aitems, {"cid" => "2", "ctexto" => "Auxiliar"})
         AAdd(:aitems, {"cid" => "3", "ctexto" => "Antiguedad de Saldos"})
         AAdd(:aitems, {"cid" => "4", "ctexto" => "Analiticas Todos los Clientes"})
         AAdd(:aitems, {"cid" => "5", "ctexto" => "Auxiliar Todos los Clientes"})
         AAdd(:aitems, {"cid" => "6", "ctexto" => "Antiguedad Todos los Clientes"})
*/
   if appdata:hreporte["ctipo_tablero"]="1" .or. appdata:hreporte["ctipo_tablero"]="4"

      genera_pdf_cartera_clientes_analiticos()

      return(nil)

   endif

   if appdata:hreporte["ctipo_tablero"]="2" .or. appdata:hreporte["ctipo_tablero"]="5"

      genera_pdf_cartera_clientes_auxiliar()

      return(nil)

   endif

   if appdata:hreporte["ctipo_tablero"]="3" .or. appdata:hreporte["ctipo_tablero"]="6"

      genera_pdf_cartera_clientes_antiguedad()

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

FUNCTION genera_pdf_cartera_clientes_antiguedad()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, hcolumna:= {=>}, ohoja, nrenglon

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

   LogFile(DToC(Date())+" "+Time()+"        carma="+hpaso["carma"])

   hpaso["adant"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   hadoerror["cquery"]          := hpaso["carma"]
   hadoerror["ctitulo_mensaje"] := "Oops...   El sistema no puede obtener datos de antiguedad de cliente"

   if error_ado(hadoerror)
      quit
   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Inicializando documento PDF...")

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
   * BUSCAMOS LA IMPRESORA PDF, SI NO LA ENCONTRAMOS MANDAMOS MENSAJE A USUARIO Y GUARDAMOS REGISTRO DEL ERROR
   *=================================================================================================================================================

   hpaso["napu"]:=Ascan( Printer:aPrinterNames, {|x| x == "PDF24" } )

   if hpaso["napu"] = 0

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "Error no se encontro impresora PDF24"
      hestado["cTituloTextoUsuario"]       := "El servidor no pudo inicializar documento PDF"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] := HB_JsonEncode( hestado )

      hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

      appdata:hreporte["laborta_proceso"] := .f.

      quit

   endif

   *=================================================================================================================================================
   * NOS POSICIONAMOS EN IMPRESORES E INICIAMOS LA CRACION DE PDF
   *=================================================================================================================================================

   hpaso["noldindex"]:=Printer:nPrinterIndex

   Printer:lPreview         := .f.
   Printer:nPrinterIndex    := hpaso["napu"]
   Printer:StartDoc(appdata:hreporte["cnombre_entregable"])
   Printer:nOrientation     :=2        /*Horizontal*/
   Printer:oCanvas:nMapMode := mmHIMETRICS

   WITH OBJECT Printer:oCanvas
      :oFont       := TFont():Create( "Arial", 12, 1, 400 )
      :oPen        := TPen():New( PS_SOLID, 1, CLR_BLACK )
      :oFont:lbold := .f.
      :lTransparent:= .t.
   END WITH

   *=================================================================================================================================================================
   * VARIABLES GENERALES ESPECIFICAS DE IMPRESION
   *=================================================================================================================================================================

   appdata:himpresion["ltitulos"]               :=.t.
   appdata:himpresion["nhoja"]                  := 0
   appdata:himpresion["ntope_derecha"]          := (Printer:nPaperWidth / 100) -.5
   appdata:himpresion["ntope_abajo"]            := (Printer:nPaperLength / 100) -.5
   appdata:himpresion["ntope_brinco"]           := appdata:himpresion["ntope_abajo"]-2
   appdata:himpresion["nrenglon"]               := 1
   appdata:himpresion["ninter"]                 := 0.35
   appdata:himpresion["nmargen_superior"]       := 0.06

   LogFile(DToC(Date())+" "+Time()+"            *** definiendo variable de impresion ntope_abajo="+LTrim(str(appdata:himpresion["ntope_abajo"],12,4))+" ntope_brinco="+LTrim(str(appdata:himpresion["ntope_brinco"],12,4)))

   *=================================================================================================================================================================
   * COMPORTAMIENTO GENERAL DE COLOR Y RAYADO
   *=================================================================================================================================================================

   appdata:himpresion["color"]                  := dgs_color_negro
   appdata:himpresion["rayado"]                 := "raya_abajo"

   *=================================================================================================================================================================
   * VARIABLES DE CONTROL DE PAUTADOS
   *=================================================================================================================================================================

   appdata:himpresion["lpautado"]               := .t.
   appdata:himpresion["ncolor_pautado_alterno"] := dgs_excel_color_alt_fondo_datos
   appdata:himpresion["nalterno"]               := 0

   *=================================================================================================================================================================
   * VARIABLES DE CONTROL PARA DATOS QUE REQUIERAN SER DIVIDIDOS EN VARIOS RENGLONES POR EXCEDER EL ANCHO DE IMPRESION ASIGNADO
   *=================================================================================================================================================================

   appdata:himpresion["media_wrap"]    := 0
   appdata:himpresion["avance_maximo"] := 0

   *=================================================================================================================================================================
   * ARMADO DE COLUMNAS DE REPORTE
   *=================================================================================================================================================================

   appdata:himpresion["acolus"]:= {}

   AAdd(appdata:himpresion["acolus"], {"titulo" => "FACTURA"               , "area" => {"izq" =>  0.6, "der" =>  2.3}, "alineacion" => "izq", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "FECHA"                 , "area" => {"izq" =>  2.3, "der" =>  3.8}, "alineacion" => "izq", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "CONCEPTO"              , "area" => {"izq" =>  3.8, "der" => 11.0}, "alineacion" => "izq", "lwrap" => .t., "fuente_titulo" => 7, "fuente_dato" => 6})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "IMPORTE"               , "area" => {"izq" => 11.0, "der" => 13.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "PAGOS"                 , "area" => {"izq" => 13.0, "der" => 15.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "SALDO"                 , "area" => {"izq" => 15.0, "der" => 17.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "NO VENCIDO"            , "area" => {"izq" => 17.0, "der" => 19.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "VENCIDO     1-30 DIAS" , "area" => {"izq" => 19.0, "der" => 21.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 6, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "VENCIDO 31-60 DIAS"    , "area" => {"izq" => 21.0, "der" => 23.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 6, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "VENCIDO 61-90 DIAS"    , "area" => {"izq" => 23.0, "der" => 25.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 6, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "VENCIDO MAS DE 90 DIAS", "area" => {"izq" => 25.0, "der" => 27.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 6, "fuente_dato" => 7})

   * FORMATO
   *     1        2        3          4        5       6       7          8       9       10     11
   *|         |       |          |         |       |       | NO      | VENC   | VENC  |  VENC | VENC |
   *| FACTURA | FECHA | CONCEPTO | IMPORTE | PAGOS | SALDO | VENCIDO | 1 - 30 | 31-60 | 61-90 | > 90 |

   configura_media()

   for each hrenglon in hpaso["adant"]

      if appdata:hreporte["ctipo_tablero"]="3" .and. hrenglon["clase"]="04"
         loop
      endif

      LogFile(DToC(Date())+" "+Time()+"            *** un nuevo renglon      nrenglon="+LTrim(str(appdata:himpresion["nrenglon"],12,4))+" ntope_abajo="+LTrim(str(appdata:himpresion["ntope_abajo"],12,4))+" ntope_brinco="+LTrim(str(appdata:himpresion["ntope_brinco"],12,4)))

      if appdata:himpresion["ltitulos"]

         if appdata:himpresion["nhoja"]>0
            Printer:EndPage()
         endif

         Printer:StartPage()

         appdata:himpresion["nhoja"] += 1

         appdata:himpresion["nrenglon"]:= 0.5

         escribe({"nfuente"  => 16 ,  "ncolor" => dgs_color_base   ,  "ncolumna" => 0.50, "ctexto"   => appdata:hreporte["cnombre_empresa"]})

         linea_simple({"narriba" => 1.2               , "nizquierda" => 0.5, "nabajo" =>  1.2              , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         escribe({"ncolumna" => appdata:himpresion["ntope_derecha"]  ,  "nfuente" => 7  ,  "calineacion" => "der"  ,  "ctexto"   => "Hoja # "+ToString(appdata:himpresion["nhoja"])})

         appdata:himpresion["nrenglon"]:= 1.6

         escribe_dato({"ncolumna" => 1.0  ,  "ctexto" => "REPORTE"         ,  "cdato" => "ANTIGUEDAD DE SALDOS"})

         if appdata:hreporte["ctipo_tablero"]="3"

            escribe_dato({"ncolumna" => 8.0  ,  "ctexto" => "CLIENTE"         ,  "cdato" => appdata:hreporte["cnombre_cliente"]})

         endif

         escribe_dato({"ncolumna" => 18.0 ,  "ctexto" => "FECHA"           ,  "cdato" => "AL: " + fectal(appdata:hreporte["cfecha_final"])})

         linea_simple({"narriba" => 2.7, "nizquierda" => 0.5, "nabajo" => 2.7, "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         appdata:himpresion["nrenglon"] := 2.7

         escribe_encabezado(dgs_color_base)

         appdata:himpresion["nrenglon"] += appdata:himpresion["ninter"]

         linea_simple({"narriba" => appdata:himpresion["nrenglon"]-0.01 , "nizquierda" => 0.5, "nabajo" => appdata:himpresion["nrenglon"]-0.01 , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         *=================================================================================================================================================
         * AJUSTE POR EL BRINCO QUE VENDRA AL IMPRIMIR LAS LINEAS DE DATOS
         *=================================================================================================================================================

         appdata:himpresion["nrenglon"] -= appdata:himpresion["ninter"]

         appdata:himpresion["ltitulos"]:=.f.

         LogFile(DToC(Date())+" "+Time()+"                !!!  imprimi titulos   quedo  nrenglon="+LTrim(str(appdata:himpresion["nrenglon"],12,4))+" ntope_abajo="+LTrim(str(appdata:himpresion["ntope_abajo"],12,4))+" ntope_brinco="+LTrim(str(appdata:himpresion["ntope_brinco"],12,4)))

      endif

      hpaso["cfactura"]                    := iif(hrenglon["factura"]=nil                   , "", hrenglon["factura"])
      hpaso["cfecha"]                      := iif(hrenglon["fecha"]=nil                     , "", fechaammdd(hrenglon["fecha"]))
      hpaso["cconcepto"]                   := iif(hrenglon["descripcion"]=nil               , "", hrenglon["descripcion"])
      hpaso["cimporte"]                    := Transform(hrenglon["importe"]                 , "999,999.99")
      hpaso["cpagos"]                      := transform(hrenglon["pagos"]                   , "999,999.99")
      hpaso["csaldo"]                      := transform(hrenglon["saldo"]                   , "999,999.99")
      hpaso["cno_vencido"]                 := transform(hrenglon["no_vencido"]              , "999,999.99")
      hpaso["cvencido_30"]                 := transform(hrenglon["vencido_01_30_dias"]      , "999,999.99")
      hpaso["cvencido_60"]                 := transform(hrenglon["vencido_31_60_dias"]      , "999,999.99")
      hpaso["cvencido_90"]                 := transform(hrenglon["vencido_61_90_dias"]      , "999,999.99")
      hpaso["cvencido_mas"]                := transform(hrenglon["vencido_mas_de_90_dias"]  , "999,999.99")

      appdata:himpresion["acolus"][1]["dato"]  := hpaso["cfactura"]
      appdata:himpresion["acolus"][2]["dato"]  := hpaso["cfecha"]
      appdata:himpresion["acolus"][3]["dato"]  := hpaso["cconcepto"]
      appdata:himpresion["acolus"][4]["dato"]  := hpaso["cimporte"]
      appdata:himpresion["acolus"][5]["dato"]  := hpaso["cpagos"]
      appdata:himpresion["acolus"][6]["dato"]  := hpaso["csaldo"]
      appdata:himpresion["acolus"][7]["dato"]  := hpaso["cno_vencido"]
      appdata:himpresion["acolus"][8]["dato"]  := hpaso["cvencido_30"]
      appdata:himpresion["acolus"][9]["dato"]  := hpaso["cvencido_60"]
      appdata:himpresion["acolus"][10]["dato"] := hpaso["cvencido_90"]
      appdata:himpresion["acolus"][11]["dato"] := hpaso["cvencido_mas"]

      appdata:himpresion["nrenglon"] += appdata:himpresion["ninter"]

      do case

         case appdata:hreporte["ctipo_tablero"]="6" .and. hrenglon["clase"]="04"

            *=================================================================================================================================================
            * NOMBRE DE CLIENTE EN OPCION TODOS LOS CLIENTES
            *=================================================================================================================================================

            appdata:himpresion["nrenglon"] += appdata:himpresion["ninter"]

            escribe({"ncolumna" => 2.4  ,  "nfuente" => 9 ,  "calineacion" => "izq"  ,  "ctexto"   => hrenglon["nombre_cliente"], "lnegrita" => .t., "ncolor" => dgs_color_base})

            appdata:himpresion["nrenglon"] += 0.2

         case hrenglon["clase"]="12"

            *=================================================================================================================================================
            * TOTAL CLIENTE TANTO PARA UN CLIENTE COMO TODOS LOS CLIENTES
            *=================================================================================================================================================

            appdata:himpresion["acolus"][1]["dato"]  := ""
            appdata:himpresion["acolus"][2]["dato"]  := ""
            appdata:himpresion["acolus"][3]["dato"]  := IIF(appdata:hreporte["ctipo_tablero"]="3", "TOTALES", "TOTAL")

            linea_simple({"narriba" => appdata:himpresion["nrenglon"] , "nizquierda" => 3.8, "nabajo" =>  appdata:himpresion["nrenglon"] , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

            escribe_linea_datos({"lnegrita" => .t., "ncolor" => dgs_color_base, "nfuente" => 8, "lsin_raya" => .t.})

            linea_simple({"narriba" => appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nizquierda" => 3.8, "nabajo" =>  appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

            appdata:himpresion["nrenglon"] += 0.1

         case appdata:hreporte["ctipo_tablero"]="6" .and. hrenglon["clase"]="16"

            *=================================================================================================================================================
            * TOTAL GENERAL OPCION TODOS LOS CLIENTES
            *=================================================================================================================================================

            appdata:himpresion["acolus"][1]["dato"]  := ""
            appdata:himpresion["acolus"][2]["dato"]  := ""
            appdata:himpresion["acolus"][3]["dato"]  := "TOTAL GENERAL"

            linea_simple({"narriba" => appdata:himpresion["nrenglon"] , "nizquierda" => 0.5, "nabajo" =>  appdata:himpresion["nrenglon"] , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

            escribe_linea_datos({"lnegrita" => .t., "ncolor" => dgs_color_base, "nfuente" => 9, "lsin_raya" => .t.})

            linea_simple({"narriba" => appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nizquierda" => 0.5, "nabajo" =>  appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

            appdata:himpresion["nrenglon"] += 0.1

         otherwise

            *=================================================================================================================================================
            * LINEA TIPICA
            *=================================================================================================================================================

            escribe_linea_datos({"nada" => ""})

      endcase

      * FORMATO
      *     1        2        3          4        5       6       7          8       9       10     11
      *|         |       |          |         |       |       | NO      | VENC   | VENC  |  VENC | VENC |
      *| FACTURA | FECHA | CONCEPTO | IMPORTE | PAGOS | SALDO | VENCIDO | 1 - 30 | 31-60 | 61-90 | > 90 |

      LogFile(DToC(Date())+" "+Time()+"            ### revisando si brinco de hoja nrenglon="+LTrim(str(appdata:himpresion["nrenglon"],12,4))+" ntope_abajo="+LTrim(str(appdata:himpresion["ntope_abajo"],12,4))+" ntope_brinco="+LTrim(str(appdata:himpresion["ntope_brinco"],12,4)))

      *=================================================================================================================================================
      * BUSCAMOS SI LA LINEA QUE SIGUE ES DE TOTALES PARA EVITAR LA REVISION DE BRINCO DE HOJA
      *=================================================================================================================================================

      if hrenglon:__enumindex < Len(hpaso["adant"])
         if hpaso["adant"][hrenglon:__enumindex +1]["clase"] = "12"
            loop
         endif
      endif

      *=================================================================================================================================================
      * DETECTAMOS SI YA REBASO EL TOPE DE ABAJO PARA PRENDER LA VARIABLE appdata:himpresion[LTITULOS"] PARA QUE EN LA SIGUIENTE VUELTA IMPRIMA TITULOS
      *=================================================================================================================================================

      if appdata:himpresion["nrenglon"] >= appdata:himpresion["ntope_brinco"]

         hpaso["nrenglon"] := appdata:himpresion["ntope_abajo"]-0.5

         escribe({"nrenglon" => hpaso["nrenglon"], "nfuente" => 9  ,  "calineacion" => "der"  ,  "ncolor" => dgs_color_base  ,  "ncolumna" => 21.0  ,  "ctexto" => "===> "+ToString(appdata:himpresion["nhoja"]+1)})

         appdata:himpresion["ltitulos"]:=.t.

      endif

   next

   *=================================================================================================================================================
   * QUE FALTA
   *=================================================================================================================================================
   *
   *=================================================================================================================================================

   Printer:EndPage()
   Printer:EndDoc()

   Printer:nPrinterIndex    := hpaso["noldindex"]

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Finalizando Documento...")

   *=================================================================================================================================================
   * REALIZAMOS LA COPIA DEL ARCHIVO A LA RUTA DONDE EL CGI LO ESPERA PARA DESCARGARLO
   *=================================================================================================================================================

   hpaso["cfile_generado"] := "C:\apache24\htdocs\pdf\"+appdata:hreporte["cnombre_entregable"]

   LogFile(DToC(Date())+" "+Time()+"   ******** cfile_generado "+hpaso["cfile_generado"]+"    "+Str(seconds(),12,4))

   *=================================================================================================================================================
   * HAREMOS UN LOOP DE ESPERA PARA DAR OPORTUNIDAD DE QUE SE GRABE EL ARCHIVO FISICAMENTE, LA EJECUCION DE CODIGO ES MAS RAPIDO QUE LAS
   * ACCIONES FISICAS EN DISCO, EL TIEMPO DE ESPERA SERA DE 10 SEGUNDOS
   *=================================================================================================================================================

   hpaso["tiempo_inicial"]:= Seconds()
   hpaso["lse_pudo"]:= .f.

   do while .t.

      if File(hpaso["cfile_generado"])
         hpaso["lse_pudo"]:= .t.
         exit
      endif

      hpaso["tiempo_actual"] := Seconds()

      if hpaso["tiempo_actual"] - hpaso["tiempo_inicial"] >10
         exit
      endif

   enddo

   LogFile(DToC(Date())+" "+Time()+"   ******** supero loop de revision de existencia de archivo en disco    "+Str(seconds(),12,4))

   *=================================================================================================================================================
   * REVISAMOS SI SE PUDO
   *=================================================================================================================================================

   if !hpaso["lse_pudo"]

      *=================================================================================================================================================
      * POR ALGUNA RAZON NO SE PUDO DETECTAR EN 10 SEGUNDOS LA CREACION FISICA DEL ARCHIVO PDF, ABORTAMOS Y MANDAMOS SEҁL AL USUARIO/CGGI
      *=================================================================================================================================================

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "Error dentro de API, excedio de 10 segundos la grabacion de archivo fisico pdf"
      hestado["cTituloTextoUsuario"]       := "El servidor no pudo grabar documento pdf (time exceeded)"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] := HB_JsonEncode( hestado )

      hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

      appdata:hreporte["laborta_proceso"] := .f.

      LogFile(DToC(Date())+" "+Time()+"   ******** Finalizando !!!!!SIN RESULTADO!!!! (mas de 10 segundos para creacion de pdf) de funcion genera_pdf_cartera_clientes_antiguedad_un_cliente   "+Str(seconds(),12,4))

      quit

   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Iniciando Envio de Documento...")

   *DELETE FILE ''+hpaso["cfile_generado"]+''

   *--------------------------------------------------------------------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *--------------------------------------------------------------------------------------------------------------------------------------

   mensaje_avance("Gestionando Entrega de Documento Solicitado")

   hpaso["cgraba"] := HB_JsonEncode(appdata:hestado )

   hb_memowrit( appdata:hreporte["carchivo_finalizacion"], hpaso["cgraba"] )

   *borra_temporales()

   LogFile(DToC(Date())+" "+Time()+"   ******** Finalizando y saliendo de funcion imprime_una_requisicion   "+Str(seconds(),12,4))

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION genera_pdf_cartera_clientes_auxiliar()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, hcolumna:= {=>}, ohoja, nrenglon

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

   LogFile(DToC(Date())+" "+Time()+"   carma=" + hpaso["carma"])

   hpaso["adant"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   hadoerror["cquery"]          := hpaso["carma"]
   hadoerror["ctitulo_mensaje"] := "Oops...   El sistema no puede obtener datos de antiguedad de cliente"

   if error_ado(hadoerror)
      quit
   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Inicializando documento PDF...")

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
   * BUSCAMOS LA IMPRESORA PDF, SI NO LA ENCONTRAMOS MANDAMOS MENSAJE A USUARIO Y GUARDAMOS REGISTRO DEL ERROR
   *=================================================================================================================================================

   hpaso["napu"]:=Ascan( Printer:aPrinterNames, {|x| x == "PDF24" } )

   if hpaso["napu"] = 0

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "Error no se encontro impresora PDF24"
      hestado["cTituloTextoUsuario"]       := "El servidor no pudo inicializar documento PDF"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] := HB_JsonEncode( hestado )

      hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

      appdata:hreporte["laborta_proceso"] := .f.

      quit

   endif

   *=================================================================================================================================================
   * NOS POSICIONAMOS EN IMPRESORES E INICIAMOS LA CRACION DE PDF
   *=================================================================================================================================================

   hpaso["noldindex"]:=Printer:nPrinterIndex

   Printer:lPreview         := .f.
   Printer:nPrinterIndex    := hpaso["napu"]
   Printer:nPaperSizeType   := 1          /* 1 - letter    2- legal (oficio) */
   Printer:StartDoc(appdata:hreporte["cnombre_entregable"])
   Printer:nOrientation     :=1        /*Vertical*/
   Printer:oCanvas:nMapMode := mmHIMETRICS

   WITH OBJECT Printer:oCanvas
      :oFont       := TFont():Create( "Arial", 12, 1, 400 )
      :oPen        := TPen():New( PS_SOLID, 1, CLR_BLACK )
      :oFont:lbold := .f.
      :lTransparent:= .t.
   END WITH

   *=================================================================================================================================================================
   * VARIABLES GENERALES ESPECIFICAS DE IMPRESION
   *=================================================================================================================================================================

   appdata:himpresion["ltitulos"]               :=.t.
   appdata:himpresion["nhoja"]                  := 0
   appdata:himpresion["ntope_derecha"]          := (Printer:nPaperWidth / 100) -.5
   appdata:himpresion["ntope_abajo"]            := (Printer:nPaperLength / 100) -.5
   appdata:himpresion["ntope_brinco"]           := appdata:himpresion["ntope_abajo"]-2
   appdata:himpresion["nrenglon"]               := 1
   appdata:himpresion["ninter"]                 := 0.38
   appdata:himpresion["nmargen_superior"]       := 0.05

   *=================================================================================================================================================================
   * COMPORTAMIENTO GENERAL DE COLOR Y RAYADO
   *=================================================================================================================================================================

   appdata:himpresion["color"]                  := dgs_color_negro
   appdata:himpresion["rayado"]                 := "raya_abajo"

   *=================================================================================================================================================================
   * VARIABLES DE CONTROL DE PAUTADOS
   *=================================================================================================================================================================

   appdata:himpresion["lpautado"]               := .t.
   appdata:himpresion["ncolor_pautado_alterno"] := dgs_excel_color_alt_fondo_datos
   appdata:himpresion["nalterno"]               := 0

   *=================================================================================================================================================================
   * VARIABLES DE CONTROL PARA DATOS QUE REQUIERAN SER DIVIDIDOS EN VARIOS RENGLONES POR EXCEDER EL ANCHO DE IMPRESION ASIGNADO
   *=================================================================================================================================================================

   appdata:himpresion["media_wrap"]    := 0
   appdata:himpresion["avance_maximo"] := 0

   *=================================================================================================================================================================
   * COLORES ESPECIALES SOLO PARA USAR EN ESTA FUNCION
   *=================================================================================================================================================================

   hpaso["ncafe"] := rgb(179, 079, 015)

   *=================================================================================================================================================================
   * ARMADO DE COLUMNAS DE REPORTE
   *=================================================================================================================================================================

   appdata:himpresion["acolus"]:= {}

   AAdd(appdata:himpresion["acolus"], {"titulo" => "FACTURA"               , "area" => {"izq" =>  0.6, "der" =>  2.3}, "alineacion" => "izq", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "FECHA"                 , "area" => {"izq" =>  2.3, "der" =>  3.8}, "alineacion" => "izq", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "CONCEPTO"              , "area" => {"izq" =>  3.8, "der" => 13.0}, "alineacion" => "izq", "lwrap" => .t., "fuente_titulo" => 7, "fuente_dato" => 6})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "SALDO ANTERIOR"        , "area" => {"izq" => 13.0, "der" => 15.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "CARGOS"                , "area" => {"izq" => 15.0, "der" => 17.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "ABONOS"                , "area" => {"izq" => 17.0, "der" => 19.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "SALDO ACTUAL"          , "area" => {"izq" => 19.0, "der" => 21.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})

   * FORMATO
   *     1        2        3          4         5        6        7
   *|         |       |          | SALDO    |        |        | SALDO  |
   *| FACTURA | FECHA | CONCEPTO | ANTERIOR | CARGOS | ABONOS | ACTUAL |

   *=================================================================================================================================================================
   * OBTENEMOS LA VARIABLE HCOLUMNA[AJUSTE_ALTO] PARA CADA ELEMENTO DE appdata:himpresion[ACOLUS] PARA LOGRAR IMPRESIONES EN MEDIO CUANDO LAS COLUMNAS TIENEN DIFERENTES
   * TAMAÑOS PERO DEBEN APARECER EN LA MISMA LINEA DE DATOS, SE OBTENDRA UNA ALTURA ADECUADA PARA QUE TODOS LOS TEXTOS COINCIDAN EN UN PUNTO MEDIO
   *=================================================================================================================================================================

   configura_media(@appdata:himpresion)

   *=================================================================================================================================================================
   * VARIABLES ESPECIALES PARA OBTENER LOS SALDOS ANTERIORES
   *=================================================================================================================================================================

   hpaso["nanterior_cliente"] := 0
   hpaso["nanterior_general"] := 0

   *=================================================================================================================================================================
   * INICIAMOS EL RECORRIDO EN EL CONJUNTO DE DATOS A IMPRIMIR
   *=================================================================================================================================================================

   for each hrenglon in hpaso["adant"]

      if appdata:hreporte["ctipo_tablero"]="3" .and. hrenglon["clase"]="04"
         loop
      endif

      if appdata:himpresion["ltitulos"]

         if appdata:himpresion["nhoja"]>0
            Printer:EndPage()
         endif

         Printer:StartPage()

         appdata:himpresion["nhoja"] += 1

         appdata:himpresion["nrenglon"]:= 0.5

         escribe({"nfuente"  => 16 ,  "ncolor" => dgs_color_base   ,  "ncolumna" => 0.50, "ctexto"   => appdata:hreporte["cnombre_empresa"]})

         linea_simple({"narriba" => 1.2               , "nizquierda" => 0.5, "nabajo" =>  1.2              , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         escribe({"ncolumna" => appdata:himpresion["ntope_derecha"]  ,  "nfuente" => 7  ,  "calineacion" => "der"  ,  "ctexto"   => "Hoja # "+ToString(appdata:himpresion["nhoja"])})

         appdata:himpresion["nrenglon"]:= 1.4

         escribe_dato({"ncolumna"      => 1.0 , ;
                       "ctexto"        => "REPORTE" , ;
                       "cdato"         => "AUXILIAR DE CARTERA", ;
                       "nfuente_dato"  => 9, ;
                       "lnegrita_dato" => .f.}, appdata:himpresion)

         if appdata:hreporte["ctipo_tablero"]="2"

            escribe_dato({"ncolumna"      => 7.0 , ;
                          "ctexto"        => "CLIENTE" , ;
                          "cdato"         => appdata:hreporte["cnombre_cliente"], ;
                          "nfuente_dato"  => 9, ;
                          "lnegrita_dato" => .f.}, appdata:himpresion)

         endif

         appdata:himpresion["nrenglon"] := 2.2

         escribe_dato({"ncolumna"      => 1.0     , ;
                       "ctexto"        => "FECHA" , ;
                       "cdato"         => "DEL: " + fectal(appdata:hreporte["cfecha_inicial"])+ " AL: " + fectal(appdata:hreporte["cfecha_final"]), ;
                       "nfuente_dato"  => 9, ;
                       "lnegrita_dato" => .f.}, appdata:himpresion)

         appdata:himpresion["nrenglon"] := 2.9

         linea_simple({"narriba" => appdata:himpresion["nrenglon"], "nizquierda" => 0.5, "nabajo" => appdata:himpresion["nrenglon"], "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         escribe_encabezado(dgs_color_base)

         appdata:himpresion["nrenglon"] += appdata:himpresion["ninter"]

         linea_simple({"narriba" => appdata:himpresion["nrenglon"]-0.01 , "nizquierda" => 0.5, "nabajo" => appdata:himpresion["nrenglon"]-0.01 , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         *=================================================================================================================================================
         * AJUSTE POR EL BRINCO QUE VENDRA AL IMPRIMIR LAS LINEAS DE DATOS
         *=================================================================================================================================================

         appdata:himpresion["nrenglon"] -= appdata:himpresion["ninter"]

         appdata:himpresion["ltitulos"]:=.f.

      endif

      *=================================================================================================================================================
      * RENGLON DE SALDO ANTERIOR EN OPCION DE UN CLIENTE
      *=================================================================================================================================================

      if appdata:hreporte["ctipo_tablero"]="2" .and. hrenglon["clase"]="03"
         hpaso["nanterior_cliente"] := hrenglon["actual"]
         loop
      endif

      hpaso["cfactura"]                    := iif(hrenglon["esfactura"]=nil                 , "", hrenglon["esfactura"])
      hpaso["cfecha"]                      := iif(hrenglon["fecha"]=nil                     , "", fechaammdd(hrenglon["fecha"]))
      hpaso["cconcepto"]                   := iif(hrenglon["descripcion"]=nil               , "", hrenglon["descripcion"])
      hpaso["canterior"]                   := ""
      hpaso["ccargo"]                      := transform(hrenglon["cargo"]                   , "9,999,999.99")
      hpaso["cabono"]                      := transform(hrenglon["abono"]                   , "9,999,999.99")
      hpaso["cactual"]                     := transform(hrenglon["actual"]                  , "9,999,999.99")

      appdata:himpresion["acolus"][1]["dato"]  := hpaso["cfactura"]
      appdata:himpresion["acolus"][2]["dato"]  := hpaso["cfecha"]
      appdata:himpresion["acolus"][3]["dato"]  := hpaso["cconcepto"]
      appdata:himpresion["acolus"][4]["dato"]  := hpaso["canterior"]
      appdata:himpresion["acolus"][5]["dato"]  := hpaso["ccargo"]
      appdata:himpresion["acolus"][6]["dato"]  := hpaso["cabono"]
      appdata:himpresion["acolus"][7]["dato"]  := hpaso["cactual"]

      appdata:himpresion["nrenglon"] += appdata:himpresion["ninter"]

      do case

         case appdata:hreporte["ctipo_tablero"]="5" .and. hrenglon["clase"]="03"

            *=================================================================================================================================================
            * NOMBRE DE CLIENTE EN OPCION TODOS LOS CLIENTES
            *=================================================================================================================================================

            appdata:himpresion["nrenglon"] += appdata:himpresion["ninter"]

            escribe({"ncolumna" => 2.4  ,  "nfuente" => 9 ,  "calineacion" => "izq"  ,  "ctexto"   => hrenglon["nombre_cliente"], "lnegrita" => .t., "ncolor" => dgs_color_base})

            appdata:himpresion["nrenglon"] += 0.2

            hpaso["nanterior_cliente"] := hrenglon["actual"]

         case hrenglon["clase"]="09" .or. hrenglon["clase"]="10"

            *=================================================================================================================================================
            * TOTAL CLIENTE LAS 2 OPCIONES UN CLIENTE O TODOS LOS CLIENTES
            *=================================================================================================================================================

            appdata:himpresion["acolus"][1]["dato"]  := ""
            appdata:himpresion["acolus"][2]["dato"]  := ""
            appdata:himpresion["acolus"][4]["dato"]  := transform(hpaso["nanterior_cliente"]     , "9,999,999.99")
            appdata:himpresion["acolus"][3]["dato"]  := IIF(appdata:hreporte["ctipo_tablero"]="2", "TOTALES", "TOTAL CLIENTE")

            hpaso["nanterior_general"] += hpaso["nanterior_cliente"]

            appdata:himpresion["nrenglon"] += 0.1

            linea_simple({"narriba" => appdata:himpresion["nrenglon"] , "nizquierda" => 3.8, "nabajo" =>  appdata:himpresion["nrenglon"] , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 0.6})

            escribe_linea_datos({"lnegrita" => .t., "ncolor" => dgs_color_base, "nfuente" => 8, "lsin_raya" => .t., "lconfirma_pautado" => .f.})

            linea_simple({"narriba" => appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nizquierda" => 3.8, "nabajo" =>  appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 0.6})

            appdata:himpresion["nrenglon"] += 0.1

         case appdata:hreporte["ctipo_tablero"]="5" .and. hrenglon["clase"]="14"

            *=================================================================================================================================================
            * TOTAL GENERAL OPCION TODOS LOS CLIENTES
            *=================================================================================================================================================

            appdata:himpresion["acolus"][1]["dato"]  := ""
            appdata:himpresion["acolus"][2]["dato"]  := ""
            appdata:himpresion["acolus"][3]["dato"]  := "TOTAL GENERAL"
            appdata:himpresion["acolus"][4]["dato"]  := transform(hpaso["nanterior_general"]     , "9,999,999.99")

            linea_simple({"narriba" => appdata:himpresion["nrenglon"] , "nizquierda" => 0.5, "nabajo" =>  appdata:himpresion["nrenglon"] , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

            escribe_linea_datos({"lnegrita" => .t., "ncolor" => dgs_color_base, "nfuente" => 9, "lsin_raya" => .t., "lconfirma_pautado" => .f.})

            linea_simple({"narriba" => appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nizquierda" => 0.5, "nabajo" =>  appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

            appdata:himpresion["nrenglon"] += 0.1

         otherwise

            *=================================================================================================================================================
            * LINEA TIPICA CARGOS
            *=================================================================================================================================================

            if hrenglon["clase"] = "07" .or. hrenglon["clase"] = "08"
               appdata:himpresion["acolus"][5]["dato"]  := ""
               escribe_linea_datos({"ncolor" => hpaso["ncafe"]})
            else
               appdata:himpresion["acolus"][6]["dato"]  := ""
               escribe_linea_datos({"nada" => ""})
            endif


      endcase

      * FORMATO
      *     1        2        3          4        5       6       7          8       9       10     11
      *|         |       |          |         |       |       | NO      | VENC   | VENC  |  VENC | VENC |
      *| FACTURA | FECHA | CONCEPTO | IMPORTE | PAGOS | SALDO | VENCIDO | 1 - 30 | 31-60 | 61-90 | > 90 |

      *=================================================================================================================================================
      * BUSCAMOS SI LA LINEA QUE SIGUE ES DE TOTALES PARA EVITAR LA REVISION DE BRINCO DE HOJA
      *=================================================================================================================================================

      if hrenglon:__enumindex < Len(hpaso["adant"])
         if hpaso["adant"][hrenglon:__enumindex +1]["clase"] = "09"
            loop
         endif
      endif

      *=================================================================================================================================================
      * DETECTAMOS SI YA REBASO EL TOPE DE ABAJO PARA PRENDER LA VARIABLE appdata:himpresion[LTITULOS"] PARA QUE EN LA SIGUIENTE VUELTA IMPRIMA TITULOS
      *=================================================================================================================================================

      if appdata:himpresion["nrenglon"] >= appdata:himpresion["ntope_brinco"]

         hpaso["nrenglon"] := appdata:himpresion["ntope_abajo"]-0.5

         escribe({"nrenglon" => hpaso["nrenglon"], "nfuente" => 9  ,  "calineacion" => "der"  ,  "ncolor" => dgs_color_base  ,  "ncolumna" => 21.0  ,  "ctexto" => "===> "+ToString(appdata:himpresion["nhoja"]+1)})

         appdata:himpresion["ltitulos"]:=.t.

      endif

   next

   *=================================================================================================================================================
   * QUE FALTA
   *=================================================================================================================================================
   *
   *=================================================================================================================================================

   Printer:EndPage()
   Printer:EndDoc()

   Printer:nPrinterIndex    := hpaso["noldindex"]

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Finalizando Documento...")

   *=================================================================================================================================================
   * REALIZAMOS LA COPIA DEL ARCHIVO A LA RUTA DONDE EL CGI LO ESPERA PARA DESCARGARLO
   *=================================================================================================================================================

   hpaso["cfile_generado"] := "C:\apache24\htdocs\pdf\"+appdata:hreporte["cnombre_entregable"]

   *=================================================================================================================================================
   * HAREMOS UN LOOP DE ESPERA PARA DAR OPORTUNIDAD DE QUE SE GRABE EL ARCHIVO FISICAMENTE, LA EJECUCION DE CODIGO ES MAS RAPIDO QUE LAS
   * ACCIONES FISICAS EN DISCO, EL TIEMPO DE ESPERA SERA DE 10 SEGUNDOS
   *=================================================================================================================================================

   hpaso["tiempo_inicial"]:= Seconds()
   hpaso["lse_pudo"]:= .f.

   do while .t.

      if File(hpaso["cfile_generado"])
         hpaso["lse_pudo"]:= .t.
         exit
      endif

      hpaso["tiempo_actual"] := Seconds()

      if hpaso["tiempo_actual"] - hpaso["tiempo_inicial"] >10
         exit
      endif

   enddo

   *=================================================================================================================================================
   * REVISAMOS SI SE PUDO
   *=================================================================================================================================================

   if !hpaso["lse_pudo"]

      *=================================================================================================================================================
      * POR ALGUNA RAZON NO SE PUDO DETECTAR EN 10 SEGUNDOS LA CREACION FISICA DEL ARCHIVO PDF, ABORTAMOS Y MANDAMOS SEҁL AL USUARIO/CGGI
      *=================================================================================================================================================

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "Error dentro de API, excedio de 10 segundos la grabacion de archivo fisico pdf"
      hestado["cTituloTextoUsuario"]       := "El servidor no pudo grabar documento pdf (time exceeded)"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] := HB_JsonEncode( hestado )

      hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

      appdata:hreporte["laborta_proceso"] := .f.

      quit

   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Iniciando Envio de Documento...")

   *DELETE FILE ''+hpaso["cfile_generado"]+''

   *--------------------------------------------------------------------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *--------------------------------------------------------------------------------------------------------------------------------------

   mensaje_avance("Gestionando Entrega de Documento Solicitado")

   hpaso["cgraba"] := HB_JsonEncode(appdata:hestado )

   hb_memowrit( appdata:hreporte["carchivo_finalizacion"], hpaso["cgraba"] )

   *borra_temporales()

   LogFile(DToC(Date())+" "+Time()+"   ******** Finalizando y saliendo de funcion genera_pdf_cartera_clientes_auxiliar   "+Str(seconds(),12,4))

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION genera_pdf_cartera_clientes_analiticos()

   LOCAL hpaso:={=>}, hestado:= {=>}, hadoerror:= {=>}
   LOCAL hrenglon:= {=>}, hcolumna:= {=>}, ohoja, nrenglon

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

   LogFile(DToC(Date())+" "+Time()+"   carma=" + hpaso["carma"])

   hpaso["adana"] :=appdata:origen_bd:QueryArrayHash(hpaso["carma"])

   hadoerror["cquery"]          := hpaso["carma"]
   hadoerror["ctitulo_mensaje"] := "Oops...   El sistema no puede obtener datos de antiguedad de cliente"

   if error_ado(hadoerror)
      quit
   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Inicializando documento PDF...")

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
   * BUSCAMOS LA IMPRESORA PDF, SI NO LA ENCONTRAMOS MANDAMOS MENSAJE A USUARIO Y GUARDAMOS REGISTRO DEL ERROR
   *=================================================================================================================================================

   hpaso["napu"]:=Ascan( Printer:aPrinterNames, {|x| x == "PDF24" } )

   if hpaso["napu"] = 0

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "Error no se encontro impresora PDF24"
      hestado["cTituloTextoUsuario"]       := "El servidor no pudo inicializar documento PDF"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] := HB_JsonEncode( hestado )

      hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

      appdata:hreporte["laborta_proceso"] := .f.

      quit

   endif

   *=================================================================================================================================================
   * NOS POSICIONAMOS EN IMPRESORES E INICIAMOS LA CRACION DE PDF
   *=================================================================================================================================================

   hpaso["noldindex"]:=Printer:nPrinterIndex

   Printer:lPreview         := .f.
   Printer:nPrinterIndex    := hpaso["napu"]
   Printer:nPaperSizeType   := 1          /* 1 - letter    2- legal (oficio) */
   Printer:StartDoc(appdata:hreporte["cnombre_entregable"])
   Printer:nOrientation     :=1        /*Vertical*/
   Printer:oCanvas:nMapMode := mmHIMETRICS

   WITH OBJECT Printer:oCanvas
      :oFont       := TFont():Create( "Arial", 12, 1, 400 )
      :oPen        := TPen():New( PS_SOLID, 1, CLR_BLACK )
      :oFont:lbold := .f.
      :lTransparent:= .t.
   END WITH

   *=================================================================================================================================================================
   * VARIABLES GENERALES ESPECIFICAS DE IMPRESION
   *=================================================================================================================================================================

   appdata:himpresion["ltitulos"]               :=.t.
   appdata:himpresion["nhoja"]                  := 0
   appdata:himpresion["ntope_derecha"]          := (Printer:nPaperWidth / 100) -.5
   appdata:himpresion["ntope_abajo"]            := (Printer:nPaperLength / 100) -.5
   appdata:himpresion["ntope_brinco"]           := appdata:himpresion["ntope_abajo"]-2
   appdata:himpresion["nrenglon"]               := 1
   appdata:himpresion["ninter"]                 := 0.38
   appdata:himpresion["nmargen_superior"]       := 0.05

   *=================================================================================================================================================================
   * COMPORTAMIENTO GENERAL DE COLOR Y RAYADO
   *=================================================================================================================================================================

   appdata:himpresion["color"]                  := dgs_color_negro
   appdata:himpresion["rayado"]                 := "raya_abajo"

   *=================================================================================================================================================================
   * VARIABLES DE CONTROL DE PAUTADOS
   *=================================================================================================================================================================

   appdata:himpresion["lpautado"]               := .t.
   appdata:himpresion["ncolor_pautado_alterno"] := dgs_excel_color_alt_fondo_datos
   appdata:himpresion["nalterno"]               := 0

   *=================================================================================================================================================================
   * VARIABLES DE CONTROL PARA DATOS QUE REQUIERAN SER DIVIDIDOS EN VARIOS RENGLONES POR EXCEDER EL ANCHO DE IMPRESION ASIGNADO
   *=================================================================================================================================================================

   appdata:himpresion["media_wrap"]    := 0
   appdata:himpresion["avance_maximo"] := 0

   *=================================================================================================================================================================
   * COLORES ESPECIALES SOLO PARA USAR EN ESTA FUNCION
   *=================================================================================================================================================================

   hpaso["ncafe"] := rgb(179, 079, 015)

   *=================================================================================================================================================================
   * ARMADO DE COLUMNAS DE REPORTE
   *=================================================================================================================================================================

   appdata:himpresion["acolus"]:= {}

   AAdd(appdata:himpresion["acolus"], {"titulo" => "CLIENTE"               , "area" => {"izq" =>  0.6, "der" =>  2.3}, "alineacion" => "izq", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "NOMBRE"                , "area" => {"izq" =>  2.3, "der" => 13.0}, "alineacion" => "izq", "lwrap" => .t., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "SALDO ANTERIOR"        , "area" => {"izq" => 13.0, "der" => 15.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "CARGOS"                , "area" => {"izq" => 15.0, "der" => 17.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "ABONOS"                , "area" => {"izq" => 17.0, "der" => 19.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})
   AAdd(appdata:himpresion["acolus"], {"titulo" => "SALDO ACTUAL"          , "area" => {"izq" => 19.0, "der" => 21.0}, "alineacion" => "der", "lwrap" => .f., "fuente_titulo" => 7, "fuente_dato" => 7})

   * FORMATO
   *     1         2        3          4        5       6
   *|         |        | SALDO    |        |        | SALDO  |
   *| CLIENTE | NOMBRE | ANTERIOR | CARGOS | ABONOS | ACTUAL |

   *=================================================================================================================================================================
   * OBTENEMOS LA VARIABLE HCOLUMNA[AJUSTE_ALTO] PARA CADA ELEMENTO DE appdata:himpresion[ACOLUS] PARA LOGRAR IMPRESIONES EN MEDIO CUANDO LAS COLUMNAS TIENEN DIFERENTES
   * TAMAÑOS PERO DEBEN APARECER EN LA MISMA LINEA DE DATOS, SE OBTENDRA UNA ALTURA ADECUADA PARA QUE TODOS LOS TEXTOS COINCIDAN EN UN PUNTO MEDIO
   *=================================================================================================================================================================

   configura_media(@appdata:himpresion)

   *=================================================================================================================================================================
   * INICIAMOS EL RECORRIDO EN EL CONJUNTO DE DATOS A IMPRIMIR
   *=================================================================================================================================================================

   for each hrenglon in hpaso["adana"]

      if appdata:hreporte["ctipo_tablero"]="3" .and. hrenglon["clase"]="04"
         loop
      endif

      if appdata:himpresion["ltitulos"]

         if appdata:himpresion["nhoja"]>0
            Printer:EndPage()
         endif

         Printer:StartPage()

         appdata:himpresion["nhoja"] += 1

         appdata:himpresion["nrenglon"]:= 0.5

         escribe({"nfuente"  => 16 ,  "ncolor" => dgs_color_base   ,  "ncolumna" => 0.50, "ctexto"   => appdata:hreporte["cnombre_empresa"]})

         linea_simple({"narriba" => 1.2               , "nizquierda" => 0.5, "nabajo" =>  1.2              , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         escribe({"ncolumna" => appdata:himpresion["ntope_derecha"]  ,  "nfuente" => 7  ,  "calineacion" => "der"  ,  "ctexto"   => "Hoja # "+ToString(appdata:himpresion["nhoja"])})

         appdata:himpresion["nrenglon"]:= 1.4

         escribe_dato({"ncolumna"      => 1.0 , ;
                       "ctexto"        => "REPORTE" , ;
                       "cdato"         => "RELACIONES ANALITICAS", ;
                       "nfuente_dato"  => 9, ;
                       "lnegrita_dato" => .f.}, appdata:himpresion)

         appdata:himpresion["nrenglon"] := 2.2

         escribe_dato({"ncolumna"      => 1.0     , ;
                       "ctexto"        => "FECHA" , ;
                       "cdato"         => "Del " + fectal(appdata:hreporte["cfecha_inicial"])+ " Al " + fectal(appdata:hreporte["cfecha_final"]), ;
                       "nfuente_dato"  => 9, ;
                       "lnegrita_dato" => .f.}, appdata:himpresion)

         appdata:himpresion["nrenglon"] := 2.9

         linea_simple({"narriba" => appdata:himpresion["nrenglon"], "nizquierda" => 0.5, "nabajo" => appdata:himpresion["nrenglon"], "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         escribe_encabezado(dgs_color_base)

         appdata:himpresion["nrenglon"] += appdata:himpresion["ninter"]

         linea_simple({"narriba" => appdata:himpresion["nrenglon"]-0.01 , "nizquierda" => 0.5, "nabajo" => appdata:himpresion["nrenglon"]-0.01 , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

         *=================================================================================================================================================
         * AJUSTE POR EL BRINCO QUE VENDRA AL IMPRIMIR LAS LINEAS DE DATOS
         *=================================================================================================================================================

         appdata:himpresion["nrenglon"] -= appdata:himpresion["ninter"]

         appdata:himpresion["ltitulos"]:=.f.

      endif

      *=================================================================================================================================================
      * PREPARACION DE LINEA DE DATOS A IMPRIMIR
      *=================================================================================================================================================

      hpaso["ccliente"]                    := iif(hrenglon["cliente"]=nil                 , "", hrenglon["cliente"])
      hpaso["cnombre"]                     := iif(hrenglon["nombre"]=nil                  , "", hrenglon["nombre"])
      hpaso["canterior"]                   := transform(hrenglon["saldo_anterior"]        , "9,999,999.99")
      hpaso["ccargo"]                      := transform(hrenglon["total_cargos"]          , "9,999,999.99")
      hpaso["cabono"]                      := transform(hrenglon["total_abonos"]          , "9,999,999.99")
      hpaso["cactual"]                     := transform(hrenglon["saldo_actual"]          , "9,999,999.99")

      appdata:himpresion["acolus"][1]["dato"]  := hpaso["ccliente"]
      appdata:himpresion["acolus"][2]["dato"]  := hpaso["cnombre"]
      appdata:himpresion["acolus"][3]["dato"]  := hpaso["canterior"]
      appdata:himpresion["acolus"][4]["dato"]  := hpaso["ccargo"]
      appdata:himpresion["acolus"][5]["dato"]  := hpaso["cabono"]
      appdata:himpresion["acolus"][6]["dato"]  := hpaso["cactual"]

      appdata:himpresion["nrenglon"] += appdata:himpresion["ninter"]

      * ====================================================================================================================
      * CAMPO CLASE
      * ====================================================================================================================
      * 06 - SALDOS CLIENTE
      * 14 - TOTAL GENERAL
      * ====================================================================================================================

      do case

         case appdata:hreporte["ctipo_tablero"]="4" .and. hrenglon["clase"]="14"

            *=================================================================================================================================================
            * TOTAL GENERAL OPCION TODOS LOS CLIENTES
            *=================================================================================================================================================

            appdata:himpresion["acolus"][1]["dato"]  := ""
            appdata:himpresion["acolus"][2]["dato"]  := "TOTAL GENERAL"

            linea_simple({"narriba" => appdata:himpresion["nrenglon"] , "nizquierda" => 0.5, "nabajo" =>  appdata:himpresion["nrenglon"] , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

            escribe_linea_datos({"lnegrita" => .t., "ncolor" => dgs_color_base, "nfuente" => 9, "lsin_raya" => .t., "lconfirma_pautado" => .f.})

            linea_simple({"narriba" => appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nizquierda" => 0.5, "nabajo" =>  appdata:himpresion["nrenglon"]+appdata:himpresion["ninter"]+0.1 , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_base, "ngrosor" => 1.2})

            appdata:himpresion["nrenglon"] += 0.1

         otherwise

            *=================================================================================================================================================
            * LINEA TIPICA
            *=================================================================================================================================================

            escribe_linea_datos({"nada" => ""})

      endcase

      * FORMATO
      *     1         2        3          4        5       6
      *|         |        | SALDO    |        |        | SALDO  |
      *| CLIENTE | NOMBRE | ANTERIOR | CARGOS | ABONOS | ACTUAL |

      *=================================================================================================================================================
      * DETECTAMOS SI YA REBASO EL TOPE DE ABAJO PARA PRENDER LA VARIABLE appdata:himpresion[LTITULOS"] PARA QUE EN LA SIGUIENTE VUELTA IMPRIMA TITULOS
      *=================================================================================================================================================

      if appdata:himpresion["nrenglon"] >= appdata:himpresion["ntope_brinco"]

         hpaso["nrenglon"] := appdata:himpresion["ntope_abajo"]-0.5

         escribe({"nrenglon" => hpaso["nrenglon"], "nfuente" => 9  ,  "calineacion" => "der"  ,  "ncolor" => dgs_color_base  ,  "ncolumna" => 21.0  ,  "ctexto" => "===> "+ToString(appdata:himpresion["nhoja"]+1)})

         appdata:himpresion["ltitulos"]:=.t.

      endif

   next

   *=================================================================================================================================================
   * QUE FALTA
   *=================================================================================================================================================
   *
   *=================================================================================================================================================

   Printer:EndPage()
   Printer:EndDoc()

   Printer:nPrinterIndex    := hpaso["noldindex"]

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Finalizando Documento...")

   *=================================================================================================================================================
   * REALIZAMOS LA COPIA DEL ARCHIVO A LA RUTA DONDE EL CGI LO ESPERA PARA DESCARGARLO
   *=================================================================================================================================================

   hpaso["cfile_generado"] := "C:\apache24\htdocs\pdf\"+appdata:hreporte["cnombre_entregable"]

   *=================================================================================================================================================
   * HAREMOS UN LOOP DE ESPERA PARA DAR OPORTUNIDAD DE QUE SE GRABE EL ARCHIVO FISICAMENTE, LA EJECUCION DE CODIGO ES MAS RAPIDO QUE LAS
   * ACCIONES FISICAS EN DISCO, EL TIEMPO DE ESPERA SERA DE 10 SEGUNDOS
   *=================================================================================================================================================

   hpaso["tiempo_inicial"]:= Seconds()
   hpaso["lse_pudo"]:= .f.

   do while .t.

      if File(hpaso["cfile_generado"])
         hpaso["lse_pudo"]:= .t.
         exit
      endif

      hpaso["tiempo_actual"] := Seconds()

      if hpaso["tiempo_actual"] - hpaso["tiempo_inicial"] >10
         exit
      endif

   enddo

   *=================================================================================================================================================
   * REVISAMOS SI SE PUDO
   *=================================================================================================================================================

   if !hpaso["lse_pudo"]

      *=================================================================================================================================================
      * POR ALGUNA RAZON NO SE PUDO DETECTAR EN 10 SEGUNDOS LA CREACION FISICA DEL ARCHIVO PDF, ABORTAMOS Y MANDAMOS SEҁL AL USUARIO/CGGI
      *=================================================================================================================================================

      inicializa_estructura_estado_peticion_api( @hestado )

      hestado["cEstado"]                   := "Cancelado"
      hestado["cDescripcion"]              := "Error dentro de API, excedio de 10 segundos la grabacion de archivo fisico pdf"
      hestado["cTituloTextoUsuario"]       := "El servidor no pudo grabar documento pdf (time exceeded)"
      hestado["cTextoUsuario"]             := "Reporte este mensaje a su asesor de sistemas"

      hpaso["cgraba"] := HB_JsonEncode( hestado )

      hb_memowrit( appdata:hreporte["carchivo_error"], hpaso["cgraba"] )

      appdata:hreporte["laborta_proceso"] := .f.

      quit

   endif

   *=================================================================================================================================================
   * MANDAMOS MENSAJE A USUARIO DE INICIO DE PREPARACION DE PETICION
   *=================================================================================================================================================

   mensaje_avance("Iniciando Envio de Documento...")

   *DELETE FILE ''+hpaso["cfile_generado"]+''

   *--------------------------------------------------------------------------------------------------------------------------------------
   * MENSAJE DE AVANCE PARA USUARIO
   *--------------------------------------------------------------------------------------------------------------------------------------

   mensaje_avance("Gestionando Entrega de Documento Solicitado")

   hpaso["cgraba"] := HB_JsonEncode(appdata:hestado )

   hb_memowrit( appdata:hreporte["carchivo_finalizacion"], hpaso["cgraba"] )

   *borra_temporales()

   LogFile(DToC(Date())+" "+Time()+"   ******** Finalizando y saliendo de funcion genera_pdf_cartera_clientes_analiticos   "+Str(seconds(),12,4))

RETURN NIL


//------------------------------------------------------------------------------

