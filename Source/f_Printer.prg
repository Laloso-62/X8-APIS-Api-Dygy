/*
 * Proyecto: Api-Rezta
 * Fichero: f_Printer.prg
 * Descripci�n:
 * Autor:
 * Fecha: 01/11/2022
 */

#include "Xailer.ch"
#include "Digio.ch"

//------------------------------------------------------------------------------

FUNCTION escribe ( hdato )

   LOCAL hpaso:= {=>}, nrencal, ncolcal, nesali, npix, aconve, nesto, neste

   *======================================================================================================================================
   * DEFINICION DE VALORES DEFAULT, CON ESTO PERMITIREMOS AHORRAR CODIGO A LAS FUNCIONES QUE UTILICEN ESTA FUNCION
   *======================================================================================================================================

   if !hb_HHasKey( hdato, "calineacion" )
      hdato["calineacion"] := "izq"
   endif

   if !hb_HHasKey( hdato, "lnegrita" )
      hdato["lnegrita"] := .f.
   endif

   if !hb_HHasKey( hdato, "nrenglon" )
      hdato["nrenglon"] := appdata:himpresion["nrenglon"]
   endif

   if !hb_HHasKey( hdato, "ncolor" )
      hdato["ncolor"] := dgs_color_negro
   endif

   if !hb_HHasKey( hdato, "nfuente" )
      hdato["nfuente"] := 8
   endif

   if !hb_HHasKey( hdato, "lwrap" )
      hdato["lwrap"] := .f.
   endif

   *--------------------------------------------------------------------------------------------------------------------------------------

   if hdato["ncolor"] = nil
      hdato["ncolor"] :=rgb( 000 , 000 , 000 )
   endif

   nrencal:=hdato["nrenglon"]*100
   ncolcal:=hdato["ncolumna"]*100

   if hdato["lwrap"]

      *======================================================================================================================================
      * EN CASO DE SER UN DATO WRAP REQUERIMOS EL DATO 'NANCHO' PARA PODER CALCULAR LA PARTICION DE LINEAS DEL TEXTO A IMPRIMIR, EN CASO DE
      * NO CONTAR CON ESTE DATO PONDREMOS UN ANCHO DE 5, SOLO PARA EVITAR UN ERROR
      *======================================================================================================================================

      if !hb_HHasKey( hdato, "nancho" )
         hdato["nancho"] := 5
      endif

      *======================================================================================================================================
      * EN CASO DE SER UN DATO WRAP REQUERIMOS EL DATO 'LMUEVE_RENGLON' PARA ACTUALIZAR EL DATO DE RENGLON QUE SIGUE, EN CASO DE NO CONTAR
      * CON ESTE DATO PONDREMOS COMO VALOR DEFAULT .F.
      * NOTA: PARA QUE ESTO FUCIONES ES NECESARIO QUE LA FUNCION SCRIBE RECIBA appdata:himpresion (@HPASO) CON ARROBA PARA QUE SE ACTUALICE BIEN
      *======================================================================================================================================

      if !hb_HHasKey( hdato, "lmueve_renglon" )
         hdato["lmueve_renglon"] := .f.
      endif

      Printer:oCanvas:oFont:nSize := hdato["nfuente"]
      Printer:oCanvas:oFont:lbold := .f.

      hpaso["ancho"]:=Printer:oCanvas:textwidth(hdato["ctexto"])

      hpaso["centi"]:=Printer:oCanvas:PixelsToMapMode(hpaso["ancho"], 0)
      hpaso["centi"][1]:=hpaso["centi"][1]/100

      hpaso["nwidth"]:=hdato["nancho"]

      *=================================================================================================
      * FORZOSAMENTE TENEMOS QUE PARTIR EL TEXTO PARA IMPRIMIR EN VARIOS RENGLONES
      *=================================================================================================

      hpaso["alineas"]:=parte_texto(hdato["ctexto"], hdato["nancho"])

      hpaso["nrenglon"]:= appdata:himpresion["nrenglon"]-appdata:himpresion["ninter"]

      *LogFile(DToC(Date())+" "+Time()+"   ******** imprimir dato partido texto = " + hcolumna["dato"]+"                 "+Str(seconds(),12,4))

      for neste:=1 to Len(hpaso["alineas"])

         hpaso["nrenglon"] += appdata:himpresion["ninter"] + appdata:himpresion["nmargen_superior"]

         *LogFile(DToC(Date())+" "+Time()+"        *** imprimiendo parcial partido renglon "+ToString(neste)+" hpaso[nrenglon] = " + Str(hpaso["nrenglon"], 14,6)+"                 "+Str(seconds(),12,4))

         nrencal:=hpaso["nrenglon"]*100

         WITH OBJECT Printer:oCanvas
            :lTransparent  := .t.
            :TextOut( ncolcal, nrencal, hpaso["alineas"][neste,01] , , hdato["ncolor"])
         END WITH

      next neste

      if hdato["lmueve_renglon"]
         appdata:himpresion["nrenglon"]:=hpaso["nrenglon"]
      endif

   endif

   if !hdato["lwrap"]

      if hdato["calineacion"]=="der" .or. hdato["calineacion"]=="Der"

         WITH OBJECT Printer:oCanvas
            :oFont:nSize := hdato["nfuente"]
            :oFont:lbold := hdato["lnegrita"]
            :TextOut( 1, 1, "." , , hdato["ncolor"])
         END WITH

         npix:=Printer:oCanvas:textwidth(hdato["ctexto"])
         aconve:=Printer:oCanvas:PixelsToMapMode(0,npix)
         nesto:=aconve[2]/100

         ncolcal:=hdato["ncolumna"]-nesto

         ncolcal:=ncolcal*100

      endif

      WITH OBJECT Printer:oCanvas
         :oFont:nSize := hdato["nfuente"]
         :oFont:lbold := hdato["lnegrita"]
         :lTransparent:= .t.
         :TextOut( ncolcal, nrencal, hdato["ctexto"] , , hdato["ncolor"])
      END WITH

   endif

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION escribe_dato ( hdato )

   LOCAL hpaso:= {=>}

   *--------------------------------------------------------------------------------------------------------------------------------------
   * DEFINICION DE VALORES DEFAULT, CON ESTO PERMITIREMOS AHORRAR CODIGO A LAS FUNCIONES QUE UTILICEN ESTA FUNCION
   *--------------------------------------------------------------------------------------------------------------------------------------

   if !hb_HHasKey( hdato, "nrenglon" )
      hdato["nrenglon"] := appdata:himpresion["nrenglon"]
   endif

   if !hb_HHasKey( hdato, "ninter" )
      hdato["ninter"] := 0.24
   endif

   if !hb_HHasKey( hdato, "nfuente_texto" )
      hdato["nfuente_texto"] := 6
   endif

   if !hb_HHasKey( hdato, "nfuente_dato" )
      hdato["nfuente_dato"] := 10
   endif

   if !hb_HHasKey( hdato, "calineacion" )
      hdato["calineacion"] := "izq"
   endif

   if !hb_HHasKey( hdato, "ncolor_texto" )
      hdato["ncolor_texto"] := dgs_color_gris
   endif

   if !hb_HHasKey( hdato, "ncolor_dato" )
      hdato["ncolor_dato"] := dgs_color_negro
   endif

   if !hb_HHasKey( hdato, "lnegrita_dato" )
      hdato["lnegrita_dato"] := .f.
   endif

   escribe({"nfuente"  => hdato["nfuente_texto"] ,  "calineacion" => hdato["calineacion"]  ,  "lnegrita" => .f.  ,                 ;
            "ncolor"   => hdato["ncolor_texto"]  ,  "nrenglon"    => hdato["nrenglon"]     ,  "ncolumna" => hdato["ncolumna"],     ;
            "ctexto"   => hdato["ctexto"]})

   hdato["nrenglon"] += hdato["ninter"]

   escribe({"nfuente"  => hdato["nfuente_dato"]  ,  "calineacion" => hdato["calineacion"]  ,  "lnegrita" => hdato["lnegrita_dato"] ,  ;
            "ncolor"   => hdato["ncolor_dato"]   ,  "nrenglon"    => hdato["nrenglon"]     ,  "ncolumna" => hdato["ncolumna"],        ;
            "ctexto"   => hdato["cdato"]})

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION escribe_encabezado(ncolor)

   LOCAL hpaso:= {=>}, hcolumna:= {=>}, neste:= 0

   *=================================================================================================================================================
   * PRIMERO DETECAMOS SI ALGUN ENCABEZADO OCUPARA SER DIVIDIDO PORQUE SU AREA DE IMPRESION NO ALCANCE A DESPLEGAR SU TEXTO COMPLETO
   * AGREGAREMOS UN DATO MAS EN EL HASH HCOLUS, EL CUAL SERVIRA SOLO PARA ESTA RUTINA, EN CASO DE QUE NO
   *=================================================================================================================================================

   appdata:himpresion["avance_maximo"] := appdata:himpresion["nrenglon"]
   appdata:himpresion["media_wrap"]    := appdata:himpresion["nrenglon"]

   hpaso["lwrapeado"] := .f.

   for each hcolumna in appdata:himpresion["acolus"]

      hcolumna["lya"] := .f.

      Printer:oCanvas:oFont:nSize := hcolumna["fuente_titulo"]
      Printer:oCanvas:oFont:lbold := .t.

      hpaso["ancho"]:=Printer:oCanvas:textwidth(hcolumna["titulo"])

      hpaso["centi"]:=Printer:oCanvas:PixelsToMapMode(hpaso["ancho"], 0)
      hpaso["centi"][1]:=hpaso["centi"][1]/100

      hpaso["nwidth"]:=hcolumna["area"]["der"]-hcolumna["area"]["izq"]

      if hpaso["nwidth"] >= hpaso["centi"][1]
         loop
      endif

      *=================================================================================================
      * FORZOSAMENTE TENEMOS QUE PARTIR EL TEXTO PARA IMPRIMIR EN VARIOS RENGLONES
      *=================================================================================================

      hpaso["lwrapeado"] := .t.

      hcolumna["lya"] := .t.

      hpaso["alineas"]:=parte_texto(hcolumna["titulo"], hpaso["nwidth"])

      hpaso["nrenglon"]:= appdata:himpresion["nrenglon"]-appdata:himpresion["ninter"]

      *LogFile(DToC(Date())+" "+Time()+"   ******** imprimir dato partido texto = " + hcolumna["dato"]+"                 "+Str(seconds(),12,4))

      for neste:=1 to Len(hpaso["alineas"])

         hpaso["nrenglon"] += appdata:himpresion["ninter"] + appdata:himpresion["nmargen_superior"]

         *LogFile(DToC(Date())+" "+Time()+"        *** imprimiendo parcial partido renglon "+ToString(neste)+" hpaso[nrenglon] = " + Str(hpaso["nrenglon"], 14,6)+"                 "+Str(seconds(),12,4))

         hpaso["ncolumna"] := IIF(hcolumna["alineacion"]="izq", hcolumna["area"]["izq"], hcolumna["area"]["der"])

         escribe({"nfuente"  => hcolumna["fuente_titulo"] ,  "calineacion" => hcolumna["alineacion"]  ,  "lnegrita" => .t.              ,     ;
                  "ncolor"   => ncolor                    ,  "nrenglon"    => hpaso["nrenglon"]       ,  "ncolumna" => hpaso["ncolumna"],     ;
                  "ctexto"   => hpaso["alineas"][neste,01]})

         *=================================================================================================================================================================
         * REGISTRAMOS EN appdata:himpresion[AVANCE_MAXIMO] EL RENGLON CON VALOR MAS ALTO PARA SABER HASTA DONDE LLEGO LA IMPRESION DE LAS LINEAS
         *=================================================================================================================================================================

         appdata:himpresion["avance_maximo"] := IIF(appdata:himpresion["avance_maximo"]<hpaso["nrenglon"], hpaso["nrenglon"], appdata:himpresion["avance_maximo"])

      next neste

      *LogFile(DToC(Date())+" "+Time()+"        *** appdata:himpresion[nrenglon] = " + Str(appdata:himpresion["nrenglon"], 14,6)+"                 "+Str(seconds(),12,4))
      *LogFile(DToC(Date())+" "+Time()+"        *** hpaso[nrenglon] = " + Str(hpaso["nrenglon"], 14,6)+"                 "+Str(seconds(),12,4))

      hpaso["media_wrap_columna"]:=((hpaso["nrenglon"] - appdata:himpresion["nrenglon"])/2)+appdata:himpresion["nrenglon"]

      if hpaso["media_wrap_columna"]>appdata:himpresion["media_wrap"]
         appdata:himpresion["media_wrap"]:=hpaso["media_wrap_columna"]
         hpaso["wrapeado"]:= .t.
      endif

      appdata:himpresion["avance_maximo"] := IIF(appdata:himpresion["avance_maximo"]<hpaso["nrenglon"], hpaso["nrenglon"], appdata:himpresion["avance_maximo"])

      *LogFile(DToC(Date())+" "+Time()+"        *** appdata:himpresion[media_wrap] = " + Str(appdata:himpresion["media_wrap"], 14,6)+"                 "+Str(seconds(),12,4))

   next

   *=================================================================================================================================================================
   * EN CASO DE QUE appdata:himpresion[MEDIA_WRAP] CONTENGA UN VALOR SUPERIOR A CERO, ACTUALIZAREMOS EN appdata:himpresion[NRENGLON] CON EL DATO DE appdata:himpresion[MEDIA_WRAP] PARA LAS
   * IMPRESIONES DEL SIGUIENTE FOR-NEXT, EL RENGLON DE appdata:himpresion[NRENGLON] SERA ACTUALIZADO HASTA EL FINAL DE ESTA FUNCION
   *=================================================================================================================================================================

   hpaso["nrenglon"]:=appdata:himpresion["nrenglon"]

   if appdata:himpresion["media_wrap"] >0

      hpaso["nrenglon"]:=appdata:himpresion["media_wrap"] + appdata:himpresion["nmargen_superior"]

   endif

   for each hcolumna in appdata:himpresion["acolus"]

      if hcolumna["lya"]
         loop
      endif

      *LogFile(DToC(Date())+" "+Time()+"        *** al imprimir columna " + hcolumna["titulo"] + "  hpaso[nrenglon] =  " + Str(hpaso["nrenglon"], 14,6)+"                 "+Str(seconds(),12,4))

      hpaso["ncolumna"] := IIF(hcolumna["alineacion"]="izq", hcolumna["area"]["izq"], hcolumna["area"]["der"])

      escribe({"nfuente"  => hcolumna["fuente_titulo"] ,  "calineacion" => hcolumna["alineacion"]  ,  "lnegrita" => .t.              ,     ;
               "ncolor"   => ncolor                    ,  "nrenglon"    => hpaso["nrenglon"]       ,  "ncolumna" => hpaso["ncolumna"],     ;
               "ctexto"   => hcolumna["titulo"]})

      *=================================================================================================================================================================
      * REGISTRAMOS EN appdata:himpresion[AVANCE_MAXIMO] EL RENGLON CON VALOR MAS ALTO PARA SABER HASTA DONDE LLEGO LA IMPRESION DE LAS LINEAS
      *=================================================================================================================================================================

      appdata:himpresion["avance_maximo"] := IIF(appdata:himpresion["avance_maximo"]<hpaso["nrenglon"], hpaso["nrenglon"], appdata:himpresion["avance_maximo"])

   next

   if hpaso["lwrapeado"]

      appdata:himpresion["nrenglon"] := appdata:himpresion["avance_maximo"]

   endif

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION configura_media()

   LOCAL hpaso:= {=>}, hcolumna:= {=>}

   *=================================================================================================================================================================
   * ESTA FUNCION TIENE EL OBJETIVO DE ENCONTRAR EL PUNTO MEDIO DE IMPRESION CUANDO EL TAMAÑO DE LAS FUENTES NO ES LA MISMA, ASI EVITAREMOS QUE ALGUNOS TEXTOS QUE
   * COMPARTEN LA MISMA LINEA DE DATOS SE VEAN DESPROPORCIONADOS, BUSCAREMOS UNA ALTURA ADECUADA PARA QUE TODOS LOS TEXTOS COINCIDAN EN UN PUNTO MEDIO
   * A CADA COLUMNA O ELEMENTO DE appdata:himpresion[appdata:himpresion["acolus"]] SE LE AGREGARA LA VARIABLE HCOLUMNA[AJUSTE_ALTO] PARA SU USO EN IMPRESIONES QUE VENDRAN MAS ADELANTE
   *=================================================================================================================================================================

   *=================================================================================================================================================================
   * PRIMERO BUSCAMOS LA FUENTE MAYOR, ESTA FUENTE SENTARA LA BASE DEL AREA VERTICAL MAXIMA QUE SE REQUIERE
   *=================================================================================================================================================================

   hpaso["nmax"]:=0


   for each hcolumna in appdata:himpresion["acolus"]

      hpaso["nmax"]:=IIF(hcolumna["fuente_dato"] > hpaso["nmax"], hcolumna["fuente_dato"], hpaso["nmax"])

   next

   *LogFile(DToC(Date())+" "+Time()+"   ******** en ajuste_media  fuenta mayor = " + Str(hpaso["nmax"],16,6)+"                 "+Str(seconds(),12,4))

   *=================================================================================================================================================================
   * AHORA OBTENEMOS LA MEDIDA DE LA FUENTE MAYOR PARA TOMARLA COMO BASE
   *=================================================================================================================================================================

   Printer:oCanvas:oFont:nSize := hpaso["nmax"]
   Printer:oCanvas:oFont:lbold := .f.

   hpaso["alto"]:=Printer:oCanvas:TextHeight("EJEMPLO")

   *LogFile(DToC(Date())+" "+Time()+"        *** alto detectado = " + Str(hpaso["alto"],16,6)+"                 "+Str(seconds(),12,4))

   hpaso["centi"]:=Printer:oCanvas:PixelsToMapMode(0, hpaso["alto"])
   hpaso["centi"][2]:=hpaso["centi"][2]/100

   appdata:himpresion["alto_base"] := hpaso["centi"][2]

   *LogFile(DToC(Date())+" "+Time()+"        *** alto en mm = " + Str(appdata:himpresion["alto_base"],16,6)+"                 "+Str(seconds(),12,4))

   *=================================================================================================================================================================
   * CON EL DATO DEL appdata:himpresion[ALTO_BASE] HAREMOS UN CALCULO COLUMNA POR COLUMNA PARA DETERMINAR LA ALTURA IDEAL DONDE DEBE INICIAR LA IMPRESION PARA LOGRAR EL
   * PUNTO MEDIO
   *=================================================================================================================================================================

   for each hcolumna in appdata:himpresion["acolus"]

      *LogFile(DToC(Date())+" "+Time()+"        ****** ajustando columna " + hcolumna["titulo"]+"                 "+Str(seconds(),12,4))

      Printer:oCanvas:oFont:nSize := hcolumna["fuente_dato"]
      Printer:oCanvas:oFont:lbold := .f.

      *=================================================================================================================================================================
      * CALCULAMOS EL ALTO O MEDIDA VERTICAL SEGUN LA FUENTE DE LA COLUMNA EN CUESTION
      *=================================================================================================================================================================

      hpaso["alto"]:=Printer:oCanvas:TextHeight("EJEMPLO")

      *LogFile(DToC(Date())+" "+Time()+"        ****** alto detectado " + Str(hpaso["alto"],16,6)+"                 "+Str(seconds(),12,4))

      hpaso["centi"]:=Printer:oCanvas:PixelsToMapMode(0, hpaso["alto"])
      hpaso["centi"][2]:=hpaso["centi"][2]/100

      *LogFile(DToC(Date())+" "+Time()+"        ****** convertido en mm " + Str(hpaso["centi"][2],16,6)+"                 "+Str(seconds(),12,4))

      *=================================================================================================================================================================
      * CALCULAMOS LA DIFERENCIA ENTRE LA ALTURA MAYOR Y LA COLUMNA QUE ESTAMOS RECORRIENDO PARA DIVIDIRLA ENTRE DOS Y ASI OBTENER LA COMPENSACION PARA INICIO DE ALTURA
      *=================================================================================================================================================================

      hcolumna["ajuste_alto"]:=(appdata:himpresion["alto_base"]-hpaso["centi"][2])/2

      *LogFile(DToC(Date())+" "+Time()+"        ****** ajuste determinado " + Str(hcolumna["ajuste_alto"],16,6)+"                 "+Str(seconds(),12,4))

   next

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION escribe_linea_datos(hespeciales)

   LOCAL hpaso:= {=>}, hcolumna:= {=>}, neste

   *=================================================================================================================================================================
   * VARIABLES GENERALES QUE PUEDEN VENIR EN HESPECIALES
   *=================================================================================================================================================================

   if !hb_HHasKey( hespeciales, "lconfirma_pautado" )
      hespeciales["lconfirma_pautado"] := .t.
   endif

   *=================================================================================================================================================================
   * VARIABLES PARA CONTROLAR DATOS QUE NECESITEN VARIAS LINEAS PARA DESPLEGARSE EN EL AREA DESIGNADA PARA SU IMPRESION
   *=================================================================================================================================================================

   appdata:himpresion["media_wrap"]    := 0
   appdata:himpresion["avance_maximo"] := appdata:himpresion["nrenglon"]
   hpaso["nrenglon_recibido"]  := appdata:himpresion["nrenglon"]
   hpaso["lwrapeado"]          := .f.

   *=================================================================================================================================================================
   * VARIABLES PARA CONTROLAR EL PAUTADO CUANDO ESTE ES CONFIGURADO EN LA VARIABLE appdata:himpresion[LPAUTADO]
   *=================================================================================================================================================================

   appdata:himpresion["nalterno"] := IIF( appdata:himpresion["nalterno"] = 1, 0, 1)

   hpaso["nmax_pautado"]:= 0

   *=================================================================================================================================================================
   * VARIABLES QUE PUEDEN MODIFICAR VALORES COMO COLOR, FUENTE, NEGRITA EN LA IMPRESION, SE UTILIZA SOBRE TODO PARA CUANDO QUERAMOS LINEAS DE DATOS DE TOTALES
   *=================================================================================================================================================================

   hpaso["lnegrita"]:=IIF(hb_HHasKey( hespeciales, "lnegrita" ), hespeciales["lnegrita"], .f.)
   hpaso["ncolor"]  :=IIF(hb_HHasKey( hespeciales, "ncolor"   ), hespeciales["ncolor"]  , appdata:himpresion["color"])
   hpaso["nfuente"] :=IIF(hb_HHasKey( hespeciales, "nfuente"  ), hespeciales["nfuente"] , 0)

   *=================================================================================================================================================================
   * PRIMER RECORRIDO BUSCANDO DATOS QUE EXCEDEN EL ANCHO ASIGNADO PARA IMPRESION, LOS DATOS QUE NO OCUPAN PARTIR EN LINEAS SU IMPRESION SE RTATARA EN OTRO FOR-NEXT
   *=================================================================================================================================================================

   for each hcolumna in appdata:himpresion["acolus"]

      hcolumna["lya"] := .f.

      *=================================================================================================================================================================
      * SI LA COLUMNA NO ESTA CONFIGURADA PARA DIVIDIR SU IMPRESION NO ES TOMADA EN CUENTA
      *=================================================================================================================================================================

      if !hcolumna["lwrap"]
         loop
      endif

      *=================================================================================================================================================================
      * DETERMINAMOS SI EL DATO CABE EN SU AREA DE IMPRESION, SI DETECTAMOS QUE SI ES POSIBLE SU IMPRESION SALTAMOS A LA SIGUINTE COLUMNA
      *=================================================================================================================================================================

      Printer:oCanvas:oFont:nSize := hcolumna["fuente_dato"]
      Printer:oCanvas:oFont:lbold := .f.

      hpaso["ancho"]:=Printer:oCanvas:textwidth(hcolumna["dato"])

      hpaso["centi"]:=Printer:oCanvas:PixelsToMapMode(hpaso["ancho"], 0)
      hpaso["centi"][1]:=hpaso["centi"][1]/100

      hpaso["nwidth"]:=hcolumna["area"]["der"]-hcolumna["area"]["izq"]

      if hpaso["nwidth"] >= hpaso["centi"][1]
         loop
      endif

      *=================================================================================================================================================================
      * COMO EL DATO SI NECESITA SER PARTIDO EN VARIAS LINEAS INICIAMOS SU TRATAMIENTO, MARCAMOS EN HPASO[LWRAPEADO] CON .T. PARA SABER QUE CUANDO MENOS UNA COLUMNA
      * REQUIRIO SER WRAPEADA, SE CONSERVARAN OTROS DATOS MAS ADELANTE COMO EL AVANCE MAXIMO, Y EL TOPE DEL AREA PAUTADA PARA SU USO EN EL FOR-NEXT DE IMPRESION DE
      * LAS COLUMNAS QUE NO REQUIEREN EL WRAPEDADO
      *=================================================================================================================================================================

      hpaso["lwrapeado"] := .t.

      hcolumna["lya"]    := .t.

      hpaso["alineas"]   :=parte_texto(hcolumna["dato"], hpaso["nwidth"])

      hpaso["nrenglon"]  := appdata:himpresion["nrenglon"]-appdata:himpresion["ninter"]

      *=================================================================================================================================================================
      * HPASO[ALINEAS] CONTIENE EL DATO A IMPRIMIR PARTIDO EN TROZOS QUE CABEN EN EL AREA DE IMPRESION DESIGNADA PARA LA COLUMNA
      *=================================================================================================================================================================

      for neste:=1 to Len(hpaso["alineas"])

         *=================================================================================================================================================================
         * AVANZAMOS UN RENGLON EN HPASO[NRENGLON] CONSIDERANDO EL DATO appdata:himpresion[NMARGEN_SUPERIOR]
         *=================================================================================================================================================================

         hpaso["nrenglon"] += appdata:himpresion["ninter"]

         hpaso["ncolumna"] := IIF(hcolumna["alineacion"]="izq", hcolumna["area"]["izq"], hcolumna["area"]["der"])

         *=================================================================================================================================================================
         * SI appdata:himpresion[LPAUTADO] ESTA ACTIVO INTENTAMOS PAUTAR ANTES DE CUALQUIER IMPRESION DE TEXTOS, GUARDAMOS EL AVANCE MAXIMO DE PAUTADO PARA CONSIDERARLO EN OTRAS
         * COLUMNAS DE ESTE FOR-NEXT ASI COMO EN EL SIGUIENTE FOR-NEXT DE LAS COLUMNAS QUE NO REQUIEREN SER WRAPEADAS O PARTIDAS
         *=================================================================================================================================================================

         if appdata:himpresion["lpautado"] .and. appdata:himpresion["nalterno"] = 1 .and. hespeciales["lconfirma_pautado"]

            if hpaso["nmax_pautado"] < hpaso["nrenglon"] - appdata:himpresion["nmargen_superior"]

               *=================================================================================================================================================================
               * HASTA AQUI DETECTAMOS QUE EL PAUTADO VA, CALCULAMOS EL TOPE DE RENGLON ABAJO HPASO[NTOPE_BAJO] PARA MANDAR LA FUNCION CUADRO_LLENO, Y GUARDAMOS EN LA VARIABLE
               * HPASO[NMAX_PAUTADO] EL AREA DE PAUTADO YA CUBIERTA PARA USOS FUTUROS
               *=================================================================================================================================================================

               hpaso["ntope_alto"]  := hpaso["nrenglon"]
               hpaso["ntope_abajo"] := hpaso["nrenglon"] +appdata:himpresion["ninter"]

               cuadro_lleno({"narriba" => hpaso["ntope_alto"], "nizquierda" => 0.5, "nabajo" =>  hpaso["ntope_abajo"]  , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => appdata:himpresion["ncolor_pautado_alterno"]})

               hpaso["nmax_pautado"]:= IIF(hpaso["nmax_pautado"] < hpaso["nrenglon"] - appdata:himpresion["nmargen_superior"], hpaso["nrenglon"] - appdata:himpresion["nmargen_superior"], hpaso["nmax_pautado"])

            endif

         endif

         *=================================================================================================================================================================
         * REVISAMOS SI RECIBIMOS EN HESPECIALES CAMBIOS EN FUENTE, COLOR Y NEGRITA PARA QUE SEAN TOMADAS EN CUENTA EN VARIABLES DENTRO DE HPASO, RECSUELTAS ALGUNAS AL
         * INICIO DE ESTA FUNCION
         *=================================================================================================================================================================

         hpaso["nsize"]:=IIF(hpaso["nfuente"] = 0, hcolumna["fuente_dato"], hpaso["nfuente"])

         escribe({"nfuente"  => hpaso["nsize"]    ,  "calineacion" => hcolumna["alineacion"]  ,  "lnegrita" => hpaso["lnegrita"],                 ;
                  "ncolor"   => hpaso["ncolor"]   ,  "nrenglon"    => hpaso["nrenglon"] + appdata:himpresion["nmargen_superior"]+hcolumna["ajuste_alto"], ;
                  "ncolumna" => hpaso["ncolumna"] ,  "ctexto"      => hpaso["alineas"][neste,01]})

      next neste

      *=================================================================================================================================================================
      * CALCULAMOS EL DATO HPASO[MEDIA_WRAP_COLUMNA] PARA QUE LOS DATOS DE LAS OTRAS COLUMNAS NO SUJETAS A WRAPEADO SE IMPRIMAN EN LA PARTE MEDIA DEL AREA UTILIZADA
      * PARA LA IMPRESION DE VARIAS LINEAS, COMPARAMOS CON LA VARIABLE appdata:himpresion[MEDIA_WRAP] PARA EN SU CASO ESTA SEA ACTUALIZADA, YA QUE ESTA VARIABLE ES LA QUE
      * UTILIZARAN TODAS LAS DEMAS COLUMNAS EN EL SIGUIENTE FOR-NEXT
      *=================================================================================================================================================================

      hpaso["media_wrap_columna"]:=((hpaso["nrenglon"] - appdata:himpresion["nrenglon"])/2)+appdata:himpresion["nrenglon"]

      if hpaso["media_wrap_columna"]>appdata:himpresion["media_wrap"]
         appdata:himpresion["media_wrap"]:=hpaso["media_wrap_columna"]
         hpaso["ajuste_alto"] := hcolumna["ajuste_alto"]
      endif

      *=================================================================================================================================================================
      * REGISTRAMOS EN appdata:himpresion[AVANCE_MAXIMO] EL RENGLON CON VALOR MAS ALTO PARA SABER HASTA DONDE LLEGO LA IMPRESION DE LAS LINEAS
      *=================================================================================================================================================================

      appdata:himpresion["avance_maximo"] := IIF(appdata:himpresion["avance_maximo"]<hpaso["nrenglon"], hpaso["nrenglon"], appdata:himpresion["avance_maximo"])

   next

   *=================================================================================================================================================================
   * EN CASO DE QUE appdata:himpresion[MEDIA_WRAP] CONTENGA UN VALOR SUPERIOR A CERO, ACTUALIZAREMOS EN appdata:himpresion[NRENGLON] CON EL DATO DE appdata:himpresion[MEDIA_WRAP] PARA LAS
   * IMPRESIONES DEL SIGUIENTE FOR-NEXT, EL RENGLON DE appdata:himpresion[NRENGLON] SERA ACTUALIZADO HASTA EL FINAL DE ESTA FUNCION
   *=================================================================================================================================================================

   hpaso["nrenglon"]:=appdata:himpresion["nrenglon"]

   if appdata:himpresion["media_wrap"] >0

      hpaso["nrenglon"]:=appdata:himpresion["media_wrap"] + appdata:himpresion["nmargen_superior"]

   endif

   *=================================================================================================================================================================
   * INICIAMOS CON LA IMPRESION DE LAS COLUMNAS QUE NO TIENEN NECESIDAD DE SER WRAPEADAS PERO QUE SE IMPRIMIRAN EN LA PARTE MEDIA DEL AREA QUE UTILIZARON LAS
   * COLUMNAS WRAPEADAS
   *=================================================================================================================================================================

   for each hcolumna in appdata:himpresion["acolus"]

      *=================================================================================================================================================================
      * BRINCAMOS LAS COLUMNAS YA WRAPEDAS E IMPRESAS EN EL FOR-NEXT PREVIO
      *=================================================================================================================================================================

      if hcolumna["lwrap"] .and. hcolumna["lya"]
         loop
      endif

      Printer:oCanvas:oFont:nSize := hcolumna["fuente_titulo"]
      Printer:oCanvas:oFont:lbold := .t.

      hpaso["alto"]:=Printer:oCanvas:TextHeight(hcolumna["titulo"])

      hpaso["centi"]:=Printer:oCanvas:PixelsToMapMode(0, hpaso["alto"])
      hpaso["centi"][2]:=hpaso["centi"][2]/100

      hpaso["media"]:= hpaso["centi"][2] /2

      *=================================================================================================================================================================
      * SI LE TOCA RENGLON PAUTADO REVISAMOS SI TODAVIA NO SE A PUESTO EL FONDO
      *=================================================================================================================================================================

      if appdata:himpresion["lpautado"] .and. appdata:himpresion["nalterno"] = 1 .and. hespeciales["lconfirma_pautado"]

         if hpaso["nmax_pautado"] < hpaso["nrenglon"] - appdata:himpresion["nmargen_superior"] - hcolumna["ajuste_alto"]

            *=================================================================================================================================================================
            * CALCULAMOS EL AREA DEL PAUTADO, CONSIDERANDO QUITAR appdata:himpresion[NMARGEN_SUPERIOR] Y HCOLUMNA[AJUSTE_ALTO]
            *=================================================================================================================================================================

            hpaso["ntope_alto"]  := hpaso["nrenglon"]
            hpaso["ntope_abajo"] := hpaso["nrenglon"] + appdata:himpresion["ninter"]

            cuadro_lleno({"narriba" => hpaso["ntope_alto"], "nizquierda" => 0.5, "nabajo" =>  hpaso["ntope_abajo"]  , "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => appdata:himpresion["ncolor_pautado_alterno"]})

            hpaso["nmax_pautado"]:= IIF(hpaso["nmax_pautado"] < hpaso["nrenglon"] - appdata:himpresion["nmargen_superior"], hpaso["nrenglon"] - appdata:himpresion["nmargen_superior"], hpaso["nmax_pautado"])

         endif

      endif

      *=================================================================================================================================================================
      * REVISAMOS SI RECIBIMOS EN HESPECIALES CAMBIOS EN FUENTE, COLOR Y NEGRITA PARA QUE SEAN TOMADAS EN CUENTA EN VARIABLES DENTRO DE HPASO, RECSUELTAS ALGUNAS AL
      * INICIO DE ESTA FUNCION
      *=================================================================================================================================================================

      hpaso["ncolumna"] := IIF(hcolumna["alineacion"]="izq", hcolumna["area"]["izq"], hcolumna["area"]["der"])

      hpaso["nsize"]:=IIF(hpaso["nfuente"] = 0, hcolumna["fuente_dato"], hpaso["nfuente"])

      escribe({"nfuente"  => hpaso["nsize"]     ,  "calineacion" => hcolumna["alineacion"]  ,  "lnegrita" => hpaso["lnegrita"],                 ;
               "ncolor"   => hpaso["ncolor"]    ,  "nrenglon"    => hpaso["nrenglon"] + appdata:himpresion["nmargen_superior"]+hcolumna["ajuste_alto"], ;
               "ncolumna" => hpaso["ncolumna"]  ,  "ctexto"   => hcolumna["dato"]})

   next

   *=================================================================================================================================================================
   * RESOLVEMOS EL RAYADO
   *=================================================================================================================================================================

   hpaso["ntope_alto"]  := appdata:himpresion["nrenglon"]
   hpaso["ntope_abajo"] := appdata:himpresion["avance_maximo"]+appdata:himpresion["ninter"]

   if !hb_HHasKey( hespeciales, "lsin_raya")

      if appdata:himpresion["rayado"]="raya_abajo"

         linea_simple({"narriba" => hpaso["ntope_abajo"], "nizquierda" => 0.5, "nabajo" => hpaso["ntope_abajo"], "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_gris_bajo, "ngrosor" => 0.3})

      endif

      if appdata:himpresion["rayado"]="cuadro_externo"

         linea_simple({"narriba" => hpaso["ntope_alto"], "nizquierda" => 0.5, "nabajo" => hpaso["ntope_alto"], "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_gris, "ngrosor" => 0.3})

         linea_simple({"narriba" => hpaso["ntope_alto"], "nizquierda" => 0.5, "nabajo" => hpaso["ntope_abajo"], "nderecha" => 0.5                        , "ncolor" => dgs_color_gris, "ngrosor" => 0.3})

         linea_simple({"narriba" => hpaso["ntope_alto"], "nizquierda" => appdata:himpresion["ntope_derecha"], "nabajo" => hpaso["ntope_abajo"], "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_gris, "ngrosor" => 0.3})

      endif

      if appdata:himpresion["rayado"]="cuadro_total"

         linea_simple({"narriba" => hpaso["ntope_abajo"], "nizquierda" => 0.5, "nabajo" => hpaso["ntope_abajo"], "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_gris, "ngrosor" => 0.3})

         linea_simple({"narriba" => hpaso["ntope_alto"] , "nizquierda" => 0.5, "nabajo" => hpaso["ntope_abajo"], "nderecha" => 0.5  , "ncolor" => dgs_color_gris, "ngrosor" => 0.3})

         linea_simple({"narriba" => hpaso["ntope_alto"] , "nizquierda" => appdata:himpresion["ntope_derecha"], "nabajo" => hpaso["ntope_abajo"], "nderecha" => appdata:himpresion["ntope_derecha"], "ncolor" => dgs_color_gris, "ngrosor" => 0.3})

         for each hcolumna in appdata:himpresion["acolus"]

            if hcolumna["alineacion"] = "izq" .and. hcolumna:__enumindex = 1
               loop
            endif

            if hcolumna["alineacion"] = "der" .and. hcolumna:__enumindex = Len(appdata:himpresion["acolus"])
               loop
            endif

            if hcolumna["alineacion"] = "izq"
               linea_simple({"narriba" => hpaso["ntope_alto"]      , "nizquierda" => hcolumna["area"]["izq"]-0.05 , "nabajo" => hpaso["ntope_abajo"], "nderecha" => hcolumna["area"]["izq"]-0.05 , "ncolor" => dgs_color_gris, "ngrosor" => 0.3})
            endif

            if hcolumna["alineacion"] = "der"

               if hcolumna:__enumindex > 1 .and. appdata:himpresion["acolus"][hcolumna:__enumindex-1]["alineacion"] = "izq"
                  linea_simple({"narriba" => hpaso["ntope_alto"]      , "nizquierda" => hcolumna["area"]["izq"]      , "nabajo" => hpaso["ntope_abajo"], "nderecha" => hcolumna["area"]["izq"]      , "ncolor" => dgs_color_gris, "ngrosor" => 0.3})
               endif

               if hcolumna:__enumindex < Len(appdata:himpresion["acolus"]) .and. appdata:himpresion["acolus"][hcolumna:__enumindex +1, "alineacion"] = "izq"
                  loop
               endif

               linea_simple({"narriba" => hpaso["ntope_alto"]      , "nizquierda" => hcolumna["area"]["der"]+0.05 , "nabajo" => hpaso["ntope_abajo"], "nderecha" => hcolumna["area"]["der"]+0.05 , "ncolor" => dgs_color_gris, "ngrosor" => 0.3})
            endif

         next

      endif

   endif

   *=================================================================================================================================================================
   * SI EXISTE LA VARIABLE LSPARADOR_IZQUIERDA EN HESPECIALES PONDREMOS UNA RAYA A LA IZQUIERDA
   *=================================================================================================================================================================

   if hb_HHasKey( hespeciales, "lseparador_izquierda")

      hpaso["ntope_alto"]  := appdata:himpresion["nrenglon"] + 0.05
      hpaso["ntope_abajo"] := appdata:himpresion["avance_maximo"]+appdata:himpresion["ninter"] - 0.05

      *LogFile(DToC(Date())+" "+Time()+"            *** al poner separador izquierdo  ntope_alto="+LTrim(str(hpaso["ntope_alto"],12,4))+" ntope_abajo="+LTrim(str(hpaso["ntope_abajo"],12,4)))

      linea_simple({"narriba" => hpaso["ntope_alto"]  , "nizquierda" => 0.4, "nabajo" => hpaso["ntope_abajo"], "nderecha" => 0.4, "ncolor" => dgs_color_base, "ngrosor" => 1.2})

   endif

   *=================================================================================================================================================================
   * ACTUALIZAMOS EL RENGLON REAL QUE VIENE EN appdata:himpresion EN CASO DE TENER AVANCES DE MAS DE UN RENGLON
   *=================================================================================================================================================================

   appdata:himpresion["nrenglon"]:=appdata:himpresion["avance_maximo"]

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION parte_texto(cestexto,nesluz)

   LOCAL hpaso:= {=>}, neste:= 0, ares:= {}
   LOCAL nvan, cparte, cres, nancho, b, aconve:={}, ncenti
   LOCAL r,k,va

   *LogFile(DToC(Date())+" "+Time()+"   ******** iniciando a partir texto = " + cestexto+"                 "+Str(seconds(),12,4))
   *LogFile(DToC(Date())+" "+Time()+"       **** nesluz = " + Str(nesluz,16,6)+"                 "+Str(seconds(),12,4))

   nvan:=1
   cparte:=""
   cres:=""

   for neste:=1 TO LEN(cestexto)

      b=subs(cestexto,neste,1)

      if asc(b)=13
         cres=cres+cparte
         aadd(ares,{cres,0})
         cres=""
         cparte=""
         loop
       endif

      if asc(b)=10
         loop
      endif

      if len(trim(b))=0

         *LogFile(DToC(Date())+" "+Time()+"       **** encontre un blanco cres = " + cres + " cparte = " + cparte + "                 "+Str(seconds(),12,4))

         nancho:=printer:ocanvas:textwidth(cres+cparte)
         aconve:=printer:ocanvas:PixelsToMapMode(nancho, 0)
         ncenti:=aconve[1]
         ncenti:=ncenti/100

         *LogFile(DToC(Date())+" "+Time()+"       **** ancho en cm de cres+cparte = " + Str(ncenti,16,6)+"                 "+Str(seconds(),12,4))

         if ncenti>nesluz
            aadd(ares,{cres,0})
            cres=""
            cparte:=right(cparte,len(cparte)-1)
            *LogFile(DToC(Date())+" "+Time()+"       **** texto se pasa agregue a pila solo cres quedando cres = " + cres + " cparte = " + cparte + "                 "+Str(seconds(),12,4))
         endif

         cres=cres+cparte
         cparte=""

      endif

      cparte:=cparte+b

   next neste

   *LogFile(DToC(Date())+" "+Time()+"       **** llegue al final  de la particion cres = " + cres + " cparte = " + cparte + "                 "+Str(seconds(),12,4))

   nancho:=printer:ocanvas:textwidth(cres+cparte)
   aconve:=printer:ocanvas:PixelsToMapMode(nancho, 0)
   ncenti:=aconve[1]
   ncenti:=ncenti/100

   *LogFile(DToC(Date())+" "+Time()+"       **** ancho en cm de cres+cparte = " + Str(ncenti,16,6)+"                 "+Str(seconds(),12,4))

   if ncenti>nesluz
     aadd(ares,{cres,0})
     cres=""
     *LogFile(DToC(Date())+" "+Time()+"       **** texto se pasa agregue a pila solo cres quedando cres = " + cres + " cparte = " + cparte + "                 "+Str(seconds(),12,4))
   endif

   cres=cres+cparte
   AADD(ares,{cres,0})

   *LogFile(DToC(Date())+" "+Time()+"       **** finalmente agregue a pila cres (cparte) remanente cres = " + cres + "                 "+Str(seconds(),12,4))

RETURN (ares)

//------------------------------------------------------------------------------

FUNCTION linea_simple( hdato)

   LOCAL narrical, nizqcal, nabacal, ndercal, open

   if hdato["ncolor"] = nil
      hdato["ncolor"] :=rgb( 000 , 000 , 000 )
   endif

   if hdato["ngrosor"] = nil
      hdato["ngrosor"] :=2.3
   endif

   open  := TPen():New( PS_SOLID, 1,  hdato["ncolor"] )
   open:nwidth:=hdato["ngrosor"]
   open:ncolor:=hdato["ncolor"]

   narrical:=hdato["narriba"]*100
   nizqcal :=hdato["nizquierda"]*100
   nabacal :=hdato["nabajo"]*100
   ndercal :=hdato["nderecha"]*100

  WITH OBJECT Printer:oCanvas
   :lTransparent   := .t.
    :open          := open
    :MoveTo(nizqcal,narrical)
    :LineTo(ndercal,nabacal)
  END WITH

  open:destroy()

RETURN NIL

//------------------------------------------------------------------------------

FUNCTION cuadro_lleno( hparams)

   LOCAL nanterior

   if !hb_HHasKey( hparams, "ncolor" )
      hparams["ncolor"] := rgb( 190 , 190 , 190 )
   endif

   hparams["narrical"]:=hparams["narriba"]*100
   hparams["nizqcal"] :=hparams["nizquierda"]*100
   hparams["nabacal"] :=hparams["nabajo"]*100
   hparams["ndercal"] :=hparams["nderecha"]*100

   WITH OBJECT Printer:oCanvas
      nanterior     :=:nClrPane
      :lTransparent := .t.
      :nClrPane     :=hparams["ncolor"]
      :fillrect({ hparams["nizqcal"] , hparams["narrical"] , hparams["ndercal"] , hparams["nabacal"] })
      :nClrPane     :=nanterior
   END WITH

RETURN NIL


//------------------------------------------------------------------------------


