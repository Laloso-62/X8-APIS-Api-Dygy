/*
 * Proyecto: Api-Rezta
 * Fichero: f_Excel.prg
 * Descripci�n:
 * Autor:
 * Fecha: 29/08/2022
 */

#include "Xailer.ch"
#include "Digio.ch"

//------------------------------------------------------------------------------

FUNCTION dameletra( nescolumna )

  LOCAL aletras:={}, cesletra:="A"

  aadd(aletras,"A")
  aadd(aletras,"B")
  aadd(aletras,"C")
  aadd(aletras,"D")
  aadd(aletras,"E")
  aadd(aletras,"F")
  aadd(aletras,"G")
  aadd(aletras,"H")
  aadd(aletras,"I")
  aadd(aletras,"J")
  aadd(aletras,"K")
  aadd(aletras,"L")
  aadd(aletras,"M")
  aadd(aletras,"N")
  aadd(aletras,"O")
  aadd(aletras,"P")
  aadd(aletras,"Q")
  aadd(aletras,"R")
  aadd(aletras,"S")
  aadd(aletras,"T")
  aadd(aletras,"U")
  aadd(aletras,"V")
  aadd(aletras,"W")
  aadd(aletras,"X")
  aadd(aletras,"Y")
  aadd(aletras,"Z")

  aadd(aletras,"AA")
  aadd(aletras,"AB")
  aadd(aletras,"AC")
  aadd(aletras,"AD")
  aadd(aletras,"AE")
  aadd(aletras,"AF")
  aadd(aletras,"AG")
  aadd(aletras,"AH")
  aadd(aletras,"AI")
  aadd(aletras,"AJ")
  aadd(aletras,"AK")
  aadd(aletras,"AL")
  aadd(aletras,"AM")
  aadd(aletras,"AN")
  aadd(aletras,"AO")
  aadd(aletras,"AP")
  aadd(aletras,"AQ")
  aadd(aletras,"AR")
  aadd(aletras,"AS")
  aadd(aletras,"AT")
  aadd(aletras,"AU")
  aadd(aletras,"AV")
  aadd(aletras,"AW")
  aadd(aletras,"AX")
  aadd(aletras,"AY")
  aadd(aletras,"AZ")

  aadd(aletras,"BA")
  aadd(aletras,"BB")
  aadd(aletras,"BC")
  aadd(aletras,"BD")
  aadd(aletras,"BE")
  aadd(aletras,"BF")
  aadd(aletras,"BG")
  aadd(aletras,"BH")
  aadd(aletras,"BI")
  aadd(aletras,"BJ")
  aadd(aletras,"BK")
  aadd(aletras,"BL")
  aadd(aletras,"BM")
  aadd(aletras,"BN")
  aadd(aletras,"BO")
  aadd(aletras,"BP")
  aadd(aletras,"BQ")
  aadd(aletras,"BR")
  aadd(aletras,"BS")
  aadd(aletras,"BT")
  aadd(aletras,"BU")
  aadd(aletras,"BV")
  aadd(aletras,"BW")
  aadd(aletras,"BX")
  aadd(aletras,"BY")
  aadd(aletras,"BZ")

  cesletra:=aletras[nescolumna]

RETURN (cesletra)

//------------------------------------------------------------------------------

FUNCTION resuelve_pestana(cpestana)

   LOCAL hpaso:= {=>}, hrenglon:= {=>}, hpestana:= {=>}, oexcel, ohoja

   *--------------------------------------------------------------------------------------------------------------------------------------
   * AGREGAMOS UNA PESTAÑA AL OBJETO EXCEL
   *--------------------------------------------------------------------------------------------------------------------------------------

   ohoja                  := appdata:oexcel:Sheets:add()
   ohoja:name             := appdata:hexcel[cpestana]["nombre"]
   ohoja:Cells:Font:Name  := "Calibri Light"
   ohoja:Cells:Font:Size  := 10

   *--------------------------------------------------------------------------------------------------------------------------------------
   * COPIAMOS LOS DATOS AL CLIPBOARD Y LUEGO LO COPIAMOS EN LA PESTAÑA
   *--------------------------------------------------------------------------------------------------------------------------------------

	WITH OBJECT TClipBoard():Create()
      IF :Open()
         :Empty()
         :SetText( appdata:hexcel[cpestana]["cdatos"] )
      ENDIF
      :End()
   END

   ohoja:paste()

   *--------------------------------------------------------------------------------------------------------------------------------------
   * CONFIGURAMOS LAS COLUMNAS DE LA PESTAÑA
   *--------------------------------------------------------------------------------------------------------------------------------------

   appdata:hexcel[cpestana]["cletra_tope"] := dameletra(Len(appdata:hexcel[cpestana]["acolus"]))

   for each hrenglon in appdata:hexcel[cpestana]["acolus"]

      ohoja:Columns( hrenglon:__enumindex ):ColumnWidth:= hrenglon["ancho"]

      do case

         case hrenglon["alineacion"] = "izq"
            ohoja:Columns( hrenglon:__enumindex ):HorizontalAlignment:=-4131

         case hrenglon["alineacion"] = "der"
            ohoja:Columns( hrenglon:__enumindex ):HorizontalAlignment:=-4152

         case hrenglon["alineacion"] = "cen"
            ohoja:Columns( hrenglon:__enumindex ):HorizontalAlignment:=-4108

         otherwise
            ohoja:Columns( hrenglon:__enumindex ):HorizontalAlignment:=-4131

      endcase

      if Len(hrenglon["formato"])>0
         ohoja:Columns( hrenglon:__enumindex ):NumberFormat:=hrenglon["formato"]
      endif

      if hrenglon["wrap"]
         ohoja:Columns( hrenglon:__enumindex ):WrapText         :=.t.
      endif

      if hb_HHasKey( hrenglon, "nfuente" )
         ohoja:Columns( hrenglon:__enumindex ):Font:Size        := hrenglon["nfuente"]
      endif

   next

   *--------------------------------------------------------------------------------------------------------------------------------------
   * REALIZAMOS ARREGLOS EN LOS PRIMEROS RENGLONES, TITULOS Y SUBTITULOS
   *--------------------------------------------------------------------------------------------------------------------------------------

   *--------------------------------------------------------------------------------------------------------------------------------------
   *-- PRIMERO ASIGNAMOS EL FONDO EN BLANCO A TODA EL AREA DE DATOS, TAMBIEN PONEMOS POR DEFAULT ALINEACION VERTICAL CENTRADA
   *--------------------------------------------------------------------------------------------------------------------------------------

   hpaso["crango"]:="A1:"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(appdata:hexcel[cpestana]["nrenglon_tope"])

   ohoja:Range(hpaso["crango"]):Font:color    := dgs_excel_color_texto_datos
   ohoja:Range(hpaso["crango"]):Interior:Color:= dgs_excel_color_fondo_datos
   ohoja:Range(hpaso["crango"]):VerticalAlignment:=-4108

   *--------------------------------------------------------------------------------------------------------------------------------------
   *-- FORMATO PAUTADO PARA TODA EL AREA DE DATOS, CONSIDERANDO DESDE LOS ENCABEZADOS, LUEGO SE LES CAMBIARA EL FORMATO
   *--------------------------------------------------------------------------------------------------------------------------------------
   *-- HACEMOS LA ASIGNACION EN LOS  PRIMEROS RENGLONES Y LUEGO COPIAMOS A TODA EL AREA DE DATOS
   *--------------------------------------------------------------------------------------------------------------------------------------
	
   hpaso["crango"]:="A"+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"])+":"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"])

   ohoja:Range(hpaso["crango"]):Font:color    := dgs_excel_color_texto_datos
   ohoja:Range(hpaso["crango"]):Interior:Color:= dgs_excel_color_fondo_datos

   hpaso["crango"]:="A"+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"]+1)+":"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"]+1)

   ohoja:Range(hpaso["crango"]):Font:color    := dgs_excel_color_texto_datos
   ohoja:Range(hpaso["crango"]):Interior:Color:= dgs_excel_color_alt_fondo_datos

   hpaso["crango"]:="A"+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"])+":"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"]+1)

   ohoja:Range(hpaso["crango"]):copy()

   hpaso["crango"]:="A"+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"])+":"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(appdata:hexcel[cpestana]["nrenglon_tope"])

   ohoja:Range(hpaso["crango"]):PasteSpecial(4)

   ohoja:Range("A2:A2"):select()

   *--------------------------------------------------------------------------------------------------------------------------------------
   * FORMATO DE TITULO
   *--------------------------------------------------------------------------------------------------------------------------------------

   hpaso["crango"]:="A1:" + appdata:hexcel[cpestana]["cletra_tope"] + "1"

   ohoja:Range(hpaso["crango"]):Font:color       := dgs_excel_color_titulos
   ohoja:Range(hpaso["crango"]):Interior:Color   := dgs_excel_color_interior_titulos
   ohoja:Range(hpaso["crango"]):Font:Bold        := .t.
   ohoja:Range(hpaso["crango"]):Font:size        := dgs_excel_fuente_titulos
   ohoja:Range(hpaso["crango"]):RowHeight        := dgs_excel_altura_renglon_titulos
   ohoja:Range(hpaso["crango"]):VerticalAlignment:=-4108

   ohoja:cells(1, 1):HorizontalAlignment:=-4131

   *--------------------------------------------------------------------------------------------------------------------------------------
   * FORMATO DE SUB-TITULOS
   *--------------------------------------------------------------------------------------------------------------------------------------

   for each hrenglon in appdata:hexcel[cpestana]["asubtitulos"]

      hpaso["crango"]:="A"+ToString(hrenglon["nrenglon"])+":"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(hrenglon["nrenglon"])

      ohoja:Range(hpaso["crango"]):Font:color       := dgs_excel_color_subtitulos
      ohoja:Range(hpaso["crango"]):Interior:Color   := dgs_excel_color_fondo_subtitulos
      ohoja:Range(hpaso["crango"]):Font:Bold        := .t.
      ohoja:Range(hpaso["crango"]):Font:size        := dgs_excel_fuente_subtitulos
      ohoja:Range(hpaso["crango"]):RowHeight        := dgs_excel_altura_subtitulos
      ohoja:Range(hpaso["crango"]):VerticalAlignment:=-4108

      ohoja:cells(hrenglon["nrenglon"], 1):HorizontalAlignment:=-4131


   next

   *--------------------------------------------------------------------------------------------------------------------------------------
   * FORMATO DE ENCABEZADOS
   *--------------------------------------------------------------------------------------------------------------------------------------

   hpaso["crango"]:="A"+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"])+":"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"])

   ohoja:Range(hpaso["crango"]):Font:color       := dgs_excel_color_encabezados
   ohoja:Range(hpaso["crango"]):Interior:Color   := dgs_excel_color_fondo_encabezados
   ohoja:Range(hpaso["crango"]):Font:Bold        := .t.
   ohoja:Range(hpaso["crango"]):Font:size        := dgs_excel_fuente_encabezados
   ohoja:Range(hpaso["crango"]):VerticalAlignment:=-4108
   ohoja:Range(hpaso["crango"]):WrapText         :=.t.

   *--------------------------------------------------------------------------------------------------------------------------------------
   * ULTIMO RENGLON O TOTALES
   *--------------------------------------------------------------------------------------------------------------------------------------

   hpaso["crango"]:="A"+ToString(appdata:hexcel[cpestana]["nrenglon_tope"])+":"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(appdata:hexcel[cpestana]["nrenglon_tope"])

   ohoja:Range(hpaso["crango"]):Font:color       := dgs_excel_color_encabezados
   ohoja:Range(hpaso["crango"]):Interior:Color   := dgs_excel_color_fondo_encabezados
   ohoja:Range(hpaso["crango"]):Font:Bold        := .t.
   ohoja:Range(hpaso["crango"]):Font:size        := dgs_excel_fuente_encabezados
   ohoja:Range(hpaso["crango"]):VerticalAlignment:=-4108
   ohoja:Range(hpaso["crango"]):WrapText         :=.t.

   *--------------------------------------------------------------------------------------------------------------------------------------
   * FILTRO EN TODOS LOS DATOS DE PESTAÑA (OPCIONAL)
   *--------------------------------------------------------------------------------------------------------------------------------------

   if appdata:hexcel[cpestana]["lfiltro"]

      hpaso["crango"]:="A"+ToString(appdata:hexcel[cpestana]["nrenglon_encabezado"])+":"+appdata:hexcel[cpestana]["cletra_tope"]+ToString(appdata:hexcel[cpestana]["nrenglon_tope"])

      ohoja:Range(hpaso["crango"]):autofilter()

   endif

   ohoja:cells(1,1):select()

RETURN (.t.)

//------------------------------------------------------------------------------
