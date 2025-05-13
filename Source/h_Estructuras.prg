/*
 * Proyecto: Api-Dygy
 * Fichero: h_Estructuras.prg
 * Descripción:
 * Autor:
 * Fecha: 11/01/2023
 */

#include "Xailer.ch"

//--------------------------------------------------------------------------------------------------------------------------------------

FUNCTION inicializa_estructura_reportes( hreporte )

   hreporte:={=>}

   hreporte["cfuncion"]                    := ""
   hreporte["cformato"]                    := ""
   hreporte["cfile"]                       := ""
   hreporte["cnombre_completo_entregable"] := ""
   hreporte["cnombre_entregable"]          := ""
   hreporte["cruta_entregable"]            := ""
   hreporte["cprograma_api"]               := ""
   hreporte["carchivo_json"]               := ""
   hreporte["carchivo_error"]              := ""
   hreporte["carchivo_avances"]            := ""
   hreporte["carchivo_finalizacion"]       := ""
   hreporte["cnombre_empresa"]             := ""

RETURN (nil)

//--------------------------------------------------------------------------------------------------------------------------------------

FUNCTION inicializa_estructura_estado_peticion_api( hestado )

   hestado:={=>}

   hestado["cEstado"]                   := ""
   hestado["cDescripcion"]              := ""
   hestado["cTituloTextoUsuario"]       := ""
   hestado["cTextoUsuario"]             := ""
   hestado["nPorcentaje"]               := 0

RETURN (nil)

//--------------------------------------------------------------------------------------------------------------------------------------

FUNCTION inicializa_estructura_hadoerror( hadoerror )

   hadoerror:={=>}

   hadoerror["cquery"]          := ""
   hadoerror["hgenerales"]      := {=>}
   hadoerror["ctitulo_mensaje"] := ""
   hadoerror["ctexto_mensaje"]  := "Reporte este mensaje con su asesor de sistemas"
   hadoerror["ctipo"]           := "error"

RETURN (nil)

//--------------------------------------------------------------------------------------------------------------------------------------
