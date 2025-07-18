/*
 * Proyecto: Api-Dygy
 * Fichero: Api-Dygy.prg
 * Descripci�n: M�dulo de entrada a la aplicaci�n
 * Autor:
 * Fecha: 17/01/2023
 */

#include "Xailer.ch"

Procedure Main(archivo_json)

   SET DATE FORMAT TO "dd/mm/yyyy"

   Application:cTitle := "Api-Dygy"

   AppData:AddData("cfile_json"           , archivo_json)
   AppData:AddData("oreceptor"            , nil)
   AppData:AddData("origen_bd"            , nil)
   AppData:AddData("nidempresa"           , {})
   AppData:AddData("cnombre_empresa"      , {})
   AppData:AddData("apuntos"              , {})
   AppData:AddData("lproduccion"          , .f.)
   AppData:AddData("lno_cierres"          , .f.)
   AppData:AddData("htimer"               , {=>})
   AppData:AddData("hreporte"             , {=>})
   AppData:AddData("himpresion"           , {=>})
   AppData:AddData("hestado"              , {=>})
   AppData:AddData("hexcel"               , {=>})
   AppData:AddData("oexcel"               , nil)
   AppData:AddData("opdf"                 , nil)
   AppData:AddData("oFileLog"             , nil)
   AppData:AddData("oConexion_nube"       , nil)

   principal():New( Application ):Show()

   Application:Run()

Return
