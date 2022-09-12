<!--#include file="conexion.asp"-->
<!--#include file="funciones/pdfTable.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

Emitido = right("0"&month(now()),2)&"/"&year(now)

Set theDoc = Server.CreateObject("ABCpdf7.Doc")

theDoc.MediaBox = "0 0 500 900"

Set cabeceraPagina = Server.CreateObject("ABCpdf7.Image")
cabeceraPagina.SetFile Server.MapPath("images/Diploma_CCHC.jpg")
Set CodQr = Server.CreateObject("ABCpdf7.Image")

w = theDoc.MediaBox.Width
h = theDoc.MediaBox.Height
l = theDoc.MediaBox.Left
b = theDoc.MediaBox.Bottom 
theDoc.Transform.Rotate 90, l, b
theDoc.Transform.Translate w, 0

theDoc.Rect.Width = h
theDoc.Rect.Height = w

histTrab="select f.Empresa,f.Obra,f.IdentificadorUnico,'FechaTexto'=right('00' + convert(varchar(4),DAY(f.[Submission Date])),2) +' de '+"&_
	"DATENAME(M, DATEADD(M, 0, f.[Submission Date]))+' de '+convert(varchar(4),year(f.[Submission Date])), "&_
	"f.RutUsuario,f.submission_id,f.Id_jotform from Formularios_Verificacion_Cumplimientos f where f.IdentificadorUnico='"&Request("id")&"'"

set rsHistTrab = conn.execute (histTrab)

Tabla(theDoc)

theCount = theDoc.PageCount

theID = theDoc.GetInfo(theDoc.Root, "Pages")
theDoc.SetInfo theID, "/Rotate", "90"

sArchivo = "pdf/Diploma_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function Tabla(theDoc)
	theDoc.Rect = "0 0 900 500"
	theDoc.AddImageObject cabeceraPagina, True

	theDoc.Rect = "320 2 790 330"
	theDoc.FontSize = 20
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.Color = "31 73 125"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	'theTable.AddText "Empresa"
	theDoc.AddHtml "<i>"&rsHistTrab("Empresa")&"</i>"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	'theTable.AddText ""
	theDoc.AddHtml "<i>"&rsHistTrab("Obra")&"</i>"

	theDoc.Rect = "0 0 900 110"
	theDoc.FontSize = 12
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.Color = "0 0 0"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText rsHistTrab("FechaTexto")

	theDoc.FontSize = 7

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	'theTable.AddText "Verificado: XX/2020"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Emitido: "&Emitido

	theTable.NextRow
	CodQr.SetFile Server.MapPath("Google Drive/Versión 2. Lista de verificación de cumplimiento/"&rsHistTrab("RutUsuario")&"_"&rsHistTrab("Id_jotform")&"_"&rsHistTrab("submission_id")&"/26_image.png")	
	theDoc.Rect = "760 180 840 100"
	theDoc.AddImageObject CodQr, False

	theTable.NextRow

	theDoc.Rect = "720 0 880 90"
	theDoc.FontSize = 7
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8

	theTable.Padding = 2

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.Color = "0 0 0"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "ID: "&rsHistTrab("IdentificadorUnico")

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Verificable en:"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "verificacioncovid.mutualasesorias.cl"

end function
%>