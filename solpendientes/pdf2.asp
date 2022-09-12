<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<!--#include file="../funciones/fecha.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
mes=""

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoMutual.jpg")
theDoc.Font = theDoc.EmbedFont("Arial")
theDoc.FontSize =8
theDoc.Rect = "50 70 570 660"
sLn = "                   "
theDoc.VPos = 0
Tabla(theDoc)

theCount = theDoc.PageCount
'/******** HEADER **********/
theDoc.Rect = "50 700 260 740"
theDoc.FontSize = 8
For i = 1 To theCount
  theDoc.PageNumber = i
  theDoc.AddImageObject theImg, False
Next

theDoc.Rect = "10 640 610 720"
theDoc.HPos = 0.5
theDoc.VPos = 0.5
theDoc.FontSize =12
For i = 1 To theCount
  theDoc.PageNumber = i
  theDoc.AddHtml ("<b>Facturar</b>")
	'theDoc.FrameRect
Next

'/******** FOOTER **********/
theCount = theDoc.PageCount
For i = 1 To theCount
	theDoc.Rect = "490 50 570 70"
	theDoc.HPos = 1.0
	theDoc.VPos = 0.5
	theDoc.FontSize = 8
 	theDoc.PageNumber = i	
  	theDoc.AddText "Página " & i & " de " & theCount
	theDoc.Rect = "50 50 570 70"
	theDoc.HPos = 0.5
	'theDoc.AddText "Agencia Antofagasta | Washington 2701 Piso 3 Antofagasta – Chile | Tel.: (55) 251585 | www.mutual.cl"
Next

sArchivo = "../pdf/Facturar_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)
function Tabla(theDoc)
	vprog=Request("prog")
	
	empresa=""
	if(Request("empresa")<>"0")then
	empresa=" and AUTORIZACION.ID_EMPRESA="&Request("empresa")
	end if
	
	
	sql = "select EMPRESAS.RUT as rut,EMPRESAS.R_SOCIAL as empresa,CURRICULO.CODIGO as codigo, "
	sql = sql&" CURRICULO.NOMBRE_CURSO as nombre, AUTORIZACION.ORDEN_COMPRA as orden, AUTORIZACION.VALOR_OC as valor "
	sql = sql&" from AUTORIZACION "
	sql = sql&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "
	sql = sql&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
	sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
	sql = sql&"  where AUTORIZACION.FACTURA is null and AUTORIZACION.ID_PROGRAMA="&vprog&empresa
	sql = sql&"  order by rut asc "

'RESPONSE.Write(sql)
'RESPONSE.End()
	set rs =  conn.execute(sql)
	theDoc.FontSize = 9
	Set theTable = New Table
	theTable.Focus theDoc, 6
	theTable.Width(0) = 1
	theTable.Width(1) = 3
	theTable.Width(2) = 1
	theTable.Width(3) = 2
	theTable.Width(4) = 1
	theTable.Width(5) = 1
	theTable.Padding = 2
	'/******* CABECERA *********/
	theTable.Frame True, False, False, False
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	'theTable.Frame True,True,False,True
	theTable.AddText "Rut"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Empresa"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Código de Curso"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Nombre de Curso"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Orden de Compra"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Valor"
	
	'/********* DATOS **********/
	n=0
	i=0
	while not rs.eof
		i=i+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText FormatNumber(mid(rs("RUT"), 1,len(rs("RUT"))-2),0)&mid(rs("RUT"), len(rs("RUT"))-1,len(rs("RUT")))
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs("empresa")
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs("codigo")
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs("nombre")
		theTable.NextCell
		theDoc.HPos = 0  
		theTable.AddText FormatNumber(rs("orden"),0)
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText "$ "&FormatNumber(rs("valor"),0)
		theTable.SelectRow theTable.Row
		
		If i = 1 Then theTable.Frame True, False, False, False
		If i Mod 2 = 1 Then theTable.Fill "174 174 174"
		n=n+1
		rs.Movenext
	wend
	'theTable.SelectRow theTable.Row
	theTable.Frame False, True, False, False
	'theDoc.FontSize = 15
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theTable.AddText ""
end function

%>