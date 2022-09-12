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
  theDoc.AddHtml ("<b>Solicitudes Pendientes</b>")
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

sArchivo = "../pdf/Solicitudes_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)
function Tabla(theDoc)
	vfechI=""
    if(Request("fechaI")<>"")then	
	    vfechI=" and SOLICITUD.fecha>=CONVERT(datetime,'"&Request("fechaI")&" 00:00:00:000"&"',105) "
	end if
	
	vfechT=""
	if(Request("fechaT")<>"")then
	vfechT=" and SOLICITUD.fecha<=CONVERT(datetime,'"&Request("fechaT")&" 23:59:59:000"&"',105) "
	end if
	
	sql = "SELECT SOLICITUD.id_solicitud, EMPRESAS.RUT, EMPRESAS.R_SOCIAL as empresa, MUTUALES.nomb_mutual, OTIC.R_SOCIAL, "
	sql = sql&"CURRICULO.CODIGO,CURRICULO.NOMBRE_CURSO, CONVERT(VARCHAR(10),SOLICITUD.fecha, 105) as fecha, SOLICITUD.participantes "
	sql = sql&" FROM SOLICITUD "
	sql = sql&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=SOLICITUD.id_empresa "
	sql = sql&" INNER JOIN MUTUALES ON MUTUALES.Mutual_id =EMPRESAS.MUTUAL "
	sql = sql&" INNER JOIN OTIC ON OTIC.ID_OTIC=EMPRESAS.ID_OTIC "
	sql = sql&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=SOLICITUD.id_mutual "
	sql = sql&" where SOLICITUD.estado=1 "&vfechI&vfechT
	sql = sql&" order by SOLICITUD.id_solicitud asc"

'RESPONSE.Write(sql)
'RESPONSE.End()
	set rs =  conn.execute(sql)
	theDoc.FontSize = 9
	Set theTable = New Table
	theTable.Focus theDoc, 6
	theTable.Width(0) = 1
	theTable.Width(1) = 3
	theTable.Width(2) = 1.5
	theTable.Width(3) = 1
	theTable.Width(4) = 1
	theTable.Width(5) = 2
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
	theDoc.HPos = 0.5
	theTable.AddText "Fecha Requerida"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "N° Part."
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Rut"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Nombre Trabajador"
	
	'/********* DATOS **********/
	n=0
	i=0
	while not rs.eof
		i=i+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs("RUT")
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs("empresa")
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText rs("fecha")
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText rs("participantes")
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectRow theTable.Row
		
		If i = 1 Then theTable.Frame True, False, False, False
		
		sol = "select SOLICITUD_TRABAJADOR.id_trabajador from SOLICITUD_TRABAJADOR "
		sol = sol&" inner join SOLICITUD on SOLICITUD.id_solicitud=SOLICITUD_TRABAJADOR.id_solicitud "
		sol = sol&" where SOLICITUD.id_solicitud='"&rs("id_solicitud")&"'"
		sol = sol&" order by SOLICITUD_TRABAJADOR.id_trabajador asc "
		
		'response.Write(sol)
		'response.End()
		
		set rsSol =  conn.execute(sol)
		
		while not rsSol.eof
			sql2 = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD "
			sql2 = sql2&" from TRABAJADOR "
			sql2 = sql2&" where TRABAJADOR.ID_TRABAJADOR="&rsSol("id_trabajador")
			set rsTrab = conn.execute (sql2)
			theTable.SelectRow theTable.Row
			If i Mod 2 = 1 Then theTable.Fill "174 174 174"
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText rsTrab("RUT")
			theTable.NextCell
			theDoc.HPos = 0
			'theTable.AddText Left(rsTrab("RUT")&"0000000000000",13)&rsTrab("NOMBRES")
            'theTable.AddText Left(rsTrab("RUT")&"             ",13)&rsTrab("NOMBRES")
			theTable.AddText rsTrab("NOMBRES")
			rsSol.Movenext
		wend
		
		theTable.SelectRow theTable.Row
		If i Mod 2 = 1 Then theTable.Fill "174 174 174"
		'If i = 1 Then theTable.Frame True, False, False, False
		'If i = 5 Then theTable.Frame True, True, False, False
		'If i Mod 2 = 1 Then theTable.Fill "174 174 174"
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