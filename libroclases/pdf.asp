<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
mes=""

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoMutual.jpg")

theDoc.Font = theDoc.EmbedFont("Arial")
theDoc.FontSize =8

dim tot
tot="select COUNT(distinct ID_EMPRESA) as total from HISTORICO_CURSOS "
tot=tot&" where HISTORICO_CURSOS.ID_PROGRAMA="&Request("prog")
tot=tot&" and HISTORICO_CURSOS.ID_EMPRESA="&Request("empresa")

set Total = conn.execute (tot)

set rsEmp = conn.execute ("select distinct ID_EMPRESA from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_PROGRAMA="&Request("prog")&" and HISTORICO_CURSOS.ID_EMPRESA="&Request("empresa"))

dim empresa_id
cont=0
while not rsEmp.eof
cont=cont+1
empresa_id=rsEmp("ID_EMPRESA")
theDoc.Rect = "50 70 570 660"
theDoc.VPos = 0
Tabla(theDoc)
theDoc.Rect = "50 70 570 560"
Tabla2(theDoc)
'theDoc.Rect = "50 70 570 250"
'Tabla3(theDoc)
theDoc.Rect = "50 70 570 405"
TablaParticipantes(theDoc)
theDoc.Rect = "50 70 570 90"
TablaFirma(theDoc)
if(cont<Total("total"))then
theDoc.Page = theDoc.AddPage()

end if
rsEmp.Movenext
wend


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
  'if(i mod 2 = 1)then
	  theDoc.PageNumber = i
	  theDoc.AddHtml ("<b>CERTIFICADO ASISTENCIA A CURSO DE CAPACITACIÓN</b>")
  'end if
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
  	theDoc.AddText ""'"Página " & i & " de " & theCount
	theDoc.Rect = "50 50 570 70"
	theDoc.HPos = 0.5
	theDoc.AddText "Agencia Antofagasta | Washington 2701 Piso 3 Antofagasta – Chile | Tel.: (55) 251585 | www.mutual.cl"
Next

sArchivo = "../pdf/Certificado_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)
function Tabla(theDoc)
	theDoc.FontSize = 10
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 1.5
	theTable.Width(1) = 5.5
	theTable.Padding = 2	
	
	sql = "select * from INSTITUCION WHERE ID_INSTITUCION=1"
	set rs = conn.execute (sql)

	
	'/******* CABECERA *********/
	theTable.Frame True, False, False, False
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "NOMBRE OTEC"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rs("R_SOCIAL")

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "RUT OTEC"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rs("RUT")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "TELEFONO"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rs("TELEFONO")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "FAX"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " "
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "FACTURA Nº"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "FECHA"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText right("0"&day(now()),2)&"/"&right("0"&month(now()),2)&"/"&year(now)
	theTable.NextRow
	theTable.SelectRow theTable.Row
	theTable.Frame True, False, False, False

end function

function Tabla2(theDoc)
	theDoc.FontSize = 10
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 1.5
	theTable.Width(1) = 5
	theTable.Padding = 2	
	
	sql2 = "select distinct CURRICULO.CODIGO as codigo, "
	sql2 = sql2&" CURRICULO.NOMBRE_CURSO as nombre_curso,CURRICULO.DESCRIPCION, "
	sql2 = sql2&" CURRICULO.HORAS as horas,EMPRESAS.RUT as rut, EMPRESAS.R_SOCIAL as r_social, "
	sql2 = sql2&" CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as fechaInicio, "
	sql2 = sql2&" CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105)as fechaTermino, " 
	sql2 = sql2&" SEDES.DIRECCION+' '+SEDES.CIUDAD as direccion "
	sql2 = sql2&" from HISTORICO_CURSOS "
	sql2 = sql2&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
	sql2 = sql2&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA "
	sql2 = sql2&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
	sql2 = sql2&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=HISTORICO_CURSOS.RELATOR "
	sql2 = sql2&" inner join SEDES on SEDES.ID_SEDE=HISTORICO_CURSOS.SEDE "
	sql2 = sql2&" where HISTORICO_CURSOS.ID_PROGRAMA="&Request("prog")&" and HISTORICO_CURSOS.ID_EMPRESA="&empresa_id
	'sql2 = sql2&" and HISTORICO_CURSOS.RELATOR="&Request("relator")
	
	'response.Write(sql2)
	'response.End()

	set rs2 = conn.execute (sql2)
	
	'/******* CABECERA *********/
	'while not rs2.eof
	theTable.Frame True, False, False, False
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "EMPRESA"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rs2("r_social")
	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "RUT"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText replace(FormatNumber(mid(rs2("rut"), 1,len(rs2("rut"))-2),0)&mid(rs2("rut"), len(rs2("rut"))-1,len(rs2("rut"))),",",".")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "CÓDIGO SENCE"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rs2("codigo")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "NOMBRE CURSO"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rs2("DESCRIPCION")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "N° HORASº "
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rs2("horas")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "FECHA INICIO"
	theTable.NextCell
	theDoc.HPos = 0 
	theTable.AddText replace(rs2("fechaInicio"),"-","/")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "FECHA TERMINO"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText replace(rs2("fechaTermino"),"-","/")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "LUGAR EJECUCION"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rs2("direccion")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "ANEXO FACTURA"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.SelectRow theTable.Row
	theTable.Frame True, False, False, False

'rs2.Movenext
	'wend
end function

function Tabla3(theDoc)
	sql = "select * from INSTITUCION WHERE ID_INSTITUCION=1"
	set rs = conn.execute (sql)

	theDoc.FontSize = 10
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 7
	theTable.Padding = 2	
	
	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "YO, "&rs("CONTACTO")&" RUT "&rs("RUT_CONTACTO")

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "Representante  del "&rs("R_SOCIAL")&","
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	
	if(month(now()) = "1")then
	mes="ENERO"
	end if
	
	if(month(now()) = "2")then
	mes="FEBRERO"
	end if
	
	if(month(now()) = "3")then
	mes="MARZO"
	end if
	
	if(month(now()) = "4")then
	mes="ABRIL"
	end if
	
	if(month(now()) = "5")then
	mes="MAYO"
	end if
	
	if(month(now()) = "6")then
	mes="JUNIO"
	end if
	
	if(month(now()) = "7")then
	mes="JULIO"
	end if
	
	if(month(now()) = "8")then
	mes="AGOSTO"
	end if
	
	if(month(now()) = "9")then
	mes="SEPTIEMBRE"
	end if
	
	if(month(now()) = "10")then
	mes="OCTUBRE"
	end if
	
	if(month(now()) = "11")then
	mes="NOVIEMBRE"
	end if
	
	if(month(now()) = "12")then
	mes="DICIEMBRE"
	end if
	
	theF1 = theDoc.AddFont("Arial")
	theF2 = theDoc.AddFont("Arial Black")
	
	theDoc.Font = theF1
	theDoc.AddText " RUT  "&rs("RUT")&", con fecha "
	theDoc.Font = theF2
	theDoc.AddText  right("0"&day(now()),2)&" de "&mes&" de "&year(now)
	theDoc.Font = theF1
	theDoc.AddText  ", certifico que los datos consignados son fidedignos."
end function

function TablaParticipantes(theDoc)
	theDoc.FontSize = 10
	Set theTable = New Table
	theTable.Focus theDoc, 4
	theTable.Width(0) = 1
	theTable.Width(1) = 4
	theTable.Width(2) = 1.5
	theTable.Width(3) = 1.5
	theTable.Padding = 2	
	
	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	
	theF1 = theDoc.AddFont("Arial")
	theF2 = theDoc.AddFont("Arial Black")
		
	theDoc.Font = theF2
	
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 1
	theTable.AddText "NOMINA DE PARTICIPANTES"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "RUT"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "NOMBRE PARTICIPANTE"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "% ASISTENCIA"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "NOTA %"
	theTable.SelectRow theTable.Row
	theTable.Frame True, False, False, False
	theDoc.Font = theF1
	
	'/********* DATOS **********/
	
	trabajador=""
	if(Request("trabajador")<>"0")then
	  trabajador=" and HISTORICO_CURSOS.ID_TRABAJADOR="&Request("trabajador")
	end if
	
	sql3 = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,HISTORICO_CURSOS.ASISTENCIA,TRABAJADOR.NACIONALIDAD,"
	sql3 = sql3&"TRABAJADOR.ID_EXTRANJERO,HISTORICO_CURSOS.CALIFICACION "
	sql3 = sql3&" from HISTORICO_CURSOS "
	sql3 = sql3&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
	sql3 = sql3&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA "
	sql3 = sql3&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
	sql3 = sql3&" where HISTORICO_CURSOS.ID_PROGRAMA="&Request("prog")&" and HISTORICO_CURSOS.ID_EMPRESA="&empresa_id&trabajador

	set rs3 = conn.execute (sql3)
	
	n=0
	i=0
	while not rs3.eof
		i=i+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		
		if(rs3("NACIONALIDAD")="0")then
			theTable.AddText replace(FormatNumber(mid(rs3("RUT"), 1,len(rs3("RUT"))-2),0)&mid(rs3("RUT"), len(rs3("RUT"))-1,len(rs3("RUT"))),",",".")
		else
			theTable.AddText rs3("ID_EXTRANJERO")
		end if
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs3("NOMBRES")
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText rs3("ASISTENCIA")
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText rs3("CALIFICACION")

		if(i=6)then
			theTable.SelectRow theTable.Row
			theTable.Frame False, True, False, False
			theDoc.Rect = "50 70 570 660"
			theDoc.Page = theDoc.AddPage()
			theDoc.FontSize = 10
			Set theTable = New Table
			theTable.Focus theDoc, 4
			theTable.Width(0) = 1
			theTable.Width(1) = 4
			theTable.Width(2) = 1.5
			theTable.Width(3) = 1.5
			theTable.Padding = 2	
			
			'/******* CABECERA *********/
			theTable.NextRow
			theTable.NextCell
			
			theF1 = theDoc.AddFont("Arial")
			theF2 = theDoc.AddFont("Arial Black")
				
			theDoc.Font = theF2
			
			theDoc.HPos = 0	
			theTable.AddText ""
			theTable.NextCell
			theDoc.HPos = 1
			theTable.AddText "NOMINA DE PARTICIPANTES"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0	
			theTable.AddText "RUT"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText "NOMBRE PARTICIPANTE"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "% ASISTENCIA"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NOTA %"
			theTable.SelectRow theTable.Row
			theTable.Frame True, False, False, False
			theDoc.Font = theF1
		end if
		
		if(i=6)then
			theTable.SelectRow theTable.Row
			theTable.Frame False, True, False, False
		end if
		
		theTable.SelectRow theTable.Row
		
		If i = 1 Then theTable.Frame True, False, False, False
		If i Mod 2 = 1 Then theTable.Fill "174 174 174"
		n=n+1
		
		rs3.Movenext
	wend
	
	theTable.SelectRow theTable.Row
	theTable.Frame False, True, False, False
	
	theDoc.Rect = "50 70 570 250"
    Tabla3(theDoc)
	
	'theTable.Frame False, True, False, False
end function

function TablaFirma(theDoc)
	theDoc.FontSize = 10
	Set theTable = New Table
	theTable.Focus theDoc, 4
	theTable.Width(0) = 2
	theTable.Width(1) = 2
	theTable.Width(2) = 1
	theTable.Width(3) = 2
	theTable.Padding = 2	
	
	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText " Nombre y firma Representante "
	theTable.Frame True, False, False, False
end function
%>