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

dim datosProg
datosProg="select C.DESCRIPCION,C.CODIGO,C.HORAS,CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"
datosProg=datosProg&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,(CASE WHEN C.VIGENCIA='1' then '6 Meses' "
datosProg=datosProg&" WHEN C.VIGENCIA='2' then '12 Meses' WHEN C.VIGENCIA='3' then '18 Meses' WHEN C.VIGENCIA='4' then '24 Meses' "
datosProg=datosProg&" WHEN C.VIGENCIA='5' then '48 Meses' END) as 'VIGENCIA' from PROGRAMA P "
datosProg=datosProg&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL WHERE P.ID_PROGRAMA="&Request("prog")

set rsDatosProg = conn.execute (datosProg)

nomCurso = rsDatosProg("DESCRIPCION")
codCurso = rsDatosProg("CODIGO")
hrsCurso = rsDatosProg("HORAS")
fechaICurso = replace(rsDatosProg("FECHA_INICIO"),"-","/")
fechaTCurso = replace(rsDatosProg("FECHA_TERMINO"),"-","/")
vigCurso = rsDatosProg("VIGENCIA")

dim trabSelec
trabSelec=""
if(Request("trabajador")<>"0")then
	trabSelec=" and H.ID_TRABAJADOR="&Request("trabajador")
end if

dim relSelec
relSelec=""
if(Request("relator")<>"0")then
	relSelec=" and H.RELATOR='"&Request("relator")&"'"
	'relSelec=""
end if

dim empSelec
empSelec=""
if(Request("empresa")<>"0")then
	empSelec=" and H.ID_EMPRESA="&Request("empresa")
end if

dim tot
tot="select COUNT(distinct ID_EMPRESA) as total from HISTORICO_CURSOS H "
tot=tot&" where H.EVALUACION='Aprobado' and H.ID_PROGRAMA="&Request("prog")
tot=tot&empSelec&relSelec

set Total = conn.execute (tot)

set rsEmp = conn.execute ("select distinct ID_EMPRESA from HISTORICO_CURSOS H where H.EVALUACION='Aprobado' and H.ID_PROGRAMA="&Request("prog")&empSelec&relSelec)

dim empresa_id
dim histTrab
dim histTrabCount
dim countHistTrab
cont=0
countHistTrab=0

mes=""
			
if(month(now())=1)then
	mes="Enero"
end if
			
if(month(now())=2)then
	mes="Febrero"
end if
			
if(month(now())=3)then
	mes="Marzo"
end if
			
if(month(now())=4)then
	mes="Abril"
end if
			
if(month(now())=5)then
	mes="Mayo"
end if
			
if(month(now())=6)then
	mes="Junio"
end if
			
if(month(now())=7)then
	mes="Julio"
end if
			
if(month(now())=8)then
	mes="Agosto"
end if
			
if(month(now())=9)then
	mes="Septiembre"
end if
			
if(month(now())=10)then
	mes="Octubre"
end if
			
if(month(now())=11)then
	mes="Noviembre"
end if
			
if(month(now())=12)then
	mes="Diciembre"
end if

cadena = day(now())&" de "&mes&" de "&year(now)
'split(FormatDateTime(now(), 1), ",")
while not rsEmp.eof
	cont=cont+1
	countHistTrab=0
	empresa_id=rsEmp("ID_EMPRESA")
	
	histTrabCount="select count(*) as totTrab from HISTORICO_CURSOS H where H.EVALUACION='Aprobado' "
	histTrabCount=histTrabCount&" AND H.ID_PROGRAMA="&Request("prog")&" and H.ID_EMPRESA="&empresa_id&relSelec&trabSelec
	
	set rsCountHistTrab = conn.execute (histTrabCount)

	histTrab="select E.RUT,UPPER(E.R_SOCIAL) as R_SOCIAL, (CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO "
	histTrab=histTrab&" WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',T.NACIONALIDAD,UPPER(T.NOMBRES) as NOMBRES, "
	histTrab=histTrab&" H.COD_AUTENFICACION from HISTORICO_CURSOS H "
	histTrab=histTrab&" inner join EMPRESAS E on E.ID_EMPRESA=H.ID_EMPRESA "
	histTrab=histTrab&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "
	histTrab=histTrab&" where H.EVALUACION='Aprobado' AND H.ID_PROGRAMA="&Request("prog")&" and H.ID_EMPRESA="&empresa_id
	histTrab=histTrab&relSelec&trabSelec
	
	set rsHistTrab = conn.execute (histTrab)
	
	while not rsHistTrab.eof
		countHistTrab=countHistTrab+1
		theDoc.Rect = "30 50 570 660"
		Tabla(theDoc)
		
		if(countHistTrab<rsCountHistTrab("totTrab"))then
			theDoc.Page = theDoc.AddPage()
	    end if
		
		rsHistTrab.Movenext
	wend

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
	  theDoc.PageNumber = i
	  theDoc.AddHtml ("<b>CERTIFICADO DE ASISTENCIA A CURSO DE CAPACITACIÓN</b>")
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
	theDoc.AddHtml ("<font color=""#999999"">Agencia Antofagasta | Washington 2701 Piso 3 Antofagasta – Chile | Tel.: (55) 251585 | www.mutual.cl</font>")
Next

sArchivo = "../pdf/Certificado_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function Tabla(theDoc)
	theDoc.FontSize = 11
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 1
	theDoc.Font = theDoc.EmbedFont("Arial")
	theDoc.AddHtml ("<b>Antofagasta, "&cadena&"</b>")
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Mutual de Seguridad Capacitación S.A., Rut 76.410.180-4 certifica que la empresa,"
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 1
	theTable.Width(1) = 3

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "RUT"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Razón Social"
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	'**** Tabla Empresa ********

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText replace(FormatNumber(mid(rsHistTrab("RUT"), 1,len(rsHistTrab("RUT"))-2),0)&mid(rsHistTrab("RUT"), len(rsHistTrab("RUT"))-1,len(rsHistTrab("RUT"))),",",".")
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsHistTrab("R_SOCIAL")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true

	theTable.NextRow
	
	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.Padding = 4

	'/******* CABECERA *********/

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Ha contratado el siguiente curso de capacitación:"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 1
	theTable.Width(1) = 3

	theTable.Padding = 9

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Curso"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	theDoc.HPos = 0.015
	theDoc.VPos = 0.5
	theDoc.AddHtml ("<font color=""#ffffff"">-</font>"&nomCurso)
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true

	theDoc.VPos = 0
	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 4
	theTable.Width(0) = 2
	theTable.Width(1) = 2
	theTable.Width(2) = 2
	theTable.Width(3) = 2

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Código del Curso"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText codCurso
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial Black")	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Nº de Horas"
	theTable.SelectCell(2)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText hrsCurso
	theTable.SelectCell(3)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Fecha de Inicio"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial")	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText fechaICurso
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial Black")	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Fecha de Termino"
	theTable.SelectCell(2)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial")	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText fechaTCurso
	theTable.SelectCell(3)
	theTable.Frame True, True, true, true

	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 1
	theTable.Width(1) = 3

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Vigencia del Curso"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theDoc.Font = theDoc.EmbedFont("Arial")	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText vigCurso
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.Padding = 4

	'/******* CABECERA *********/

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Y certifica que asistió el siguiente participante obteniendo una calificación positiva:"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 1
	theTable.Width(1) = 3

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "RUT"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Nombre del Participante"
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5

	if(rsHistTrab("NACIONALIDAD")="0")then
	theTable.AddText replace(FormatNumber(mid(rsHistTrab("TrabId"), 1,len(rsHistTrab("TrabId"))-2),0)&mid(rsHistTrab("TrabId"), len(rsHistTrab("TrabId"))-1,len(rsHistTrab("TrabId"))),",",".")
	else
	theTable.AddText rsHistTrab("TrabId")
	end if
	
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsHistTrab("NOMBRES")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.Padding = 4

	'/******* CABECERA *********/

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml ("<p align=""justify"">- Este certificado es válido sólo para documentar los requisitos de control de acceso a faenas y NO es válido para documentar franquicias Sence.</p>")

	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""	

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml ("<p align=""justify"">- Este certificado es válido sólo para ser presentado por la empresa contratante del curso y sólo tiene vigencia si el asistente indicado tiene relación con dicha empresa.</p>")
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca MEL y que pueden cambiar sin previo aviso.</p>")
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 3
	theTable.Width(1) = 3

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Código de Autenticación de Certificado:"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsHistTrab("COD_AUTENFICACION")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.Padding = 4

	'/******* CABECERA *********/

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml ("<p align=""justify"">Este certificado ha sido emitido electrónicamente y su validez puede ser verificada ingresando al sistema <font color=""#0000cc""><u>http://norte.otecmutual.cl</u></font> donde se puede consultar el Código de Autenticación de Certificado.</p>")

end function

%>