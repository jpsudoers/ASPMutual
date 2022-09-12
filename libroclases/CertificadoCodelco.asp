<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set cabeceraPagina = Server.CreateObject("ABCpdf7.Image")
Set FirmaRel = Server.CreateObject("ABCpdf7.Image")

if(Request("empresa")<>"2841")then
	cabeceraPagina.SetFile Server.MapPath("../images/fondoCertCodelco2018.jpg")
else
	cabeceraPagina.SetFile Server.MapPath("../images/fondoCertCodelco2017c.jpg")
end if


theDoc.Font = theDoc.EmbedFont("Arial")
theDoc.FontSize =8

w = theDoc.MediaBox.Width
h = theDoc.MediaBox.Height
l = theDoc.MediaBox.Left
b = theDoc.MediaBox.Bottom 
theDoc.Transform.Rotate 90, l, b
theDoc.Transform.Translate w, 0

theDoc.Rect.Width = h
theDoc.Rect.Height = w

dim trabSelec
trabSelec=""
if(Request("trabajador")<>"0")then
	trabSelec=" and H.ID_TRABAJADOR="&Request("trabajador")
end if

dim relSelec
relSelec=""
if(Request("relator")<>"0")then
	relSelec=" and H.RELATOR='"&Request("relator")&"'"
end if

dim empSelec
empSelec=""
if(Request("empresa")<>"0")then
	empSelec=" and H.ID_EMPRESA="&Request("empresa")
end if

	histTrab="select (CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO  "
	histTrab=histTrab&" WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',UPPER(T.NOMBRES) as NOMBRES,  "
	histTrab=histTrab&" H.COD_AUTENFICACION, "
	histTrab=histTrab&"right('00' + CONVERT(varchar, Day(P.FECHA_TERMINO)), 2) as 'Dia',"
	histTrab=histTrab&"Case Month(P.FECHA_TERMINO) "
	histTrab=histTrab&"When 1 then 'Enero' "
	histTrab=histTrab&"When 2 then 'Febrero' "
	histTrab=histTrab&"When 3 then 'Marzo' "
	histTrab=histTrab&"When 4 then 'Abril' "
	histTrab=histTrab&"When 5 then 'Mayo' "
	histTrab=histTrab&"When 6 then 'Junio' "
	histTrab=histTrab&"When 7 then 'Julio' "
	histTrab=histTrab&"When 8 then 'Agosto' "
	histTrab=histTrab&"When 9 then 'Septiembre' "
	histTrab=histTrab&"When 10 then 'Octubre' "
	histTrab=histTrab&"When 11 then 'Noviembre' "
	histTrab=histTrab&"When 12 then 'Diciembre' End as Mes,Year(P.FECHA_TERMINO) as 'Año', "	
	histTrab=histTrab&" CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO, "
	histTrab=histTrab&" CONVERT(VARCHAR(10),DATEADD(year,2,convert(date,P.FECHA_TERMINO)), 105) as Vig, "
	histTrab=histTrab&" UPPER(ir.NOMBRES+' '+ir.A_PATERNO+' '+ir.A_MATERNO) as instructor,IR.RUT AS RutRel,"	
	histTrab=histTrab&" H.COD_AUTENFICACION,T.NACIONALIDAD from HISTORICO_CURSOS H "
	histTrab=histTrab&" inner join PROGRAMA p on p.ID_PROGRAMA=h.ID_PROGRAMA "
	histTrab=histTrab&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "
	histTrab=histTrab&" inner join bloque_programacion bq on bq.id_bloque=H.ID_BLOQUE "   
	histTrab=histTrab&" inner join INSTRUCTOR_RELATOR ir on ir.ID_INSTRUCTOR=bq.id_relator "	
	histTrab=histTrab&" where H.EVALUACION='Aprobado' AND H.ID_PROGRAMA="&Request("prog")
	histTrab=histTrab&relSelec&trabSelec&empSelec
	
	set rsHistTrab = conn.execute (histTrab)

'theDoc.Rect = "2 2 790 310"
Tabla(theDoc)

theCount = theDoc.PageCount
'/******** HEADER **********/
'theDoc.Rect = "2 2 790 610"
'theDoc.FontSize = 8
'For i = 1 To theCount
  'theDoc.PageNumber = i
  'theDoc.AddImageObject cabeceraPagina, False
'Next

'theDoc.Rect = "10 640 610 780"
'theDoc.HPos = 0.5
'theDoc.VPos = 0.5
'theDoc.FontSize =12
'For i = 1 To theCount
	  'theDoc.PageNumber = i
	  'theDoc.AddHtml ("<b>CERTIFICADO DE ASISTENCIA Y APROBACIÓN A CURSO DE CAPACITACIÓN</b>")
'Next

'/******** FOOTER **********/
'theCount = theDoc.PageCount
'For i = 1 To theCount
	'theDoc.Rect = "490 50 570 70"
	''theDoc.HPos = 1.0
	'theDoc.VPos = 0.5
	'theDoc.FontSize = 8
 	'theDoc.PageNumber = i	
  	'theDoc.AddText ""'"Página " & i & " de " & theCount
	'theDoc.Rect = "50 50 570 70"
	'theDoc.AddImageObject piePagina, False
	'theDoc.HPos = 0.5
	'theDoc.AddHtml ("<font color=""#999999"">Agencia Antofagasta | Washington 2701 Piso 3 Antofagasta – Chile | Tel.: (55) 251585 | www.mutual.cl</font>")
'Next

theID = theDoc.GetInfo(theDoc.Root, "Pages")
theDoc.SetInfo theID, "/Rotate", "90"

sArchivo = "../pdf/Certificado_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function Tabla(theDoc)
	theDoc.Rect = "1 1 790 610"
	theDoc.AddImageObject cabeceraPagina, False

	theDoc.Rect = "160 2 790 350"
	theDoc.Color = "64 64 64"
	theDoc.FontSize = 20
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Cambria")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5

	'if(rsHistTrab("NACIONALIDAD")="0")then
		'theTable.AddText rsHistTrab("NOMBRES")&", Rut "&replace(FormatNumber(mid(rsHistTrab("TrabId"), 1,len(rsHistTrab("TrabId"))-2),0)&mid(rsHistTrab("TrabId"), len(rsHistTrab("TrabId"))-1,len(rsHistTrab("TrabId"))),",",".")
	'else
		theTable.AddText rsHistTrab("NOMBRES")'&", Rut "&rsHistTrab("TrabId")
	'end if
	
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
	
		if(rsHistTrab("NACIONALIDAD")="0")then
			theTable.AddText replace(FormatNumber(mid(rsHistTrab("TrabId"), 1,len(rsHistTrab("TrabId"))-2),0)&mid(rsHistTrab("TrabId"), len(rsHistTrab("TrabId"))-1,len(rsHistTrab("TrabId"))),",",".")
		else
			theTable.AddText rsHistTrab("TrabId")
		end if	
	
	
		theTable.NextRow
		theDoc.Font = theDoc.EmbedFont("Cambria")
		theDoc.Rect = "160 2 790 109"
		
		theDoc.FontSize = 16
		
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 1	
							
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText rsHistTrab("instructor")
		
		'theTable.NextRow
		'theTable.NextCell
		'theDoc.HPos = 0.5
		'theTable.AddText "Relator/Facilitador"
		
		'theTable.NextRow
		'theTable.NextCell
		'theDoc.HPos = 0.5	
		'theTable.AddText "Mutual de Seguridad Capacitación S.A."	
		
	theTable.NextRow

	theDoc.Rect = "10 2 790 120"
	
	theDoc.FontSize = 66
	Set theTable = New Table
	theTable.Focus theDoc, 4
	theTable.Width(0) = 4.5
	'theTable.Width(1) = 5
	'theTable.Width(2) = 4
	'theTable.Width(3) = 4
	
	theTable.Padding = 4
	theDoc.Color = "255 255 255"
	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsHistTrab("Dia")
	
	
	theDoc.Rect = "90 2 790 105"
	
	theDoc.FontSize = 12
	Set theTable = New Table
	theTable.Focus theDoc, 4
	theTable.Width(0) = 4.5
	theTable.Width(1) = 5
	'theTable.Width(2) = 7

	
	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Cambria")
	theTable.NextRow
	'theTable.NextCell
	'theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText UCase(rsHistTrab("Mes"))	
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""'"Validez:"	
	theTable.NextRow
	'theTable.NextCell
	'theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsHistTrab("Año")
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText rsHistTrab("Dia")&" "&rsHistTrab("Mes")&" "&cstr(cint(rsHistTrab("Año")+2))		
		
	theTable.NextRow

	theDoc.Rect = "70 2 790 20"
	
	theDoc.FontSize = 8
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8

	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Código de Autenticación : "&rsHistTrab("COD_AUTENFICACION")
	
	if(rsHistTrab("RutRel")<>"11616799-9" and rsHistTrab("RutRel")<>"5940032-0")then
	theTable.NextRow
	FirmaRel.SetFile Server.MapPath("../images/FR-"&rsHistTrab("RutRel")&".jpg")		
	theDoc.Rect = "400 177 580 118"
	theDoc.AddImageObject FirmaRel, False
	end if
end function
%>