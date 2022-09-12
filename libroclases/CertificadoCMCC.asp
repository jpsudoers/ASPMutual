<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
Set piePagina = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoCertCMCC.jpg")
piePagina.SetFile Server.MapPath("../images/timbreCMCC2.jpg")

theDoc.Font = theDoc.EmbedFont("Arial")

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

cadena = ""'day(now())&" de "&mes&" de "&year(now)

	histTrab="select (CASE WHEN (h.ID_AUTORIZACION='49759' or h.ID_AUTORIZACION='50843' or h.ID_AUTORIZACION='51370' or h.ID_AUTORIZACION='51262') then "
    histTrab=histTrab&"'76412470-7' WHEN (h.ID_AUTORIZACION='51619') then '58091269-9' WHEN (h.ID_AUTORIZACION='60474') then '76642320-5' WHEN (h.ID_AUTORIZACION='58951') then '92177000-6' WHEN (h.ID_AUTORIZACION='51833' or h.ID_AUTORIZACION='51864') then '76877410-2' WHEN (h.ID_AUTORIZACION='52646' or h.ID_AUTORIZACION='52536') then '99580830-7' WHEN (h.ID_AUTORIZACION='63626' or h.ID_AUTORIZACION='63625') then '77276280-1' WHEN (h.ID_AUTORIZACION='69264' and h.ID_HISTORICO_CURSO='182213') then '76108126-8' else E.RUT END) as 'RUT', "
	histTrab=histTrab&"(CASE WHEN (h.ID_AUTORIZACION='49759' or h.ID_AUTORIZACION='50843' or h.ID_AUTORIZACION='51370' or h.ID_AUTORIZACION='51262') then "
    histTrab=histTrab&"'TRANSPORTES ALBERTO DIAZ PARRAGUEZ' WHEN (h.ID_AUTORIZACION='60474') then 'Compañia De Transportes Ventrosa Spa' WHEN (h.ID_AUTORIZACION='51619') then 'Giw Industries Inc.' WHEN (h.ID_AUTORIZACION='51833' or h.ID_AUTORIZACION='51864') then 'SERVICIOS DE TRANSPORTES DANIEL TAPIA JOFRE EIRL' WHEN (h.ID_AUTORIZACION='58951') then 'Lavanderia Le Grand Chic S.A.' WHEN (h.ID_AUTORIZACION='67281') then 'CIMM TECNOLOGIA Y SERVICIOS S.A.' WHEN (h.ID_AUTORIZACION='52646' or h.ID_AUTORIZACION='52536') then 'NEW TECH COPPER S.P.A.' WHEN (h.ID_AUTORIZACION='63626' or h.ID_AUTORIZACION='63625') then 'INDUSTRIAL SUPPORT COMPANY LTDA.' WHEN (h.ID_AUTORIZACION='69264' and h.ID_HISTORICO_CURSO='182213') then 'IMA AUTOMATIZACION LTDA.' else UPPER(E.R_SOCIAL) END) as 'R_SOCIAL',  "
	histTrab=histTrab&"/*E.RUT,UPPER(E.R_SOCIAL) as R_SOCIAL,*/ (CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO "
	histTrab=histTrab&" WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',T.NACIONALIDAD,UPPER(T.NOMBRES) as NOMBRES, "
	histTrab=histTrab&" H.COD_AUTENFICACION,CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,H.CALIFICACION from HISTORICO_CURSOS H "
	histTrab=histTrab&" inner join EMPRESAS E on E.ID_EMPRESA=H.ID_EMPRESA "
	histTrab=histTrab&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "
	histTrab=histTrab&" inner join PROGRAMA P on p.id_programa=H.id_programa "	
	histTrab=histTrab&" where H.EVALUACION='Aprobado' AND H.ID_PROGRAMA="&Request("prog")&" and H.ID_EMPRESA="&Request("empresa")
	histTrab=histTrab&relSelec&trabSelec
	
	set rsHistTrab = conn.execute (histTrab)

	theDoc.Rect = "70 50 530 720"
	theDoc.Color = "0 0 0"
	Tabla(theDoc)

theCount = theDoc.PageCount
'/******** HEADER **********/
theDoc.Rect = "70 720 130 780"
'theDoc.FontSize = 8
'For i = 1 To theCount
  theDoc.PageNumber = i
  theDoc.AddImageObject theImg, False
  
  theDoc.Rect = "70 390 130 450"
  theDoc.AddImageObject theImg, False  
'Next

theDoc.Rect = "10 720 520 740"
theDoc.HPos = 1
theDoc.AddHtml ("<b>N° "&rsHistTrab("COD_AUTENFICACION")&"</b>")

theDoc.Rect = "10 640 610 780"
theDoc.HPos = 0.5
theDoc.VPos = 0.5
theDoc.FontSize =12
'For i = 1 To theCount
theDoc.PageNumber = i
theDoc.Color = "232 81 0"
theDoc.AddHtml ("<b>Certificado Cursos Inducción “Cero Daño” CMCC</b>")
'Next
theDoc.Color = "0 0 0"
theDoc.Rect = "10 390 520 410"
theDoc.HPos = 1
theDoc.AddHtml ("<b>N° "&rsHistTrab("COD_AUTENFICACION")&"</b>")

theDoc.Rect = "10 370 610 390"
theDoc.HPos = 0.5
theDoc.VPos = 0.5
'theDoc.FontSize =12
'For i = 1 To theCount
theDoc.PageNumber = i
theDoc.Color = "232 81 0"
theDoc.AddHtml ("<b>Certificado Cursos Inducción “Cero Daño” CMCC</b>")
'Next

'/******** FOOTER **********/
theCount = theDoc.PageCount
For i = 1 To theCount
	theDoc.Rect = "490 50 570 70"
	theDoc.HPos = 1.0
	theDoc.VPos = 0.5
	theDoc.FontSize = 8
 	theDoc.PageNumber = i	
Next

cert = "IF NOT EXISTS (select * from LOG_CERTIFICADOS C where C.CERTIFICADO='Certificado_"&fecha&".pdf') BEGIN "&_
       "insert into LOG_CERTIFICADOS (ID_EMPRESA,ID_TRABAJADOR,ID_PROGRAMA,FECHA_GENERACION,GENERADO,TIPO_ERROR,CERTIFICADO)"&_
       " values('"&Request("empresa")&"','"&Request("trabajador")&"','"&Request("prog")&"',GETDATE(),1,Null,'Certificado_"&fecha&".pdf') END "

conn.execute (cert)

sArchivo = "../pdf/Certificado_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function Tabla(theDoc)
	theDoc.FontSize = 12
	theDoc.TextStyle.Justification = 1
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	'theTable.NextRow
	'theTable.NextCell
	'theDoc.HPos = 1
	'theDoc.Font = theDoc.EmbedFont("Arial")
	'theDoc.AddHtml ("<b>Antofagasta, "&cadena&"</b>")
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	'theTable.NextRow
	'theTable.NextCell
	'theDoc.HPos = 0
	'theDoc.TextStyle.Justification = 1
	'theDoc.Rect.Inset 10, 0
	'theTable.AddText "Compañía Minera Cerro Colorado otorga al Sr(a): "&rsHistTrab("NOMBRES")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	if(rsHistTrab("NACIONALIDAD")="0")then
		theTable.AddText "Compañía Minera Cerro Colorado otorga al Sr(a): "&rsHistTrab("NOMBRES")&", Rut: "&replace(FormatNumber(mid(rsHistTrab("TrabId"), 1,len(rsHistTrab("TrabId"))-2),0)&mid(rsHistTrab("TrabId"), len(rsHistTrab("TrabId"))-1,len(rsHistTrab("TrabId"))),",",".")&". El presente certificado de aprobación del curso de inducción, realizado con fecha: "&rsHistTrab("FECHA_INICIO")&"  Nota: "&rsHistTrab("CALIFICACION")
	else
		theTable.AddText "Compañía Minera Cerro Colorado otorga al Sr(a): "&rsHistTrab("NOMBRES")&", Rut: "&rsHistTrab("TrabId")&". El presente certificado de aprobación del curso de inducción, realizado con fecha: "&rsHistTrab("FECHA_INICIO")&"  Nota: "&rsHistTrab("CALIFICACION")
	end if
	
	'theTable.NextRow
	'theTable.NextCell
	'theTable.AddText "realizado con fecha: "&rsHistTrab("FECHA_INICIO")&"  Nota: "
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText "Vigencia: Este certificado tiene una validez de 4 años, a contar de la fecha de realización de este curso, será de uso individual a quien se le otorga este certificado."	

	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""	
	
	theTable.NextRow

	Set theTable = New Table
	theTable.Focus theDoc, 3
	theTable.Width(0) = 3
	theTable.Width(1) = 2
	theTable.Width(2) = 3
	theTable.Padding = 4	

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Relator Curso Inducción"'rsHistTrab("COD_AUTENFICACION")
	theTable.SelectCell(0)
	theTable.Frame True, false, false, false
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""'rsHistTrab("COD_AUTENFICACION")
	'theTable.SelectCell(1)
	'theTable.Frame True, false, false, false	
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Especialista HSE"'rsHistTrab("COD_AUTENFICACION")
	theTable.SelectCell(2)
	theTable.Frame True, false, false, false	
	
	theTable.NextRow	

	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""		
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "COPIA INTERNA"	
	theTable.SelectCell(0)
	theTable.Frame false, True, false, false		

	theTable.NextRow

	'theDoc.FontSize = 9
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4	

	theDoc.Rect = "260 598 340 518"
	theDoc.AddImageObject piePagina, False	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""				
	
	theTable.NextRow	
	
	'------------COPIA TRABAJADOR
	'theDoc.Rect = "10 440 610 480"
	'theDoc.Color = "232 81 0"
	'theDoc.AddHtml ("<b>Certificado Cursos Inducción “Cero Daño” CMCC</b>")
	
	'theDoc.Color = "0 0 0"
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	'theTable.NextRow
	'theTable.NextCell
	'theDoc.HPos = 1
	'theDoc.Font = theDoc.EmbedFont("Arial")
	'theDoc.AddHtml ("<b>Antofagasta, "&cadena&"</b>")
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	if(rsHistTrab("NACIONALIDAD")="0")then
		theTable.AddText "Compañía Minera Cerro Colorado otorga al Sr(a): "&rsHistTrab("NOMBRES")&", Rut: "&replace(FormatNumber(mid(rsHistTrab("TrabId"), 1,len(rsHistTrab("TrabId"))-2),0)&mid(rsHistTrab("TrabId"), len(rsHistTrab("TrabId"))-1,len(rsHistTrab("TrabId"))),",",".")&". El presente certificado de aprobación del curso de inducción, realizado con fecha: "&rsHistTrab("FECHA_INICIO")&"  Nota: "&rsHistTrab("CALIFICACION")
	else
		theTable.AddText "Compañía Minera Cerro Colorado otorga al Sr(a): "&rsHistTrab("NOMBRES")&", Rut: "&rsHistTrab("TrabId")&". El presente certificado de aprobación del curso de inducción, realizado con fecha: "&rsHistTrab("FECHA_INICIO")&"  Nota: "&rsHistTrab("CALIFICACION")
	end if
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText "Vigencia: Este certificado tiene una validez de 4 años, a contar de la fecha de realización de este curso, será de uso individualizado a quien se le otorga este certificado."	

	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""	
	
	theTable.NextRow

	Set theTable = New Table
	theTable.Focus theDoc, 3
	theTable.Width(0) = 3
	theTable.Width(1) = 2
	theTable.Width(2) = 3
	theTable.Padding = 4	

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Relator Curso Inducción"'rsHistTrab("COD_AUTENFICACION")
	theTable.SelectCell(0)
	theTable.Frame True, false, false, false
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""'rsHistTrab("COD_AUTENFICACION")
	'theTable.SelectCell(1)
	'theTable.Frame True, false, false, false	
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Especialista HSE"'rsHistTrab("COD_AUTENFICACION")
	theTable.SelectCell(2)
	theTable.Frame True, false, false, false	
	
	theTable.NextRow	

	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""		
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "COPIA TRABAJADOR"	
	theTable.SelectCell(0)
	theTable.Frame false, True, false, false		

	theTable.NextRow

	'theDoc.FontSize = 9
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4	

	theDoc.Rect = "260 250 340 170"
	theDoc.AddImageObject piePagina, False			
	
end function

%>