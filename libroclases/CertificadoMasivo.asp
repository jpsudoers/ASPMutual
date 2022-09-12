<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%

	fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
	fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
	cadena = day(now())&" de "&mes&" de "&year(now)
	sArchivo = "../pdf/xCertificado_"&fecha&".pdf"

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
Set piePagina = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoMutual.jpg")
piePagina.SetFile Server.MapPath("../images/FIRMA_JA.jpg")

theDoc.Font = theDoc.EmbedFont("Arial")

H=Request("h")


My_Array=split(H,",")
cantiad = UBound(My_Array)
'response.write(cantidad):response.End()
For Each item In My_Array

if trim(item) <> "" then



	rut = ""
	r_social = ""
	trabid = ""
	nombres = ""
	autenficacion = ""
	nacionalidad = ""
	nomCurso = ""
	codCurso = ""
	hrsCurso = ""
	fechaICurso = ""
	fechaTCurso = ""
   
	Set DATOS = Server.CreateObject("ADODB.RecordSet")
	'sql = "exec CERTIFICADO_MASIVO "&item&" "  
	sql = "select top 1(CASE WHEN (h.ID_AUTORIZACION='49759' or h.ID_AUTORIZACION='50843' or h.ID_AUTORIZACION='51370' or h.ID_AUTORIZACION='51262') then '76412470-7' WHEN (h.ID_AUTORIZACION='51619') then '58091269-9' WHEN (h.ID_AUTORIZACION='60474') then '76642320-5' WHEN (h.ID_AUTORIZACION='58951') then '92177000-6' WHEN (h.ID_AUTORIZACION='51833' or h.ID_AUTORIZACION='51864') then '76877410-2' WHEN (h.ID_AUTORIZACION='52646' or h.ID_AUTORIZACION='52536') then '99580830-7' WHEN (h.ID_AUTORIZACION='63626' or h.ID_AUTORIZACION='63625') then '77276280-1' else E.RUT END) as 'RUT',  (CASE WHEN (h.ID_AUTORIZACION='49759' or h.ID_AUTORIZACION='50843' or h.ID_AUTORIZACION='51370' or h.ID_AUTORIZACION='51262') then  'TRANSPORTES ALBERTO DIAZ PARRAGUEZ' WHEN (h.ID_AUTORIZACION='60474') then 'Compa�ia De Transportes Ventrosa Spa' WHEN (h.ID_AUTORIZACION='51619') then 'Giw Industries Inc.' WHEN (h.ID_AUTORIZACION='51833' or h.ID_AUTORIZACION='51864') then 'SERVICIOS DE TRANSPORTES DANIEL TAPIA JOFRE EIRL' WHEN (h.ID_AUTORIZACION='58951') then 'Lavanderia Le Grand Chic S.A.' WHEN (h.ID_AUTORIZACION='67281') then 'CIMM TECNOLOGIA Y SERVICIOS S.A.' WHEN (h.ID_AUTORIZACION='52646' or h.ID_AUTORIZACION='52536') then 'NEW TECH COPPER S.P.A.' WHEN (h.ID_AUTORIZACION='63626' or h.ID_AUTORIZACION='63625') then 'INDUSTRIAL SUPPORT COMPANY LTDA.' else UPPER(E.R_SOCIAL) END) as 'R_SOCIAL', (CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId' ,T.NACIONALIDAD,UPPER(T.NOMBRES) as NOMBRES, H.COD_AUTENFICACION, c.CODIGO, c.DESCRIPCION, c.HORAS,p.FECHA_INICIO_, p.FECHA_TERMINO from HISTORICO_CURSOS H inner join EMPRESAS E on E.ID_EMPRESA=H.ID_EMPRESA inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR  inner join PROGRAMA p on p.ID_PROGRAMA = h.ID_PROGRAMA  inner join curriculo c on c.id_mutual = p.id_mutual where h.ID_HISTORICO_CURSO in ("&item&") "
'response.write(sql):response.End()
	DATOS.Open sql,conn
if not DATOS.eof then
	rut = DATOS("rut")
	r_social = DATOS("r_social")
	trabid = DATOS("TrabId")
	nombres = DATOS("NOMBRES")
	autenficacion = DATOS("COD_AUTENFICACION")
	nacionalidad = DATOS("NACIONALIDAD")
	
	nomCurso = DATOS("DESCRIPCION")
	codCurso = DATOS("CODIGO")
	hrsCurso = DATOS("HORAS")
	fechaICurso = replace(DATOS("FECHA_INICIO_"),"-","/")
	fechaTCurso = replace(DATOS("FECHA_TERMINO"),"-","/")	

	theDoc.Rect = "30 50 570 700"

	theDoc.FontSize = 11
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	theDoc.Rect = "50 740 220 780"
	theDoc.FontSize = 8
	theDoc.AddImageObject theImg, False
	
	theDoc.Rect = "10 640 610 780"
	theDoc.FontSize =12
	theDoc.HPos = 0.5	
	theDoc.AddHtml ("<br><br><br><br><br><br><b>CERTIFICADO DE ASISTENCIA Y APROBACIÓN A CURSO DE CAPACITACIÓN</b>")	
  
  	theDoc.Rect = "50 740 220 780"
  
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
	theTable.AddText replace(FormatNumber(mid(RUT, 1,len(RUT)-2),0)&mid(RUT, len(RUT)-1,len(RUT)),",",".")
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText r_social
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
	theTable.AddText "Nro de Horas"
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
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.Padding = 4

'	'/******* CABECERA *********/

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
	theTable.Focus theDoc, 3
	theTable.Width(0) = 1
	theTable.Width(1) = 3
	theTable.Width(2) = 0.6

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
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Estado"
	theTable.SelectCell(2)
	theTable.Frame True, True, true, true	

	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5

	if(nacionalidad="0")then
	theTable.AddText replace(FormatNumber(mid(trabid, 1,len(trabid)-2),0)&mid(trabid, len(trabid)-1,len(trabid)),",",".")

	else
	theTable.AddText trabid
	end if
	
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText Datos("NOMBRES")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Aprobado"
	theTable.SelectCell(2)
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
	
'	if(codCurso<>"2016-091")then
'		theTable.NextRow
'		theTable.NextCell
'		theTable.AddText ""
'	
'		theTable.NextRow
'		theTable.NextCell
'		theDoc.HPos = 0
'		
'		if(codCurso<>"12.37.944-899" and codCurso<>"12.37.944-900")then
'			theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca MEL y que pueden cambiar sin previo aviso.</p>")
'		else
'			theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca Mantos Cooper y que pueden cambiar sin previo aviso.</p>")
'		end if 
'	end if
'	
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
	theTable.AddText Datos("COD_AUTENFICACION")
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
	theDoc.AddHtml ("<p align=""justify"">Este certificado ha sido emitido electrónicamente y su validez puede ser verificada ingresando a <font color=""#0000cc""><u>http://norte.otecmutual.cl/cert_online.asp</u></font> donde se puede consultar el Código de Autenticación de Certificado.</p>")

		theDoc.FontSize = 9
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 4	

		theDoc.Rect = "230 20 360 120"
		theDoc.AddImageObject piePagina, False		

		theDoc.Rect = "150 10 450 65"
		theDoc.FontSize = 8
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 1	
							
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText "JORGE ALDAY MONTECINOS"	
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText "SUBGERENTE DE CAPACITACIÓN"
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5	
		theTable.AddText "MUTUAL DE SEGURIDAD CAPACITACIÓN S.A."	
		theDoc.Rect = "30 50 570 700"
		
		theDoc.Page = theDoc.AddPage()
end if

end if
Next 'For Each item In My_Array



	theDoc.Save Server.MapPath(sArchivo)	
	Response.Redirect(sArchivo)
%>