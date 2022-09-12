<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
mes=""

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
Set piePagina = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoMutual.jpg")
piePagina.SetFile Server.MapPath("../images/FIRMA_MM.jpg")

theDoc.Font = theDoc.EmbedFont("Arial")

dim datosProg
datosProg="select DESCRIPCION=(CASE WHEN c.ID_MUTUAL=70 THEN c.NOMBRE_CURSO else C.DESCRIPCION end),C.CODIGO,HORAS=(case when p.ID_Modalidad=1 or p.ID_Modalidad=2 then C.HORAS else "
datosProg=datosProg&" (DATEDIFF(hour, CONVERT(datetime, P.FECHA_INICIO_ + ' ' +  "
datosProg=datosProg&" (select hb.HORARIO from horario_bloques hb where hb.ID_HORARIO=P.BMI), 21),CONVERT(datetime,  "
datosProg=datosProg&" P.FECHA_INICIO_ + ' ' + (select hb.HORARIO from horario_bloques hb  "
datosProg=datosProg&" where hb.ID_HORARIO=P.BTF),21))*(DATEDIFF(day, P.FECHA_INICIO_, P.FECHA_TERMINO)+1)) END),CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"
datosProg=datosProg&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,(CASE WHEN C.VIGENCIA='1' then '6' "
datosProg=datosProg&" WHEN C.VIGENCIA='2' then '12' WHEN C.VIGENCIA='3' then '18' WHEN C.VIGENCIA='4' then '24' "
datosProg=datosProg&" WHEN C.VIGENCIA='5' then '48' WHEN C.VIGENCIA='6' then '36' END) as 'VIGENCIA',p.id_mutual from PROGRAMA P "
datosProg=datosProg&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL WHERE P.ID_PROGRAMA="&Request("prog")

set rsDatosProg = conn.execute (datosProg)

nomCurso = rsDatosProg("DESCRIPCION")
codCurso = rsDatosProg("CODIGO")
IdMutual = rsDatosProg("id_mutual")

if(cint(rsDatosProg("HORAS"))>8 and cint(rsDatosProg("HORAS"))<16)then
	hrsCurso = "8"
elseif(cint(rsDatosProg("HORAS"))>16)then
	hrsCurso = "16"
else
	hrsCurso = rsDatosProg("HORAS")
end if

fechaICurso = replace(rsDatosProg("FECHA_INICIO"),"-","/")
fechaTCurso = replace(rsDatosProg("FECHA_TERMINO"),"-","/")
vigCurso = Cint( rsDatosProg("VIGENCIA"))/12

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

	histTrab="select (CASE WHEN (h.ID_AUTORIZACION='49759' or h.ID_AUTORIZACION='50843' or h.ID_AUTORIZACION='51370' or h.ID_AUTORIZACION='51262') then "
    histTrab=histTrab&"'76412470-7' WHEN (h.ID_AUTORIZACION='51619') then '58091269-9' WHEN (h.ID_AUTORIZACION='60474') then '76642320-5' WHEN (h.ID_AUTORIZACION='58951') then '92177000-6' WHEN (h.ID_AUTORIZACION='51833' or h.ID_AUTORIZACION='51864') then '76877410-2' WHEN (h.ID_AUTORIZACION='52646' or h.ID_AUTORIZACION='52536') then '99580830-7' WHEN (h.ID_AUTORIZACION='63626' or h.ID_AUTORIZACION='63625') then '77276280-1' WHEN (h.ID_AUTORIZACION='69264' and h.ID_HISTORICO_CURSO='182213') then '76108126-8'  WHEN (h.ID_AUTORIZACION='96424' and h.ID_HISTORICO_CURSO='237349') or (h.ID_AUTORIZACION='84768' and h.ID_HISTORICO_CURSO='208828') or (h.ID_AUTORIZACION='73294' and h.ID_HISTORICO_CURSO='187879') then '76547741-7' else E.RUT END) as 'RUT', "
	histTrab=histTrab&"(CASE WHEN (h.ID_AUTORIZACION='49759' or h.ID_AUTORIZACION='50843' or h.ID_AUTORIZACION='51370' or h.ID_AUTORIZACION='51262') then "
    histTrab=histTrab&"'TRANSPORTES ALBERTO DIAZ PARRAGUEZ' WHEN (h.ID_AUTORIZACION='60474') then 'Compañia De Transportes Ventrosa Spa' WHEN (h.ID_AUTORIZACION='51619') then 'Giw Industries Inc.' WHEN (h.ID_AUTORIZACION='51833' or h.ID_AUTORIZACION='51864') then 'SERVICIOS DE TRANSPORTES DANIEL TAPIA JOFRE EIRL' WHEN (h.ID_AUTORIZACION='58951') then 'Lavanderia Le Grand Chic S.A.' WHEN (h.ID_AUTORIZACION='67281') then 'CIMM TECNOLOGIA Y SERVICIOS S.A.' WHEN (h.ID_AUTORIZACION='52646' or h.ID_AUTORIZACION='52536') then 'NEW TECH COPPER S.P.A.' WHEN (h.ID_AUTORIZACION='63626' or h.ID_AUTORIZACION='63625') then 'INDUSTRIAL SUPPORT COMPANY LTDA.' WHEN (h.ID_AUTORIZACION='69264' and h.ID_HISTORICO_CURSO='182213') then 'IMA AUTOMATIZACION LTDA.' WHEN (h.ID_AUTORIZACION='96424' and h.ID_HISTORICO_CURSO='237349') or (h.ID_AUTORIZACION='84768' and h.ID_HISTORICO_CURSO='208828') or (h.ID_AUTORIZACION='73294' and h.ID_HISTORICO_CURSO='187879') then 'Avant SPA' else UPPER(E.R_SOCIAL) END) as 'R_SOCIAL',  "
	histTrab=histTrab&"/*E.RUT,UPPER(E.R_SOCIAL) as R_SOCIAL,*/ (CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO "
	histTrab=histTrab&" WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',T.NACIONALIDAD,UPPER(T.NOMBRES) as NOMBRES, "
	histTrab=histTrab&"UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE( "
	histTrab=histTrab&"REPLACE(t.NOMBRES,' ','_') "
	histTrab=histTrab&",'Ñ','N') "
	histTrab=histTrab&",',','_') "
	histTrab=histTrab&",'Á ','_') "
	histTrab=histTrab&",'É','_') "
	histTrab=histTrab&",'Í','_') "
	histTrab=histTrab&",'Ó','_') "
	histTrab=histTrab&",'Ú','_')) as ARCHIVO, "
	histTrab=histTrab&" H.COD_AUTENFICACION,facilitador=dbo.MayMinTexto (Ir.NOMBRES+' '+Ir.A_PATERNO+' '+Ir.A_MATERNO),'EstApro'=H.CALIFICACION,"
	histTrab=histTrab&"'fechaTermino'=(case when P.ID_MUTUAL=2363 then "
	histTrab=histTrab&" (case when H.[FECHA_FIN_E-LEARNING] is not null then CONVERT(VARCHAR(10),H.[FECHA_FIN_E-LEARNING], 105) else CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) "
	histTrab=histTrab&" END) else CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) END) "
	histTrab=histTrab&" from HISTORICO_CURSOS H "
	histTrab=histTrab&" inner join EMPRESAS E on E.ID_EMPRESA=H.ID_EMPRESA "
	histTrab=histTrab&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "
	histTrab=histTrab&" inner join INSTRUCTOR_RELATOR ir on ir.ID_INSTRUCTOR=H.RELATOR "
	histTrab=histTrab&" inner join programa p on p.ID_PROGRAMA=H.ID_PROGRAMA  "
	histTrab=histTrab&" where H.EVALUACION='Aprobado' AND H.ID_PROGRAMA="&Request("prog")&" and H.ID_EMPRESA="&empresa_id
	histTrab=histTrab&relSelec&trabSelec
	
	set rsHistTrab = conn.execute (histTrab)
	nomArchivo=""
	while not rsHistTrab.eof
		countHistTrab=countHistTrab+1
		theDoc.Rect = "30 50 570 700"
		nomArchivo=rsHistTrab("ARCHIVO")
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


'Response.end()

theCount = theDoc.PageCount
'/******** HEADER **********/
theDoc.Rect = "50 740 220 780"
theDoc.FontSize = 8
For i = 1 To theCount
  theDoc.PageNumber = i
  theDoc.AddImageObject theImg, False
Next

theDoc.Rect = "10 640 610 780"
theDoc.HPos = 0.5
theDoc.VPos = 0.5
theDoc.FontSize =12
For i = 1 To theCount
	  theDoc.PageNumber = i
	  theDoc.AddHtml ("<b>CERTIFICADO DE ASISTENCIA Y APROBACIÓN A CURSO DE CAPACITACIÓN</b>")
Next

'/******** FOOTER **********/
theCount = theDoc.PageCount
For i = 1 To theCount
	theDoc.Rect = "490 50 570 70"
	theDoc.HPos = 1.0
	theDoc.VPos = 0.5
	theDoc.FontSize = 8
 	theDoc.PageNumber = i	
  	'theDoc.AddText ""'"Página " & i & " de " & theCount
	'theDoc.Rect = "50 50 570 70"
	'theDoc.HPos = 0.5
	'theDoc.AddHtml ("<font color=""#999999"">Agencia Antofagasta | Washington 2701 Piso 3 Antofagasta – Chile | Tel.: (55) 251585 | www.mutual.cl</font>")
Next

cert = "IF NOT EXISTS (select * from LOG_CERTIFICADOS C where C.CERTIFICADO='Certificado_"&fecha&".pdf') BEGIN "&_
       "insert into LOG_CERTIFICADOS (ID_EMPRESA,ID_TRABAJADOR,ID_PROGRAMA,FECHA_GENERACION,GENERADO,TIPO_ERROR,CERTIFICADO)"&_
       " values('"&Request("empresa")&"','"&Request("trabajador")&"','"&Request("prog")&"',GETDATE(),1,Null,'Certificado_"&fecha&".pdf') END "

conn.execute (cert)

sArchivo = "../pdf/CERTIFICADO_"&nomArchivo&".pdf"
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
	theTable.Padding = 6

	'/******* CABECERA *********/

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0

	if(IdMutual<>2352 and IdMutual<>2356)then
		theDoc.AddHTML "Ha contratado el siguiente curso de capacitación, el cual tiene una <b>vigencia de "&Cstr(vigCurso)&" años</b>."
	else
		theDoc.AddHTML "Ha contratado el siguiente curso de capacitación."
	end if

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
	'theTable.NextRow


	if(IdMutual<>2345 and IdMutual<>2352 and IdMutual<>2347 and IdMutual<>2356 and IdMutual<>2362 and IdMutual<>2361 and IdMutual<>2363)then
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Facilitador"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	theTable.AddText rsHistTrab("facilitador")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	elseif(IdMutual=2363)then
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Modalidad Curso"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	theTable.AddText "Asincrónico"
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	else
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Modalidad Curso"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	theTable.AddText "Virtual"
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	end if

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

	if(codCurso="PMCHS_1 Inducción especifica")then
		theTable.AddText "2 Horas 30 Min."
	elseif(codCurso="PMCHS_4 Refugio Minero")then
		theTable.AddText "1 Horas 30 Min."
	elseif(codCurso="AB_ ANGLO" or codCurso="RL_ANGLO" or codCurso="GPM-MASC" or codCurso="PMCM-MASC" or codCurso="TCAL_ANGLO" or codCurso="ESES_ANGLO" or codCurso="MEL_CDN_BHP")then
		theTable.AddText "Asincrónico"
	else
		theTable.AddText hrsCurso
	end if

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

if(IdMutual<>"2363")then 
	theTable.AddText fechaTCurso
else
	theTable.AddText replace(rsHistTrab("fechaTermino"),"-","/")
end if


	theTable.SelectCell(3)
	theTable.Frame True, True, true, true

	theTable.NextRow

if(codCurso="VP_CAM08")then
	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 3
	theTable.Width(0) = 2
	theTable.Width(1) = 2
	theTable.Width(2) = 4

	theTable.Padding = 2
	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	theTable.AddText " - Objetivos del curso"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true


	theTable.NextCell
	theTable.AddText " - Neumáticos"
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " - Marco regulatorio"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " - Reglamento de conducción y tipos de terreno"
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

theDoc.Font = theDoc.EmbedFont("Arial Black")

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Contenidos del curso"
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " - Definiciones"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " - Seguridad, equipamiento y mantenimiento"
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, true, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " - Características de 4x4"
	theTable.SelectCell(1)
	theTable.Frame False, true, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, true, true, true


	theTable.NextRow
end if

if(codCurso="VP_TA_03")then
	theDoc.FontSize = 8
	Set theTable = New Table
	theTable.Focus theDoc, 3
	theTable.Width(0) = 2
	theTable.Width(1) = 3.7
	theTable.Width(2) = 2.3

	theTable.Padding = 2
	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	theTable.AddText " 1. Objetivos del curso"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true


	theTable.NextCell
	theTable.AddText " 5. Margen de seguridad."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 2. Marco regulatorio."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 6. Factor de caída."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 3. ECF 2 requisitos para trabajos en altura física."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 7. Síndrome del arnés."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 4. Componentes sistemas y subsistemas personales de detención de caídas (spdc)."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 8. Plataformas fijas, móviles, temporal y su verificación."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.FontSize = 11
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Contenidos del curso"
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theDoc.FontSize = 8
	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " de caídas (spdc)."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " su verificación."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.FontSize = 11
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theDoc.FontSize = 8
	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText "  - Tipos de arnés anticaídas. - Estrobos."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText "  - Amortiguador de impacto. - Rieles verticales. - Líneas retráctiles."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText "  - Líneas de vida horizontal: permanente y temporal."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true


	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, true, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText "  - Líneas de vida vertical: permanente y temporal."
	theTable.SelectCell(1)
	theTable.Frame False, true, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, true, true, true


	theTable.NextRow
end if

if(codCurso="VP_TA_PRACTICO-Trabajo En Altura")then
	theDoc.FontSize = 8
	Set theTable = New Table
	theTable.Focus theDoc, 3
	theTable.Width(0) = 2
	theTable.Width(1) = 4
	theTable.Width(2) = 2

	theTable.Padding = 2
	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	theTable.AddText " 1.Sistemas y subsistemas personales de detención de caídas (SPDC)."
	theTable.SelectCell(1)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " 2. Margen de seguridad"
	theTable.SelectCell(2)
	theTable.Frame False, False, true, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " - Tipos de arnés anticaídas."
	theTable.SelectCell(1)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " 3. Rescate en Trabajo de Altura Física "
	theTable.SelectCell(2)
	theTable.Frame False, False, true, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " - Estrobos."
	theTable.SelectCell(1)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " 4. Factor de caída."
	theTable.SelectCell(2)
	theTable.Frame False, False, true, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.FontSize = 11
	theTable.AddText "Contenidos del curso"
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theDoc.FontSize = 8
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " - Amortiguador de impacto."
	theTable.SelectCell(1)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, False, true, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " - Rieles verticales."
	theTable.SelectCell(1)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, False, true, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " - Líneas retráctiles."
	theTable.SelectCell(1)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, False, true, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " - Líneas de vida horizontal: permanente y temporal."
	theTable.SelectCell(1)
	theTable.Frame False, False, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, False, true, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, true, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " - Líneas de vida vertical: permanente y temporal."
	theTable.SelectCell(1)
	theTable.Frame False, true, true, true

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, true, true, true

	theTable.NextRow
end if


if(codCurso="VP_RIG_20")then
	theDoc.FontSize = 8
	Set theTable = New Table
	theTable.Focus theDoc, 3
	theTable.Width(0) = 2
	theTable.Width(1) = 3
	theTable.Width(2) = 3

	theTable.Padding = 2
	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	theTable.AddText " 1. Objetivos del curso"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true


	theTable.NextCell
	theTable.AddText " 11. Qué debe tener un plan de Izaje."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 2. Marco Regulatorio: NORMA ASME B-30."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 12. RSV N°3."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 3. ECF N°7."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 13. Aseguramiento del Área de Trabajo."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 4. Qué se entiende por Rigger de carga"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 14. Clases de Rigger."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.FontSize = 11
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Contenidos del curso"
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theDoc.FontSize = 8
	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 5. Qué se entiende por señales del Rigger "
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 15. Elementos que debe utilizar el Rigger."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.FontSize = 11
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theDoc.Font = theDoc.EmbedFont("Arial")
	theDoc.FontSize = 8
	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 6. Cuáles son las señales del Rigger de Carga."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 16. Distancia de seguridad en izaje de cargas."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 7. Cuáles son los requisitos asociados al Rigger."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 17. Requisitos de la Organización."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 8. Competencias debe tener un Rigger."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	theTable.AddText " 18. Requisitos de los Equipos. "
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true


	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, False, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 9. Funciones del Rigger."
	theTable.SelectCell(1)
	theTable.Frame False, False, False, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 19. Programa de Inspección y Mantenimiento."
	theTable.SelectCell(2)
	theTable.Frame False, False, False, true

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame False, true, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText " 10. Qué es un plan de Izaje."
	theTable.SelectCell(1)
	theTable.Frame False, true, true, true

	theTable.NextCell
	'theDoc.HPos = 0.015
	'theDoc.VPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, true, true, true


	theTable.NextRow
end if

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

	if(IdMutual<>2323)then
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Y certifica que asistió el siguiente participante obteniendo una calificación positiva:"
	else
	theTable.NextRow
	theTable.NextCell
	theTable.AddText "Y certifica que asistió al curso de Inducción Gold Fields dando cumplimiento al DS 40, Art. 21 el siguiente participante, obteniendo una calificación positiva:"
	end if

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow

	theDoc.FontSize = 11
	Set theTable = New Table
	theTable.Focus theDoc, 3
	theTable.Width(0) = 0.8
	theTable.Width(1) = 3
	theTable.Width(2) = 0.8

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
	theTable.NextCell
	theDoc.HPos = 0

	if(IdMutual=2317 or IdMutual=2319 or IdMutual=2324 or IdMutual=2325 or IdMutual=2332 or IdMutual=2334 or IdMutual=2335 or IdMutual=2318 or IdMutual=2341 or IdMutual=2342 or IdMutual=2343 or IdMutual=2326 or IdMutual=2320 or IdMutual=2321 or IdMutual=2322 or IdMutual=2333 or IdMutual=2330 or IdMutual=2331 or IdMutual=2327 or IdMutual=2328 or IdMutual=82 or IdMutual=2316 or IdMutual=2337 or IdMutual=2338 or IdMutual=2339 or IdMutual=2340 or IdMutual=2345 or IdMutual=2347)then
		theTable.AddText "Aprobado "&rsHistTrab("EstApro")&"%"
	else
		theTable.AddText "Aprobado"
	end if

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

	if(codCurso<>"HSSCOVID y HIC" and codCurso<>"GRM")then
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml ("<p align=""justify"">- Este certificado es válido sólo para documentar los requisitos de control de acceso a faenas y NO es válido para documentar franquicias Sence.</p>")
	end if

	If(codCurso<>"12.37.9742-66" and codCurso<>"12.380.081-54" and codCurso<>"VP_EX_04" and codCurso<>"VP_TA_03" and codCurso<>"VP_SU_05" and codCurso<>"VP_AB_01" and codCurso<>"VP_HT_06" and codCurso<>"VP_BI_02" and codCurso<>"VP_RE_07" and codCurso<>"VP_CAM08" and codCurso<>"VP_MD_10" and codCurso<>"VP_EI_11" and codCurso<>"VP_HTMP_12" and codCurso<>"VP_GPE_13" and codCurso<>"VP_GM_14" and codCurso<>"VP_PI_15" and codCurso<>"VP_EPPG_V16" and codCurso<>"VP_EPA_09" and codCurso<>"VP_AND_17" and codCurso<>"VP_RIG_20" and codCurso<>"AB_ ANGLO" and codCurso<>"PMCHS_2 Reglamento de Tránsito" and codCurso<>"12-38-005537" and codCurso<>"GF_OP_01" and codCurso<>"GF_OP_02" and codCurso<>"RL_ANGLO" and codCurso<>"GPM-MASC" and codCurso<>"ECF CODELCO" and codCurso<>"PMCM-MASC" and codCurso<>"TCAL_ANGLO" and codCurso<>"ESES_ANGLO" and codCurso<>"HSSCOVID y HIC" and codCurso<>"GRM" and codCurso<>"VP_TG MCCA_23" and codCurso<>"VP_EC_21" and codCurso<>"VP_TA_PRACTICO-Trabajo En Altura")then 

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
	end if
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	if(codCurso<>"2016-091" and codCurso<>"250619 MIEVITA" and codCurso<>"12.380.081-54" and codCurso<>"VP_EX_04" and codCurso<>"VP_TA_03" and codCurso<>"VP_SU_05" and codCurso<>"VP_AB_01" and codCurso<>"VP_HT_06" and codCurso<>"VP_BI_02" and codCurso<>"VP_RE_07" and codCurso<>"VP_CAM08" and codCurso<>"VP_MD_10" and codCurso<>"VP_EI_11" and codCurso<>"VP_HTMP_12" and codCurso<>"VP_GPE_13" and codCurso<>"VP_GM_14" and codCurso<>"VP_PI_15" and codCurso<>"VP_EPPG_V16" and codCurso<>"VP_EPA_09" and codCurso<>"VP_AND_17" and codCurso<>"PMCHS_1 Inducción especifica" and codCurso<>"VP_AR_19" and codCurso<>"PMCHS_2 Reglamento de Tránsito" and codCurso<>"PMCHS_4 Refugio Minero" and codCurso<>"PMCHS_3 Plan de emergencia"  and codCurso<>"VP_RIG_20" and codCurso<>"VP_IO_18" and codCurso<>"AB_ ANGLO" and codCurso<>"12-38-005537" and codCurso<>"GF_OP_01" and codCurso<>"GF_OP_02" and codCurso<>"RL_ANGLO" and codCurso<>"GPM-MASC" and codCurso<>"ECF CODELCO" and codCurso<>"VP_PRI_AU" and codCurso<>"PMCM-MASC" and codCurso<>"TCAL_ANGLO" and codCurso<>"ESES_ANGLO" and codCurso<>"HSSCOVID y HIC"  and codCurso<>"GRM" and codCurso<>"VP_TG MCCA_23" and codCurso<>"VP_TG MCCA_23" and codCurso<>"VP_EC_21" and codCurso<>"VP_TA_PRACTICO-Trabajo En Altura")then 
		theTable.NextRow
		theTable.NextCell
		theTable.AddText ""
	
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		
		if(codCurso<>"12.37.944-899" and codCurso<>"12.37.944-900")then
			if(codCurso<>"5-2-A09" and codCurso<>"TMERT-GM-15052018" and codCurso<>"MMC-GM-15052018" and codCurso<>"PNS-GM-15052018" and codCurso<>"PXOR-GM-15052018" and codCurso<>"PAUX-GM-15052018" and codCurso<>"PRUV-GM-15052018")then
				if(codCurso="12.37.9414-76" or codCurso="12-37-9741-16 ")then
					theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca BHP PAMPA NORTE y que pueden cambiar sin previo aviso.</p>")
				ElseIf(codCurso="26062019-MSG")then
					theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca Sierra Gorda SCM y que pueden cambiar sin previo aviso.</p>")	
				ElseIf(codCurso="12.37.9742-66")then
					theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca Mantos Blancos y que pueden cambiar sin previo aviso.</p>")		
				ElseIf(codCurso="GF_SSO")then
					theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca Gold Fields y que pueden cambiar sin previo aviso.</p>")		
				else
					theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca MEL y que pueden cambiar sin previo aviso.</p>")
				end if
			else
					theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca Codelco y que pueden cambiar sin previo aviso.</p>")		
						
			end if
		else
				theDoc.AddHtml ("<p align=""justify"">- La vigencia del certificado (y del respectivo curso) será determinada por las reglas que establezca Mantos Cooper y que pueden cambiar sin previo aviso.</p>")
		end if 
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	end if
	
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

	theDoc.FontSize = 8
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

Set FirmaRel = Server.CreateObject("ABCpdf7.Image")
if(codCurso<>"VP_CAM08-")then
		theDoc.FontSize = 9
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 4	

		theDoc.Rect = "210 5 380 95"
		theDoc.AddImageObject piePagina, False		

		theDoc.Rect = "150 10 450 55"
		theDoc.FontSize = 8
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 1	
							
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText "MARCELO MANCINI LOCH"	
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText "SUBGERENTE DE MINERIA"
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5	
		theTable.AddText "MUTUAL DE SEGURIDAD CAPACITACIÓN S.A."

else
		theDoc.FontSize = 9
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 4	

	FirmaRel.SetFile Server.MapPath("../images/FR-"&"18107554-6"&".jpg")		
	theDoc.Rect = "230 30 360 130"
	theDoc.AddImageObject FirmaRel, False
		

		theDoc.Rect = "150 10 450 65"
		theDoc.FontSize = 8
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 1	
							
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText "DIEGO MARINADO"

		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText "RELATOR"
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5	
		theTable.AddText "MUTUAL DE SEGURIDAD CAPACITACIÓN S.A."

end if	

end function

%>