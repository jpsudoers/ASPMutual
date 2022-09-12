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
datosProg="select DESCRIPCION=(CASE WHEN c.ID_MUTUAL=70 THEN c.NOMBRE_CURSO else C.DESCRIPCION end),C.CODIGO,HORAS=(case when p.ID_Modalidad=1 then C.HORAS else "
datosProg=datosProg&" (DATEDIFF(hour, CONVERT(datetime, P.FECHA_INICIO_ + ' ' +  "
datosProg=datosProg&" (select hb.HORARIO from horario_bloques hb where hb.ID_HORARIO=P.BMI), 21),CONVERT(datetime,  "
datosProg=datosProg&" P.FECHA_INICIO_ + ' ' + (select hb.HORARIO from horario_bloques hb  "
datosProg=datosProg&" where hb.ID_HORARIO=P.BTF),21))*(DATEDIFF(day, P.FECHA_INICIO_, P.FECHA_TERMINO)+1)) END),CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"
datosProg=datosProg&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,(CASE WHEN C.VIGENCIA='1' then '6' "
datosProg=datosProg&" WHEN C.VIGENCIA='2' then '12' WHEN C.VIGENCIA='3' then '18' WHEN C.VIGENCIA='4' then '24' "
datosProg=datosProg&" WHEN C.VIGENCIA='5' then '48' END) as 'VIGENCIA',p.id_mutual from PROGRAMA P "
datosProg=datosProg&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL WHERE P.ID_PROGRAMA="&Request("prog")

set rsDatosProg = conn.execute (datosProg)





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



sArchivo = "../pdf/CERTIFICADO_"&nomArchivo&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)



%>