<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->

<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
mes=""
idMutual="0"
fila=0
Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
Set theImgGob = Server.CreateObject("ABCpdf7.Image")
Set theImgMut = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoMutual.jpg")
theImgGob.SetFile Server.MapPath("../images/logoGob3.jpg")
theImgMut.SetFile Server.MapPath("../images/mutualSeg.jpg")
theDoc.Font = theDoc.EmbedFont("Times New Roman")
theDoc.FontSize =8

w = theDoc.MediaBox.Width
h = theDoc.MediaBox.Height
l = theDoc.MediaBox.Left
b = theDoc.MediaBox.Bottom 
theDoc.Transform.Rotate 90, l, b
theDoc.Transform.Translate w, 0

theDoc.Rect.Width = h
theDoc.Rect.Height = w

Tabla(theDoc)
theDoc.Rect = "50 70 740 380"
Tabla2(theDoc)
theDoc.Rect = "40 10 760 570"
Tabla3(theDoc)
theDoc.Rect = "30 50 760 570"
Tabla6(theDoc)
theDoc.Rect = "30 50 760 570"
Tabla7(theDoc)
theDoc.Rect = "30 50 760 570"
Tabla9(theDoc)


theID = theDoc.GetInfo(theDoc.Root, "Pages")
theDoc.SetInfo theID, "/Rotate", "90"

sArchivo = "../pdf/Libro_Clases_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function Tabla(theDoc)
	theDoc.FontSize = 20
	Set theTable = New Table
	theTable.Focus theDoc, 3
	theTable.Width(0) = 3
	theTable.Width(1) = 3
	theTable.Width(2) = 3
	theTable.Padding = 2	
	
	sql = "select * from INSTITUCION WHERE ID_INSTITUCION=1"
	set rs = conn.execute (sql)

	
	'/******* CABECERA *********/
	'theTable.Frame True, False, False, False
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theDoc.Rect = "50 550 240 450"
	theDoc.AddImageObject theImgGob, False
	'theDoc.AddHtml ("<b>  SENCEdedsfdfsdfs</b>")
	theTable.NextCell
	theDoc.HPos = 0.5
	theDoc.Rect = "250 530 470 400"
	theDoc.AddHtml ("<b>SENCE</b>")
	theDoc.Rect = "250 500 470 400"
	theDoc.FontSize = 16
	theTable.AddText "SERVICIO NACIONAL DE CAPACITACIÓN Y EMPLEO SENCE II REGIÓN ANTOFAGASTA"
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Rect = "630 550 500 500"	
	theDoc.AddImageObject theImg, False
	theDoc.Rect = "730 480 600 430"
	theDoc.AddImageObject theImgMut, False
end function

function Tabla2(theDoc)
	theDoc.FontSize = 12
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 1.8
	theTable.Width(1) = 4.2
	theTable.Padding = 2	
	
	'sql3 = "select distinct PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,CURRICULO.NOMBRE_CURSO,CURRICULO.DESCRIPCION, "
	'sql3 = sql3&" INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO as instructor, "
	'sql3 = sql3&" SEDES.NOMBRE as sede,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
	'sql3 = sql3&" CONVERT(VARCHAR(10),FECHA_TERMINO, 105) as FECHA_TERMINO "
	'sql3 = sql3&" from PROGRAMA "
	'sql3 = sql3&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
	'sql3 = sql3&" inner join HISTORICO_CURSOS on HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA "  
	'sql3 = sql3&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=HISTORICO_CURSOS.RELATOR "
	'sql3 = sql3&" inner join SEDES on SEDES.ID_SEDE=HISTORICO_CURSOS.SEDE "
	'sql3 = sql3&" where PROGRAMA.ID_PROGRAMA="&Request("prog")
	'sql3 = sql3&" and HISTORICO_CURSOS.RELATOR="&Request("relator")
	
	sql3="select p.ID_PROGRAMA,c.CODIGO,c.NOMBRE_CURSO,c.DESCRIPCION,ir.NOMBRES+' '+ir.A_PATERNO+' '+ir.A_MATERNO as instructor," 
 	sql3 = sql3&"CONVERT(VARCHAR(10),p.FECHA_INICIO_, 105) as FECHA_INICIO_,"
	sql3 = sql3&"CONVERT(VARCHAR(10),p.FECHA_TERMINO, 105) as FECHA_TERMINO,"
 	sql3 = sql3&"(CASE WHEN bq.id_sede =  27 THEN bq.nom_sede "
    sql3 = sql3&" WHEN bq.id_sede <>  27 THEN s.NOMBRE END) as sede,"
    sql3 = sql3&"(CASE WHEN bq.id_rel_seg IS NOT NULL THEN ' / '+(select ri.NOMBRES+' '+ri.A_PATERNO+' '+ri.A_MATERNO "
	sql3 = sql3&" from INSTRUCTOR_RELATOR ri where ri.ID_INSTRUCTOR=bq.id_rel_seg) "
    sql3 = sql3&" WHEN bq.id_rel_seg IS NULL THEN '' END) as rel_seg,c.ID_MUTUAL "
 	sql3 = sql3&"from bloque_programacion bq "
 	sql3 = sql3&" inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "
 	sql3 = sql3&" inner join CURRICULO c on c.ID_MUTUAL=p.ID_MUTUAL "
 	sql3 = sql3&" inner join INSTRUCTOR_RELATOR ir on ir.ID_INSTRUCTOR=bq.id_relator "
 	sql3 = sql3&" inner join SEDES s on s.ID_SEDE=bq.id_sede " 
 	sql3 = sql3&" where bq.id_relator='"&Request("relator")&"' and bq.id_programa='"&Request("prog")&"'"
	
	
	set rs3 = conn.execute (sql3)
	
	idMutual=rs3("ID_MUTUAL")
	
	'/******* CABECERA *********/
	theTable.Frame True, False, False, False
	theDoc.FrameRect

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0.1
	theDoc.FontSize = 18
	theDoc.AddHtml ("<B>LIBRO DE CONTROL DE CLASES</B>")
	theDoc.FontSize = 12
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "NOMBRE  DE  ACCIÓN DE CAPACITACIÓN"
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.FontSize = 14
	theDoc.AddHtml (": <B>"&UCase(rs3("DESCRIPCION"))&"</B>")
	theDoc.FontSize = 12
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "CÓDIGO DEL CURSO"
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml (": <B>"&rs3("CODIGO")&"</B>")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "FECHA DE EJECUCIÓN"
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml (": <B>FECHA INICIO "&rs3("FECHA_INICIO_")&" FECHA TÉRMINO "&rs3("FECHA_TERMINO")&"</B>")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "DIRECCIÓN DE EJECUCIÓN"
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml (": <B>"&UCase(rs3("sede"))&"</B>")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "HORARIO "
	theTable.NextCell
	theDoc.HPos = 0
	if(idMutual<>"40")then
	theDoc.AddHtml (": <B>MAÑANA 08:30 A 13:00  -  TARDE 14:30 A 18:00</B>")
	else
	theDoc.AddHtml (": <B>INGRESO A LAS 16:30, LUEGO 08:30 Y POR ÚLTIMO 14:30</B>")	
	end if
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "NOMBRE INSTRUCTOR (ES)"
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml (": <B>"&UCase(rs3("instructor"))&UCase(rs3("rel_seg"))&"</B>")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "NOMBRE  OTEC"
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.AddHtml (": <B>MUTUAL CAPACITACIÓN S.A.</B>")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0	
	theTable.AddText "NOMBRE DE LA OTIC, CUANDO CORRESPONDA"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""

    theDoc.Page = theDoc.AddPage()
end function

function Tabla3(theDoc)

if(idMutual="40")then
theDoc.FontSize = 12
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	
	mes=""
	
	if(month(now())=1)then
	mes="ENERO"
	end if
	
	if(month(now())=2)then
	mes="FEBRERO"
	end if
	
	if(month(now())=3)then
	mes="MARZO"
	end if
	
	if(month(now())=4)then
	mes="ABRIL"
	end if
	
	if(month(now())=5)then
	mes="MAYO"
	end if
	
	if(month(now())=6)then
	mes="JUNIO"
	end if
	
	if(month(now())=7)then
	mes="JULIO"
	end if
	
	if(month(now())=8)then
	mes="AGOSTO"
	end if
	
	if(month(now())=9)then
	mes="SEPTIEMBRE"
	end if
	
	if(month(now())=10)then
	mes="OCTUBRE"
	end if
	
	if(month(now())=11)then
	mes="NOVIEMBRE"
	end if
	
	if(month(now())=12)then
	mes="DICIEMBRE"
	end if
	
	theDoc.AddHtml ("<b>HOJA DE REGISTRO DE ASISTENCIA MES DE "&mes&"</b>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.NextRow
	theDoc.FontSize = 8
	
	
	Set theTable = New Table
	theTable.Focus theDoc, 7
	theTable.Width(0) = 0.4
	theTable.Width(1) = 2.2
	theTable.Width(2) = 0.8
	theTable.Width(3) = 1
	theTable.Width(4) = 1
	theTable.Width(5) = 1
	theTable.Width(6) = 1
	'theTable.Width(7) = 1
	theTable.Padding = 5

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, false, True, false
	theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(4)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "DÍAS"
	theTable.SelectCell(5)
	theTable.Frame True, True, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(6)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(7)
	theTable.Frame True, false, false, true
		theTable.Fill "174 174 174"
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame false, True, True, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame False, True, false, false
		theTable.Fill "174 174 174"		
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(4)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "JORNADA"
	theTable.SelectCell(5)
	theTable.Frame True, True, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(6)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(7)
	theTable.Frame True, false, false, true
		theTable.Fill "174 174 174"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, True, True, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame True, True, false, false
		theTable.Fill "174 174 174"		
	theTable.NextCell
	theDoc.HPos = 0.9
	theTable.AddText "FECHA"
	theTable.SelectCell(2)
	theTable.Frame False, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(5)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(6)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(7)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.9
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, True, True, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.9
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame True, True, false, false
		theTable.Fill "174 174 174"		
	theTable.NextCell
	theDoc.HPos = 0.9
	theTable.AddText "HORARIO"
	theTable.SelectCell(2)
	theTable.Frame False, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "16:30 P.M."
	theTable.SelectCell(4)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "8:30 A.M."
	theTable.SelectCell(5)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "14:30 P.M."
	theTable.SelectCell(6)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "% Asistencia"
	theTable.SelectCell(7)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, True, true, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "NOMBRE DEL PARTICIPANTE"
	theTable.SelectCell(1)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "CÉDULA"
	theTable.SelectCell(2)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Tarde"
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Mañana"
	theTable.SelectCell(5)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Tarde"
	theTable.SelectCell(6)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(7)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"

	'**** Tabla participantes ********
	sql3 = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.NACIONALIDAD,TRABAJADOR.ID_EXTRANJERO, "
	sql3 = sql3&"HISTORICO_CURSOS.ASISTENCIA,HISTORICO_CURSOS.CALIFICACION "
	sql3 = sql3&" from HISTORICO_CURSOS "
	sql3 = sql3&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
	sql3 = sql3&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA "
	sql3 = sql3&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
	sql3 = sql3&" where HISTORICO_CURSOS.ID_PROGRAMA="&Request("prog")
	sql3 = sql3&" and HISTORICO_CURSOS.RELATOR="&Request("relator")
	sql3 = sql3&" order by TRABAJADOR.NOMBRES asc"
	set rs3 = conn.execute (sql3)

	while not rs3.eof
		fila=fila+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText fila
		theTable.SelectCell(0)
		theTable.Frame True, True, True, True
			theTable.Fill "174 174 174"
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs3("NOMBRES")
		theTable.SelectCell(1)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 1
		theTable.AddText ""
		
		if(rs3("NACIONALIDAD")="0")then
			theTable.AddText replace(FormatNumber(mid(rs3("RUT"), 1,len(rs3("RUT"))-2),0)&mid(rs3("RUT"), len(rs3("RUT"))-1,len(rs3("RUT"))),",",".")
		else
			theTable.AddText rs3("ID_EXTRANJERO")
		end if
		
		theTable.SelectCell(2)
		theTable.Frame True, True, True, True		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(4)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(5)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(6)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(7)
		theTable.Frame True, True, True, True
		
		if(fila=18 or fila=36 or fila=54)then
			theDoc.Rect = "50 50 750 100"
			Pie(theDoc)
			theDoc.Rect = "40 10 760 570"
			theDoc.Page = theDoc.AddPage()
			theDoc.FontSize = 12
			
			Set theTable = New Table
			theTable.Focus theDoc, 1
			theTable.Width(0) = 8
			theTable.Padding = 4
		
			'/******* CABECERA *********/
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			
			mes=""
			
			if(month(now())=1)then
			mes="ENERO"
			end if
			
			if(month(now())=2)then
			mes="FEBRERO"
			end if
			
			if(month(now())=3)then
			mes="MARZO"
			end if
			
			if(month(now())=4)then
			mes="ABRIL"
			end if
			
			if(month(now())=5)then
			mes="MAYO"
			end if
			
			if(month(now())=6)then
			mes="JUNIO"
			end if
			
			if(month(now())=7)then
			mes="JULIO"
			end if
			
			if(month(now())=8)then
			mes="AGOSTO"
			end if
			
			if(month(now())=9)then
			mes="SEPTIEMBRE"
			end if
			
			if(month(now())=10)then
			mes="OCTUBRE"
			end if
			
			if(month(now())=11)then
			mes="NOVIEMBRE"
			end if
			
			if(month(now())=12)then
			mes="DICIEMBRE"
			end if
			
			theDoc.AddHtml ("<b>HOJA DE REGISTRO DE ASISTENCIA MES DE "&mes&"</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.NextRow
			theDoc.FontSize = 8
			
			
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.4
			theTable.Width(1) = 2.2
			theTable.Width(2) = 0.8
			theTable.Width(3) = 1
			theTable.Width(4) = 1
			theTable.Width(5) = 1
			theTable.Width(6) = 1
			'theTable.Width(7) = 1
			theTable.Padding = 5
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, false, True, false
			theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(2)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "DÍAS"
			theTable.SelectCell(5)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, false, false, true
				theTable.Fill "174 174 174"
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame false, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame False, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "JORNADA"
			theTable.SelectCell(5)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, false, false, true
				theTable.Fill "174 174 174"
		
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText "FECHA"
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText "HORARIO"
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "16:30 P.M."
			theTable.SelectCell(4)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "8:30 A.M."
			theTable.SelectCell(5)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "14:30 P.M."
			theTable.SelectCell(6)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "% Asistencia"
			theTable.SelectCell(7)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
		
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, true, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NOMBRE DEL PARTICIPANTE"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "CÉDULA"
			theTable.SelectCell(2)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tarde"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Mañana"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tarde"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
		end if
	  rs3.Movenext
	wend

		'if(fila>1)then
		For cuen = fila+1 To 30 Step 1
		fila=fila+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText fila
		theTable.SelectCell(0)
		theTable.Frame True, True, True, True
			theTable.Fill "174 174 174"
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(1)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(2)
		theTable.Frame True, True, True, True		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(4)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(5)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(6)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(7)
		theTable.Frame True, True, True, True
			
		  if(fila=18)then
			theDoc.Rect = "50 50 750 100"
			Pie(theDoc)
			theDoc.Rect = "40 10 760 570"
			theDoc.Page = theDoc.AddPage()
			theDoc.FontSize = 12
			
			Set theTable = New Table
			theTable.Focus theDoc, 1
			theTable.Width(0) = 8
			theTable.Padding = 4
		
			'/******* CABECERA *********/
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			
			mes=""
			
			if(month(now())=1)then
			mes="ENERO"
			end if
			
			if(month(now())=2)then
			mes="FEBRERO"
			end if
			
			if(month(now())=3)then
			mes="MARZO"
			end if
			
			if(month(now())=4)then
			mes="ABRIL"
			end if
			
			if(month(now())=5)then
			mes="MAYO"
			end if
			
			if(month(now())=6)then
			mes="JUNIO"
			end if
			
			if(month(now())=7)then
			mes="JULIO"
			end if
			
			if(month(now())=8)then
			mes="AGOSTO"
			end if
			
			if(month(now())=9)then
			mes="SEPTIEMBRE"
			end if
			
			if(month(now())=10)then
			mes="OCTUBRE"
			end if
			
			if(month(now())=11)then
			mes="NOVIEMBRE"
			end if
			
			if(month(now())=12)then
			mes="DICIEMBRE"
			end if
			
			theDoc.AddHtml ("<b>HOJA DE REGISTRO DE ASISTENCIA MES DE "&mes&"</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.NextRow
			theDoc.FontSize = 8
			
			
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.4
			theTable.Width(1) = 2.2
			theTable.Width(2) = 0.8
			theTable.Width(3) = 1
			theTable.Width(4) = 1
			theTable.Width(5) = 1
			theTable.Width(6) = 1
			'theTable.Width(7) = 1
			theTable.Padding = 5
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, false, True, false
			theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(2)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "DÍAS"
			theTable.SelectCell(5)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, false, false, true
				theTable.Fill "174 174 174"
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame false, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame False, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "JORNADA"
			theTable.SelectCell(5)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, false, false, true
				theTable.Fill "174 174 174"
		
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText "FECHA"
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText "HORARIO"
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "16:30 P.M."
			theTable.SelectCell(4)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "8:30 A.M."
			theTable.SelectCell(5)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "14:30 P.M."
			theTable.SelectCell(6)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "% Asistencia"
			theTable.SelectCell(7)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
		
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, true, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NOMBRE DEL PARTICIPANTE"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "CÉDULA"
			theTable.SelectCell(2)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tarde"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Mañana"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tarde"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
		end if
		Next

	theDoc.Rect = "50 50 750 100"
	Pie(theDoc)
	
	theDoc.Page = theDoc.AddPage()
	
else

	theDoc.FontSize = 12
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	
	mes=""
	
	if(month(now())=1)then
	mes="ENERO"
	end if
	
	if(month(now())=2)then
	mes="FEBRERO"
	end if
	
	if(month(now())=3)then
	mes="MARZO"
	end if
	
	if(month(now())=4)then
	mes="ABRIL"
	end if
	
	if(month(now())=5)then
	mes="MAYO"
	end if
	
	if(month(now())=6)then
	mes="JUNIO"
	end if
	
	if(month(now())=7)then
	mes="JULIO"
	end if
	
	if(month(now())=8)then
	mes="AGOSTO"
	end if
	
	if(month(now())=9)then
	mes="SEPTIEMBRE"
	end if
	
	if(month(now())=10)then
	mes="OCTUBRE"
	end if
	
	if(month(now())=11)then
	mes="NOVIEMBRE"
	end if
	
	if(month(now())=12)then
	mes="DICIEMBRE"
	end if
	
	theDoc.AddHtml ("<b>HOJA DE REGISTRO DE ASISTENCIA MES DE "&mes&"</b>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.NextRow
	theDoc.FontSize = 8
	
	
	Set theTable = New Table
	theTable.Focus theDoc, 8
	theTable.Width(0) = 0.4
	theTable.Width(1) = 2.2
	theTable.Width(2) = 0.8
	theTable.Width(3) = 1
	theTable.Width(4) = 1
	theTable.Width(5) = 1
	theTable.Width(6) = 1
	theTable.Width(7) = 1
	theTable.Padding = 5

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, false, True, false
	theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(3)
	theTable.Frame True, false, True, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(4)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "DÍAS"
	theTable.SelectCell(5)
	theTable.Frame True, True, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(6)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(7)
	theTable.Frame True, false, false, true
		theTable.Fill "174 174 174"
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame false, True, True, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame False, True, false, false
		theTable.Fill "174 174 174"		
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(2)
	theTable.Frame False, True, false, True
		theTable.Fill "174 174 174"
    theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(3)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(4)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "JORNADA"
	theTable.SelectCell(5)
	theTable.Frame True, True, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(6)
	theTable.Frame True, false, false, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(7)
	theTable.Frame True, false, false, true
		theTable.Fill "174 174 174"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, True, True, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame True, True, false, false
		theTable.Fill "174 174 174"		
	theTable.NextCell
	theDoc.HPos = 0.9
	theTable.AddText "FECHA"
	theTable.SelectCell(2)
	theTable.Frame False, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(5)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(6)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(7)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.9
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, True, True, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.9
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame True, True, false, false
		theTable.Fill "174 174 174"		
	theTable.NextCell
	theDoc.HPos = 0.9
	theTable.AddText "HORARIO"
	theTable.SelectCell(2)
	theTable.Frame False, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "8:30 A.M."
	theTable.SelectCell(3)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "14:30 P.M."
	theTable.SelectCell(4)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "8:30 A.M."
	theTable.SelectCell(5)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "14:30 P.M."
	theTable.SelectCell(6)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "% Asistencia"
	theTable.SelectCell(7)
	theTable.Frame False, True, True, True
		theTable.Fill "174 174 174"

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, True, true, false
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "NOMBRE DEL PARTICIPANTE"
	theTable.SelectCell(1)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "CÉDULA"
	theTable.SelectCell(2)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"		
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Mañana"
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Tarde"
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Mañana"
	theTable.SelectCell(5)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Tarde"
	theTable.SelectCell(6)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText ""
	theTable.SelectCell(7)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"

	'**** Tabla participantes ********
	sql3 = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.NACIONALIDAD,TRABAJADOR.ID_EXTRANJERO, "
	sql3 = sql3&"HISTORICO_CURSOS.ASISTENCIA,HISTORICO_CURSOS.CALIFICACION "
	sql3 = sql3&" from HISTORICO_CURSOS "
	sql3 = sql3&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
	sql3 = sql3&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA "
	sql3 = sql3&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
	sql3 = sql3&" where HISTORICO_CURSOS.ID_PROGRAMA="&Request("prog")
	sql3 = sql3&" and HISTORICO_CURSOS.RELATOR="&Request("relator")
	sql3 = sql3&" order by TRABAJADOR.NOMBRES asc"
	set rs3 = conn.execute (sql3)

	while not rs3.eof
		fila=fila+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText fila
		theTable.SelectCell(0)
		theTable.Frame True, True, True, True
			theTable.Fill "174 174 174"
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs3("NOMBRES")
		theTable.SelectCell(1)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 1
		theTable.AddText ""
		
		if(rs3("NACIONALIDAD")="0")then
			theTable.AddText replace(FormatNumber(mid(rs3("RUT"), 1,len(rs3("RUT"))-2),0)&mid(rs3("RUT"), len(rs3("RUT"))-1,len(rs3("RUT"))),",",".")
		else
			theTable.AddText rs3("ID_EXTRANJERO")
		end if
		
		theTable.SelectCell(2)
		theTable.Frame True, True, True, True		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(3)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(4)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(5)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(6)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(7)
		theTable.Frame True, True, True, True
		
		if(fila=18 or fila=36 or fila=54)then
			theDoc.Rect = "50 50 750 100"
			Pie(theDoc)
			theDoc.Rect = "40 10 760 570"
			theDoc.Page = theDoc.AddPage()
			theDoc.FontSize = 12
			
			Set theTable = New Table
			theTable.Focus theDoc, 1
			theTable.Width(0) = 8
			theTable.Padding = 4
		
			'/******* CABECERA *********/
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			
			mes=""
			
			if(month(now())=1)then
			mes="ENERO"
			end if
			
			if(month(now())=2)then
			mes="FEBRERO"
			end if
			
			if(month(now())=3)then
			mes="MARZO"
			end if
			
			if(month(now())=4)then
			mes="ABRIL"
			end if
			
			if(month(now())=5)then
			mes="MAYO"
			end if
			
			if(month(now())=6)then
			mes="JUNIO"
			end if
			
			if(month(now())=7)then
			mes="JULIO"
			end if
			
			if(month(now())=8)then
			mes="AGOSTO"
			end if
			
			if(month(now())=9)then
			mes="SEPTIEMBRE"
			end if
			
			if(month(now())=10)then
			mes="OCTUBRE"
			end if
			
			if(month(now())=11)then
			mes="NOVIEMBRE"
			end if
			
			if(month(now())=12)then
			mes="DICIEMBRE"
			end if
			
			theDoc.AddHtml ("<b>HOJA DE REGISTRO DE ASISTENCIA MES DE "&mes&"</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.NextRow
			theDoc.FontSize = 8
			
			
			Set theTable = New Table
			theTable.Focus theDoc, 8
			theTable.Width(0) = 0.4
			theTable.Width(1) = 2.2
			theTable.Width(2) = 0.8
			theTable.Width(3) = 1
			theTable.Width(4) = 1
			theTable.Width(5) = 1
			theTable.Width(6) = 1
			theTable.Width(7) = 1
			theTable.Padding = 5
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, false, True, false
			theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(2)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(3)
			theTable.Frame True, false, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "DÍAS"
			theTable.SelectCell(5)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, false, false, true
				theTable.Fill "174 174 174"
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame false, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame False, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(3)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "JORNADA"
			theTable.SelectCell(5)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, false, false, true
				theTable.Fill "174 174 174"
		
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText "FECHA"
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText "HORARIO"
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "8:30 A.M."
			theTable.SelectCell(3)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "14:30 P.M."
			theTable.SelectCell(4)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "8:30 A.M."
			theTable.SelectCell(5)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "14:30 P.M."
			theTable.SelectCell(6)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "% Asistencia"
			theTable.SelectCell(7)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
		
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, true, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NOMBRE DEL PARTICIPANTE"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "CÉDULA"
			theTable.SelectCell(2)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Mañana"
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tarde"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Mañana"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tarde"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
		end if
	  rs3.Movenext
	wend

		'if(fila>1)then
		For cuen = fila+1 To 30 Step 1
		fila=fila+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText fila
		theTable.SelectCell(0)
		theTable.Frame True, True, True, True
			theTable.Fill "174 174 174"
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(1)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(2)
		theTable.Frame True, True, True, True		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(3)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(4)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(5)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(6)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(7)
		theTable.Frame True, True, True, True
			
		  if(fila=18)then
			theDoc.Rect = "50 50 750 100"
			Pie(theDoc)
			theDoc.Rect = "40 10 760 570"
			theDoc.Page = theDoc.AddPage()
			theDoc.FontSize = 12
			
			Set theTable = New Table
			theTable.Focus theDoc, 1
			theTable.Width(0) = 8
			theTable.Padding = 4
		
			'/******* CABECERA *********/
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			
			mes=""
			
			if(month(now())=1)then
			mes="ENERO"
			end if
			
			if(month(now())=2)then
			mes="FEBRERO"
			end if
			
			if(month(now())=3)then
			mes="MARZO"
			end if
			
			if(month(now())=4)then
			mes="ABRIL"
			end if
			
			if(month(now())=5)then
			mes="MAYO"
			end if
			
			if(month(now())=6)then
			mes="JUNIO"
			end if
			
			if(month(now())=7)then
			mes="JULIO"
			end if
			
			if(month(now())=8)then
			mes="AGOSTO"
			end if
			
			if(month(now())=9)then
			mes="SEPTIEMBRE"
			end if
			
			if(month(now())=10)then
			mes="OCTUBRE"
			end if
			
			if(month(now())=11)then
			mes="NOVIEMBRE"
			end if
			
			if(month(now())=12)then
			mes="DICIEMBRE"
			end if
			
			theDoc.AddHtml ("<b>HOJA DE REGISTRO DE ASISTENCIA MES DE "&mes&"</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.NextRow
			theDoc.FontSize = 8
			
			
			Set theTable = New Table
			theTable.Focus theDoc, 8
			theTable.Width(0) = 0.4
			theTable.Width(1) = 2.2
			theTable.Width(2) = 0.8
			theTable.Width(3) = 1
			theTable.Width(4) = 1
			theTable.Width(5) = 1
			theTable.Width(6) = 1
			theTable.Width(7) = 1
			theTable.Padding = 5
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, false, True, false
			theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(2)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(3)
			theTable.Frame True, false, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "DÍAS"
			theTable.SelectCell(5)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, false, false, true
				theTable.Fill "174 174 174"
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame false, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame False, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(3)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "JORNADA"
			theTable.SelectCell(5)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, false, false, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, false, false, true
				theTable.Fill "174 174 174"
		
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText "FECHA"
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, True, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText ""
			theTable.SelectCell(1)
			theTable.Frame True, True, false, false
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.9
			theTable.AddText "HORARIO"
			theTable.SelectCell(2)
			theTable.Frame False, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "8:30 A.M."
			theTable.SelectCell(3)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "14:30 P.M."
			theTable.SelectCell(4)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "8:30 A.M."
			theTable.SelectCell(5)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "14:30 P.M."
			theTable.SelectCell(6)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "% Asistencia"
			theTable.SelectCell(7)
			theTable.Frame False, True, True, True
				theTable.Fill "174 174 174"
		
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			theTable.SelectCell(0)
			theTable.Frame True, True, true, false
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NOMBRE DEL PARTICIPANTE"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "CÉDULA"
			theTable.SelectCell(2)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"		
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Mañana"
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tarde"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Mañana"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tarde"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText ""
			theTable.SelectCell(7)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
		end if
		Next

	theDoc.Rect = "50 50 750 100"
	Pie(theDoc)
	
	theDoc.Page = theDoc.AddPage()
	
	end if
end function

function Pie(theDoc)
	theDoc.FontSize = 10
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 2

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.AddText " Nota 1: El participante  debe firmar su asistencia en el casillero respectivo. No se aceptaran  enmendaduras ni correcciones."
	theTable.SelectCell(0)
	theTable.Frame True, FALSE, True, True
	
	theTable.NextRow
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText " Nota 2: La asistencia deberá ser registrada como máximo a los 20 minutos del comienzo del horario normal de clases."
	theTable.SelectCell(0)
	theTable.Frame FALSE, True, True, True
end function

function Tabla6(theDoc)
	theDoc.FontSize = 12
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theDoc.AddHtml ("<CENTER><b>DESCRIPCIÓN DE LAS ACTIVIDADES REALIZADAS</b></CENTER>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow

	theDoc.FontSize = 8
	Set theTable = New Table
	theTable.Focus theDoc, 5
	theTable.Width(0) = 1
	theTable.Width(1) = 3
	theTable.Width(2) = 3
	theTable.Width(3) = 1
	theTable.Width(4) = 1.5

	theTable.Padding = 10

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "FECHA"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "TEMA"
	theTable.SelectCell(1)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "ACTIVIDAD"
	theTable.SelectCell(2)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "HORAS"
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "FIRMA INSTRUCTOR"
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"

	'**** Tabla participantes ********
	
	theTable.Padding = 6
	for i=0 to 17 step 1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText ""
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(2)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(3)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(4)
		theTable.Frame True, True, true, true
	next


	theDoc.Page = theDoc.AddPage()
end function

function Tabla7(theDoc)
	theDoc.FontSize = 12
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theDoc.AddHtml ("<b>ANTECEDENTES DE LOS PARTICIPANTES</b>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow

	theDoc.FontSize = 8
	Set theTable = New Table
	theTable.Focus theDoc, 7
	theTable.Width(0) = 0.5
	theTable.Width(1) = 3
	theTable.Width(2) = 1.1
	theTable.Width(3) = 2
	theTable.Width(4) = 2.2
	theTable.Width(5) = 2
	theTable.Width(6) = 1

	theTable.Padding = 10

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "N°"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "NOMBRE APELLIDOS"
	theTable.SelectCell(1)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "CÉDULA IDENTIDAD"
	theTable.SelectCell(2)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "NIVEL ESCOLARIDAD"
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "EMPRESA"
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "CARGO DESEMPEÑADO"
	theTable.SelectCell(5)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "FIRMA"
	theTable.SelectCell(6)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"

	'**** Tabla participantes ********
	
	sql3 = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD,TRABAJADOR.NACIONALIDAD,"
	sql3 = sql3&"TRABAJADOR.ID_EXTRANJERO,dbo.MayMinTexto (EMPRESAS.R_SOCIAL) as R_SOCIAL "
	sql3 = sql3&" from HISTORICO_CURSOS "
	sql3 = sql3&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
	sql3 = sql3&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA "
	sql3 = sql3&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
	sql3 = sql3&" where HISTORICO_CURSOS.ID_PROGRAMA="&Request("prog")
	sql3 = sql3&" and HISTORICO_CURSOS.RELATOR="&Request("relator")
	sql3 = sql3&" order by TRABAJADOR.NOMBRES asc"
	set rs3 = conn.execute (sql3)
	
	theTable.Padding = 4
	participante=0
	while not rs3.eof
		participante=participante+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText participante
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs3("NOMBRES")
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 1
		
		if(rs3("NACIONALIDAD")="0")then
			theTable.AddText replace(FormatNumber(mid(rs3("RUT"), 1,len(rs3("RUT"))-2),0)&mid(rs3("RUT"), len(rs3("RUT"))-1,len(rs3("RUT"))),",",".")
		else
			theTable.AddText rs3("ID_EXTRANJERO")
		end if
		
		theTable.SelectCell(2)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		
		if(rs3("ESCOLARIDAD")=0)then
		escolaridad="Sin Escolaridad"
		end if
		
		if(rs3("ESCOLARIDAD")=1)then
		escolaridad="Básica Incompleta"
		end if
		
		if(rs3("ESCOLARIDAD")=2)then
		escolaridad="Básica Completa"
		end if


		if(rs3("ESCOLARIDAD")=3)then
		escolaridad="Media Incompleta"
		end if

		if(rs3("ESCOLARIDAD")=4)then
		escolaridad="Media Completa"
		end if
		
		if(rs3("ESCOLARIDAD")=5)then
		escolaridad="Superior Técnica Incompleta"
		end if

		if(rs3("ESCOLARIDAD")=6)then
		escolaridad="Superior Técnica Profesional Completa"
		end if

		if(rs3("ESCOLARIDAD")=7)then
		escolaridad="Universitaria Incompleta"
		end if
		
		if(rs3("ESCOLARIDAD")=8)then
		escolaridad="Universitaria Completa"
		end if
		
		theTable.AddText escolaridad
		theTable.SelectCell(3)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs3("R_SOCIAL")
		theTable.SelectCell(4)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs3("CARGO_EMPRESA")
		theTable.SelectCell(5)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(6)
		theTable.Frame True, True, true, true
		
		if(participante=25 or participante=50)then
			theDoc.Rect = "30 50 760 570"
		    theDoc.Page = theDoc.AddPage()
			theDoc.FontSize = 12
	
			Set theTable = New Table
			theTable.Focus theDoc, 1
			theTable.Width(0) = 8
			theTable.Padding = 4
		
			'/******* CABECERA *********/
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theDoc.Font = theDoc.EmbedFont("Arial Black")
		
			theDoc.AddHtml ("<b>ANTECEDENTES DE LOS PARTICIPANTES</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			
			theTable.NextRow
		
			theDoc.FontSize = 8
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.5
			theTable.Width(1) = 3
			theTable.Width(2) = 1.1
			theTable.Width(3) = 2
			theTable.Width(4) = 2.2
			theTable.Width(5) = 2
			theTable.Width(6) = 1
		
			theTable.Padding = 10
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "N°"
			theTable.SelectCell(0)
			theTable.Frame True, True, true, true
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NOMBRE APELLIDOS"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "CÉDULA IDENTIDAD"
			theTable.SelectCell(2)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NIVEL ESCOLARIDAD"
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "EMPRESA"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "CARGO DESEMPEÑADO"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "FIRMA"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
				
			theTable.Padding = 4
		end if
		
	rs3.Movenext
	wend
	
	For cuen = participante+1 To 30 Step 1
	    participante=participante+1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText participante
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 1
		
		theTable.AddText ""
		
		theTable.SelectCell(2)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
	
		theTable.AddText ""
		theTable.SelectCell(3)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(4)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(5)
		theTable.Frame True, True, true, true
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(6)
		theTable.Frame True, True, true, true
		
		if(participante=25)then
			theDoc.Rect = "30 50 760 570"
		    theDoc.Page = theDoc.AddPage()
			theDoc.FontSize = 12
	
			Set theTable = New Table
			theTable.Focus theDoc, 1
			theTable.Width(0) = 8
			theTable.Padding = 4
		
			'/******* CABECERA *********/
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theDoc.Font = theDoc.EmbedFont("Arial Black")
		
			theDoc.AddHtml ("<b>ANTECEDENTES DE LOS PARTICIPANTES</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			
			theTable.NextRow
		
			theDoc.FontSize = 8
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.5
			theTable.Width(1) = 3
			theTable.Width(2) = 1.1
			theTable.Width(3) = 2
			theTable.Width(4) = 2.2
			theTable.Width(5) = 2
			theTable.Width(6) = 1
		
			theTable.Padding = 10
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "N°"
			theTable.SelectCell(0)
			theTable.Frame True, True, true, true
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NOMBRE APELLIDOS"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "CÉDULA IDENTIDAD"
			theTable.SelectCell(2)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "NIVEL ESCOLARIDAD"
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "EMPRESA"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "CARGO DESEMPEÑADO"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "FIRMA"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
				
			theTable.Padding = 4
		end if
	Next
	
	theDoc.Page = theDoc.AddPage()
	theDoc.Page = theDoc.AddPage()
end function

function Tabla9(theDoc)
theDoc.FontSize = 12
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextRow
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theDoc.AddHtml ("<b>HOJA DE FISCALIZACIÓN - SENCE</b>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow
	theDoc.FontSize = 9
	Set theTable = New Table
	theTable.Focus theDoc, 5
	theTable.Width(0) = 1
	theTable.Width(1) = 1
	theTable.Width(2) = 2
	theTable.Width(3) = 2
	theTable.Width(4) = 2

	theTable.Padding = 16

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "FECHA"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "HORA"
	theTable.SelectCell(1)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "NOMBRE DEL FISCALIZADOR"
	theTable.SelectCell(2)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "OBSERVACIONES"
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "FIRMA DEL FISCALIZADOR"
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"

	'**** Tabla participantes ********
	
	theTable.Padding = 6
	for i=0 to 17 step 1
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText ""
		theTable.SelectCell(0)
		theTable.Frame false, false, True, false
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(1)
		theTable.Frame false, false, True, false
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(2)
		theTable.Frame false, false, True, false
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(3)
		theTable.Frame True, True, True, True
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(4)
		theTable.Frame false, false, false, True
		
		if(i=6 or i=12)then
		theTable.SelectRow theTable.Row
		theTable.Frame True, False, False, False
		end if
	next
	theTable.NextRow
	theTable.SelectRow theTable.Row
	theTable.Frame True, False, False, False
end function
%>