<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

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

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
Set piePagina = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoMutual.jpg")
piePagina.SetFile Server.MapPath("../images/FIRMA_AA.jpg")

dim IdAuto
IdAuto=Request("IdAuto")

theDoc.Font = theDoc.EmbedFont("Arial")


InsMac="SELECT count(*) as total FROM AUTORIZACION A "&_
       " INNER JOIN EMPRESAS E ON E.ID_EMPRESA=A.ID_EMPRESA "&_
       " INNER JOIN PROGRAMA P ON P.ID_PROGRAMA=A.ID_PROGRAMA "&_
       " WHERE E.ID_EMPRESA='"&Request("emp")&"' AND P.FECHA_INICIO_<CONVERT(DATE,'01-06-2009') and P.ID_MUTUAL='"&Request("curso")&"'"

set rsInsMac = conn.execute (InsMac)

if(cint(rsInsMac("total"))>0)then
	theDoc.Rect = "50 50 550 680"
	TablaMac(theDoc)
	
	theDoc.Rect = "50 50 550 680"
	Tabla(theDoc)	
else
	theDoc.Rect = "50 50 550 680"
	Tabla(theDoc)
end if

theCount = theDoc.PageCount
'/******** HEADER **********/
theDoc.Rect = "30 750 150 785"
theDoc.FontSize = 8
For i = 1 To theCount
  theDoc.PageNumber = i
  theDoc.AddImageObject theImg, False
Next

theDoc.Rect = "90 685 520 730"
theDoc.HPos = 0.5
'theDoc.VPos = 0.5
theDoc.FontSize =18
For i = 1 To theCount
	  theDoc.Font = theDoc.EmbedFont("Arial Black")
	  theDoc.PageNumber = i
	  theDoc.TextStyle.Justification = 1
	  theDoc.AddHtml("<b>CERTIFICADO</b>")
Next

'/******** FOOTER **********/
theCount = theDoc.PageCount
For i = 1 To theCount
	theDoc.Rect = "510 30 590 50"
	theDoc.HPos = 1.0
	theDoc.FontSize = 8
 	theDoc.PageNumber = i	
  	theDoc.AddText "Página " & i & " de " & theCount
	'theDoc.Rect = "80 10 530 70"
	'theDoc.TextStyle.Justification = 1
    'theDoc.AddImageObject piePagina, False
Next

sArchivo = "../pdf/Cert_Masivo_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function Tabla(theDoc)
	theDoc.FontSize = 12
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4	
						
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	
	set rsEmp = conn.execute ("select UPPER (E.R_SOCIAL) as R_SOCIAL, E.RUT from empresas e where e.tipo=1 and e.id_empresa='"&Request("emp")&"'")
	
	set rsCurr = conn.execute ("SELECT UPPER (C.NOMBRE_CURSO) AS NOMBRE_CURSO, C.HORAS FROM CURRICULO C WHERE C.ID_MUTUAL='"&Request("curso")&"'")
	
	theDoc.AddHtml("<p align=""justify""><b>ALEX AGUILERA OLIVARES</b>, Director de Capacitación Mutual de Seguridad Capacitación, certifica que los siguientes trabajadores de la empresa <b>"&rsEmp("R_SOCIAL")&" R.U.T. "&replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")&"</b> aprobaron el Curso <b>"""&rsCurr("NOMBRE_CURSO")&"""</b>, con duración de "&rsCurr("HORAS")&" horas.</p>")
	
	'theDoc.Rect = "100 50 500 600"
    theDoc.Rect = "70 50 530 600"
	
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.FontSize = 9
	Set theTable = New Table
						theTable.Focus theDoc, 5
						theTable.Width(0) = 0.38
						theTable.Width(1) = 0.8
						theTable.Width(2) = 3
						theTable.Width(3) = 0.9						
					    theTable.Width(4) = 0.9						
						theTable.Padding = 4	
						
						theTable.NextRow
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Nº"
						theTable.SelectCell(0)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "RUT"
						theTable.SelectCell(1)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "NOMBRE"
						theTable.SelectCell(2)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "F. CURSO"
						theTable.SelectCell(3)
						theTable.Frame True, True, true, true	
												
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "VIGENCIA"
						theTable.SelectCell(4)
						theTable.Frame True, True, true, true						

	trab=""
	
	emp=""
	
	if(Request("emp")<>"0" and Request("emp")<>"")then
		emp=" and H.ID_EMPRESA='"&Request("emp")&"'"
	end if
	
	curso=" and P.ID_MUTUAL='"&Request("curso")&"'"

	sql = "select UPPER (T.NOMBRES) as NOMBRES,C.NOMBRE_CURSO, "
	sql = sql&"(CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',T.NACIONALIDAD,"
	sql = sql&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,H.ID_PROGRAMA,H.RELATOR,"
	sql = sql&"H.ID_TRABAJADOR,(CASE WHEN H.COD_AUTENFICACION is null then 'No Aplica' "
	sql = sql&" WHEN H.COD_AUTENFICACION is not null then H.COD_AUTENFICACION END) as 'CODIGO',"
	sql = sql&"CONVERT(VARCHAR(10),DATEADD(MONTH, (CASE " 
	sql = sql&" WHEN C.VIGENCIA =  1 THEN 6 "
	sql = sql&" WHEN C.VIGENCIA =  2 THEN 12 "
	sql = sql&" WHEN C.VIGENCIA =  3 THEN 18 "
	sql = sql&" WHEN C.VIGENCIA =  4 THEN 24 "
	sql = sql&" WHEN C.VIGENCIA =  5 THEN 48 "
	sql = sql&"	WHEN C.VIGENCIA =  6 THEN 36 "	
	sql = sql&" END), CONVERT(date,DATEADD(DAY, 0, CONVERT(date,P.FECHA_TERMINO, 105)), 105)), 105) as vigencia,"
	sql = sql&"(CASE WHEN CONVERT(date,GETDATE(), 105)<=DATEADD(MONTH, (CASE " 
	sql = sql&"		 WHEN C.VIGENCIA =  1 THEN 6 "
	sql = sql&" 	 WHEN C.VIGENCIA =  2 THEN 12 "
	sql = sql&"		 WHEN C.VIGENCIA =  3 THEN 18 "
	sql = sql&"		 WHEN C.VIGENCIA =  4 THEN 24 "
	sql = sql&"		 WHEN C.VIGENCIA =  5 THEN 48 "
	sql = sql&"		 WHEN C.VIGENCIA =  6 THEN 36 "
	sql = sql&"	     END), CONVERT(date,DATEADD(DAY, 1, CONVERT(date,P.FECHA_TERMINO, 105)), 105)) THEN '#000' "
	sql = sql&"WHEN CONVERT(date,GETDATE(), 105)>DATEADD(MONTH, (CASE " 
	sql = sql&"		 WHEN C.VIGENCIA =  1 THEN 6 "
	sql = sql&"		 WHEN C.VIGENCIA =  2 THEN 12 "
	sql = sql&"		 WHEN C.VIGENCIA =  3 THEN 18 "
	sql = sql&"		 WHEN C.VIGENCIA =  4 THEN 24 "
	sql = sql&"		 WHEN C.VIGENCIA =  5 THEN 48 "
	sql = sql&"		 WHEN C.VIGENCIA =  6 THEN 36 "	
	sql = sql&"		 END), CONVERT(date,DATEADD(DAY, 1, CONVERT(date,P.FECHA_TERMINO, 105)), 105)) THEN '#ff0000' "
	sql = sql&"END) as color "  
	sql = sql&"  from HISTORICO_CURSOS H "
	sql = sql&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "
	sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=H.ID_PROGRAMA "
	sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "
	sql = sql&" where P.FECHA_INICIO_>=CONVERT(DATE,'01-06-2011') and H.EVALUACION<>'Reprobado' and H.ESTADO=0 "&trab&emp&curso&" ORDER BY P.FECHA_TERMINO,T.NOMBRES ASC"


	set rsTrab = conn.execute (sql)

	theDoc.Font = theDoc.EmbedFont("Arial")

	cont=0
	n_trab=1
	
	while not rsTrab.eof
				cont=cont+1
					if(cont=31)then
						theDoc.Page = theDoc.AddPage()
						theDoc.Rect = "70 50 530 670"
						theDoc.Font = theDoc.EmbedFont("Arial Black")
						theDoc.FontSize = 9
						Set theTable = New Table
						theTable.Focus theDoc, 5
						theTable.Width(0) = 0.38
						theTable.Width(1) = 0.8
						theTable.Width(2) = 3
						theTable.Width(3) = 0.9						
					    theTable.Width(4) = 0.9	
						theTable.Padding = 4	
						
						theTable.NextRow
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Nº"
						theTable.SelectCell(0)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "RUT"
						theTable.SelectCell(1)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "NOMBRE"
						theTable.SelectCell(2)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "F. CURSO"
						theTable.SelectCell(3)
						theTable.Frame True, True, true, true	
												
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "VIGENCIA"
						theTable.SelectCell(4)
						theTable.Frame True, True, true, true							
						
						theDoc.Font = theDoc.EmbedFont("Arial")					
					cont=1
				end if				
			    theTable.NextRow
				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText n_trab
				theTable.SelectCell(0)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 1
				theTable.AddText rsTrab("TrabId")
				theTable.SelectCell(1)
				theTable.Frame True, True, true, true	
			
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsTrab("NOMBRES")
				theTable.SelectCell(2)
				theTable.Frame True, True, true, true	

				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText rsTrab("FECHA_TERMINO")
				theTable.SelectCell(3)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText rsTrab("vigencia")
				theTable.SelectCell(4)
				theTable.Frame True, True, true, true					
				
				n_trab=n_trab+1
			rsTrab.Movenext
	wend	

		if(cont<=15)then
			theDoc.Rect = "50 50 550 300"
		else
			theDoc.Page = theDoc.AddPage()	
			theDoc.Rect = "50 50 550 300"
		end if		

		theDoc.FontSize = 12
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 4	
							
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theDoc.AddHtml("")'<p align=""justify"">Se extiende el presente certificado a solicitud de empresa CELULOSA ARAUCO Y CONSTITUCIÓN S.A.</p>	
		
		theDoc.Rect = "200 163 390 280"
		theDoc.AddImageObject piePagina, False		

		theDoc.Rect = "200 20 400 185"
		theDoc.FontSize = 9
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 1	
							
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText "ALEX AGUILERA OLIVARES"	
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText "DIRECTOR DE CAPACITACIÓN"
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5	
		theTable.AddText "MUTUAL DE SEGURIDAD CAPACITACIÓN S.A."	

		theDoc.Rect = "50 20 550 120"

		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 1	
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0	
		theTable.AddText "Antofagasta, "&cadena			
		
		'theTable.NextRow
		'theTable.NextCell
		'theDoc.HPos = 0	
		'theTable.AddText "005/ 11"
		
		'theTable.NextRow
		'theTable.NextCell
		'theDoc.HPos = 0	
		'theTable.AddText "FCP/cjm."									
		
end function

function TablaMac(theDoc)
	theDoc.FontSize = 12
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4	
						
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	
	set rsEmp = conn.execute ("select UPPER (E.R_SOCIAL) as R_SOCIAL, E.RUT from empresas e where e.tipo=1 and e.id_empresa='"&Request("emp")&"'")
	
	set rsCurr = conn.execute ("SELECT UPPER (C.NOMBRE_CURSO) AS NOMBRE_CURSO, C.HORAS FROM CURRICULO C WHERE C.ID_MUTUAL='"&Request("curso")&"'")
	
	theDoc.AddHtml("<p align=""justify"">En base a información recibida de Celulosa Arauco y Constitución S.A., el siguiente listado de personas pertenecientes a la empresa <b>"&rsEmp("R_SOCIAL")&" R.U.T. "&replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")&"</b> aprobaron el Curso <b>"""&rsCurr("NOMBRE_CURSO")&"""</b>, desarrollado por la empresa <b>MA&C CONSULTORES LIMITADA</b>.</p>")
	
	'theDoc.Rect = "100 50 500 600"
    theDoc.Rect = "70 50 530 600"
	
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.FontSize = 9
	Set theTable = New Table
						theTable.Focus theDoc, 5
						theTable.Width(0) = 0.38
						theTable.Width(1) = 0.8
						theTable.Width(2) = 3
						theTable.Width(3) = 0.9						
					    theTable.Width(4) = 0.9						
						theTable.Padding = 4	
						
						theTable.NextRow
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Nº"
						theTable.SelectCell(0)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "RUT"
						theTable.SelectCell(1)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "NOMBRE"
						theTable.SelectCell(2)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "F. CURSO"
						theTable.SelectCell(3)
						theTable.Frame True, True, true, true	
												
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "VIGENCIA"
						theTable.SelectCell(4)
						theTable.Frame True, True, true, true						

	trab=""
	
	emp=""
	
	if(Request("emp")<>"0" and Request("emp")<>"")then
		emp=" and H.ID_EMPRESA='"&Request("emp")&"'"
	end if
	
	curso=" and P.ID_MUTUAL='"&Request("curso")&"'"

	sql = "select UPPER (T.NOMBRES) as NOMBRES,C.NOMBRE_CURSO, "
	sql = sql&"(CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',T.NACIONALIDAD,"
	sql = sql&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,H.ID_PROGRAMA,H.RELATOR,"
	sql = sql&"H.ID_TRABAJADOR,(CASE WHEN H.COD_AUTENFICACION is null then 'No Aplica' "
	sql = sql&" WHEN H.COD_AUTENFICACION is not null then H.COD_AUTENFICACION END) as 'CODIGO',"
	sql = sql&"CONVERT(VARCHAR(10),DATEADD(MONTH, (CASE " 
	sql = sql&" WHEN C.VIGENCIA =  1 THEN 6 "
	sql = sql&" WHEN C.VIGENCIA =  2 THEN 12 "
	sql = sql&" WHEN C.VIGENCIA =  3 THEN 18 "
	sql = sql&" WHEN C.VIGENCIA =  4 THEN 24 "
	sql = sql&" WHEN C.VIGENCIA =  5 THEN 48 "
	sql = sql&"	WHEN C.VIGENCIA =  6 THEN 36 "	
	sql = sql&" END), CONVERT(date,DATEADD(DAY, 0, CONVERT(date,P.FECHA_TERMINO, 105)), 105)), 105) as vigencia,"
	sql = sql&"(CASE WHEN CONVERT(date,GETDATE(), 105)<=DATEADD(MONTH, (CASE " 
	sql = sql&"		 WHEN C.VIGENCIA =  1 THEN 6 "
	sql = sql&" 	 WHEN C.VIGENCIA =  2 THEN 12 "
	sql = sql&"		 WHEN C.VIGENCIA =  3 THEN 18 "
	sql = sql&"		 WHEN C.VIGENCIA =  4 THEN 24 "
	sql = sql&"		 WHEN C.VIGENCIA =  5 THEN 48 "
	sql = sql&"		 WHEN C.VIGENCIA =  6 THEN 36 "	
	sql = sql&"	     END), CONVERT(date,DATEADD(DAY, 1, CONVERT(date,P.FECHA_TERMINO, 105)), 105)) THEN '#000' "
	sql = sql&"WHEN CONVERT(date,GETDATE(), 105)>DATEADD(MONTH, (CASE " 
	sql = sql&"		 WHEN C.VIGENCIA =  1 THEN 6 "
	sql = sql&"		 WHEN C.VIGENCIA =  2 THEN 12 "
	sql = sql&"		 WHEN C.VIGENCIA =  3 THEN 18 "
	sql = sql&"		 WHEN C.VIGENCIA =  4 THEN 24 "
	sql = sql&"		 WHEN C.VIGENCIA =  5 THEN 48 "
	sql = sql&"		 WHEN C.VIGENCIA =  6 THEN 36 "	
	sql = sql&"		 END), CONVERT(date,DATEADD(DAY, 1, CONVERT(date,P.FECHA_TERMINO, 105)), 105)) THEN '#ff0000' "
	sql = sql&"END) as color "  
	sql = sql&"  from HISTORICO_CURSOS H "
	sql = sql&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR "
	sql = sql&" inner join PROGRAMA P on P.ID_PROGRAMA=H.ID_PROGRAMA "
	sql = sql&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "
	sql = sql&" where P.FECHA_INICIO_<CONVERT(DATE,'01-06-2011') and H.EVALUACION<>'Reprobado' and H.ESTADO=0 "&trab&emp&curso&" ORDER BY P.FECHA_TERMINO,T.NOMBRES ASC"


	set rsTrab = conn.execute (sql)

	theDoc.Font = theDoc.EmbedFont("Arial")

	cont=0
	n_trab=1
	
	while not rsTrab.eof
				cont=cont+1
					if(cont=31)then
						theDoc.Page = theDoc.AddPage()
						theDoc.Rect = "70 50 530 670"
						theDoc.Font = theDoc.EmbedFont("Arial Black")
						theDoc.FontSize = 9
						Set theTable = New Table
						theTable.Focus theDoc, 5
						theTable.Width(0) = 0.38
						theTable.Width(1) = 0.8
						theTable.Width(2) = 3
						theTable.Width(3) = 0.9						
					    theTable.Width(4) = 0.9	
						theTable.Padding = 4	
						
						theTable.NextRow
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Nº"
						theTable.SelectCell(0)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "RUT"
						theTable.SelectCell(1)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "NOMBRE"
						theTable.SelectCell(2)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "F. CURSO"
						theTable.SelectCell(3)
						theTable.Frame True, True, true, true	
												
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "VIGENCIA"
						theTable.SelectCell(4)
						theTable.Frame True, True, true, true							
						
						theDoc.Font = theDoc.EmbedFont("Arial")					
					cont=1
				end if				
			    theTable.NextRow
				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText n_trab
				theTable.SelectCell(0)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 1
				theTable.AddText rsTrab("TrabId")
				theTable.SelectCell(1)
				theTable.Frame True, True, true, true	
			
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsTrab("NOMBRES")
				theTable.SelectCell(2)
				theTable.Frame True, True, true, true	

				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText rsTrab("FECHA_TERMINO")
				theTable.SelectCell(3)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText rsTrab("vigencia")
				theTable.SelectCell(4)
				theTable.Frame True, True, true, true					
				
				n_trab=n_trab+1
			rsTrab.Movenext
	wend	

		if(cont<=15)then
			theDoc.Rect = "50 50 550 300"
		else
			theDoc.Page = theDoc.AddPage()	
			theDoc.Rect = "50 50 550 300"
		end if		

		theDoc.FontSize = 12
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 4	
							
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theDoc.AddHtml("<p align=""justify"">Se extiende el presente certificado a solicitud de empresa CELULOSA ARAUCO Y CONSTITUCIÓN S.A.</p>")	
		
		theDoc.Rect = "210 190 390 240"
		theDoc.AddImageObject theImg, False	

		theDoc.Rect = "200 20 400 185"
		theDoc.FontSize = 9
		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 1	
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5	
		theTable.AddText "MUTUAL DE SEGURIDAD CAPACITACIÓN S.A."	

		theDoc.Rect = "50 20 550 120"

		Set theTable = New Table
		theTable.Focus theDoc, 1
		theTable.Width(0) = 8
		theTable.Padding = 1	
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0	
		theTable.AddText "Chillán, "&cadena			
		
		theDoc.Page = theDoc.AddPage()		
end function
%>