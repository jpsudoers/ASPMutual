<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
Set piePagina = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoMutual.jpg")
piePagina.SetFile Server.MapPath("../images/logoCertSence.jpg")

dim IdAuto
IdAuto=Request("IdAuto")

theDoc.Font = theDoc.EmbedFont("Arial Black")

theDoc.Rect = "100 50 570 700"
Tabla(theDoc)

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
theDoc.FontSize =12
For i = 1 To theCount
	  theDoc.PageNumber = i
	  theDoc.TextStyle.Justification = 1
	  theDoc.AddHtml("<b>CERTIFICADO DE ASISTENCIA ACTIVIDAD OTEC, CFT O ENTIDAD NIVELADORA DE ESTUDIOS, IMPUTADA EN FORMA TOTAL O PARCIAL A FRANQUICIA TRIBUTARIA DE CAPACITACIÓN</b>")
Next

'/******** FOOTER **********/
theCount = theDoc.PageCount
For i = 1 To theCount
	theDoc.Rect = "510 30 590 50"
	theDoc.HPos = 1.0
	theDoc.FontSize = 8
 	theDoc.PageNumber = i	
  	theDoc.AddText "Página " & i & " de " & theCount
	theDoc.Rect = "80 10 530 70"
	theDoc.TextStyle.Justification = 1
    theDoc.AddImageObject piePagina, False
Next

sArchivo = "../pdf/Certificado_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function Tabla(theDoc)
	theDoc.FontSize = 10
	
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 0.3
	theTable.Width(1) = 7.7
	theTable.Padding = 4

	'/******* CABECERA *********/
	
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.AddText "X"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	
	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "   Actividad dentro del año calendario"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, False
	
	theTable.Padding = 0
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theTable.AddText ""
	
	theTable.Padding = 4
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	
	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "   Actividad parcial"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, False
	
	theTable.Padding = 0
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theTable.AddText ""
	
	theTable.Padding = 4
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.AddText ""
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
	
	theDoc.Font = theDoc.EmbedFont("Arial")

	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "   Actividad complementaria"
	theTable.SelectCell(1)
	theTable.Frame False, False, False, False

'	theTable.NextRow	
	
    theDoc.Rect = "70 50 570 580"	
	
	theDoc.FontSize = 10
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Se extiende el presente certificado de asistencia correspondiente a la actividad de capacitación que a continuación se señala:"
	
	dim datosProg
datosProg="select E.RUT as rutEmpresa,UPPER(E.R_SOCIAL) as nomEmpresa,C.DESCRIPCION,C.CODIGO,CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"
datosProg=datosProg&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,C.HORAS,A.N_REG_SENCE,"
datosProg=datosProg&"(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN "  
datosProg=datosProg&"(select eotic.RUT from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) " 
datosProg=datosProg&" ELSE 'No Aplica' end) as rutOtic,"
datosProg=datosProg&"(CASE WHEN A.CON_FRANQUICIA=1 and A.CON_OTIC=1 and A.ID_OTIC<>0 THEN "  
datosProg=datosProg&"(select dbo.MayMinTexto (eotic.R_SOCIAL) "
datosProg=datosProg&" from EMPRESAS eotic where eotic.ID_EMPRESA=A.ID_OTIC) " 
datosProg=datosProg&" ELSE 'No Aplica' end) as nomOtic from AUTORIZACION A "
datosProg=datosProg&" INNER JOIN PROGRAMA P ON P.ID_PROGRAMA=A.ID_PROGRAMA "
datosProg=datosProg&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "
datosProg=datosProg&" INNER JOIN EMPRESAS E ON E.ID_EMPRESA=A.ID_EMPRESA "
datosProg=datosProg&" where A.ID_AUTORIZACION='"&IdAuto&"'"

set rsDatosProg = conn.execute (datosProg)

set rsMutual = conn.execute ("SELECT RUT,R_SOCIAL,CONTACTO,RUT_CONTACTO FROM INSTITUCION")
	
	theDoc.Rect = "30 50 570 545"	
	
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.FontSize = 8
	
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 3.1
    theTable.Width(1) = 4.9
	theTable.Padding = 4	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Razón social OTEC, CFT o entidad niveladora"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsMutual("R_SOCIAL")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "RUT OTEC, CFT o entidad niveladora"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsMutual("RUT")
	'theTable.AddText replace(FormatNumber(mid(rsMutual("RUT"), 1,len(rsMutual("RUT"))-2),0)&mid(rsMutual("RUT"), len(rsMutual("RUT"))-1,len(rsMutual("RUT"))),",",".")		
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Razón social empresa                                              "
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ucase(rsDatosProg("nomEmpresa"))
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "RUT empresa"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true		

	theTable.NextCell
	theDoc.HPos = 0
	'theTable.AddText rsDatosProg("RUT")
	theTable.AddText replace(FormatNumber(mid(rsDatosProg("rutEmpresa"), 1,len(rsDatosProg("rutEmpresa"))-2),0)&mid(rsDatosProg("rutEmpresa"), len(rsDatosProg("rutEmpresa"))-1,len(rsDatosProg("rutEmpresa"))),",",".")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Razón social OTIC (si corresponde a actividad intermediada por éste)"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ucase(rsDatosProg("nomOtic"))
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "RUT OTIC (si corresponde a actividad intermediada por éste)"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true		

	theTable.NextCell
	theDoc.HPos = 0
	'theTable.AddText rsDatosProg("RUT")
	if(rsDatosProg("rutOtic")<>"No Aplica")then
		theTable.AddText replace(FormatNumber(mid(rsDatosProg("rutOtic"), 1,len(rsDatosProg("rutOtic"))-2),0)&mid(rsDatosProg("rutOtic"), len(rsDatosProg("rutOtic"))-1,len(rsDatosProg("rutOtic"))),",",".")
	else
		theTable.AddText ucase(rsDatosProg("rutOtic"))
	end if
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Nombre de la actividad (1)                                           "
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsDatosProg("DESCRIPCION")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Código Sence"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsDatosProg("CODIGO")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true		
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Fecha inicio"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsDatosProg("FECHA_INICIO")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Fecha de término"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsDatosProg("FECHA_TERMINO")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Nº de horas (para actividades parciales o complementarias indicar número efectivo de horas realizadas en el año correspondiente)"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText rsDatosProg("HORAS")
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Nº de factura"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Nº registro de acción Sence "
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true		
	
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Participantes:"
	
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""			

	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	
	theTable.NextCell
	theTable.AddText ""	
	
	theTable.NextRow
	
	theDoc.Font = theDoc.EmbedFont("Arial Black")
		theDoc.FontSize = 8
	Set theTable = New Table
	theTable.Focus theDoc, 6
	theTable.Width(0) = 0.38
    theTable.Width(1) = 1
	theTable.Width(2) = 1.8
    theTable.Width(3) = 1.8
	theTable.Width(4) = 2.72
    theTable.Width(5) = 1.1	
	theTable.Padding = 4	
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Nº                "
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "RUT"
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true	

	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Apellido paterno"
	theTable.SelectCell(2)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Apellido materno"
	theTable.SelectCell(3)
	theTable.Frame True, True, true, true		
	
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Nombres"
	theTable.SelectCell(4)
	theTable.Frame True, True, true, true	
	
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Porcentaje de asistencia"
	theTable.SelectCell(5)
	theTable.Frame True, True, true, true	

	dim datosTrab
	datosTrab="select (CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',"
	datosTrab=datosTrab&"T.APATERTRAB,T.AMATERTRAB,T.NOM_TRAB,H.ASISTENCIA,NACIONALIDAD from HISTORICO_CURSOS H "
	datosTrab=datosTrab&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR " 
	datosTrab=datosTrab&" WHERE H.ID_AUTORIZACION='"&IdAuto&"'"

	set rsTrab = conn.execute (datosTrab)

	theDoc.Font = theDoc.EmbedFont("Arial")

	cont=0
	
	while not rsTrab.eof
				cont=cont+1
					if(cont=3 or cont=31)then
					  'if(cont=10)then
						theDoc.Page = theDoc.AddPage()
						theDoc.Rect = "30 50 570 670"	
						theDoc.Font = theDoc.EmbedFont("Arial Black")
						theDoc.FontSize = 8
						Set theTable = New Table
						theTable.Focus theDoc, 6
						theTable.Width(0) = 0.38
						theTable.Width(1) = 1
						theTable.Width(2) = 1.8
						theTable.Width(3) = 1.8
						theTable.Width(4) = 2.72
						theTable.Width(5) = 1.1	
						theTable.Padding = 4	
						
						theTable.NextRow
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Nº                "
						theTable.SelectCell(0)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "RUT"
						theTable.SelectCell(1)
						theTable.Frame True, True, true, true	
					
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Apellido paterno"
						theTable.SelectCell(2)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Apellido materno"
						theTable.SelectCell(3)
						theTable.Frame True, True, true, true		
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Nombres"
						theTable.SelectCell(4)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Porcentaje de asistencia"
						theTable.SelectCell(5)
						theTable.Frame True, True, true, true	
						
						theDoc.Font = theDoc.EmbedFont("Arial")					
					'end if	
				end if				
			    theTable.NextRow
				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText cont
				theTable.SelectCell(0)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 1
				
				if(rsTrab("NACIONALIDAD")="0")then
					theTable.AddText replace(FormatNumber(mid(rsTrab("TrabId"), 1,len(rsTrab("TrabId"))-2),0)&mid(rsTrab("TrabId"), len(rsTrab("TrabId"))-1,len(rsTrab("TrabId"))),",",".")	
				else
					theTable.AddText rsTrab("TrabId")
				end if			

				theTable.SelectCell(1)
				theTable.Frame True, True, true, true	
			
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsTrab("APATERTRAB")
				theTable.SelectCell(2)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsTrab("AMATERTRAB")
				theTable.SelectCell(3)
				theTable.Frame True, True, true, true		
				
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsTrab("NOM_TRAB")
				theTable.SelectCell(4)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText ""&rsTrab("ASISTENCIA")
				theTable.SelectCell(5)
				theTable.Frame True, True, true, true	
				

			rsTrab.Movenext
	wend	

		theDoc.Rect = "30 50 570 180"
	
		theDoc.Font = theDoc.EmbedFont("Arial Black")
		theDoc.FontSize = 7
		
		Set theTable = New Table
		theTable.Focus theDoc, 4
		theTable.Width(0) = 2
		theTable.Width(1) = 3
		theTable.Width(2) = 1
		theTable.Width(3) = 2
		theTable.Padding = 4	
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText "Firma representante legal OTEC, CFT o entidad niveladora"
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true	
		
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText ""
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true	
	
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText ""
		theTable.SelectCell(2)
		theTable.Frame False, False, False, False	
		
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText ""
		theTable.SelectCell(3)
		theTable.Frame False, False, False, False	
		
			theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText "Nombre representante legal OTEC, CFT o entidad niveladora"
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true	
		
		theTable.NextCell
		theDoc.VPos = 0.5
		theTable.AddText rsMutual("CONTACTO")
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true	
	
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText ""
		theTable.SelectCell(2)
		theTable.Frame False, False, False, False	
		
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText ""
		theTable.SelectCell(3)
		theTable.Frame False, False, False, False
	
		theDoc.VPos = 0	
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText "RUT representante legal OTEC, CFT o entidad niveladora"
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true	
		
		theTable.NextCell
		theDoc.HPos = 0
		theDoc.VPos = 0.5	
		theTable.AddText rsMutual("RUT_CONTACTO")
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true	
	
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText ""
		theTable.SelectCell(2)
		theTable.Frame False, False, False, False	
		
		theTable.NextCell
		theDoc.HPos = 0.5
		theDoc.VPos = 0.5
			theDoc.Font = theDoc.EmbedFont("Arial")
		theTable.AddText "Fecha emisión: "&right("0"&day(now()),2)&"/"&right("0"&month(now()),2)&"/"&year(now)
		theTable.SelectCell(3)
		theTable.Frame True, True, true, true
		
		theDoc.VPos = 0
			theTable.NextRow
		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText ""
	
		theTable.NextCell
		theTable.AddText ""
	
		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextRow
		
		Set theTable = New Table
		theTable.Focus theDoc, 3
		theTable.Width(0) = 1
		theTable.Width(1) = 2.5
		theTable.Width(2) = 4.5
		theTable.Padding = 2			
		
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""	
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""	
		theTable.SelectCell(1)
		theTable.Frame False, True, False, False	
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""	
end function
%>