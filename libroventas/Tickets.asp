<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set marco = Server.CreateObject("ABCpdf7.Image")
marco.SetFile Server.MapPath("../images/marcos6.jpg")

dim IdAuto
IdAuto=Request("id")
NTickets=Request("ntickets")
nfactura=Request("nfactura")

conn.execute ("UPDATE AUTORIZACION set N_TICKETS='"&NTickets&"' WHERE N_TICKETS is null and ID_AUTORIZACION='"&IdAuto&"'")

theDoc.Font = theDoc.EmbedFont("Arial")
theDoc.Rect = "80 50 500 770"
Tabla(theDoc)

theCount = theDoc.PageCount
'/******** HEADER **********/
'theDoc.Rect = "30 750 150 785"
'theDoc.FontSize = 8
For i = 1 To theCount
  theDoc.PageNumber = i
  'theDoc.AddImageObject theImg, False
Next

'/******** FOOTER **********/
'theCount = theDoc.PageCount
'For i = 1 To theCount
	'theDoc.Rect = "70 110 170 210"
	'theDoc.TextStyle.Justification = 1
    'theDoc.AddImageObject piePagina, False
'Next

sArchivo = "../pdf/Tickets_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function Tabla(theDoc)
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4	
	
	theDoc.FontSize = 12
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 1
	theTable.AddText "N° de Factura : "&nfactura
	
	theTable.NextRow
	
	theDoc.Rect = "80 50 500 750"
		
    theDoc.AddImageObject marco, False

	theDoc.FontSize = 18
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 29

	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theDoc.Color = "255 255 255"
		'theDoc.Rect.Inset 0, 0
	'theDoc.FrameRect
	'theDoc.AddHtml("<p><font color='white'>Notificación de Ingreso de Tickets</font></p>")
	theTable.AddText "Notificación de Ingreso de Tickets"
	theTable.SelectCell(0)
	'theTable.Fill "0 0 0"
	'theTable.Frame True, True, true, true

	theDoc.Color = "0 0 0"
	
	theTable.Padding = 8
	theDoc.FontSize = 10

	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	'theDoc.TextStyle.LineSpacing = 20
	'theDoc.AddHtml("<p><font color='black'>Estimado (a): CARRASCO,</font></p>")
	theTable.AddText "Estimado (a): CARRASCO,"
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, true	
	
	theDoc.Font = theDoc.EmbedFont("Arial")	
	'theTable.NextRow
	'theTable.NextCell
	'theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, true		
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText "Junto con saludar, informamos que hemos recibido su solicitud y se registró con los siguientes datos,"
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, true		
	
	theTable.NextRow
	
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 2
	theTable.Width(1) = 6
	theTable.Padding = 8	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText "N° Ticket"
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False		
	theTable.NextCell
	theTable.AddText NTickets
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true			

	theTable.NextRow
	theTable.NextCell
	theTable.AddText "Nombre"
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "LAURA CARRASCO MENESES"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true			

'theDoc.TextStyle.LineSpacing = 20


	theTable.NextRow
	theTable.NextCell
	theTable.AddText "Email"
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theDoc.TextStyle.Underline = True
	theDoc.Color = "0 0 255"	
	theTable.AddText "lcarrasco@mutual.cl"	
	'theTable.Padding = 25
	'theDoc.AddHtml("<br/>&nbsp;&nbsp;&nbsp;<a href='mailto:lcarrasco@mutual.cl'>lcarrasco@mutual.cl</a>")
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	theDoc.TextStyle.Underline = False
	set rsRegSence = conn.execute ("SELECT A.N_REG_SENCE FROM AUTORIZACION A WHERE A.ID_AUTORIZACION="&IdAuto)
	
	theDoc.Color = "0 0 0"
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText "Asunto..."
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "Consultas Sence | ID de acción "&rsRegSence("N_REG_SENCE")&" | 76410180-4"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText ""	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "Nueva solicitud de soporte Sence "	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "--------------------------------------------"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText "Descripción..."
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "OPERADOR"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "Nombre: Laura Carrasco"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "Rut: 12490739-K"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "Email: lcarrasco@mutual.cl"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText ""	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "__________________________________________"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true					
	
	theTable.NextRow

	theDoc.FontSize = 10
	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 0.2
	theTable.Width(1) = 7.8
	theTable.Padding = 8

	'/******* CABECERA *********/
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""	
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	'theTable.AddText 
	theDoc.AddHtml("<font color='#FF9900'>NOTICIA:</font> Lo invitamos a conocer nuestro sitio web de <font color='#0033CC'><u>cursos e-learning</u></font>. Encontrará paso a paso como utilizar el sistema, así como la descripción de las tareas a realizar previas al inicio de la toma de asistencia.")
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, true		
theTable.Padding = 40
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""	
	theTable.NextCell
	theTable.AddText ""	
	


	Set theTable = New Table
	theTable.Focus theDoc, 2
	theTable.Width(0) = 2.8
	theTable.Width(1) = 5.2
	theTable.Padding = 4	

	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText ""	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "Mesa de Ayuda"
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theDoc.Font = theDoc.EmbedFont("Arial")	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText ""	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theDoc.Color = "83 86 90"
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "Tel.: (56-2) 24968130 opción 2"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theDoc.TextStyle.Underline = True
	theDoc.Color = "0 0 255"
		
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "lce@sence.cl"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText "www.sence.cl"	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText ""	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true	
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, False, true, False	
	theTable.NextCell
	theTable.AddText ""	
	'theTable.SelectCell(1)
	'theTable.Frame False, False, False, true		
	
	theTable.NextRow
	theTable.NextCell
	theTable.AddText ""
	'theTable.SelectCell(0)
	'theTable.Frame False, True, true, False	
	theTable.NextCell
	theTable.AddText ""	
	'theTable.SelectCell(1)
	'theTable.Frame False, True, False, true		
end function
%>