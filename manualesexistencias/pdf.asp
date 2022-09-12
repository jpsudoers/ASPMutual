<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->

<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

	dim mes(12) 
	mes(0)="Enero"
	mes(1)="Febrero"
	mes(2)="Marzo"
	mes(3)="Abril"
	mes(4)="Mayo"
	mes(5)="Junio"
	mes(6)="Julio"
	mes(7)="Agosto"
	mes(8)="Septiembre"
	mes(9)="Octubre"
	mes(10)="Noviembre"
	mes(11)="Diciembre"

	fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
	fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
	
	Set theDoc = Server.CreateObject("ABCpdf7.Doc")
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

	theDoc.Rect = "30 50 760 570"
	TablaIngresos(theDoc)

	theCount = theDoc.PageCount
	For i = 1 To theCount
		theDoc.Rect = "650 40 750 60"
		theDoc.HPos = 1.0
		theDoc.VPos = 0.5
		theDoc.FontSize = 8
	  theDoc.PageNumber = i	
	  theDoc.AddText "Página " & i & " de " & theCount
	Next

theID = theDoc.GetInfo(theDoc.Root, "Pages")
theDoc.SetInfo theID, "/Rotate", "90"

sArchivo = "../pdf/Informe_Existencias_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function TablaIngresos(theDoc)
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 8

	'/******* CABECERA *********/
	
	theDoc.FontSize = 12
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theDoc.AddHtml ("<b>Informe de Existencias</b>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow

	theDoc.FontSize = 7
	Set theTable = New Table
	theTable.Focus theDoc, 4
	theTable.Width(0) = 0.3
	theTable.Width(1) = 5.5
	theTable.Width(2) = 1
	theTable.Width(3) = 1

	theTable.Padding = 10

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText " Items                                                                                       "
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
		theTable.Fill "174 174 174"	
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Nombre Articulo"
	theTable.SelectCell(1)
	theTable.Frame True, True, true, true
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Salidas desde el "&mid(request("f_ini"),1,2)&" de "&mes(mid(request("f_ini"),4,2)-1)&" de "&mid(request("f_ini"),7,4)&" Hasta el "&mid(request("f_fin"),1,2)&" de "&mes(mid(request("f_fin"),4,2)-1)&" de "&mid(request("f_fin"),7,4)
	theTable.SelectCell(2)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Stock Actual (Hasta el "&right("0"&day(now()),2)&" de "&mes(month(now())-1)&" de "&year(now)&")"
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	

	'**** Tabla participantes ********
	
	qsMan="select a.DESC_ARTICULO as 'Nombre Articulo',isnull((SELECT SUM(CANTIDAD) "
	qsMan=qsMan&" FROM MOVIMIENTOS m "
	qsMan=qsMan&" inner join bloque_programacion bq on bq.id_bloque=m.ID_PROG_BLOQUE "
	qsMan=qsMan&" inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "
	qsMan=qsMan&" WHERE (m.TIPO_MOVIMIENTO = 3) and m.ID_ARTICULO=a.ID_ARTICULO "
	qsMan=qsMan&" and CONVERT(date,p.FECHA_INICIO_)>=CONVERT(date,'"&request("f_ini")&"') "
	qsMan=qsMan&" and CONVERT(date,p.FECHA_INICIO_)<=CONVERT(date,'"&request("f_fin")&"')),0) as 'S_Hasta', "
	qsMan=qsMan&" isnull((SELECT SUM(CANTIDAD) FROM MOVIMIENTOS "
	qsMan=qsMan&" WHERE (TIPO_MOVIMIENTO = 1) and ID_ARTICULO=a.ID_ARTICULO),0)+ "
	qsMan=qsMan&" isnull((SELECT SUM(CANTIDAD) FROM MOVIMIENTOS "
	qsMan=qsMan&" WHERE (TIPO_MOVIMIENTO = 3) and ID_ARTICULO=a.ID_ARTICULO),0) as 'S_Actual' from ARTICULOS a WHERE A.ESTADO_ARTICULO=1"

	set rsMan =  conn.execute(qsMan)
	
	theTable.Padding = 4

	items=0
	while not rsMan.eof
				items=items+1
				theTable.NextRow
				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText items
				theTable.SelectCell(0)
				theTable.Frame True, True, true, true
									
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsMan("Nombre Articulo")
				theTable.SelectCell(1)
				theTable.Frame True, True, true, true
				
				theTable.NextCell
				theDoc.HPos = 1
				theTable.AddText replace(FormatNumber(rsMan("S_Hasta"),0),",",".")
				theTable.SelectCell(2)
				theTable.Frame True, True, true, true
				
				theTable.NextCell
				theDoc.HPos = 1
				theTable.AddText replace(FormatNumber(rsMan("S_Actual"),0),",",".")
				theTable.SelectCell(3)
				theTable.Frame True, True, true, true
			rsMan.Movenext
	wend
end function
%>