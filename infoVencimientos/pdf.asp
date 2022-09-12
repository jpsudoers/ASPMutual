<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

Set theDoc = Server.CreateObject("ABCpdf7.Doc")
Set theImg = Server.CreateObject("ABCpdf7.Image")
theImg.SetFile Server.MapPath("../images/logoMutual.jpg")
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

theDoc.Rect = "50 70 740 520"
Tabla(theDoc)

theCount = theDoc.PageCount
'/******** HEADER **********/
theDoc.Rect = "50 600 260 540"
theDoc.FontSize = 8
For i = 1 To theCount
  theDoc.PageNumber = i
  theDoc.AddImageObject theImg, False
Next

theDoc.Rect = "10 640 740 450"
theDoc.HPos = 0.5
theDoc.VPos = 0.5
theDoc.FontSize =12
For i = 1 To theCount
  theDoc.PageNumber = i
  theDoc.Color = "0 0 0"
  theDoc.AddHtml ("<b>Informe de Vencimientos</b>")
  theDoc.Color = "0 0 0"
	'theDoc.FrameRect
Next

'/******** FOOTER **********/
theCount = theDoc.PageCount
For i = 1 To theCount
	theDoc.Rect = "490 50 680 70"
	theDoc.HPos = 1.0
	theDoc.VPos = 0.5
	theDoc.FontSize = 8
 	theDoc.PageNumber = i	
  	theDoc.AddText "Página " & i & " de " & theCount
	theDoc.Rect = "50 50 570 70"
	theDoc.HPos = 0.5
	'theDoc.AddText "Agencia Antofagasta | Washington 2701 Piso 3 Antofagasta – Chile | Tel.: (55) 251585 | www.mutual.cl"
Next

theID = theDoc.GetInfo(theDoc.Root, "Pages")
theDoc.SetInfo theID, "/Rotate", "90"

sArchivo = "../pdf/Info_Vencimientos_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)
function Tabla(theDoc)
	 sql = "select FC.MONTO as VALOR_OC,FC.FACTURA,UPPER(E.R_SOCIAL) as R_SOCIAL,(select isnull(SUM(I.monto),0) from INGRESOS I "
	 sql = sql&" where I.id_factura=FC.ID_FACTURA)as cancelado, "
	 sql = sql&" (Select Datediff(""d"", Min(FC2.FECHA_VENCIMIENTO), Max(CONVERT(date,GETDATE(), 105))) "
	 sql = sql&" from FACTURAS FC2 where FC2.ID_FACTURA=FC.ID_FACTURA) as dias "
	 sql = sql&" from FACTURAS FC inner join EMPRESAS E on E.ID_EMPRESA=FC.ID_EMPRESA "
	 sql = sql&" where FC.FECHA_VENCIMIENTO<CONVERT(date,GETDATE(), 105) and FC.ESTADO_CANCELACION=0" 

	set rs =  conn.execute(sql)
	
	theDoc.FontSize = 9
	Set theTable = New Table
	theTable.Focus theDoc, 5
	theTable.Width(0) = 3
	theTable.Width(1) = 1
	theTable.Width(2) = 1
	theTable.Width(3) = 1
	theTable.Width(4) = 1

	theTable.Padding = 2
	theDoc.Font = theDoc.AddFont("Arial Black")
	
	'/******* CABECERA *********/
	theTable.Frame True, False, False, False
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Empresa"
	theTable.SelectCell(0)
	theTable.Frame True, True, True, True
		'theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Factura"
	theTable.SelectCell(1)
	theTable.Frame True, True, True, True
		'theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Monto"
	theTable.SelectCell(2)
	theTable.Frame True, True, True, True
		'theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Cancelado"
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		'theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Saldo"
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		'theTable.Fill "174 174 174"	

	theDoc.Font = theDoc.AddFont("Arial")
	
	'/********* DATOS **********/
	hasta=0
	color="0"
	while not rs.eof

		    if(cdbl(rs("dias"))>0 and cdbl(rs("dias"))<=15)then
			   color="244 244 162"
			end if
			
			if(cdbl(rs("dias"))>=16 and cdbl(rs("dias"))<=30)then
			    color="235 170 121"
			end if
			
			if(cdbl(rs("dias"))>=31)then
			  color="255 109 109"
			end if
		
		    if(Request("dias")=15 or Request("dias")=30)then
				if(cdbl(rs("dias"))>=cdbl(cdbl(Request("dias")) - cdbl(14)) and cdbl(rs("dias"))<=cdbl(Request("dias")))then
					theTable.NextRow		
					theTable.NextCell
					theDoc.HPos = 0
					theTable.AddText rs("R_SOCIAL")
					theTable.NextCell
					theDoc.HPos = 1
					theTable.AddText rs("FACTURA")
					theTable.NextCell
					theDoc.HPos = 1
					theTable.AddText "$ "&replace(FormatNumber(cdbl(rs("VALOR_OC")),0),",",".")
					theTable.NextCell
					theDoc.HPos = 1
					theTable.AddText "$ "&replace(FormatNumber(cdbl(rs("cancelado")),0),",",".")
					theTable.NextCell
					theDoc.HPos = 1
					theTable.AddText "$ "&replace(FormatNumber(cdbl(cdbl(rs("VALOR_OC")) - cdbl(rs("cancelado"))),0),",",".")
		
					theTable.SelectRow theTable.Row
					theTable.Fill color
				end if
			end if
			
			 if(Request("dias")=45)then
				if(cdbl(rs("dias"))>=cdbl(cdbl(Request("dias")) - cdbl(14)))then
					theTable.NextRow		
					theTable.NextCell
					theDoc.HPos = 0
					theTable.AddText rs("R_SOCIAL")
					theTable.NextCell
					theDoc.HPos = 1
					theTable.AddText rs("FACTURA")
					theTable.NextCell
					theDoc.HPos = 1
					theTable.AddText "$ "&replace(FormatNumber(cdbl(rs("VALOR_OC")),0),",",".")
					theTable.NextCell
					theDoc.HPos = 1
					theTable.AddText "$ "&replace(FormatNumber(cdbl(rs("cancelado")),0),",",".")
					theTable.NextCell
					theDoc.HPos = 1
					theTable.AddText "$ "&replace(FormatNumber(cdbl(cdbl(rs("VALOR_OC")) - cdbl(rs("cancelado"))),0),",",".")
		
					theTable.SelectRow theTable.Row
					theTable.Fill color
				end if
			end if
			
			if(cdbl(Request("dias"))=0)then
				theTable.NextRow		
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rs("R_SOCIAL")
				theTable.NextCell
				theDoc.HPos = 1
				theTable.AddText rs("FACTURA")
				theTable.NextCell
				theDoc.HPos = 1
				theTable.AddText "$ "&replace(FormatNumber(cdbl(rs("VALOR_OC")),0),",",".")
				theTable.NextCell
				theDoc.HPos = 1
				theTable.AddText "$ "&replace(FormatNumber(cdbl(rs("cancelado")),0),",",".")
				theTable.NextCell
				theDoc.HPos = 1
				theTable.AddText "$ "&replace(FormatNumber(cdbl(cdbl(rs("VALOR_OC")) - cdbl(rs("cancelado"))),0),",",".")
	
				theTable.SelectRow theTable.Row
				theTable.Fill color
			end if
		rs.Movenext
	wend
	theDoc.Color = "0 0 0"
	theTable.NextRow
	theTable.SelectRow theTable.Row
	theTable.Frame True, False, False, False
end function
%>