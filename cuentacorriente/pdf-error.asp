<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->

<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
mes=""

emp_id=""

if(Request("empresa")<>"")then
emp_id=" where FC.ID_EMPRESA="&Request("empresa")
end if

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
  theDoc.AddHtml ("<b>Estado de Cuenta Corriente</b>")
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

sArchivo = "../pdf/Cuenta_Corriente_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)
function Tabla(theDoc)
	sql = "select distinct FC.ID_EMPRESA as empresa_id,E.RUT as rut,UPPER(E.R_SOCIAL) as empresa from FACTURAS FC "
    sql = sql&" inner join EMPRESAS E on E.ID_EMPRESA=FC.ID_EMPRESA "&emp_id

	set rs =  conn.execute(sql)
	
	theDoc.FontSize = 8
	Set theTable = New Table
	theTable.Focus theDoc, 4
	theTable.Width(0) = 3.8
	theTable.Width(1) = 2.2
	theTable.Width(2) = 2.2
	theTable.Width(3) = 0.8

	theTable.Padding = 2
	
	'/******* CABECERA *********/
	theTable.Frame True, False, False, False
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Empresa"
	theTable.SelectCell(0)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Factura"
	theTable.SelectCell(1)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Pago"
	theTable.SelectCell(2)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Saldo"
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	
	theTable.NextRow
	
	Set theTable = New Table
	theTable.Focus theDoc, 9
	theTable.Width(0) = 0.7
	theTable.Width(1) = 3.1
	theTable.Width(2) = 0.7
	theTable.Width(3) = 0.5
	theTable.Width(4) = 1
	theTable.Width(5) = 0.7
	theTable.Width(6) = 0.5
	theTable.Width(7) = 1
	theTable.Width(8) = 0.8

	theTable.Padding = 2
	'/******* CABECERA *********/
	theTable.Frame True, False, False, False
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Rut"
	theTable.SelectCell(0)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Razón Social"
	theTable.SelectCell(1)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Fecha"
	theTable.SelectCell(2)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "#"
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "$"
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Fecha"
	theTable.SelectCell(5)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "#"
	theTable.SelectCell(6)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "$"
	theTable.SelectCell(7)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.8
	theTable.AddText ""
	theTable.SelectCell(8)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	
	'/********* DATOS **********/
	n=0
	totSaldo=0
	filas = 0
	while not rs.eof
	n=0

		If theTable.NextRow = False Then
	
			If theTable.RowTruncated Then
			  theTable.DeleteLastRow
			  filas = filas - 1
			End If
		
			filas = filas - 1
			If filas >= 0 Then theRows(filas) = theRows(0)
				
			theDoc.Flatten
			theTable.NewPage
		End If

		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText replace(FormatNumber(mid(rs("rut"), 1,len(rs("rut"))-2),0)&mid(rs("rut"), len(rs("rut"))-1,len(rs("rut"))),",",".")
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rs("empresa")

		theTable.SelectRow theTable.Row
		
		theTable.Frame True, False, False, False
		
		fac = "select FC.FACTURA,CONVERT(VARCHAR(10),FC.FECHA_EMISION, 105) as FECHA_EMISION,FC.MONTO as VALOR_OC,FC.ID_FACTURA "
        fac = fac&" from FACTURAS FC where FC.ID_EMPRESA="&rs("empresa_id")
		
		set rsFac =  conn.execute(fac)
		totFac=0
		totPag=0
		n_pagos=1
		totSaldoParcial=0
			while not rsFac.eof
			n_pagos=1
					if(n=1)then
						theTable.NextCell
						theDoc.HPos = 0
						theTable.AddText ""'
						theTable.NextCell
						theDoc.HPos = 0
						theTable.AddText ""'
					end if
					
						theTable.NextCell
						theDoc.HPos = 0
						theTable.AddText rsFac("FECHA_EMISION")
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText rsFac("FACTURA")
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "$ "&replace(FormatNumber(rsFac("VALOR_OC"),0),",",".")
						totFac=cdbl(rsFac("VALOR_OC"))
					
						pag = "select I.comprobante,I.monto,CONVERT(VARCHAR(10),I.fecha_pago, 105) as fecha_pago from INGRESOS I "
 						pag = pag&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=I.empresa " 
 						pag = pag&" where I.id_factura="&rsFac("ID_FACTURA")&" and EMPRESAS.ID_EMPRESA="&rs("empresa_id")  
				
						set rsPag =  conn.execute(pag)
						
					   PagCount="select COUNT(*) as totalPag from INGRESOS I inner join EMPRESAS E on E.ID_EMPRESA=I.empresa where"
					   PagCount=PagCount&" I.id_factura="&rsFac("ID_FACTURA")&" and E.ID_EMPRESA="&rs("empresa_id")
						
						set rsPagCount =  conn.execute(PagCount)
		
						if not rsPag.eof and not rsPag.bof then
						totSaldoParcial=0
								while not rsPag.eof
								
									theTable.NextCell
									theDoc.HPos = 0.5
									theTable.AddText rsPag("fecha_pago")
									theTable.NextCell
									theDoc.HPos = 0.5
									theTable.AddText rsPag("comprobante")
									theTable.NextCell
									theDoc.HPos = 0.5
									theTable.AddText "$ "&replace(FormatNumber(rsPag("monto"),0),",",".")
									totPag=cdbl(rsPag("monto"))
									
									theTable.NextCell
									theDoc.HPos = 1
									
									if(n_pagos>1)then
									totSaldoParcial=cdbl(cdbl(totSaldoParcial) - cdbl(totPag))
									else
									totSaldoParcial=cdbl(cdbl(totSaldoParcial) + cdbl(cdbl(totFac) - cdbl(totPag)))
									end if
									
									theTable.AddText "$ "&replace(FormatNumber(totSaldoParcial,0),",",".")
									theTable.NextRow
									n=1
									
									if(n_pagos<cdbl(rsPagCount("totalPag")))then
									    theTable.NextRow
									    theTable.NextCell
										theDoc.HPos = 0.5
										theTable.AddText ""
										theTable.NextCell
										theDoc.HPos = 0.5
										theTable.AddText ""
										theTable.NextCell
										theDoc.HPos = 0.5
										theTable.AddText ""
										theTable.NextCell
										theDoc.HPos = 0.5
										theTable.AddText ""
										theTable.NextCell
										theDoc.HPos = 0.5
										theTable.AddText ""
									end if
									n_pagos=n_pagos+1
									rsPag.Movenext
								wend
								theTable.NextRow
						else
							theTable.NextCell
							theDoc.HPos = 0.5
							theTable.AddText "-"
							theTable.NextCell
							theDoc.HPos = 0.5
							theTable.AddText "-"
							theTable.NextCell
							theDoc.HPos = 0.5
							theTable.AddText "$ 0"
							
							theTable.NextCell
							theDoc.HPos = 1
							
							if(n_pagos>1)then
									totSaldoParcial=cdbl(cdbl(totSaldoParcial) - cdbl(totPag))
									else
									totSaldoParcial=cdbl(cdbl(totSaldoParcial) + cdbl(cdbl(totFac) - cdbl(totPag)))
							end if
							
							theTable.AddText "$ "&replace(FormatNumber(totSaldoParcial,0),",",".")
							
							theTable.NextRow
							n=1
						end if
				rsFac.Movenext
				totSaldo=totSaldo + cdbl(totSaldoParcial)
			wend
		rs.Movenext
	wend
	theTable.SelectRow theTable.Row
	theTable.Frame False, True, False, False
	
	theTable.NextRow
	theDoc.FontSize = 14
	Set theTable = New Table
	theTable.Focus theDoc, 4
	theTable.Width(0) = 4
	theTable.Width(1) = 2
	theTable.Width(2) = 3
	theTable.Width(3) = 2

	theTable.Padding = 2
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
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theTable.AddText ""
	theTable.NextCell
	theDoc.HPos = 1
	theTable.AddText "Total"
	theTable.NextCell
	theDoc.HPos = 1
	theTable.AddText "$ "&replace(FormatNumber(totSaldo,0),",",".")
end function
%>