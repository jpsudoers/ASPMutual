<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->

<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

	dim totIngresos
	dim totEgresos
	dim totAjustesNeg
	dim totAjustesPos
	Dim totGeneral
	
	totIngresos=0
	totEgresos=0
	totAjustesNeg=0
	totAjustesPos=0
	totGeneral=0

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
fila=0
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

if(Request("tipo")="0")then	
	theDoc.Rect = "30 50 760 570"
	TablaIngresos(theDoc)
	
	theDoc.Rect = "30 50 760 570"
	TablaSalidas(theDoc)
	
	theDoc.Rect = "30 50 760 570"
	TablaAjuste(theDoc)
else
	if(Request("tipo")="1")then	
		theDoc.Rect = "30 50 760 570"
		TablaIngresos(theDoc)
	end if
	
	if(Request("tipo")="2")then		
		theDoc.Rect = "30 50 760 570"
		TablaSalidas(theDoc)
	end if
	
	if(Request("tipo")="3")then		
		theDoc.Rect = "30 50 760 570"
		TablaAjuste(theDoc)
	end if	
end if

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

sArchivo = "../pdf/Informe_Movimientos_"&fecha&".pdf"
theDoc.Save Server.MapPath(sArchivo)
Response.Redirect(sArchivo)

function TablaIngresos(theDoc)
	theDoc.FontSize = 14
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	
	if(Request("tipo")="0" or Request("tipo")="1")then	
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theDoc.Font = theDoc.EmbedFont("Arial Black")
	
		theDoc.AddHtml ("<b>Informe de Movimientos</b>")
		theDoc.Font = theDoc.EmbedFont("Arial")
		
		theTable.NextRow
		theTable.NextCell
		theTable.AddText ""
	end if
	
	theDoc.FontSize = 12
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theDoc.AddHtml ("<b>Ingresos</b>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow

	theDoc.FontSize = 7
	Set theTable = New Table
	theTable.Focus theDoc, 7
	theTable.Width(0) = 0.52
	theTable.Width(1) = 0.95
	theTable.Width(2) = 1.5
	theTable.Width(3) = 0.38
	theTable.Width(4) = 1.25
	theTable.Width(5) = 1.25
	theTable.Width(6) = 1.35

	theTable.Padding = 10

	'/******* CABECERA *********/
	theDoc.Font = theDoc.EmbedFont("Arial Black")
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Fecha"
	theTable.SelectCell(0)
	theTable.Frame True, True, true, true
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Tipo de Movimiento"
	theTable.SelectCell(1)
	theTable.Frame True, True, false, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Articulo"
	theTable.SelectCell(2)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Cant."
	theTable.SelectCell(3)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Razones"
	theTable.SelectCell(4)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Fecha / Relator"
	theTable.SelectCell(5)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"
	theTable.NextCell
	theDoc.HPos = 0.5
	theTable.AddText "Usuario"
	theTable.SelectCell(6)
	theTable.Frame True, True, True, True
		theTable.Fill "174 174 174"				

	'**** Tabla participantes ********
	
	qsIng="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Mov_Articulo,CANTIDAD,"
	qsIng=qsIng&" dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_User "
	qsIng=qsIng&" FROM MOVIMIENTOS M "
	qsIng=qsIng&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO " 
	qsIng=qsIng&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsIng=qsIng&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO " 
	qsIng=qsIng&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsIng=qsIng&" WHERE M.TIPO_MOVIMIENTO IN (1,2) AND M.MODULO=1 AND M.ESTADO=1 " 
	qsIng=qsIng&" ORDER BY M.FECHA DESC "
	
	set rsIng = conn.execute (qsIng)
	
	theTable.Padding = 4
	ingresos=0
	while not rsIng.eof
		totIngresos=totIngresos+cdbl(rsIng("CANTIDAD"))
		ingresos=ingresos+1
		theTable.NextRow
		
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText rsIng("FECHA_MOV")
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsIng("NOMBRE_MOV")
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsIng("Mov_Articulo")
		theTable.SelectCell(2)
		theTable.Frame True, True, true, true

		theTable.NextCell
		theDoc.HPos = 1
		theTable.AddText FormatNumber(rsIng("CANTIDAD"),0)
		theTable.SelectCell(3)
		theTable.Frame True, True, true, true
		
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
		theTable.AddText rsIng("Mov_User")
		theTable.SelectCell(6)
		theTable.Frame True, True, true, true

		
		if(ingresos=24)then
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
			theDoc.HPos = 0
			theDoc.Font = theDoc.EmbedFont("Arial Black")
		
			theDoc.AddHtml ("<b>Ingresos</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			
			theTable.NextRow
		
			theDoc.FontSize = 7
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.52
			theTable.Width(1) = 0.95
			theTable.Width(2) = 1.5
			theTable.Width(3) = 0.38
			theTable.Width(4) = 1.25
			theTable.Width(5) = 1.25
			theTable.Width(6) = 1.35
		
			theTable.Padding = 10
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha"
			theTable.SelectCell(0)
			theTable.Frame True, True, true, true
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tipo de Movimiento"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Articulo"
			theTable.SelectCell(2)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Cant."
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Razones"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha / Relator"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Usuario"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"	
				
			ingresos=0	
			theTable.Padding = 4
		end if
	rsIng.Movenext
	wend
	
		theDoc.FontSize = 9
		theTable.NextRow
		
		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText ""

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
		theTable.AddText ""

		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText ""

		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText "Total Ingresos : "&FormatNumber(totIngresos,0)
	
	if(Request("tipo")="0")then	
		theDoc.Page = theDoc.AddPage()
	end if
end function

function TablaSalidas(theDoc)
	theDoc.FontSize = 14
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	if(Request("tipo")="2")then	
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theDoc.Font = theDoc.EmbedFont("Arial Black")
	
		theDoc.AddHtml ("<b>Informe de Movimientos</b>")
		theDoc.Font = theDoc.EmbedFont("Arial")
		
		theTable.NextRow
		theTable.NextCell
		theTable.AddText ""
	end if
	
	theDoc.FontSize = 12
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theDoc.AddHtml ("<b>Salidas</b>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow

			theDoc.FontSize = 7
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.52
			theTable.Width(1) = 0.95
			theTable.Width(2) = 1.5
			theTable.Width(3) = 0.38
			theTable.Width(4) = 1.25
			theTable.Width(5) = 1.25
			theTable.Width(6) = 1.35
		
			theTable.Padding = 10
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha"
			theTable.SelectCell(0)
			theTable.Frame True, True, true, true
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tipo de Movimiento"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Articulo"
			theTable.SelectCell(2)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Cant."
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Razones"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha / Relator"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Usuario"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"	

	'**** Tabla participantes ********
	
 	qsEgres="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Nom_Articulo,CANTIDAD, "
	qsEgres=qsEgres&"dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_Nom, "
	qsEgres=qsEgres&"CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105)+', '+"
	qsEgres=qsEgres&"dbo.MayMinTexto(IR.NOMBRES+' '+IR.A_PATERNO+' '+IR.A_MATERNO) AS FECHA_NOM_RELATOR "
	qsEgres=qsEgres&" FROM MOVIMIENTOS M " 
	qsEgres=qsEgres&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO "
	qsEgres=qsEgres&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsEgres=qsEgres&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO " 
	qsEgres=qsEgres&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsEgres=qsEgres&" inner join bloque_programacion BP ON BP.id_bloque=M.ID_PROG_BLOQUE "
	qsEgres=qsEgres&" INNER JOIN PROGRAMA P ON P.ID_PROGRAMA=BP.id_programa "
	qsEgres=qsEgres&" INNER JOIN INSTRUCTOR_RELATOR IR ON IR.ID_INSTRUCTOR=BP.id_relator "
	qsEgres=qsEgres&" WHERE M.TIPO_MOVIMIENTO IN (3) AND M.MODULO=2 AND M.ESTADO=1" 
	qsEgres=qsEgres&" ORDER BY M.FECHA DESC "
	
	set rsEgres =  conn.execute(qsEgres)
	
	theTable.Padding = 4
	egresos=0
	while not rsEgres.eof
		totEgresos=totEgresos+cdbl(rsEgres("CANTIDAD"))
		egresos=egresos+1
		theTable.NextRow
		
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText rsEgres("FECHA_MOV")
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsEgres("NOMBRE_MOV")
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsEgres("Nom_Articulo")
		theTable.SelectCell(2)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 1
		theTable.AddText FormatNumber(rsEgres("CANTIDAD"),0)
		theTable.SelectCell(3)
		theTable.Frame True, True, true, true

		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(4)
		theTable.Frame True, True, true, true

		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsEgres("FECHA_NOM_RELATOR")
		theTable.SelectCell(5)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsEgres("Mov_Nom")
		theTable.SelectCell(6)
		theTable.Frame True, True, true, true

		if(egresos=24)then
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
			theDoc.HPos = 0
			theDoc.Font = theDoc.EmbedFont("Arial Black")
		
			theDoc.AddHtml ("<b>Salidas</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			
			theTable.NextRow
		
			theDoc.FontSize = 7
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.52
			theTable.Width(1) = 0.95
			theTable.Width(2) = 1.5
			theTable.Width(3) = 0.38
			theTable.Width(4) = 1.25
			theTable.Width(5) = 1.25
			theTable.Width(6) = 1.35
		
			theTable.Padding = 10
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha"
			theTable.SelectCell(0)
			theTable.Frame True, True, true, true
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tipo de Movimiento"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Articulo"
			theTable.SelectCell(2)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Cant."
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Razones"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha / Relator"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Usuario"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"	
				
			egresos=0	
			theTable.Padding = 4
		end if

	rsEgres.Movenext
	wend
		
		theDoc.FontSize = 9
		theTable.NextRow
		
		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText ""

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
		theTable.AddText ""

		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText ""

		theTable.NextCell
		theTable.AddText ""		
		
		theTable.NextCell
		theTable.AddText "Total Salidas : "&FormatNumber(totEgresos,0)
	
	if(Request("tipo")="0")then		
		theDoc.Page = theDoc.AddPage()
	end if
end function

function TablaAjuste(theDoc)
	theDoc.FontSize = 14
	
	Set theTable = New Table
	theTable.Focus theDoc, 1
	theTable.Width(0) = 8
	theTable.Padding = 4

	'/******* CABECERA *********/
	if(Request("tipo")="3")then	
		theTable.NextRow
		theTable.NextCell
		theDoc.HPos = 0.5
		theDoc.Font = theDoc.EmbedFont("Arial Black")
	
		theDoc.AddHtml ("<b>Informe de Movimientos</b>")
		theDoc.Font = theDoc.EmbedFont("Arial")
		
		theTable.NextRow
		theTable.NextCell
		theTable.AddText ""
	end if
	theDoc.FontSize = 12
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theDoc.Font = theDoc.EmbedFont("Arial Black")

	theDoc.AddHtml ("<b>Ajustes</b>")
	theDoc.Font = theDoc.EmbedFont("Arial")
	
	theTable.NextRow
	theTable.NextCell
	theDoc.HPos = 0
	theTable.AddText ""
	
	theTable.NextRow

			theDoc.FontSize = 7
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.52
			theTable.Width(1) = 0.95
			theTable.Width(2) = 1.5
			theTable.Width(3) = 0.38
			theTable.Width(4) = 1.25
			theTable.Width(5) = 1.25
			theTable.Width(6) = 1.35
		
			theTable.Padding = 10
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha"
			theTable.SelectCell(0)
			theTable.Frame True, True, true, true
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tipo de Movimiento"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Articulo"
			theTable.SelectCell(2)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Cant."
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Razones"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha / Relator"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Usuario"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"	

	'**** Tabla participantes ********
	
	qsAjuste="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Nom_Articulo,CANTIDAD, "
	qsAjuste=qsAjuste&"M.EXPLICACION,dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_Nom,M.TIPO_AJUSTE " 
	qsAjuste=qsAjuste&" FROM MOVIMIENTOS M " 
	qsAjuste=qsAjuste&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO " 
	qsAjuste=qsAjuste&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsAjuste=qsAjuste&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO "
	qsAjuste=qsAjuste&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsAjuste=qsAjuste&" WHERE M.TIPO_MOVIMIENTO IN (4,5,6) AND M.MODULO=3 AND M.ESTADO=1 " 
	qsAjuste=qsAjuste&" ORDER BY M.FECHA DESC " 

	set rsAjuste =  conn.execute(qsAjuste)
	
	theTable.Padding = 4
	ajustes=0
	theDoc.FontSize = 7
	
	while not rsAjuste.eof
		if(rsAjuste("TIPO_AJUSTE")="2")then
			totAjustesNeg=totAjustesNeg+cdbl(rsAjuste("CANTIDAD"))
		else
			totAjustesPos=totAjustesPos+cdbl(rsAjuste("CANTIDAD"))
		end if
		
		ajustes=ajustes+1
		theTable.NextRow
		
		theTable.NextCell
		theDoc.HPos = 0.5
		theTable.AddText rsAjuste("FECHA_MOV")
		theTable.SelectCell(0)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsAjuste("NOMBRE_MOV")
		theTable.SelectCell(1)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsAjuste("Nom_Articulo")
		theTable.SelectCell(2)
		theTable.Frame True, True, true, true

		theTable.NextCell
		theDoc.HPos = 1
		theTable.AddText FormatNumber(rsAjuste("CANTIDAD"),0)
		theTable.SelectCell(3)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsAjuste("EXPLICACION")
		theTable.SelectCell(4)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText ""
		theTable.SelectCell(5)
		theTable.Frame True, True, true, true
		
		theTable.NextCell
		theDoc.HPos = 0
		theTable.AddText rsAjuste("Mov_Nom")
		theTable.SelectCell(6)
		theTable.Frame True, True, true, true

		if(ajustes=23)then
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
			theDoc.HPos = 0
			theDoc.Font = theDoc.EmbedFont("Arial Black")
		
			theDoc.AddHtml ("<b>Ajustes</b>")
			theDoc.Font = theDoc.EmbedFont("Arial")
			
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0
			theTable.AddText ""
			
			theTable.NextRow
		
			theDoc.FontSize = 7
			Set theTable = New Table
			theTable.Focus theDoc, 7
			theTable.Width(0) = 0.52
			theTable.Width(1) = 0.95
			theTable.Width(2) = 1.5
			theTable.Width(3) = 0.38
			theTable.Width(4) = 1.25
			theTable.Width(5) = 1.25
			theTable.Width(6) = 1.35
		
			theTable.Padding = 10
		
			'/******* CABECERA *********/
			theDoc.Font = theDoc.EmbedFont("Arial Black")
			theTable.NextRow
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha"
			theTable.SelectCell(0)
			theTable.Frame True, True, true, true
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Tipo de Movimiento"
			theTable.SelectCell(1)
			theTable.Frame True, True, false, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Articulo"
			theTable.SelectCell(2)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Cant."
			theTable.SelectCell(3)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Razones"
			theTable.SelectCell(4)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Fecha / Relator"
			theTable.SelectCell(5)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"
			theTable.NextCell
			theDoc.HPos = 0.5
			theTable.AddText "Usuario"
			theTable.SelectCell(6)
			theTable.Frame True, True, True, True
				theTable.Fill "174 174 174"		
				
			ajustes=0	
			theTable.Padding = 4
			theDoc.FontSize = 7
		end if

	rsAjuste.Movenext
	wend
	
		theTable.NextRow
		
		theDoc.FontSize = 9
		
		Set theTable = New Table
		theTable.Focus theDoc, 3
		theTable.Width(0) = 2.6
		theTable.Width(1) = 2.6
		theTable.Width(2) = 2
		
		theTable.Padding = 4

		theTable.NextRow
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
		theTable.AddText "Total Ajustes Negativos : "&FormatNumber(totAjustesNeg,0)
		
		theTable.NextRow
		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText ""
		
		theTable.NextCell
		theTable.AddText "Total Ajustes Positivos : "&FormatNumber(totAjustesPos,0)
	
end function
%>