<!--#include file="../conexion.asp"-->
<%
dim archivo,ref,objFile
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

set archivo = CreateObject("Scripting.FileSystemObject")
set ref = archivo.OpenTextFile(server.mappath("TXT_SFT_"&fecha&".txt"), 2, True)

	qsHc=" SELECT CONVERT(VARCHAR(10),F.FECHA_EMISION,105) AS FECHA_FACTURA,"&_
	 "e.RUT,UPPER(E.R_SOCIAL) as 'Empresa',UPPER(E.DIRECCION) AS DIRECCION,"&_
	 "UPPER(E.DIRECCION+', '+E.CIUDAD) AS DIR, E.ID_EMPRESA,"&_
	 "DBO.MayMinTexto(E.COMUNA) AS COMUNA, DBO.MayMinTexto(E.CIUDAD) AS CIUDAD,"&_
	 "E.FONO, UPPER(E.GIRO) AS GIRO,(CASE WHEN a.TIPO_DOC='0' then 'O/C '+a.ORDEN_COMPRA "&_
	 " WHEN a.TIPO_DOC='1' then 'Vale Vista N° '+a.ORDEN_COMPRA "&_
	 " WHEN a.TIPO_DOC='2' then 'DV '+a.ORDEN_COMPRA "&_
	 " WHEN a.TIPO_DOC='3' then 'Transferencia N° '+a.ORDEN_COMPRA "&_
	 " WHEN a.TIPO_DOC='4' then 'CONTRA FACTURA '+a.ORDEN_COMPRA "&_
	 " END) as 'Tipo Documento',f.MONTO,f.VALOR_CURSO,"&_
	 "CONVERT(VARCHAR(10),P.FECHA_INICIO_,105) AS FECHA_INICIO,"&_
	 "CONVERT(VARCHAR(10),P.FECHA_TERMINO,105) AS FECHA_TERMINO, "&_
	 "CONVERT(VARCHAR(10),GETDATE(),105) AS hoy, "&_	 
	 "C.HORAS,C.DESCRIPCION,C.COD_SOFTLAND,c.CODIGO,"&_
	 "ISNULL(convert(varchar,A.N_REG_SENCE),'') as N_REG_SENCE, "&_ 
	 "(select (CASE WHEN bq.ID_SEDE =  27 THEN bq.nom_sede "&_
	 "WHEN bq.ID_SEDE <>  27 THEN S.DIRECCION+', '+S.CIUDAD END) "&_
	 "from bloque_programacion bq "&_
	 " inner join SEDES s on s.ID_SEDE=bq.id_sede "&_
	 " inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "&_
	 " where bq.id_bloque=A.ID_BLOQUE) as SEDE, "&_
	 " ASIS=(SELECT COUNT(*) FROM HISTORICO_CURSOS HC  "&_
	 " WHERE HC.ID_AUTORIZACION=A.ID_AUTORIZACION AND HC.ASISTENCIA<>0), "&_
	 " INASIS=(SELECT COUNT(*) FROM HISTORICO_CURSOS HC  "&_
	 " WHERE HC.ID_AUTORIZACION=A.ID_AUTORIZACION AND HC.ASISTENCIA=0),F.ID_FACTURA "&_ 
	 " FROM FACTURAS F "&_
	 " inner join AUTORIZACION a on a.ID_AUTORIZACION=f.ID_AUTORIZACION "&_
	 " inner join empresas e on e.ID_EMPRESA=f.ID_EMPRESA "&_
	 " inner join programa p on p.ID_PROGRAMA=a.ID_PROGRAMA "&_
	 " inner join CURRICULO c on c.ID_MUTUAL=p.ID_MUTUAL "&_
	 " WHERE F.FECHA_EMISION>=CONVERT(date, '"&request("i")&"') and  "&_
	 " F.FECHA_EMISION<=CONVERT(date, '"&request("t")&"') "&_
	 " and F.FACTURA IS NULL order by f.FACTURA asc"

	set rsHc =  conn.execute(qsHc)
	
	Descripcion=""
    while not rsHc.eof
	Descripcion=rsHc("SEDE")&"#1"&rsHc("CODIGO")&"#2"&replace(cstr(rsHc("HORAS")),"-","/")&_
				"#3"&replace(cstr(rsHc("FECHA_INICIO")),"-","/")&"#4"&replace(cstr(rsHc("FECHA_TERMINO")),"-","/")&_
				"#5"&rsHc("ASIS")&"#6"&_
				rsHc("INASIS")&"#7"&rsHc("VALOR_CURSO")&"#8"&rsHc("Tipo Documento")&"#9"&rsHc("N_REG_SENCE")&"$1"&rsHc("Empresa")&"$2"&rsHc("RUT")&"$3"
	
			ref.writeline("22;"&rsHc("ID_FACTURA")&";"&replace(cstr(rsHc("FECHA_FACTURA")),"-","/")&";01;"&Descripcion&";"&mid(rsHc("RUT"), 1,len(rsHc("RUT"))-2)&";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;"&rsHc("COD_SOFTLAND")&";"&rsHc("MONTO")&";;1;;;;;;;;;;;;;;;;;")

        rsHc.MoveNext
    wend

ref.close() 

Set objFile = archivo.GetFile(server.mappath("TXT_SFT_"&fecha&".txt"))

Response.Buffer = True

Const adTypeBinary = 1

Response.Clear

Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type = adTypeBinary
objStream.LoadFromFile Server.MapPath("TXT_SFT_"&fecha&".txt")

	ContentType = "application/octet-stream"

    Response.AddHeader "Content-Disposition", "attachment; filename=" & "TXT_SFT_"&fecha&".txt"
    Response.AddHeader "Content-Length", objFile.Size

    Response.Charset = "UTF-8"
    Response.ContentType = ContentType

    Response.BinaryWrite objStream.Read
    Response.Flush

objStream.Close
Set objStream = Nothing

Private Sub DownloadFile(file)
 	Dim strAbsFile
    Dim objFSO
    Dim objFile
    Dim objStream
  '-- set absolute file location
    strAbsFile = Server.MapPath(file)
  '-- create FSO object to check if file exists and get properties
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")	
	
	Set objFile = objFSO.GetFile(strAbsFile)
	'-- first clear the response, and then set the appropriate headers
	Response.Clear
	'-- the filename you give it will be the one that is shown
	' to the users by default when they save
	Response.AddHeader "Content-Disposition", "attachment; filename=" & objFile.Name
	Response.AddHeader "Content-Length", objFile.Size
	Response.ContentType = "application/octet-stream"
	Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Open
	'-- set as binary
	objStream.Type = 1
	Response.CharSet = "UTF-8"
	'-- load into the stream the file
	objStream.LoadFromFile(strAbsFile)
	'-- send the stream in the response
	Response.BinaryWrite(objStream.Read)
	objStream.Close
	Set objStream = Nothing
	Set objFile = Nothing
End Sub

Function Rpad (sValue, sPadchar, iLength)
  Rpad = sValue & string(iLength - Len(sValue), sPadchar)
End Function

Function Lpad (sValue, sPadchar, iLength)
  Lpad = string(iLength - Len(sValue),sPadchar) & sValue
End Function

%>