<!--#include file="../funciones/xelupload.asp"-->
<!--#include file="../conexion.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload
objUpload.Upload()

vsolId = objUpload.Form("solId")

set objFich = objUpload.Ficheros("txtDoc")

aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="Doc_"&fecha&"."&ext

vcorrelativo = "1"

sql = "select isnull(max(CORRELATIVO) + 1,1) as correlativo  from SOLICITUDES_CREDITO_ARCHIVOS "

SET rs = conn.execute(sql)
While not rs.EOF
	vcorrelativo = rs("correlativo")
	rs.MoveNext
wend

objFich.GuardarComo nom_arch, Server.MapPath("Documentos")

query = "insert into  SOLICITUDES_CREDITO_ARCHIVOS (ID_SOLICITUD,CORRELATIVO,NOMBRE_ARCHIVO,FECHA_ARCHIVO)   "
query = query & " values ("& vsolId &","& vcorrelativo &",'"& nom_arch &"', GETDATE()) "

conn.execute (query)

set oFich = nothing
set objUpload = nothing

%>