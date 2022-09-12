<!--#include file="../funciones/xelupload.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload

objUpload.Upload()

'fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
'fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
'set objFich = objUpload.Ficheros("txtXLS")

fecha = objUpload.Form("txtXLSPreins")
	on error resume next
		For each objFich in objUpload.Ficheros.Items
			aux=split(objFich.Nombre,".")
			ext=aux(ubound(aux))
			nom_arch="Excel_"&fecha&"."&ext
			objFich.GuardarComo nom_arch,Server.MapPath("../cargas")
		next
		if err.number <> 0 then
			Response.Write("{"&chr(13))
			Response.Write("error: '"&err.description&"',"&chr(13))
			Response.Write("msg: '',"&chr(13))
			Response.Write("}")
		else
			Response.Write("{"&chr(13))
			Response.Write("error: '',"&chr(13))
			Response.Write("msg: '"&nom_arch&"',"&chr(13))
			Response.Write("}")
		end if
	on error goto 0

'set oFich = nothing
'set objUpload = nothing
%>