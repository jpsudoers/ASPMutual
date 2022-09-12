<!--#include file="../funciones/xelupload.asp"-->
<!--#include file="../conexion.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload
objUpload.Upload()

set objFich = objUpload.Ficheros("txtDoc")

aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="SolicitudCredito_"&fecha&"."&ext

objFich.GuardarComo nom_arch,Server.MapPath("Documentos")

vtxRut = objUpload.Form("txRut")
vtxtNombreEmpresa = objUpload.Form("txtNombreEmpresa")

vtxtCreditoPlazo = objUpload.Form("txtCreditoPlazo")
vtxtCreditoSol = objUpload.Form("txtCreditoSol")
vtxtCreditoDia = objUpload.Form("txtCreditoDia")
vtxtCreditoEncargado1 = objUpload.Form("txtCreditoEncargado1")
vtxtCreditoCorreo1 = objUpload.Form("txtCreditoCorreo1")
vtxtCreditoFono1 = objUpload.Form("txtCreditoFono1")
vtxtCreditoEncargado2 = objUpload.Form("txtCreditoEncargado2")
vtxtCreditoCorreo2 = objUpload.Form("txtCreditoCorreo2")
vtxtCreditoFono2 = objUpload.Form("txtCreditoFono2")
vtxtCreditoRepLegal = objUpload.Form("txtCreditoRepLegal")
vtxtCreditoRepLegalCorreo = objUpload.Form("txtCreditoRepLegalCorreo")
vtxtCreditoRepLegalFono = objUpload.Form("txtCreditoRepLegalFono")

vtxtCreditoHorario = "NULL"
if(objUpload.Form("txtCreditoHorario")<>"")then
	vtxtCreditoHorario = "'"&objUpload.Form("txtCreditoHorario")&"'"
end if

vtxtCreditoInfo = "NULL"
if(objUpload.Form("txtCreditoInfo")<>"")then
	vtxtCreditoInfo = "'"&objUpload.Form("txtCreditoInfo")&"'"
end if

vcond_OC = "NULL"
if(objUpload.Form("checkcond_OC")<>"")then
	vcond_OC = "'"&objUpload.Form("checkcond_OC")&"'"
end if

vcond_HES = "NULL"
if(objUpload.Form("checkcond_HES")<>"")then
	vcond_HES = "'"&objUpload.Form("checkcond_HES")&"'"
end if

vcond_MIGO = "NULL"
if(objUpload.Form("checkcond_MIGO")<>"")then
	vcond_MIGO = "'"&objUpload.Form("checkcond_MIGO")&"'"
end if

vcond_OTRO = "NULL"
if(objUpload.Form("checkcond_OTRO")<>"")then
	vcond_OTRO = "'"&objUpload.Form("checkcond_OTRO")&"'"
end if

vcond_OTROTxt = "NULL"
if(objUpload.Form("cond_OTROTxt")<>"")then
	vcond_OTROTxt = "'"&objUpload.Form("cond_OTROTxt")&"'"
end if

query = "insert into SOLICITUDES_CREDITO (RUT_EMPRESA,NOMBRE_EMPRESA,FECHA_SOLICITUD,ID_ESTADO,USUARIO_INGRESO,NOMBRE_ARCHIVO,CreditoPlazo,CreditoSol,CreditoDia,CreditoEncargado1,CreditoCorreo1,CreditoFono1,CreditoEncargado2,CreditoCorreo2,CreditoFono2,CreditoRepLegal,CreditoRepLegalCorreo,CreditoRepLegalFono,CreditoHorario,CreditoInfo,cond_OC,cond_HES,cond_MIGO,cond_OTRO,cond_OTROTxt)  "
query = query & " values ('"& vtxRut &"','"& vtxtNombreEmpresa &"', GETDATE(),1, '"& Session("usuario")&"','"&nom_arch&"','"&vtxtCreditoPlazo&"',"
query = query & "'"&vtxtCreditoSol&"','"&vtxtCreditoDia&"','"&vtxtCreditoEncargado1&"','"&vtxtCreditoCorreo1&"','"&vtxtCreditoFono1&"','"&vtxtCreditoEncargado2&"',"
query = query & "'"&vtxtCreditoCorreo2&"','"&vtxtCreditoFono2&"','"&vtxtCreditoRepLegal&"','"&vtxtCreditoRepLegalCorreo&"','"&vtxtCreditoRepLegalFono&"',"
query = query &vtxtCreditoHorario&","&vtxtCreditoInfo&","&vcond_OC&","&vcond_HES&","&vcond_MIGO&","&vcond_OTRO&","&vcond_OTROTxt&") "

'response.write(query)
'response.end()

conn.execute (query)

set oFich = nothing
set objUpload = nothing

%>