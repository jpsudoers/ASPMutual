<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vid = Request("ins_del") 

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

' send one copy with Google SMTP server (with autentication)
schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com" 
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "notificaciones@otecmutual.cl"
Flds.Item(schema & "sendpassword") =  "admin_2015"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

'creamos el nombre del archivo 
archivo= Server.MapPath("mailEliminaInscripcion.html")
'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 
'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 
'cerramos el fichero 
fich.close() 

dim AutoEmpresa
AutoEmpresa="select EMPRESAS.ID_EMPRESA,EMPRESAS.EMAIL,EMPRESAS.R_SOCIAL,PROGRAMA.ID_PROGRAMA,"
AutoEmpresa=AutoEmpresa&"AUTORIZACION.N_PARTICIPANTES,AUTORIZACION.VALOR_OC,AUTORIZACION.DOCUMENTO_COMPROMISO,"
AutoEmpresa=AutoEmpresa&"AUTORIZACION.TIPO_DOC,AUTORIZACION.ORDEN_COMPRA from AUTORIZACION "
AutoEmpresa=AutoEmpresa&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.id_empresa " 
AutoEmpresa=AutoEmpresa&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.id_programa " 
AutoEmpresa=AutoEmpresa&" where AUTORIZACION.ID_AUTORIZACION='"&vid&"'"

set rsAutoEmpresa = conn.execute (AutoEmpresa)

texto_fichero = Replace(texto_fichero,"<!--#empresa-->",rsAutoEmpresa("R_SOCIAL"))

dim queryProg
queryProg = "select CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_ from PROGRAMA "
queryProg = queryProg&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL where PROGRAMA.ID_PROGRAMA='"&rsAutoEmpresa("ID_PROGRAMA")&"'"

set rsQueryProg = conn.execute (queryProg)

texto_fichero = Replace(texto_fichero,"<!--#curso-->",rsQueryProg("NOMBRE_CURSO"))
texto_fichero = Replace(texto_fichero,"<!--#fecha-->",rsQueryProg("FECHA_INICIO_"))

texto_fichero = Replace(texto_fichero,"<!--#asistentes-->",rsAutoEmpresa("N_PARTICIPANTES"))  
texto_fichero = Replace(texto_fichero,"<!--#valor-->","$ "&replace(FormatNumber(rsAutoEmpresa("VALOR_OC"),0),",","."))

if(rsAutoEmpresa("TIPO_DOC")="0")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Orden de Compra N° "&rsAutoEmpresa("ORDEN_COMPRA"))
end if

if(rsAutoEmpresa("TIPO_DOC")="1")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Vale Vista N° "&rsAutoEmpresa("ORDEN_COMPRA")) 
end if

if(rsAutoEmpresa("TIPO_DOC")="2")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Depósito Cheque N° "&rsAutoEmpresa("ORDEN_COMPRA")) 
end if

if(rsAutoEmpresa("TIPO_DOC")="3")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Transferencia N° "&rsAutoEmpresa("ORDEN_COMPRA")) 
end if

if(rsAutoEmpresa("TIPO_DOC")="4")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Carta Compromiso N° "&rsAutoEmpresa("ORDEN_COMPRA"))
end if

texto_fichero = Replace(texto_fichero,"<!--#doc_tipo-->","http://norte.otecmutual.cl/ordenes/"&rsAutoEmpresa("DOCUMENTO_COMPROMISO"))

dim queryTabla
queryTabla="select TRABAJADOR.RUT,dbo.MayMinTexto(TRABAJADOR.NOMBRES) as NOMBRES from HISTORICO_CURSOS "
queryTabla=queryTabla&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.id_trabajador "
queryTabla=queryTabla&" where HISTORICO_CURSOS.ID_AUTORIZACION='"&vid&"' order by TRABAJADOR.NOMBRES asc "

set rsQueryTabla = conn.execute (queryTabla)

dim textoTabla
textoTabla=""
while not rsQueryTabla.eof
	textoTabla=textoTabla&"<tr>"
	textoTabla=textoTabla&"<td>"&replace(FormatNumber(mid(rsQueryTabla("RUT"), 1,len(rsQueryTabla("RUT"))-2),0)&mid(rsQueryTabla("RUT"), len(rsQueryTabla("RUT"))-1,len(rsQueryTabla("RUT"))),",",".")&"</td>"
	textoTabla=textoTabla&"<td>"&rsQueryTabla("NOMBRES")&"</td>"
	textoTabla=textoTabla&"</tr>"
rsQueryTabla.Movenext
wend
texto_fichero = Replace(texto_fichero,"<!--#tabla-->",textoTabla)  

With iMsg
.To = rsAutoEmpresa("EMAIL")
.BCC = "notificaciones@otecmutual.cl"
.From = "Mutual Capacitación <notificaciones@otecmutual.cl>"
.Subject = "Inscripción a Curso - Eliminada"
.HTMLBody = texto_fichero
.Sender = "Mutual Capacitación"
.Organization = "Mutual Capacitación"
Set .Configuration = iConf
SendEmailGmail = .Send
End With

set iMsg = nothing
set iConf = nothing
set Flds = nothing

conn.execute ("delete from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_AUTORIZACION="&vid)

conn.execute ("delete from AUTORIZACION where AUTORIZACION.ID_AUTORIZACION="&vid)

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</insertar>") 
%>