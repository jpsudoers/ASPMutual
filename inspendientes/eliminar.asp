<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vid = Request("idElimPreins") 

dim upPreins
upPreins="UPDATE preinscripciones set estado=0, situacion=3, situacion_fecha=GETDATE (), "
upPreins=upPreins&"situacion_encargado='"&Session("usuarioMutual")&"', "
upPreins=upPreins&"situacion_razon='"&Request("razonElim")&"' where id_preinscripcion="&vid

conn.execute (upPreins)

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
Flds.Item(schema & "sendpassword") =  "admin_2019"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

'creamos el nombre del archivo 
archivo= Server.MapPath("mailEliminaSolicitud.html")
'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 
'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 
'cerramos el fichero 
fich.close() 

dim preInsEmpresa
preInsEmpresa="select EMPRESAS.ID_EMPRESA,EMPRESAS.EMAIL,EMPRESAS.R_SOCIAL,PROGRAMA.ID_PROGRAMA,"
preInsEmpresa=preInsEmpresa&"preinscripciones.participantes,preinscripciones.valor_total,preinscripciones.doc_compromiso, "
preInsEmpresa=preInsEmpresa&"tipo_compromiso,numero_compromiso "
preInsEmpresa=preInsEmpresa&" from preinscripciones "
preInsEmpresa=preInsEmpresa&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=preinscripciones.id_empresa "
preInsEmpresa=preInsEmpresa&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=preinscripciones.id_programacion "
preInsEmpresa=preInsEmpresa&" where preinscripciones.id_preinscripcion="&vid

set rsPreInsEmpresa = conn.execute (preInsEmpresa)

texto_fichero = Replace(texto_fichero,"<!--#empresa-->",rsPreInsEmpresa("R_SOCIAL"))

dim queryProg
queryProg = "select CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_ from PROGRAMA "
queryProg = queryProg&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL where PROGRAMA.ID_PROGRAMA='"&rsPreInsEmpresa("ID_PROGRAMA")&"'"

set rsQueryProg = conn.execute (queryProg)

texto_fichero = Replace(texto_fichero,"<!--#curso-->",rsQueryProg("NOMBRE_CURSO"))
texto_fichero = Replace(texto_fichero,"<!--#fecha-->",rsQueryProg("FECHA_INICIO_"))

texto_fichero = Replace(texto_fichero,"<!--#asistentes-->",rsPreInsEmpresa("participantes"))  
texto_fichero = Replace(texto_fichero,"<!--#valor-->","$ "&replace(FormatNumber(rsPreInsEmpresa("valor_total"),0),",","."))

if(rsPreInsEmpresa("tipo_compromiso")="0")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Orden de Compra N° "&rsPreInsEmpresa("numero_compromiso"))
end if

if(rsPreInsEmpresa("tipo_compromiso")="1")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Vale Vista N° "&rsPreInsEmpresa("numero_compromiso")) 
end if

if(rsPreInsEmpresa("tipo_compromiso")="2")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Depósito Cheque N° "&rsPreInsEmpresa("numero_compromiso")) 
end if

if(rsPreInsEmpresa("tipo_compromiso")="3")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Transferencia N° "&rsPreInsEmpresa("numero_compromiso")) 
end if

if(rsPreInsEmpresa("tipo_compromiso")="4")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Carta Compromiso N° "&rsPreInsEmpresa("numero_compromiso"))
end if

texto_fichero = Replace(texto_fichero,"<!--#doc_tipo-->","http://norte.otecmutual.cl/ordenes/"&rsPreInsEmpresa("doc_compromiso"))
texto_fichero = Replace(texto_fichero,"<!--#razon-->",Request("razonElim"))
'response.End()

dim queryTabla
queryTabla="select TRABAJADOR.RUT,dbo.MayMinTexto(TRABAJADOR.NOMBRES) as NOMBRES from PREINSCRIPCION_TRABAJADOR "
queryTabla=queryTabla&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=PREINSCRIPCION_TRABAJADOR.id_trabajador "
queryTabla=queryTabla&" where PREINSCRIPCION_TRABAJADOR.id_preinscripcion='"&vid&"' order by TRABAJADOR.NOMBRES asc "

set rsQueryTabla = conn.execute (queryTabla)

dim textoTabla
textoTabla=""
while not rsQueryTabla.eof
	textoTabla=textoTabla&"<tr>"
	textoTabla=textoTabla&"<td>"&replace(FormatNumber(mid(rsQueryTabla("RUT"), 1,len(rsQueryTabla("RUT"))-2),0)&mid(rsQueryTabla("RUT"), len(rsQueryTabla("RUT"))-1,len(rsQueryTabla("RUT"))),",",".")&"</td>"
	textoTabla=textoTabla&"<td colspan='2'>"&rsQueryTabla("NOMBRES")&"</td>"
	textoTabla=textoTabla&"</tr>"
rsQueryTabla.Movenext
wend
texto_fichero = Replace(texto_fichero,"<!--#tabla-->",textoTabla)  

With iMsg
.To = rsPreInsEmpresa("EMAIL")
.BCC = "notificaciones@otecmutual.cl"
.From = "Mutual Capacitación <notificaciones@otecmutual.cl>"
.Subject = "Solicitud de Inscripción ha Curso - Eliminada"
.HTMLBody = texto_fichero
.Sender = "Mutual Capacitación"
.Organization = "Mutual Capacitación"
Set .Configuration = iConf
SendEmailGmail = .Send
End With

set iMsg = nothing
set iConf = nothing
set Flds = nothing

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>