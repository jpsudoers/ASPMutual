<%Response.CodePage = 65001
Response.CharSet = "utf-8"
'Response.ContentType="text/xml"
'Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
'Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vrut = Request("txRut")
vrSocial=Request("txtRsoc")
vgiro=Request("txtGiro")
vdir=Request("txtDir")
vcom=Request("txtCom")
vciu=Request("txtCiu")
vfon=Request("txtFon")

vfax="NULL"
if(Request("txtFax")<>"")then
	vfax="'"&Request("txtFax")&"'"
end if
'vcontratista=Request("Contratista")
vmut=Request("txtMutual")

vnomb=Request("txtNomb")
vmail=Request("txtEmail") 
vfono=Request("txtFono")
vcargo=Request("txtCargo")  

vnombcont=Request("txtNombCont")
vcargocont=Request("txtCargoCont")  
vemailcont=Request("txtEmailCont")
vfonocont=Request("txtFonoCont")

vigual=Request("contactoIgual")

FM1=Request("FM1")
FM2=Request("FM2")
FM3=Request("FM3")

dim query
query = "IF NOT EXISTS (select * from EMPRESAS e where e.rut='"&vrut&"' and e.ESTADO=1 and (e.PREINSCRITA=0 or e.PREINSCRITA=1)) BEGIN "
query = query&"insert into EMPRESAS (RUT,R_SOCIAL,GIRO,DIRECCION,COMUNA,CIUDAD,FONO,FAX,MUTUAL,ESTABLE,ESTADO "
query = query&",NOMBRES,EMAIL,CARGO,ID_OTIC,NOMBRE_CONTA,CARGO_CONTA,EMAIL_CONTA,FONO_CONTACTO,FONO_CONTABILIDAD,TIPO,PREINSCRITA"
query = query&",TIPO_CONTACTO,ACTIVA,FECHA_SOLICITUD)"',FONO_FM,FONO_CONTACTO_FM,FONO_CONTABILIDAD_FM
query = query&" values('"&vrut&"',dbo.MayMinTexto('"&vrSocial&"'),dbo.MayMinTexto('"&vgiro&"'),"
query = query&"dbo.MayMinTexto('"&vdir&"'),dbo.MayMinTexto('"&vcom&"'),dbo.MayMinTexto('"&vciu&"'),'"&vfon&"',"&vfax&","
query = query&"'"&vmut&"',0,1,dbo.MayMinTexto('"&vnomb&"'),LOWER('"&vmail&"'),dbo.MayMinTexto('"&vcargo&"'),0,"
query = query&"dbo.MayMinTexto('"&vnombcont&"'),dbo.MayMinTexto('"&vcargocont&"'),LOWER('"&vemailcont&"'),"
query = query&"'"&vfono&"','"&vfonocont&"',1,0,'"&vigual&"',1,GETDATE ()) END "'  ,"&FM1&","&FM2&","&FM3&""
conn.execute (query)
'RESPONSE.WRITE(query)
'RESPONSE.End()
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
archivo= Server.MapPath("mail.html")
'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 
'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 
'cerramos el fichero 
fich.close() 
texto_fichero = Replace(texto_fichero,"<!--#rut-->",vrut)
texto_fichero = Replace(texto_fichero,"<!--#empresa-->",vrSocial)
texto_fichero = Replace(texto_fichero,"<!--#giro-->",vgiro)
texto_fichero = Replace(texto_fichero,"<!--#direccion-->",vdir)
texto_fichero = Replace(texto_fichero,"<!--#comuna-->",vcom)
texto_fichero = Replace(texto_fichero,"<!--#ciudad-->",vciu)
texto_fichero = Replace(texto_fichero,"<!--#contacto-->",vnomb)
texto_fichero = Replace(texto_fichero,"<!--#fono-->",vfon)
texto_fichero = Replace(texto_fichero,"<!--#mail-->",vmail)

'response.End()

With iMsg
.To = "notificaciones@otecmutual.cl"
.From = "Solicitud de Nueva Empresa <notificaciones@otecmutual.cl>"
.Subject = "Solicitud de Nueva Empresa"
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
Response.Write("</insertar>") 
%>