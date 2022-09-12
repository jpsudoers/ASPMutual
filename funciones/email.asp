<%
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
archivo= Server.MapPath("mail.html")
'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 
'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 
'cerramos el fichero 
fich.close() 
texto_fichero = Replace(texto_fichero,"<!--#empresa-->","mario")
'response.End()

With iMsg
.To = "notificaciones@otecmutual.cl"
.From = "Solicitud Inscripción <notificaciones@otecmutual.cl>"
.Subject = "Solicitud Inscripción"
.HTMLBody = texto_fichero
.Sender = "Mutual Capacitación"
.Organization = "Mutual Capacitación"
Set .Configuration = iConf
SendEmailGmail = .Send
End With

set iMsg = nothing
set iConf = nothing
set Flds = nothing
Response.Redirect("../index.asp")
%>
