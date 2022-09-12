<!--#include file="../conexion.asp"-->
<%
on error resume next
dim query
query = "select em.ID_EMPRESA,LOWER (em.EMAIL) as EMAIL from EMPRESAS em "
query = query&" where em.TIPO=1 and em.PREINSCRITA=1 order by em.ID_EMPRESA asc"

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
Flds.Item(schema & "sendpassword") = "admin_2015"
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

dim dirCorreos

	dirCorreos=""
	set rsQueryTabla = conn.execute (query)
	while not rsQueryTabla.eof
		dirCorreos=dirCorreos&rsQueryTabla("EMAIL")&";"
		'dirCorreos="mgonzalez@subcontrataley.cl; mario.gonzalez20@hotmail.com"
		rsQueryTabla.MoveNext
	wend
	response.Write(dirCorreos)
	response.End()
	
	With iMsg
		'.To = "mgonzalez@subcontrataley.cl; mario.gonzalez20@hotmail.com"
		.BCC = dirCorreos
		.From = "Mutual Capacitación <notificaciones@otecmutual.cl>"
		.Subject = "Mutual Capacitación"
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
%>