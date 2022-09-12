<!--#include file="../conexion.asp"-->
<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%
on error resume next
vempresa = trim(Request("resRutEmpresa"))
vcorreo = trim(Request("resEmailEmpresa"))

query = "select(select COUNT(*) from EMPRESAS where RUT='"&vempresa&"' and EMAIL='"&vcorreo&"' "
query = query&" and PASSWORD_COORDINACION is not null and PASSWORD_COORDINACION<>'' and PREINSCRITA=1) as cordinacion, "
query = query&"(select COUNT(*) from EMPRESAS  where RUT='"&vempresa&"' and EMAIL_CONTA='"&vcorreo&"' "
query = query&" and PASSWORD_CONTA is not Null  and PASSWORD_CONTA<>'' and PREINSCRITA=1) as contabilidad "

set rsResContra = conn.execute (query)

Dim my_num,max,min
min=10001
max=50001

Randomize
my_num=int((max-min+1)*rnd+min)

dim enviar
enviar=0
dim userMail
userMail=""
dim validoRc
validoRc="0"
		if(rsResContra("cordinacion")="1")then
		   conn.execute ("UPDATE EMPRESAS SET PASSWORD_COORDINACION='"&my_num&"' WHERE rut='"&vempresa&"' and PREINSCRITA=1")
		   set rsUser = conn.execute ("select dbo.MayMinTexto(NOMBRES) as NOMBRES from EMPRESAS where RUT='"&vempresa&"'")
		   userMail=rsUser("NOMBRES")
		   enviar=1
		end if
		
		if(rsResContra("contabilidad")="1")then
		   conn.execute ("UPDATE EMPRESAS SET PASSWORD_CONTA='"&my_num&"' WHERE rut='"&vempresa&"' and PREINSCRITA=1")
		   set rsUser = conn.execute ("select dbo.MayMinTexto(NOMBRE_CONTA) as NOMBRES from EMPRESAS where RUT='"&vempresa&"'")
		   userMail=rsUser("NOMBRES")
		   enviar=1
		end if
	
	    if(enviar=1)then
		    validoRc="1"
		
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
			
			texto_fichero = Replace(texto_fichero,"<!--#usuarioCorreo-->",userMail)
			texto_fichero = Replace(texto_fichero,"<!--#rut_empresa-->",vempresa) 
			texto_fichero = Replace(texto_fichero,"<!--#usuario-->",vcorreo)
			texto_fichero = Replace(texto_fichero,"<!--#pass-->",my_num)
			
			With iMsg
			.To = vcorreo
			.BCC = "notificaciones@otecmutual.cl"
			.From = "Mutual Capacitación <notificaciones@otecmutual.cl>"
			.Subject = "Restablecer Contraseña"
			.HTMLBody = texto_fichero
			.Sender = "Mutual Capacitación"
			.Organization = "Mutual Capacitación"
			Set .Configuration = iConf
			SendEmailGmail = .Send
			End With
			
			set iMsg = nothing
			set iConf = nothing
			set Flds = nothing
	end if
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<DATOS>"&chr(13)) 
Response.Write("<row>"&chr(13))
Response.Write("<Valido>"&validoRc&"</Valido>"&chr(13))
Response.Write("</row>"&chr(13))
Response.Write("</DATOS>") 
%>