<%
			Response.ContentType = "text/xml"
			Response.AddHeader "Cache-control", "private"
			Response.AddHeader "Expires", "-1"
			Response.CodePage = 65001
			Response.CharSet = "utf-8"

			on error resume next

			vempresa = trim(Request("rutEmpresa"))
			vcorreo = LCase(trim(Request("correoUser")))
			userMail = ULCASE(trim(Request("nomUser")))
			my_num = LCase(trim(Request("passUser")))
			
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
			.To = "notificaciones@otecmutual.cl"
			.BCC = vcorreo
			.From = "Mutual Capacitación <notificaciones@otecmutual.cl>"
			.Subject = "Mutual Capacitación Informa"
			.HTMLBody = texto_fichero
			.Sender = "Mutual Capacitación"
			.Organization = "Mutual Capacitación"
			Set .Configuration = iConf
			SendEmailGmail = .Send
			End With
			
			set iMsg = nothing
			set iConf = nothing
			set Flds = nothing
			
			Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
			Response.Write("<DATOS>"&chr(13)) 
			Response.Write("<row>"&chr(13))
			Response.Write("<Valido>1</Valido>"&chr(13))
			Response.Write("</row>"&chr(13))
			Response.Write("</DATOS>") 
			
			Function ULCASE(cad)
				CadConversion=""
				Cadena=cad
				cont=1
				siguiente=1
				While cont<=Len(Cadena)
					car= Mid(Cadena,cont,1)
					If siguiente=1 Then
						CadConversion= CadConversion & UCase(car)
					Else
						CadConversion= CadConversion & LCase(car)
					End If
					
					If (car=" ") Then
						siguiente=1
					Else
						siguiente=0
					End If
					
					cont= cont+1
				Wend
				ULCASE= CadConversion
			End Function
%>