<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/pdfTable.asp"-->
<!--#include file="Repor_Cierre.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

Server.ScriptTimeout = 1000000000

dim estado_eval,estado_eval2

For i = 1 To Request("countFilas") Step 1
if(Request("E"&i)="A")then
	estado_eval="Aprobado"
elseif(Request("E"&i)="R")then
	estado_eval="Reprobado"
elseif(Request("E"&i)="C")then
	estado_eval="Con Obs."
end if

if(Request("Er"&i)="A")then
	estado_eval2="Aprobado"
elseif(Request("Er"&i)="R")then
	estado_eval2="Reprobado"
elseif(Request("Er"&i)="C")then
	estado_eval2="Con Obs."
end if

if(Request("A"&i)<>"" AND Request("C"&i)<>"")then
upHistorico = "update HISTORICO_CURSOS set ASIS_CDN='"&Request("A"&i)&"', "&_
              "CAL_CDN='"&Request("C"&i)&"', EVA_CDN='"&estado_eval&"' where ID_HISTORICO_CURSO='"&Request("H"&i)&"'"

conn.execute (upHistorico)			  
end if

if(Request("E_DT"&i)="1")then
	ModCdn="select ASIS_REL,CAL_REL,EVA_REL from HISTORICO_CURSOS hc where hc.ID_HISTORICO_CURSO='"&Request("H"&i)&"'"
	
	set rsModCdn = conn.execute (ModCdn)
		
	if not rsModCdn.eof and not rsModCdn.bof then 
		movCdn = "insert into DETALLE_ACTIVIDAD_CIERRE_CURSO(ID_HISTORICO_CURSOS,ASISTENCIA,CALIFICACION,EVALUACION,"&_	
			     "ASISTENCIA_NUEVA,CALIFICACION_NUEVA,EVALUACION_NUEVA,FECHA_REALIZACION,ID_USUARIO)"&_
		         " values('"&Request("H"&i)&"','"&rsModCdn("ASIS_REL")&"','"&rsModCdn("CAL_REL")&"',"&_
		  "'"&rsModCdn("EVA_REL")&"','"&Request("Ar"&i)&"','"&Request("Cr"&i)&"','"&estado_eval2&"',GETDATE(),"&Request("u")&")"
		
		conn.execute (movCdn)
	end if

	upHistorico2 = "update HISTORICO_CURSOS set ASIS_REL='"&Request("Ar"&i)&"', "&_
              "CAL_REL='"&Request("Cr"&i)&"', EVA_REL='"&estado_eval2&"' where ID_HISTORICO_CURSO='"&Request("H"&i)&"'"
			  
	conn.execute (upHistorico2)		
end if

 if(Request("ec")="1")then
	if(Request("A"&i)<>"" AND Request("C"&i)<>"")then
		query = "update HISTORICO_CURSOS set ASISTENCIA='"&Request("A"&i)&"', "
		query = query&"CALIFICACION='"&Request("C"&i)&"', EVALUACION='"&estado_eval&"', ESTADO=2 "
		query = query&" where ID_HISTORICO_CURSO='"&Request("H"&i)&"'"
	else
		query = "update HISTORICO_CURSOS set ASISTENCIA='"&Request("Ar"&i)&"', "
		query = query&"CALIFICACION='"&Request("Cr"&i)&"', EVALUACION='"&estado_eval2&"', ESTADO=2 "
		query = query&" where ID_HISTORICO_CURSO='"&Request("H"&i)&"'"
	end if

	conn.execute (query)	
 end if
Next	

	   conn.execute ("exec dbo.ACTUALIZA_ESTATUS")
	
if(Request("ec")="1")then
	upBloque="UPDATE bloque_programacion SET estado_eva_cdn=1, id_usuario_cdn="&Request("u")&", fecha_rev_cdn=GETDATE(), doc_cierre='"&"Reporte_Cierre_Curso_"&fecha&".pdf"&"' WHERE ID_BLOQUE="&Request("b")&";"
	
	conn.execute (upBloque)
	
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
	Flds.Item(schema & "sendpassword") = "admin_2019"
	Flds.Item(schema & "smtpusessl") = 1
	Flds.Update
	
	'creamos el nombre del archivo 
	archivo= Server.MapPath("mailReporte.html")
	'conectamos con el FSO 
	set confile = createObject("scripting.filesystemobject") 
	'volvemos a abrir el fichero para lectura 
	set fich = confile.OpenTextFile(archivo) 
	
	'leemos el contenido del fichero 
	texto_fichero = fich.readAll() 
	'cerramos el fichero 
	fich.close() 
	
	DTprog = "select  dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO, "&_
    "CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "&_
    "CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, "&_
    "REL=dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO), "&_
    "(CASE WHEN SEDES.ID_SEDE =  27 THEN bloque_programacion.nom_sede  "&_
    "WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as DIR  "&_
    " from bloque_programacion   "&_
    " inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=bloque_programacion.id_programa "&_
    " inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL 	 "&_
    " inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede    "&_
    " inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator	 "&_
    " where bloque_programacion.ID_BLOQUE="&Request("b")
	
	
	set rsQueryProg = conn.execute (DTprog)
	
	texto_fichero = Replace(texto_fichero,"<!--#curso-->",rsQueryProg("NOMBRE_CURSO"))
	texto_fichero = Replace(texto_fichero,"<!--#urlExcel-->","http://norte.otecmutual.cl/evaluacioncierre/Repor_Cierre_xls.asp?b="&Request("b"))
	texto_fichero = Replace(texto_fichero,"<!--#inicio-->",rsQueryProg("FECHA_INICIO_"))
	texto_fichero = Replace(texto_fichero,"<!--#termino-->",rsQueryProg("FECHA_TERMINO"))
	texto_fichero = Replace(texto_fichero,"<!--#lugar-->",rsQueryProg("DIR"))
	texto_fichero = Replace(texto_fichero,"<!--#relator-->",rsQueryProg("rel"))

	TrabAuto = "select T.RUT,T.NOMBRES,T.NACIONALIDAD,T.ID_EXTRANJERO, "&_
			 " hc.ASISTENCIA,hc.CALIFICACION,hc.EVALUACION, "&_
			 " hc.ASIS_CDN,hc.CAL_CDN,hc.EVA_CDN, "&_
			 " hc.ASIS_REL,hc.CAL_REL,hc.EVA_REL "&_
			 " from HISTORICO_CURSOS hc "&_
			 " inner join trabajador t on t.ID_TRABAJADOR=hc.ID_TRABAJADOR "&_
			 " inner join bloque_programacion bq on bq.id_bloque=hc.ID_BLOQUE "&_
			 " where bq.ID_BLOQUE="&Request("b")
	
	set rsQueryTabla = conn.execute (TrabAuto)

	textoTabla=""
	while not rsQueryTabla.eof
		textoTabla=textoTabla&"<tr>"
		textoTabla=textoTabla&"<td>"&replace(FormatNumber(mid(rsQueryTabla("RUT"), 1,len(rsQueryTabla("RUT"))-2),0)&mid(rsQueryTabla("RUT"), len(rsQueryTabla("RUT"))-1,len(rsQueryTabla("RUT"))),",",".")&"</td>"
		textoTabla=textoTabla&"<td>"&rsQueryTabla("NOMBRES")&"</td>"
		textoTabla=textoTabla&"<td>"&rsQueryTabla("ASIS_CDN")&"</td>"
		textoTabla=textoTabla&"<td>"&rsQueryTabla("CAL_CDN")&"</td>"
		textoTabla=textoTabla&"<td>"&rsQueryTabla("EVA_CDN")&"</td>"
		textoTabla=textoTabla&"<td>"&rsQueryTabla("ASIS_REL")&"</td>"
		textoTabla=textoTabla&"<td>"&rsQueryTabla("CAL_REL")&"</td>"
		textoTabla=textoTabla&"<td>"&rsQueryTabla("EVA_REL")&"</td>"		
		textoTabla=textoTabla&"</tr>"
	rsQueryTabla.Movenext
	wend
	
	 texto_fichero = Replace(texto_fichero,"<!--#tabla-->",textoTabla) 	
	
	
	 TrabAutoMod = "select T.RUT,T.NOMBRES,T.NACIONALIDAD,T.ID_EXTRANJERO,"&_
     "dt.ASISTENCIA,dt.CALIFICACION,dt.EVALUACION,"&_
     "dt.ASISTENCIA_NUEVA,dt.CALIFICACION_NUEVA,dt.EVALUACION_NUEVA,dt.FECHA_REALIZACION"&_
     " from HISTORICO_CURSOS hc "&_
     " inner join trabajador t on t.ID_TRABAJADOR=hc.ID_TRABAJADOR "&_
     " inner join bloque_programacion bq on bq.id_bloque=hc.ID_BLOQUE "&_
     " inner join DETALLE_ACTIVIDAD_CIERRE_CURSO dt on dt.ID_HISTORICO_CURSOS=hc.ID_HISTORICO_CURSO "&_
     " where bq.ID_BLOQUE="&Request("b")&_ 
	 " order by dt.FECHA_REALIZACION desc"
	 
	set rsQueryTabla2 = conn.execute (TrabAutoMod)
		
	textoTabla2=""
	while not rsQueryTabla2.eof
		textoTabla2=textoTabla2&"<tr>"
		textoTabla2=textoTabla2&"<td>"&rsQueryTabla2("FECHA_REALIZACION")&"</td>"		
		textoTabla2=textoTabla2&"<td>"&replace(FormatNumber(mid(rsQueryTabla2("RUT"), 1,len(rsQueryTabla2("RUT"))-2),0)&mid(rsQueryTabla2("RUT"), len(rsQueryTabla2("RUT"))-1,len(rsQueryTabla2("RUT"))),",",".")&"</td>"
		textoTabla2=textoTabla2&"<td>"&rsQueryTabla2("NOMBRES")&"</td>"
		textoTabla2=textoTabla2&"<td>"&rsQueryTabla2("ASISTENCIA")&"</td>"
		textoTabla2=textoTabla2&"<td>"&rsQueryTabla2("CALIFICACION")&"</td>"
		textoTabla2=textoTabla2&"<td>"&rsQueryTabla2("EVALUACION")&"</td>"
		textoTabla2=textoTabla2&"<td>"&rsQueryTabla2("ASISTENCIA_NUEVA")&"</td>"
		textoTabla2=textoTabla2&"<td>"&rsQueryTabla2("CALIFICACION_NUEVA")&"</td>"
		textoTabla2=textoTabla2&"<td>"&rsQueryTabla2("EVALUACION_NUEVA")&"</td>"		
		textoTabla2=textoTabla2&"</tr>"
	rsQueryTabla2.Movenext
	wend
	
	texto_fichero = Replace(texto_fichero,"<!--#tabla2-->",textoTabla2) 	
	
	
	With iMsg
	.To = "gpcontreras@mutualcapacitacion.cl, alaguilera@mutualasesorias.cl"
	.BCC = "notificaciones@otecmutual.cl"
	.From = "Cierre de Curso <notificaciones@otecmutual.cl>"
	.Subject = "Cierre de Curso"
	.HTMLBody = texto_fichero
	.AddAttachment "C:\inetpub\wwwroot\Sitios\norte\mas.subcontrataley.cl\pdf\Reporte_Cierre_Curso_"&fecha&".pdf"
	.Sender = "Mutual Capacitación"
	.Organization = "Mutual Capacitación"
	Set .Configuration = iConf
	SendEmailGmail = .Send
	End With
	
	set iMsg = nothing
	set iConf = nothing
	set Flds = nothing
	
end if



'Response.Write("<sql>"&query&"</sql>")
'if err.number <> 0 then
	'Response.Write("<commit>false</commit>")
'else
	'Response.Write("<commit>true</commit>")
'end if
'Response.Write("</modificar>") 
%>