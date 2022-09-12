
<!--#include file="../conexion.asp"-->

<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
  
'on error resume next

voperacion = Request("operacion")

if  voperacion = "guardarsolicitud" then

	Response.ContentType="text/xml"
	Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
	Response.Write("<data>")

	vIU = "ACTUALIZO"
	vsolId = Request("solId")
	vtxRut= Request("txRut")
	vtxtNombreEmpresa= Request("txtNombreEmpresa")

	if vsolId = ""  then

		query = "insert into SOLICITUDES_CREDITO (RUT_EMPRESA,NOMBRE_EMPRESA,FECHA_SOLICITUD,ID_ESTADO,USUARIO_INGRESO)  "
		query = query & " values ('"& vtxRut &"','"& vtxtNombreEmpresa &"', GETDATE(),1, '"& Session("usuario") &"') "

		conn.execute (query)

		'sql = "select max(ID_SOLICITUD) as solId  from SOLICITUDES_CREDITO "
		
		'SET rs = conn.execute(sql)
		'While not rs.EOF
		'	vsolId = rs("solId")
		'	rs.MoveNext
		'wend

	else

		query = "update SOLICITUDES_CREDITO SET RUT_EMPRESA = '" & vtxRut & "', NOMBRE_EMPRESA = '" & vtxtNombreEmpresa & "', "
		query = query & " FECHA_SOLICITUD = GETDATE(), ID_ESTADO = 1, USUARIO_INGRESO = '"& Session("usuario") &"'  "
		query = query &" where ID_SOLICITUD = " & vsolId

		conn.execute (query)

	end if
	Response.Write("<error>false</error>")
	Response.Write("</data>") 

elseif voperacion = "delSolicitud" then
		
		Response.ContentType="text/xml"
		Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
		Response.Write("<data>")

		vsolId = Request("idSolicitud")

		if vsolId = "" then
			Response.Write("<error>true</error>")
		else

			query = "delete from SOLICITUDES_CREDITO_ARCHIVOS  where ID_SOLICITUD = " & vsolId
			conn.execute (query)

			query = "delete from SOLICITUDES_CREDITO  where ID_SOLICITUD = " & vsolId
			conn.execute (query)

		end if
	Response.Write("</data>") 

elseif voperacion = "delDocumento" then
		
		Response.ContentType="text/xml"
		Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
		Response.Write("<data>")

		vcorrelativo = Request("correlativo")

		if vcorrelativo = "" then
			Response.Write("<borro>false</borro>")
		else

			query = "delete from SOLICITUDES_CREDITO_ARCHIVOS  where CORRELATIVO = " & vcorrelativo
			conn.execute (query)
			Response.Write("<borro>true</borro>")

		end if
	Response.Write("</data>") 

elseif voperacion = "upEstado" then
		vsolId = Request("solId")
		vestado = Request("selEstadoSolicitud")

		if vsolId = "" then
			Response.Write("<error>true</error>")
		else

			query = "update SOLICITUDES_CREDITO SET  FECHA_MODIFICACION = GETDATE(), ID_ESTADO = "& vestado &", USUARIO_MODIFICA = '"& Session("usuario") &"'  "
			query = query &" where ID_SOLICITUD = " & vsolId

			conn.execute (query)

		end if

		Response.Write("</data>") 
elseif voperacion = "sugEmpresa" then
	query = "Select  * "
	query = query&" FROM EMPRESAS"
	query = query&" WHERE EMPRESAS.RUT LIKE '%"&Request("txt")&"%' AND EMPRESAS.TIPO<>3 AND EMPRESAS.PREINSCRITA=1 "
	query = query&" AND EMPRESAS.ESTADO=1 "
	query = query&" OR EMPRESAS.R_SOCIAL LIKE '%"&Request("txt")&"%' AND EMPRESAS.TIPO<>3 AND EMPRESAS.PREINSCRITA=1 "
	query = query&" AND EMPRESAS.ESTADO=1 "

	SET rs = conn.execute(query)
	While not rs.EOF
		Response.Write("<li onclick=""fill('"&rs("R_SOCIAL")&"');"">("&rs("RUT")&") "&ucase(rs("R_SOCIAL"))&"</li>")
		rs.MoveNext
	wend

elseif voperacion = "listaDocumentos" then

	vsolId = Request("solId")
	vborrar = Request("borrar")

	if vsolId <> "" then

		query = " Select  * "
		query = query & " FROM SOLICITUDES_CREDITO_ARCHIVOS "
		query = query & " WHERE ID_SOLICITUD = "& vsolId  
		strconten = ""

		SET rs = conn.execute(query)
		While not rs.EOF

			funBorrar = ""
			relativo = "solicitud/"
			if vborrar = "1" then
				funBorrar = "<a href=""javascript:EliminarDoc('"& rs("CORRELATIVO") &"'); "" class='Spr2-SV'> " 
				funBorrar = funBorrar & "<div class='eliminar_img_doc'><img src='images/ico_limpiar.png' width='24' height='24' alt=''/></div></a>"
			else
				relativo ="solicitud/"
			end if


			nomarchivo = rs("NOMBRE_ARCHIVO")
			extension = Right(nomarchivo, 3)
			vRuta = relativo & "Documentos/" & rs("NOMBRE_ARCHIVO")
			xRecurso = ""
			

			if(ucase(extension) = "JPG" OR ucase(extension) = "PEG" OR ucase(extension) = "PNG" OR ucase(extension) = "GIF") then

				xRecurso = "style=""background-image: url('"& vRuta &"');  background-size: 100px 100px; """

			end if

			strconten = strconten & "<table id='DOC_CORR_"&rs("CORRELATIVO")&"' name='DOC_CORR_"&rs("CORRELATIVO")&"' class='fto_adjunta_table'><tr><td>" 
			strconten = strconten &	"<div class='fto_adjunta' "& xRecurso &"	> "
			strconten = strconten &	 funBorrar &" <a href='"& vRuta &"' target='_blank'><div  style='height:100%; width:100%'></div></a></div>" 
			strconten = strconten & "</td></tr><tr><td></td></tr></table>" 
			'strconten = strconten & "</td></tr><tr><td>"& nomarchivo &"</td></tr></table>"

			rs.MoveNext
		wend
		 
	end if

	RESPONSE.WRITE(strconten)

elseif voperacion = "listaEstados" then
	query = "Select  ID_ESTADO, NOMBRE_ESTADO FROM  SOLICITUDES_CREDITO_ESTADOS "

	SET rs = conn.execute(query)

	Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
	Response.Write("<estado>"&chr(13)) 

	While not rs.EOF

		Response.Write("<row>"&chr(13))
		Response.Write("<ID_ESTADO>"&rs("ID_ESTADO")&"</ID_ESTADO>"&chr(13))
		Response.Write("<NOMBRE_ESTADO>"&rs("NOMBRE_ESTADO")&"</NOMBRE_ESTADO>"&chr(13))
		Response.Write("</row>"&chr(13))

		rs.MoveNext
		
	wend

Response.Write("</estado>") 

elseif voperacion = "generarSolicitud" then

		query = "insert into SOLICITUDES_CREDITO (RUT_EMPRESA,NOMBRE_EMPRESA,FECHA_SOLICITUD,ID_ESTADO,USUARIO_INGRESO)  "
		query = query & " values ('1','', GETDATE(),0, '"& Session("usuario") &"') "

		conn.execute (query)

		sql = "select max(ID_SOLICITUD) as solId  from SOLICITUDES_CREDITO "
		
		SET rs = conn.execute(sql)
		While not rs.EOF
			vsolId = rs("solId")
			rs.MoveNext
		wend

		RESPONSE.WRITE(vsolId)


end if

'RESPONSE.WRITE(query)
'RESPONSE.End()



'if err.number <> 0 then
'	Response.Write("<salida>false</salida>")
'else
'	Response.Write("<salida>true</salida>")
'end if
'if voperacion <> "sugEmpresa" AND voperacion <> "listaDocumentos" AND voperacion <> "updocumento"  then
'	
'end if
%>