<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"

pagina=1
if(Request("page")<>"")then pagina = CInt(Request("page"))

limite = 1
if(Request("rows")<>"")then limite = CInt(Request("rows"))

orden = "DESC"
if(Request("sord") <> "")then orden = Request("sord")

columna = "CURRICULO.NOMBRE_CURSO"
if(Request("sidx") <> "")then columna = Request("sidx")

node=0
if(request("nodeid")<>"")then node=cInt(request("nodeid"))

%>

<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,CURRICULO.SENCE as SENCE, "
sql = sql&"CURRICULO.NOMBRE_CURSO,(PROGRAMA.CUPOS - (select COUNT(HISTORICO_CURSOS.ID_PROGRAMA) from HISTORICO_CURSOS "
sql = sql&"where HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA)) as ins_vacantes, "
sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_CIERRE, 105) as FECHA_CIERRE, "
sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_ "
sql = sql&" from PROGRAMA "
sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
sql = sql&" where PROGRAMA.ESTADO=1 AND PROGRAMA.VIGENCIA=1 "
sql = sql&" and PROGRAMA.FECHA_CIERRE>=CONVERT(date,GETDATE(), 105) "
sql = sql&" ORDER BY "&columna&" "&orden

'RESPONSE.Write(sql)
'RESPONSE.End()

DATOS.Open sql,oConn
total_pages = 0
if( DATOS.RecordCount >0 )then 
	IF((DATOS.RecordCount MOD limite)>0)THEN
		total_pages = FIX(DATOS.RecordCount/limite )+1
	ELSE
		total_pages = (DATOS.RecordCount/limite)
	END IF	
END IF	



if (pagina > total_pages) then pagina=total_pages
inicio = (limite*pagina) - limite 
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<page>"&pagina&"</page>"&chr(13))
Response.Write("<total>"&total_pages&"</total>"&chr(13))
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

'response.Write(sql)
'response.end()
lugar=""
fila=0
WHILE NOT DATOS.EOF
   if(DATOS("ins_vacantes")<>"0")then
		if(fila>=inicio AND fila<(limite*pagina))then
			Response.Write("<row id="""&fila&""">"&chr(13))
			Response.Write("<cell><![CDATA["&right("000000000"&DATOS("ID_PROGRAMA"),5)&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("CODIGO")&"]]></cell>"&chr(13))
			if(DATOS("SENCE")=0)then
			Response.Write("<cell><![CDATA[Si]]></cell>"&chr(13))
			else
			Response.Write("<cell><![CDATA[No]]></cell>"&chr(13))
			end if
			
			if(DATOS("ID_PROGRAMA")="187")then
				lugar=" - Santiago"
			else
				lugar=""
			end if
			
			Response.Write("<cell><![CDATA["&DATOS("NOMBRE_CURSO")&lugar&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("ins_vacantes")&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("FECHA_INICIO_")&"]]></cell>"&chr(13))
			
			if(Session("activa_empresa")="1")then
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_PROGRAMA")&","&DATOS("ins_vacantes")&")""></a></span>]]></cell>"&chr(13))
			else
				if(Session("activa_empresa")="0")then
					Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""bloqueo();""></a></span>]]></cell>"&chr(13))
				else
				    Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""cierraSession();""></a></span>]]></cell>"&chr(13))
				end if
			end if
			
			Response.Write("</row>"&chr(13)) 
		end if
	fila=fila+1
	end if
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
