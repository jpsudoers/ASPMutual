<!--#include file="../conexion.asp"-->
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

orden = "ASC"
if(Request("sord") <> "")then orden = Request("sord")

columna = "NOMBRE"
if(Request("sidx") <> "")then columna = Request("sidx")

node=0
if(request("nodeid")<>"")then node=cInt(request("nodeid"))

%>

<%

vid = request("rutEmpresa")
if(request("rutEmpresa")="undefined") then
	vid=0
end if
sql = "select ID_EMPRESAS_USUARIOS, ID_EMPRESA, NOMBRE, CARGO, TELEFONO, EMAIL, CONTRASENA, Rut_Empresa"
sql = sql&" from EMPRESAS_USUARIOS "
sql = sql&" where ID_EMPRESA= 0 and ESTADO= 1 and Rut_Empresa='"&vid&"'"


'RESPONSE.Write(sql)
'RESPONSE.End()

set RespuestaListado =  conn.execute(sql)
total_pages = 0
if( RespuestaListado.RecordCount >0 )then 
	IF((RespuestaListado.RecordCount MOD limite)>0)THEN
		total_pages = FIX(RespuestaListado.RecordCount/limite )+1
	ELSE
		total_pages = (RespuestaListado.RecordCount/limite)
	END IF	
ELSE
		total_pages = 1	
END IF	

if (pagina > total_pages) then pagina=total_pages
inicio = (limite*pagina) - limite 
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<page>"&pagina&"</page>"&chr(13))
Response.Write("<total>"&total_pages&"</total>"&chr(13))
Response.Write("<records>"&RespuestaListado.RecordCount&"</records>"&chr(13))

fila=0
WHILE NOT RespuestaListado.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA["&RespuestaListado("Rut_Empresa")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&RespuestaListado("NOMBRE")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&RespuestaListado("CARGO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&RespuestaListado("TELEFONO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&RespuestaListado("EMAIL")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&RespuestaListado("CONTRASENA")&"]]></cell>"&chr(13))
        Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""eliminar2("&RespuestaListado("ID_EMPRESAS_USUARIOS")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
  fila=fila+1
  RespuestaListado.MoveNext
WEND
Response.Write("</rows>") 
'response.write(RespuestaListado)

%>