<!--#include file="../cnn_string.asp"-->
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
columna = "R_SOCIAL"
if(Request("sidx") <> "")then columna = Request("sidx")

node=0
if(request("nodeid")<>"")then node=cInt(request("nodeid"))

%>

<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

'sql = "SELECT ID_EMPRESA,rut,r_social,CONVERT(VARCHAR(10),FECHA_SOLICITUD, 105) as FECHA_SOLICITUD from empresas "
'sql = sql&" WHERE ESTADO=1 AND TIPO=1 AND PREINSCRITA=0 ORDER BY FECHA_SOLICITUD asc"

sql = "SELECT E.ID_EMPRESA,E.rut,E.r_social,CONVERT(VARCHAR(10),E.FECHA_SOLICITUD, 105) as FECHA_SOLICITUD,"
sql = sql&" case (select Datediff(""d"", Min(CONVERT(date,Emp.FECHA_SOLICITUD,105)), Max(CONVERT(date,GETDATE(), 105))) "
sql = sql&" from EMPRESAS Emp where Emp.ID_EMPRESA=E.ID_EMPRESA) "
sql = sql&" when 0 then '#66ff00' "
sql = sql&" when 1 then '#ffff00' "
sql = sql&" when 2 then '#ff0000' "
sql = sql&" else '#ff0000' end as color "
sql = sql&"  from empresas E "
sql = sql&"  WHERE E.ESTADO=1 AND E.TIPO=1 AND E.PREINSCRITA=0 ORDER BY E.FECHA_SOLICITUD asc "

'response.Write(sql)
'response.End()

DATOS.Open sql,oConn
total_pages = 0
if( DATOS.RecordCount >0 )then 
	IF((DATOS.RecordCount MOD limite)>0)THEN
		total_pages = FIX(DATOS.RecordCount/limite )+1
	ELSE
		total_pages = (DATOS.RecordCount/limite)
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
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
	if(fila>=inicio AND fila<(limite*pagina))then
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA[<span style=""color:#000000; position:relative;"">"&DATOS("FECHA_SOLICITUD")&"<span style=""color:"&DATOS("color")&";top:-1px;left:-1px;position:absolute;"">"&DATOS("FECHA_SOLICITUD")&"</span></span>]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&DATOS("rut")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("r_social")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" name=""aContrato"" onclick=""update("&DATOS("ID_EMPRESA")&")""></a></span>]]></cell>"&chr(13))
Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" name=""aContrato"" onclick=""eliminar("&DATOS("ID_EMPRESA")&")""></a></span>]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
	end if
	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>