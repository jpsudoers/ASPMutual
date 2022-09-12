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

columna = "TRABAJADOR.RUT"
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

sql = " SELECT E.ID_EMPRESA, ID_EMP_INS_USR=0, E.NOMBRE_CONTA NOMBRES, E.EMAIL_CONTA EMAIL, E.FONO_CONTABILIDAD FONO, "&_
	  " E.CARGO_CONTA CARGO, 'Contacto Contabilidad' COMENTARIO, VAL=0, ID_EMP_USU=0 FROM EMPRESAS E "&_
	  " WHERE E.ID_EMPRESA = "&Request("id")&_
	  " UNION "&_
	  " SELECT ID_EMPRESA, ID_EMP_INS_USR, U.NOMBRES, U.EMAIL, U.FONO, U.CARGO, COMENTARIO, 1 VAL, U.ID_EMP_USU "&_
	  " FROM EMP_INS_USR LEFT JOIN EMPRESA_USUARIOS U ON (U.ID_EMP_USU = EMP_INS_USR.ID_EMP_USU) "&_
	  " WHERE ESTADO = 1 AND TIPO_CONTACTO = 4 AND ID_EMPRESA = "&Request("id")&_
	  " ORDER BY VAL, ID_EMP_INS_USR"

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
		
		Response.Write("<cell><![CDATA["&DATOS("NOMBRES")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("EMAIL")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&DATOS("FONO")&"]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA["&DATOS("CARGO")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA["&DATOS("COMENTARIO")&"]]></cell>"&chr(13))		
		Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Modificar Registro"" class=""ui-icon ui-icon-pencil"" onclick=""update_usr("&DATOS("ID_EMP_USU")&","&DATOS("ID_EMPRESA")&","&DATOS("VAL")&")""></a></span>]]></cell>"&chr(13))
        if(DATOS("VAL")="1") then
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-trash"" onclick=""delete_usr("&DATOS("ID_EMP_USU")&","&DATOS("ID_EMPRESA")&")""></a></span>]]></cell>"&chr(13))
		else
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar Registro"" class=""ui-icon ui-icon-cancel""></a></span>]]></cell>"&chr(13))
		end if
		Response.Write("</row>"&chr(13))
	end if
  fila=fila+1
  DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>
