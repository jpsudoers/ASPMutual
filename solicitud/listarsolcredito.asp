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

orden = "DESC"
if(Request("sord") <> "")then orden = Request("sord")

columna = "SOLICITUDES_CREDITO.FECHA_MODIFICACION"
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

vFEstado = request("estado")

sqlFilestado = " AND SOLICITUDES_CREDITO.ID_ESTADO <> 0 "

if vFEstado <> "" AND vFEstado <> "0" then
	sqlFilestado = " AND SOLICITUDES_CREDITO.ID_ESTADO =  " & vFEstado
end if

sql = "select SOLICITUDES_CREDITO.ID_SOLICITUD, SOLICITUDES_CREDITO.RUT_EMPRESA, SOLICITUDES_CREDITO.NOMBRE_EMPRESA, SOLICITUDES_CREDITO.FECHA_SOLICITUD, SOLICITUDES_CREDITO.ID_ESTADO"
sql = sql&", SOLICITUDES_CREDITO.USUARIO_INGRESO, SOLICITUDES_CREDITO.USUARIO_MODIFICA, SOLICITUDES_CREDITO.FECHA_MODIFICACION, "
sql = sql&" SOLICITUDES_CREDITO_ESTADOS.NOMBRE_ESTADO as ESTADO_DESC,'Linea'=SOLICITUDES_CREDITO.[CreditoSol]+'/'+SOLICITUDES_CREDITO.[CreditoPlazo]+'/'+SOLICITUDES_CREDITO.[CreditoDia]+'/'+[CreditoEncargado1]+'/'+[CreditoCorreo1]+'/'+[CreditoFono1]+'/'+[CreditoEncargado2]+'/'+[CreditoCorreo2]+'/'+[CreditoFono2]+'/'+[CreditoRepLegal]+'/'+[CreditoRepLegalCorreo]+'/'+[CreditoRepLegalFono]+'/'+REPLACE((case when [cond_OC]='1' then 'OC, ' else '' end)+(case when [cond_HES]='1' then 'HES, '  else '' end)+(case when [cond_MIGO]='1' then 'MIGO, '  else '' end)+(case when [cond_OTRO]='1' then [cond_OTROTxt]  else '' end)+';',', ;',''),[CreditoHorario],[CreditoInfo] from SOLICITUDES_CREDITO, SOLICITUDES_CREDITO_ESTADOS "
sql = sql&" where SOLICITUDES_CREDITO.ID_ESTADO = SOLICITUDES_CREDITO_ESTADOS.ID_ESTADO "& sqlFilestado &" ORDER BY "&columna&" "&orden

DATOS.Open sql, oConn

total_filas = 0
total_pages = 0

total_filas = DATOS.RecordCount  
if( total_filas >0 )then 
	IF((total_filas MOD limite)>0)THEN
		total_pages = FIX(total_filas/limite )+1
	ELSE
		total_pages = (total_filas/limite)
	END IF	
END IF	


if (pagina > total_pages) then pagina=total_pages
inicio = (limite*pagina) - limite 
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<page>"&pagina&"</page>"&chr(13))
Response.Write("<total>"&total_pages&"</total>"&chr(13))
Response.Write("<records>"&total_filas&"</records>"&chr(13))
'Response.Write("<sql>"&sql&"</sql>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
		if(fila>=inicio AND fila<(limite*pagina))then
			Response.Write("<row id="""&fila&""">"&chr(13))
			Response.Write("<cell><![CDATA["&right("000000000"&DATOS("ID_SOLICITUD"),5)&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("RUT_EMPRESA")&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("NOMBRE_EMPRESA")&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("FECHA_SOLICITUD")&"]]></cell>"&chr(13))
			Response.Write("<cell><![CDATA["&DATOS("ESTADO_DESC")&"]]></cell>"&chr(13))
	 
			'if(Session("activa_empresa")="1")then
			Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Cambiar Estado"" class=""ui-icon ui-icon-pencil"" name=""aSolicitud"" onclick=""update("&DATOS("ID_SOLICITUD")&",'"&DATOS("RUT_EMPRESA")&"','"&DATOS("NOMBRE_EMPRESA")&"',"&DATOS("ID_ESTADO")&",'"&DATOS("Linea")&"','"&DATOS("CreditoHorario")&"','"&DATOS("CreditoInfo")&"')""></a></span>]]></cell>"&chr(13))
			'else
				'if(Session("activa_empresa")="0")then
				'	Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Cambiar Estado"" class=""ui-icon ui-icon-pencil"" name=""aSolicitud"" onclick=""bloqueo();""></a></span>]]></cell>"&chr(13))
				'else
				    Response.Write("<cell><![CDATA[<span class=""ui-state-valid"" ><a href=""#"" title=""Eliminar"" class=""ui-icon ui-icon-trash"" name=""eSolicitud"" onclick=""eliminar("&DATOS("ID_SOLICITUD")&");""></a></span>]]></cell>"&chr(13))
				'end if
			'end if
			
			Response.Write("</row>"&chr(13)) 
		end if
	fila=fila+1
	DATOS.MoveNext
WEND


Response.Write("</rows>") 
%>
