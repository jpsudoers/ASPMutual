<!--#include file="../conexion.asp"-->
<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<bloquear>") 

on error resume next

dim vestado
if(Request("idBloquearEstado")="0")then
vestado="1"
else
vestado="0"
end if

vemp=Request("idBloquearIDEmpresa")
'vuser=Request("idBloquearIDUsuario")
'vcli=Request("idBloquearIDCliente")
'vrazon=Request("razonBloquear")
'vid=Request("tabProgId")

dim query
'query = "UPDATE EMPRESAS SET ACTIVA='"&vestado&"' WHERE ID_EMPRESA='"&vemp&"'"
query="IF EXISTS (SELECT * FROM EMPRESA_TIPO_COMPROMISO WHERE ID_EMPRESA="&vemp&" AND ID_COMPROMISO_PAGO=1) "
query=query+"UPDATE EMPRESA_TIPO_COMPROMISO SET ESTADO_EMPRESA_COMPROMISO="&vestado&" WHERE ID_EMPRESA='"&vemp&"' AND ID_COMPROMISO_PAGO=1 "
query=query+"ELSE INSERT INTO EMPRESA_TIPO_COMPROMISO(ID_EMPRESA, ID_COMPROMISO_PAGO, ESTADO_EMPRESA_COMPROMISO, FECHA_EMPRESA_COMPROMISO) VALUES ("&vemp&", 1, 0, GETDATE())"
conn.execute (query)

'insSQL="IF NOT EXISTS (select * from EMPRESAS_BLOQUEOS EB where EB.ID_EMPRESA='"&vemp&"' AND EB.COD_INGRESO='"&vid&"') BEGIN "
'insSQL=insSQL&"insert into EMPRESAS_BLOQUEOS (ID_EMPRESA,ID_USUARIO,ESTADO,FECHA_ESTADO,DESCRIPCION,COD_INGRESO) values ('"&vemp&"','"&vuser&"','"&vestado&"',GETDATE(),'"&vrazon&"','"&vid&"') END"

'response.Write(insSQL)
'response.End()
'conn.execute (insSQL)

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</bloquear>") 
%>