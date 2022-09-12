<!--#include file="../conexion.asp"-->
<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<Insertar>") 

on error resume next

dim query
For i = 1 To Request("totitemCostos") Step 1
	'query = "update bloque_programacion set estado_eva_cdn="&eva_cdn&", fecha_rev_cdn=GETDATE(), id_usuario_cdn='"&Request("idCusr")&"' where id_bloque='"&Request("idCbloque")&"'"

det_Gasto="Null"
if(Request("txtIDCosto"&i)="10" and Request("txtGasto"&i)="1")then
	det_Gasto="'"&Request("txtDetItem"&i)&"'"
end if	
	
PROV=16
if(Request("id_prov"&i)<>"0" and Request("id_prov"&i)<>"")then
	PROV="'"&Request("id_prov"&i)&"'"
end if	
	
	
  query = "insert into GASTOS_PROGRAMA(ID_GASTO,ID_BLOQUE,DETALLE,MONTO,ESTADO,ID_PROVEEDORES) values('"&Request("txtIDCosto"&i)&"','"&Request("idCbloque")&"',"&det_Gasto&",'"&Request("txtMont"&i)&"',"&Request("txtGasto"&i)&","&PROV&")"
  
  conn.execute (query)
Next	
Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</Insertar>") 
%>