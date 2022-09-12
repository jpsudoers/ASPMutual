<!--#include file="../conexion.asp"-->
<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<Modificar>") 

on error resume next

dim query
For i = 1 To Request("totitemCostos") Step 1

det_Gasto="Null"
if(Request("IDCosto"&i)="10" and Request("txtGasto"&i)="1")then
	det_Gasto="'"&Request("txtDetItem"&i)&"'"
end if	

PROV=16
if(Request("id_prov"&i)<>"0" and Request("id_prov"&i)<>"" and Request("id_prov"&i)<>"16")then
	PROV="'"&Request("id_prov"&i)&"'"
end if	

query = "update GASTOS_PROGRAMA set MONTO='"&Request("txtMont"&i)&"', ESTADO="&Request("txtGasto"&i)&",DETALLE="&det_Gasto&",ID_PROVEEDORES="&PROV&" where ID_GASTOS_PROGRAMA='"&Request("txtIDCosto"&i)&"'"
 
  conn.execute (query)
Next	
Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</Modificar>") 
%>