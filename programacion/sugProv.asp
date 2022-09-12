<!--#include file="../conexion.asp"-->
<%

Response.CodePage = 65001
Response.CharSet = "utf-8"

vCliente=0
if(request("c")<>"")then vCliente=request("c") end if

query = "Select P.ID_PROVEEDORES,P.RUT,P.PROVEEDOR from PROVEEDORES p "&_
		" where p.ESTADO=1 and ((P.RUT LIKE '%"&Request("txt")&"%') "&_
		" OR (P.PROVEEDOR LIKE '%"&Request("txt")&"%'))"

SET rs = conn.execute(query)
While not rs.EOF
	Response.Write("<li onclick=""fillProv('"&rs("ID_PROVEEDORES")&"','"&rs("RUT")&"','"&request("i")&"');"">("&rs("RUT")&") "&rs("PROVEEDOR")&"</li>")
	rs.MoveNext
wend
%>