<!--#include file="../conexion.asp"-->
<%

Response.CodePage = 65001
Response.CharSet = "utf-8"

query = "Select  * "
query = query&" FROM EMPRESAS"
query = query&" WHERE EMPRESAS.RUT LIKE '%"&Request("txt")&"%' "
query = query&" OR EMPRESAS.R_SOCIAL LIKE '%"&Request("txt")&"%' AND EMPRESAS.TIPO=1"

SET rs = conn.execute(query)
While not rs.EOF
	Response.Write("<li onclick=""fill2('"&rs("ID_EMPRESA")&"','"&rs("RUT")&"');"">("&rs("RUT")&") "&rs("R_SOCIAL")&"</li>")
	rs.MoveNext
wend
%>