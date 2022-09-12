<!--#include file="../conexion.asp"-->
<%
query = "Select TOP 5 * "
query = query&" FROM EMPRESAS"
query = query&" WHERE EMPRESAS.RUT LIKE '%"&Request("txt")&"%' "
query = query&" OR EMPRESAS.R_SOCIAL LIKE '%"&Request("txt")&"%' "

SET rs = conn.execute(query)
While not rs.EOF
	Response.Write("<li onclick=""fill('"&rs("ID_EMPRESA")&"','"&rs("RUT")&"');"">("&rs("RUT")&") "&rs("R_SOCIAL")&"</li>")
	rs.MoveNext
wend
%>