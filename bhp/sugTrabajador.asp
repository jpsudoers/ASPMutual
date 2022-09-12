<!--#include file="../conexion.asp"-->
<%

Response.CodePage = 65001
Response.CharSet = "utf-8"

query = "Select ID_TRABAJADOR,RUT,UPPER(NOMBRES) as NOMBRES from TRABAJADOR "
query = query&" WHERE RUT LIKE '%"&Request("txt")&"%' "
query = query&" OR NOMBRES LIKE '%"&Request("txt")&"%' "

SET rs = conn.execute(query)
While not rs.EOF
	Response.Write("<li onclick=""fill('"&rs("ID_TRABAJADOR")&"','"&rs("RUT")&"');"">("&rs("RUT")&") "&rs("NOMBRES")&"</li>")
	rs.MoveNext
wend
%>