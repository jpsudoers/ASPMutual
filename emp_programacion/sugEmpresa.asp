<!--#include file="../conexion.asp"-->
<%

Response.CodePage = 65001
Response.CharSet = "utf-8"

query = "Select  * "
query = query&" FROM EMPRESAS"
query = query&" WHERE EMPRESAS.RUT LIKE '%"&Request("txt")&"%' AND EMPRESAS.TIPO<>3 AND EMPRESAS.PREINSCRITA=1 "
query = query&" AND EMPRESAS.ID_OTIC='"&Request("otic")&"'"
query = query&" OR EMPRESAS.R_SOCIAL LIKE '%"&Request("txt")&"%' AND EMPRESAS.TIPO<>3 AND EMPRESAS.PREINSCRITA=1 "
query = query&" AND EMPRESAS.ID_OTIC='"&Request("otic")&"'"

SET rs = conn.execute(query)
While not rs.EOF
	Response.Write("<li onclick=""fill('"&rs("ID_EMPRESA")&"','"&rs("RUT")&"');"">("&rs("RUT")&") "&rs("R_SOCIAL")&"</li>")
	rs.MoveNext
wend
%>