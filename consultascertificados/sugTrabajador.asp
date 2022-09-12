<!--#include file="../conexion.asp"-->
<%

Response.CodePage = 65001
Response.CharSet = "utf-8"

query = "select DISTINCT T.ID_TRABAJADOR,T.RUT,UPPER(T.NOMBRES) as NOMBRES from HISTORICO_CURSOS hc "
query = query&" inner join TRABAJADOR T on T.ID_TRABAJADOR=hc.ID_TRABAJADOR "
query = query&" WHERE hc.ESTADO=0 AND hc.EVALUACION<>'Reprobado' "
query = query&" AND (RUT LIKE '%"&Request("txt")&"%' "
query = query&" OR NOMBRES LIKE '%"&Request("txt")&"%') "

SET rs = conn.execute(query)
While not rs.EOF
	Response.Write("<li onclick=""fill('"&rs("ID_TRABAJADOR")&"','"&rs("RUT")&"');"">("&rs("RUT")&") "&rs("NOMBRES")&"</li>")
	rs.MoveNext
wend
%>