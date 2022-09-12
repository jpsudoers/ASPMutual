<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%

Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select * from usuarios u where u.RUT='"&Request("txtRut")&"' and U.ESTADO=1"

DATOS.Open sql,oConn

if (not(DATOS.eof)) then
	Response.Write("false")
else
	Response.Write("true")
end if
%>
