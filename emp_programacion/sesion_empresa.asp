<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"

Response.Write("<DATOS>"&chr(13)) 
Response.Write("<row>"&chr(13))
Response.Write("<sesEmp>"&Session("usuario")&"</sesEmp>"&chr(13))
Response.Write("</row>"&chr(13))
Response.Write("</DATOS>") 
%>
