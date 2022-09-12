<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 

fila=1
dim cuen
For cuen = 1 To 1
		Response.Write("<row id="""&fila&""">"&chr(13))
		Response.Write("<cell><![CDATA[]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[]]></cell>"&chr(13))
		Response.Write("<cell><![CDATA[]]></cell>"&chr(13))
		Response.Write("</row>"&chr(13))
		fila=fila+1
next
Response.Write("</rows>") 
%>
