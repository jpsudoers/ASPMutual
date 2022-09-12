<%
Dim TicketURL
TicketURL = "http://soporte.masesorias.cl/api/tickets.xml" 
APIKey = "8F074F0F4E5A91FF445324E222334052"
'XML data 
XML_SendData = ("<?xml version=""1.0"" encoding=""UTF-8""?>")&vbCrLf 
XML_SendData = XML_SendData & ("<ticket alert=""true"" autorespond=""true"" source=""API"">")&vbCrLf 
XML_SendData = XML_SendData & (" <name>API</name>")&vbCrLf 
XML_SendData = XML_SendData & (" <email>correo@domain.com</email>")&vbCrLf 
XML_SendData = XML_SendData & (" <topicId>32</topicId>")&vbCrLf 
XML_SendData = XML_SendData & (" <subject>Testing API7</subject>")&vbCrLf 
XML_SendData = XML_SendData & (" <phone>318-555-8634</phone>")&vbCrLf 
XML_SendData = XML_SendData & (" <message>Message content here</message>")&vbCrLf 
XML_SendData = XML_SendData & (" <attachments></attachments>")&vbCrLf 
XML_SendData = XML_SendData & (" <ip>186.105.122.53</ip>")&vbCrLf 
XML_SendData = XML_SendData & ("</ticket>")


'Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
'objXMLHTTP.Open "POST", TicketURL, false
'objXMLHTTP.setRequestHeader "X-API-Key", APIKey
'objXMLHTTP.setRequestHeader "Content-Type", "text/xml"
'objXMLHTTP.Send msgXML
'result=objXMLHTTP.responseText
'Set objXMLHTTP = Nothing


Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP") 'Server.CreateObject("Microsoft.XMLHTTP")
xmlHTTP.open "POST", TicketURL, false
'xmlhttp.setRequestHeader "Content-type", "text/HTML; charset=utf-8"
'xmlhttp.setRequestHeader  "Content-Type","application/x-www-form-urlencoded"
'xmlHttp.setRequestHeader "X-API-Key", APIKey
xmlHttp.setRequestHeader "X-API-Key", APIKey
xmlHttp.setRequestHeader "Content-Type", "text/xml"
xmlHttp.Send XML_SendData

strStatus = xmlHttp.Status
strRetval = xmlHttp.responseText
response.write strStatus&"<br>"
response.write strRetval&"<br>"
response.write XML_SendData

Set xmlHttp = nothing
'Set xmlDoc = nothing
%>