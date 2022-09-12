
<!--#INCLUDE file="aspJSON1.18.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
'Response.ContentType ="application/vnd.ms-excel"
'Response.AddHeader "content-disposition", "inline; filename=Carga_MCAP_"&fecha&".xls"
	

Dim oHTTP : Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
oHTTP.Open "GET", "https://stage.mcap.cl/wp-json/mcap/v1/events/j8Up987874k7676YYY67ew3/20220608/20220610", False
oHTTP.SetRequestHeader "Content-Type", "application/json"
oHTTP.Send()

'Response.write "status: " & oHTTP.status
Response.write "ResponseText: " & oHTTP.ResponseText

'Set oJSON = New aspJSON

'jsonstring = oHTTP.ResponseText

'Response.Write jsonstring

Set oJSON = New aspJSON
'oJSON.loadJSON(jsonstring)

'Response.Write oJSON.data("order_date") & "<br>"	
					
'myJSON = JSON.parse(oHTTP.ResponseText)
'response.write myJSON.id & "<br />"



Set oHTTP = Nothing

%>