
<%@LANGUAGE="VBSCRIPT"%>

<%
Set conn = Server.CreateObject("ADODB.Connection") 
conn.CommandTimeout = 0
conn.Open("Provider=SQLOLEDB; User ID=uAntof;Password=mutu4ls3g*2015;data source=127.0.0.1;Initial Catalog=dbmas")

srtsql="select * from prueba "
set rs = conn.execute (srtsql)

response.write(request("eliminar"))
if request("eliminar") = "del" then

	id_prueba = request("id_prueba")
	srtsql="delete  prueba where id_prueba = "&id_prueba&" "
	
	'response.write(srtsql):RESPONSE.End()
	 conn.execute (srtsql)

end if

if request("MODIFICA") = "MOD" then
NOMBRE = REQUEST("NOMBRE")

	id_prueba = request("id_prueba")
	srtsql="UPDATE PRUEBA SET NOMBRE = '"&NOMBRE&"' where id_prueba = "&id_prueba&""
	
	'response.write(srtsql):RESPONSE.End()
	 conn.execute (srtsql)
	 RESPONSE.REDIRECT("PRUEBA16.ASP")

end if


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento sin título</title>


</head>

<body>
<table width="489" border="1">
<form>
  <tr>
    <td width="144">Nombre</td>
    <td width="144">Apewllido</td>
    <td width="144">Rut</td>
    <td width="29">&nbsp;</td>
  </tr>
  <%
  while not rs.eof
  
  %>
  <tr>
    <td><input name="NOMBRE" type="text" value="<%=rs("nombre")%>" /></td>
    <td><input name="" type="text" value="<%=rs("apellido")%>" /></td>
    <td><input name="" type="text" value="<%=rs("rut")%>" /></td>
    <td><input name="eliminar" type="submit" value=""  />
      <input type="submit" name="MODIFICA" id="button" value="MOD" /></td>
    <input name="id_prueba" type="text" value="<%=rs("id_prueba")%>" />
    
  </tr>
  
  <%
  rs.movenext
  wend

  %>

  
 </form>
</table>
</body>
</html>
