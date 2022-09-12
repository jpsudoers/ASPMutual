<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

query= "select * from CURRICULO where ID_MUTUAL="&Request("id")
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
        <td width="200">C&oacute;digo : <%=rsEmp("CODIGO")%></td>
        <td width="600">Nombre : <%=rsEmp("NOMBRE_CURSO")%></td>        
    </tr>
	<tr>
        <td colspan="2">&nbsp;</td>        
    </tr>    
</table>
<table id="list2"></table> 
<div id="pager2"></div> 
<%
   end if
%>