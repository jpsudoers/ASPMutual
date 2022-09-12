<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vautorizacion = Request("id")
vfactura = Request("fac")
%>
 <form name="frmntickets" id="frmntickets" action="#" method="post">
    <table cellspacing="1" cellpadding="1" border=0>
     <tr>
       <td>Ingrese N&deg; de Tickets : </td>
       <td><input type="hidden" id="idAutoTickets" name="idAutoTickets" value="<%=vautorizacion%>"/><input type="hidden" id="nFac" name="nFac" value="<%=vfactura%>"/><input type="text" id="txtNTickets" name="txtNTickets" value=""/></td>
    </tr>
   </table>
   <br />
   <div id="messageBox1" style="height:60px;overflow:auto;width:250px;"> 
  	<ul></ul> 
   </div>    
</form> 
