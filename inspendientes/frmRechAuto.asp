<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<form name="frmRechAuto" id="frmRechAuto" action="inspendientes/correoRechazo.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td width="500">Esta seguro de rechazar la solicitud de inscripci&oacute;n a Curso.</td>
    </tr>
    <tr>
       <td>&nbsp;<input id="idRechPreins" name="idRechPreins" type="hidden" value="<%=Request("id")%>"/></td>
    </tr>
    <tr>
       <td>Razones : </td>
    </tr>
    <tr>
       <td><textarea name="razonRech" id="razonRech" rows="5" cols="70"></textarea></td>
    </tr>
    <tr>
       <td>&nbsp;</td>
    </tr>
   </table>
</form> 
<div id="messageBox2" style="height:30px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 