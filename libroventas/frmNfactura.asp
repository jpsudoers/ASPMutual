<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>

<form name="frmNfactura" id="frmNfactura" action="libroventas/Nfactura.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td width="500">Por favor Ingrese Número de Factura:</td>
    </tr>
    <tr><td>&nbsp;</td></tr>
    <tr>
    <td>Nº de Factura: &nbsp;<input id="Nmfactura" name="Nmfactura" type="text" />
       <input id="idFactNota" name="idFactNota" type="hidden" value="<%=Request("id_f")%>"/>
       <input id="idUSRNota" name="idUSRNota" type="hidden" value="<%=Request("id_u")%>"/>
    </td>
    </tr>
    <tr>
       <td>&nbsp;</td>
    </tr>
   </table>
</form> 
<div id="messageBox2" style="height:70px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 