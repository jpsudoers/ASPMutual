<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
%>
<form name="frmBloqueaEmp" id="frmBloqueaEmp" action="finanzasempresa/actDesEmpresa.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td width="500"><%if(Request("est")="1")then%>Esta seguro que desea <b>Bloquear</b> la empresa?.<%else%>Esta seguro que desea <b>Desbloquear</b> la empresa?.<%end if%></td>
    </tr>
    <tr>
       <td>&nbsp;<input id="idBloquearIDUsuario" name="idBloquearIDUsuario" type="hidden" value="<%=Request("u")%>"/>
       <input id="idBloquearIDEmpresa" name="idBloquearIDEmpresa" type="hidden" value="<%=Request("e")%>"/>
       <input id="idBloquearIDCliente" name="idBloquearIDCliente" type="hidden" value="<%=Request("c")%>"/>
       <input id="idBloquearEstado" name="idBloquearEstado" type="hidden" value="<%=Request("est")%>"/>
		<input id="tabProgId" name="tabProgId" type="hidden" value="<%=fecha%>"/>        
       </td>
    </tr>
    <tr>
       <td>Ingrese las Razones : </td>
    </tr>
    <tr>
       <td><textarea name="razonBloquear" id="razonBloquear" rows="5" cols="70"></textarea></td>
    </tr>
    <tr>
       <td>&nbsp;</td>
    </tr>
   </table>
</form> 
<div id="messageBox2" style="height:30px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 