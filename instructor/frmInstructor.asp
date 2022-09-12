<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select * from INSTRUCTOR_RELATOR where ID_INSTRUCTOR="&vid
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmInstructor" id="frmInstructor" action="instructor/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="110">Rut :</td>
      	<td width="200"><input id="txRut" name="txRut" type="text" tabindex="1" onKeyPress="" size="12" maxlength="11" value="<%=rsEmp("rut")%>"/></td>
        <td width="20" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("ID_INSTRUCTOR")%>" /></td>
        <td width="150">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="160">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
     <tr>
        <td>Nombres :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="2" maxlength="50" value="<%=rsEmp("nombres")%>"  size="20"/></td>
        <td></td>
        <td>Apellido Paterno :</td>
        <td><input id="txtAp_pater" name="txtAp_pater" type="text" tabindex="3" maxlength="50" value="<%=rsEmp("a_paterno")%>"  size="20"/></td>
        <td></td>
        <td>Apellido Materno :</td>
        <td><input id="txtAp_mater" name="txtAp_mater" type="text" tabindex="4" maxlength="50" value="<%=rsEmp("a_materno")%>"  size="20"/></td>
    </tr>
</table>
</form> 
<%
   else
%> <form name="frmInstructor" id="frmInstructor" action="instructor/insertar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="102">Rut :</td>
      	<td width="166"><input id="txRut" name="txRut" type="text" tabindex="1" onKeyPress="" size="12" maxlength="11" /></td>
        <td width="20" ></td>
        <td width="126">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="128">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
     <tr>
        <td>Nombres :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="2" maxlength="50" size="20"/></td>
        <td></td>
        <td>Apellido Paterno :</td>
        <td><input id="txtAp_pater" name="txtAp_pater" type="text" tabindex="3" maxlength="50" size="20"/></td>
        <td></td>
        <td>Apellido Materno :</td>
        <td><input id="txtAp_mater" name="txtAp_mater" type="text" tabindex="4" maxlength="50" size="20"/></td>
    </tr>
</table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 
