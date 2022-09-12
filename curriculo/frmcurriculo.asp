<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select * from CURRICULO where ID_MUTUAL="&vid
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmCurriculo" id="frmCurriculo" action="curriculo/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
        <td>C&oacute;digo</td>
        <td width="20">:</td>
        <td><input id="txtCod" name="txtCod" type="text" tabindex="1" size="33" maxlength="33" value="<%=rsEmp("CODIGO")%>"/></td>
    </tr>
    <tr>
        <td>Sence</td>
        <td width="20">:</td>
         <%
		   if(rsEmp("SENCE")=0)then
		   senceSi="checked"
		   else
		   senceNo="checked"
		   end if
		 %>
        <td><input type="radio" name="Sence" id="Sence" value="0" <%=senceSi%>>Si
		   <input type="radio" name="Sence" id="Sence" value="1" <%=senceNo%>>No</td>
            <td><input type="hidden" id="txtSence" name="txtSence" value="<%=rsEmp("SENCE")%>" /></td>
    </tr>
    <tr>
    	<td width="130">Nombre</td>
        <td width="10">:</td>
      	<td width="250"><input id="txtNom" name="txtNom" type="text" tabindex="2" onKeyPress="" size="102" 
        maxlength="103" value="<%=rsEmp("NOMBRE_CURSO")%>"/></td>
        <td width="20" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("ID_MUTUAL")%>"/></td>
   	</tr>
  <tr>
        <td valign="top">Descripci&oacute;n</td>
        <td valign="top" width="10">:</td>
        <td><textarea type="text" rows="3" cols="100" id="txtDesc" name="txtDesc" tabindex="3"><%=rsEmp("DESCRIPCION")%></textarea></td>
    </tr>
     <tr>
    	<td valign="top">Objetivos</td>
        <td valign="top">:</td>
        <td><textarea type="text" rows="3" cols="100" id="txtObj" name="txtObj" tabindex="4"><%=rsEmp("OBJETIVOS")%></textarea></td>
        
    </tr>
    <tr>
        <td>Audiencia</td>
        <td>:</td>
        <td><input id="txtAud" name="txtAud" type="text" tabindex="5" maxlength="20" value="<%=rsEmp("AUDIENCIA")%>" size="20"/></td>
    </tr>
     <tr>
        <td>Horas</td>
        <td>:</td>
        <td><input id="txtHor" name="txtHor" type="text" tabindex="6" size="20" maxlength="40" value="<%=rsEmp("HORAS")%>"/></td>
    </tr>
    <tr>
        <td>Valor</td>
        <td>:</td>
        <td><input id="txtValor" name="txtValor" type="text" tabindex="7" size="20" maxlength="15" value="<%=rsEmp("VALOR")%>"/></td>
    </tr>
    <tr>
        <td>Valor Afiliados</td>
        <td>:</td>
        <td><input id="txtValAfiliados" name="txtValAfiliados" type="text" tabindex="8" size="20" maxlength="15" value="<%=rsEmp("VALOR_AFILIADOS")%>"/></td>
    </tr>    
     <tr>
        <td>Vigencia</td>
        <td>:</td>
        <td><select id="txtVig" name="txtVig" tabindex="9"></select><input id="txtVigId" name="txtVigId" type="hidden" value="<%=rsEmp("VIGENCIA")%>" /></td>
    </tr>
    <tr>
        <td>Cód. Softland</td>
        <td>:</td>
        <td><input id="txtSoftland" name="txtSoftland" type="text" tabindex="10" size="20" maxlength="15" value="<%=rsEmp("COD_SOFTLAND")%>"/></td>
        </td>
    </tr>
    <tr>
        <td>Centro Costo</td>
        <td>:</td>
        <td><input id="txtCeco" name="txtCeco" type="text" tabindex="11" size="20" maxlength="15" value="<%=rsEmp("ID_CECO")%>"/></td>
        </td>
    </tr>
	  <tr>
        <td>Porcentaje de Asistencia</td>
        <td>:</td>
        <td><input id="txtAsistencia" name="txtAsistencia" type="number" tabindex="12" size="20" maxlength="3" min="1" max ="100" value="<%=rsEmp("PORCE_ASISTENCIA")%>"/></td>
        </td>
    </tr>
	  <tr>
        <td>Porcentaje de Calificacion</td>
        <td>:</td>
        <td><input id="txtCalificacion" name="txtCalificacion" type="number" tabindex="13" size="20" maxlength="3" min="1" max ="100" value="<%=rsEmp("PORCE_CALIFICACION")%>"/></td>
        </td>
    </tr>
</table>
</form> 
<%
   else
%> <form name="frmCurriculo" id="frmCurriculo" action="curriculo/insertar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
        <td>C&oacute;digo</td>
        <td>:</td>
        <td><input id="txtCod" name="txtCod" type="text" tabindex="1" size="33" maxlength="33"/></td>
    </tr>
      <tr>
        <td>Sence</td>
        <td>:</td>
  			<td><input type="radio" name="Sence" id="Sence" value="0" checked>Si
		   	<input type="radio" name="Sence" id="Sence" value="1">No</td>
    </tr>
    <tr>
    	<td width="110">Nombre</td>
        <td>:</td>
      	<td width="600"><input id="txtNom" name="txtNom" type="text" tabindex="2" size="103" maxlength="103"/></td>
        <td width="20" >&nbsp;</td>
   	</tr>
  <tr>
        <td valign="top">Descripci&oacute;n</td>
        <td valign="top">:</td>
        <td><textarea type="text" rows="3" cols="100" id="txtDesc"  tabindex="3" name="txtDesc"/></textarea></td>
    </tr>
     <tr>
    	<td valign="top">Objetivos</td>
        <td valign="top">:</td>
        <td><textarea type="text" rows="3" cols="100" id="txtObj"  tabindex="4" name="txtObj"/></textarea></td>
    </tr>
      <tr>
        <td>Audiencia</td>
        <td>:</td>
        <td><input id="txtAud" name="txtAud" type="text" tabindex="5" maxlength="20" size="20"/></td>
    </tr>
     <tr>
        <td>Horas</td>
        <td>:</td>
        <td><input id="txtHor" name="txtHor" type="text" tabindex="6" size="20" maxlength="3"/></td>
    </tr>
    <tr>
        <td>Valor</td>
        <td>:</td>
        <td><input id="txtValor" name="txtValor" type="text" tabindex="7" size="20" maxlength="15"/></td>
    </tr>
    <tr>
        <td>Valor Afiliados</td>
        <td>:</td>
        <td><input id="txtValAfiliados" name="txtValAfiliados" type="text" tabindex="8" size="20" maxlength="15"/></td>
    </tr>    
     <tr>
        <td>Vigencia</td>
        <td>:</td>
        <td><select id="txtVig" name="txtVig" tabindex="9"></select></td>
    </tr>
    <tr>
        <td>Cód. Softland</td>
        <td>:</td>
        <td><input id="txtSoftland" name="txtSoftland" type="text" tabindex="10" size="20" maxlength="30"/></td>
    </tr>
    <tr>
        <td>Centro Costo</td>
        <td>:</td>
        <td><input id="txtCeco" name="txtCeco" type="text" tabindex="11" size="20" maxlength="30"/></td>
     </tr>
	 <tr>
        <td>Porcentaje de Asistencia</td>
        <td>:</td>
        <td><input id="txtAsistencia" name="txtAsistencia" type="number" tabindex="12" size="20" maxlength="3" min="1" max ="100" value="<%=rsEmp("PORCE_ASISTENCIA")%>"/></td>
        </td>
    </tr>
	  <tr>
        <td>Porcentaje de Calificacion</td>
        <td>:</td>
        <td><input id="txtCalificacion" name="txtCalificacion" type="number" tabindex="13" size="20" maxlength="3" min="1" max ="100" value="<%=rsEmp("PORCE_CALIFICACION")%>"/></td>
        </td>
    </tr>
</table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:100px;overflow:auto;width:600px;"> 
  	<ul></ul> 
</div> 
