<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

if(Request("val")="0") then
query="SELECT E.ID_EMPRESA, E.NOMBRE_CONTA NOMBRES, E.EMAIL_CONTA EMAIL, E.FONO_CONTABILIDAD FONO, E.CARGO_CONTA CARGO,  "
query=query&"'Contacto Contabilidad' COMENTARIO, 0 VAL, ID_EMP_USU=0 FROM EMPRESAS E WHERE E.ID_EMPRESA="&Request("ide")
else
query="SELECT U.ID_EMP_USU, U.NOMBRES, U.EMAIL, U.FONO, U.CARGO, I.COMENTARIO FROM dbo.EMPRESA_USUARIOS U "
query=query&"LEFT JOIN dbo.EMP_INS_USR I ON (I.ID_EMP_USU=U.ID_EMP_USU) WHERE U.ID_EMP_USU="&Request("idc")
end if

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmTrabajador" id="frmTrabajador" action="finanzasempresa/modificar_trab.asp" method="post">
 <table cellspacing="0" cellpadding="1" border=0>
    <tr>
        <td width="115">Nombre</td>
        <td width="15">:</td>
        <td width="250">
    <input id="txtNombre" name="txtNombre" type="text" tabindex="21" maxlength="100" size="40" value="<%=rsEmp("NOMBRES")%>"/>
       	</td>
    </tr>
    <tr>
        <td>Email</td>
        <td>:</td>
        <td>
    <input id="txtEmail" name="txtEmail" type="text" tabindex="23" maxlength="100" size="40" value="<%=rsEmp("EMAIL")%>"/>
        </td>
    </tr>
    <tr>
        <td>Fono</td>
        <td>:</td>
        <td>
    <input id="txtFono" name="txtFono" type="text" tabindex="24" maxlength="100" size="40" value="<%=rsEmp("FONO")%>"/>
        </td>
    </tr>
    <tr>
        <td>Cargo</td>
        <td>:</td>
        <td>
    <input id="txtCargo" name="txtCargo" type="text" tabindex="25" maxlength="100" size="40" value="<%=rsEmp("CARGO")%>"/>
        </td>
    </tr>
    <tr id="trComentario">
        <td valign="top">Comentarios</td>
        <td valign="top">:</td>
        <td>
        <textarea id="txtComent" name="txtComent" tabindex="26" cols="38" type="text" rows="6"><%=rsEmp("COMENTARIO")%></textarea>
        </td>
    </tr>        
    <input id="idc" name="idc" type="hidden" value="<%=rsEmp("ID_EMP_USU")%>"/>
    <input id="ide" name="ide" type="hidden" value="<%=Request("ide")%>"/>
    <input id="val1" name="val1" type="hidden" value="<%=Request("val")%>"/>
    <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
   </table>
</form> 
<%
   else
%><form name="frmTrabajador" id="frmTrabajador" action="finanzasempresa/insertar_trab.asp" method="post">
 <table cellspacing="0" cellpadding="1" border=0>
     <tr>
        <td width="115">Nombre</td>
        <td width="15">:</td>
        <td width="250">
    <input id="txtNombre" name="txtNombre" type="text" tabindex="21" maxlength="100" size="40" value="<%=rsEmp("NOMBRES")%>"/>
       	</td>
    </tr>
    <tr>
        <td>Email</td>
        <td>:</td>
        <td>
    <input id="txtEmail" name="txtEmail" type="text" tabindex="23" maxlength="100" size="40" value="<%=rsEmp("EMAIL")%>"/>
        </td>
    </tr>
    <tr>
        <td>Fono</td>
        <td>:</td>
        <td>
    <input id="txtFono" name="txtFono" type="text" tabindex="24" maxlength="100" size="40" value="<%=rsEmp("FONO")%>"/>
        </td>
    </tr>
    <tr>
        <td>Cargo</td>
        <td>:</td>
        <td>
    <input id="txtCargo" name="txtCargo" type="text" tabindex="25" maxlength="100" size="40" value="<%=rsEmp("CARGO")%>"/>
        </td>
    </tr> 
    <tr>
        <td valign="top">Comentarios</td>
        <td valign="top">:</td>
        <td>
    <textarea id="txtComent" name="txtComent" tabindex="26" cols="38" type="text" rows="6"><%=rsEmp("COMENTARIO")%></textarea>
        </td>
    </tr>  
    <input id="ide" name="ide" type="hidden" value="<%=Request("ide")%>"/>            
    <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
   </table>
</form> 
<%
   end if
%>
<div id="messageBox2" style="height:90px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 