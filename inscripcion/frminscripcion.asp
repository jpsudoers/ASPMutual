<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select * from AUTORIZACION where ID_AUTORIZACION="&vid
set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frminscripcion" id="frminscripcion" action="inscripcion/modificar.asp" method="post">
    <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="155">C&oacute;digo Autorizaci&oacute;n:</td>
       <td width="200"><select id="autorizacion" name="autorizacion" tabindex="1" onchange="cargaDatos(this.value);">
       </select></td>
       <td width="20"><input type="hidden" id="txtIdAutorizacion" name="txtIdAutorizacion" value="" /></td>
        <td width="150">&nbsp;</td>
       <td width="172">&nbsp;</td>
        <td width="20" >&nbsp;</td>
       <td width="150">&nbsp;</td>
     </tr>
     <tr>
       <td>Rut Empresa :</td>
       <td><label id="txtRutEmpresa" name="txtRutEmpresa"></label></td>
       <td><input type="hidden" id="txtIdEmpresa" name="txtIdEmpresa" value="" /></td>
       <td>Raz&oacute;n Social  :</td>
         <td colspan="3"><label id="txtRSocial" name="txtRSocial"></label></td>
     </tr>
     <tr>
       <td>Nombre Curso :</td>
         <td><label id="nomCurso" name="nomCurso"></label></td>
       <td><input type="hidden" id="txtIdPrograma" name="txtIdPrograma" value="" /></td>
       <td>Descripci&oacute;n :</td>
       <td colspan="3"><label id="desCurso" name="desCurso"></label></td>
     </tr>
       <tr>
       <td>Rut Instructor :</td>
         <td><label id="txtRutInstructor" name="txtRutInstructor"></label></td>
       <td></td>
       <td>Nombre Instructor :</td>
       <td colspan="3"><label id="nomInstructor" name="nomInstructor"></label></td>
     </tr>
      <tr>
       <td>Sede :</td>
         <td><label id="txtSede" name="txtSede"></label></td>
          <td></td>
          <td>Direcci&oacute;n :</td>
          <td colspan="3"><label id="txtDir" name="txtDir"></label></td>
     </tr>
               <tr>
       <td>&nbsp;</td>
        <td>&nbsp;</td>
         <td>&nbsp;</td>
          <td>&nbsp;</td>
     </tr>
       <tr>
       <td>&nbsp;</td>
        <td>&nbsp;</td>
         <td>&nbsp;</td>
          <td><input id="datostabla" name="datostabla" type="hidden"/></td>
     </tr>
   </table>
</form> 
<%
   else
%> <form name="frminscripcion" id="frminscripcion" action="inscripcion/insertar.asp" method="post">
    <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="155">C&oacute;digo Autorizaci&oacute;n:</td>
       <td width="200"><select id="autorizacion" name="autorizacion" tabindex="1" onchange="cargaDatos(this.value);grilla(0);">
       </select></td>
       <td width="20" >&nbsp;</td>
       <td width="150">&nbsp;</td>
       <td width="172">&nbsp;</td>
        <td width="20" >&nbsp;</td>
       <td width="150">&nbsp;</td>
     </tr>
     <tr>
       <td>Rut Empresa :</td>
       <td><label id="txtRutEmpresa" name="txtRutEmpresa"></label></td>
       <td><input type="hidden" id="txtIdEmpresa" name="txtIdEmpresa" value="" /></td>
       <td>Raz&oacute;n Social :</td>
         <td colspan="3"><label id="txtRSocial" name="txtRSocial"></label></td>
     </tr>
     <tr>
       <td>Nombre Curso :</td>
         <td><label id="nomCurso" name="nomCurso"></label></td>
       <td><input type="hidden" id="txtIdPrograma" name="txtIdPrograma" value="" /></td>
       <td>Descripci&oacute;n :</td>
       <td colspan="3"><label id="desCurso" name="desCurso"></label></td>
     </tr>
       <tr>
       <td>Rut Instructor :</td>
         <td><label id="txtRutInstructor" name="txtRutInstructor"></label></td>
       <td></td>
       <td>Nombre Instructor :</td>
       <td colspan="3"><label id="nomInstructor" name="nomInstructor"></label></td>
     </tr>
      <tr>
       <td>Sede :</td>
         <td><label id="txtSede" name="txtSede"></label></td>
          <td></td>
          <td>Direcci&oacute;n :</td>
          <td colspan="3"><label id="txtDir" name="txtDir"></label></td>
     </tr>
          <tr>
       <td>&nbsp;</td>
        <td>&nbsp;</td>
         <td>&nbsp;</td>
          <td>&nbsp;</td>
     </tr>
               <tr>
       <td>&nbsp;</td>
        <td>&nbsp;</td>
         <td>&nbsp;</td>
          <td><input id="datostabla" name="datostabla" type="hidden"/></td>
     </tr>
   </table>
</form> 
<%
end if
%>
<div width="200">
<table id="list2"></table> 
<div id="pager2"></div> 
</div>
<div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 

