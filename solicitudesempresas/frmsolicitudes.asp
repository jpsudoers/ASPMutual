<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query="SELECT SOLICITUD.id_solicitud,EMPRESAS.RUT, EMPRESAS.R_SOCIAL, MUTUALES.nomb_mutual, OTIC.R_SOCIAL as otic, "
query= query&"CURRICULO.codigo, CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),SOLICITUD.fecha, 105) as fecha, SOLICITUD.participantes "
query= query&" FROM SOLICITUD "
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=SOLICITUD.id_empresa "
query= query&" INNER JOIN MUTUALES ON MUTUALES.Mutual_id =EMPRESAS.MUTUAL "
query= query&" INNER JOIN OTIC ON OTIC.ID_OTIC=EMPRESAS.ID_OTIC "
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=SOLICITUD.id_mutual "
query= query&" where SOLICITUD.id_solicitud ="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmSolicitudes" id="frmSolicitudes" action="" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="155">Rut Empresa :</td>
       <td width="200"><%=rsEmp("RUT")%></td>
       <td width="20">&nbsp;</td>
       <td width="120">Raz&oacute;n Social :</td>
       <td width="300" colspan="4"><%=rsEmp("R_SOCIAL")%></td>
     </tr>
     <tr>
       <td>Nombre Curso :</td>
       <td colspan="4"><%=rsEmp("NOMBRE_CURSO")%></td>
     </tr>
     <tr>
       <td>Fecha Requerida :</td>
       <td><%=rsEmp("fecha")%></td>
       <td></td>
       <td>N&deg; Participantes :</td>
       <td><%=rsEmp("participantes")%></td>
     </tr>
   </table>
</form> 
<%
else
%>
 <form name="frmSolicitudes" id="frmSolicitudes" action="solicitudesempresas/insertar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="155">Rut Empresa :</td>
       <td width="200"><input id="txRut" name="txRut" type="text" tabindex="1" maxlength="11" size="11" onkeyup="lookup(this.value);"/>
       <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:100px;left:150px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList">
              &nbsp;
            </div>
         </div></td>
       <td width="20"><input type="hidden" id="Empresa" name="Empresa"/></td>
       <td width="120">Raz&oacute;n Social :</td>
       <td width="300" colspan="4"><label id="txtRsoc" name="txtRsoc"></label></td>
     </tr>
     <tr>
       <td>Nombre Curso :</td>
       <td colspan="4"><select id="txtCurso" name="txtCurso" tabindex="2"></select></td>
     </tr>
     <tr>
       <td>Fecha Requerida :</td>
       <td><input id="txtFecha" name="txtFecha" type="text" tabindex="3" maxlength="50" size="12"/></td>
       <td></td>
       <td>N&deg; Participantes :</td>
       <td><input id="txtPart" name="txtPart" type="text" tabindex="4" maxlength="3" size="12"/></td>
     </tr>
   </table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 

