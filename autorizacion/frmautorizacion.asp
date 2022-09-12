<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")


dim query
query= "select AUTORIZACION.ID_AUTORIZACION,EMPRESAS.R_SOCIAL,ID_PROGRAMA,AUTORIZACION.ID_EMPRESA,AUTORIZACION.ID_OTIC,ORDEN_COMPRA,VALOR_OC,"
query= query&"CODIGO_AUTORIZACION,CONVERT(VARCHAR(10),FECHA__AUTORIZACION, 105) as FECHA__AUTORIZACION,AUTORIZADOR,VALOR_AUTORIZADO,INSCRITOS"
query= query&",AUTORIZACION.ESTADO,ORDEN_COMPRA_OTIC,VALOR_OCOMPRA_OTIC,SOLICITUD from AUTORIZACION"
query= query&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA"
query= query&" where ID_AUTORIZACION="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmAutorizacion" id="frmAutorizacion" action="autorizacion/modificar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="155">C&oacute;digo Programaci&oacute;n:</td>
       <td width="150"><select id="Curriculo" name="Curriculo" tabindex="1"></select></td>
       <td width="20" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("ID_AUTORIZACION")%>"/>
       <input type="hidden" id="txtIdCurriculo" name="txtIdCurriculo" value="<%=rsEmp("ID_PROGRAMA")%>"/></td>
       <td width="120">&nbsp;</td>
       <td width="350">&nbsp;</td>
     </tr>
     <tr>
       <td>Rut Empresa :</td>
       <td><input id="txRut" name="txRut" type="text" tabindex="2" maxlength="11" size="11" onkeyup="lookup(this.value);"/>
       <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:100px;left:150px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList">
              &nbsp;
            </div>
         </div>
       </td>
       <td><input type="HIDDEN" id="Empresa" name="Empresa" value="<%=rsEmp("ID_EMPRESA")%>"/></td>
       <td>Raz&oacute;n Social : </td>
       <td><label id="txtRsoc" name="txtRsoc"></label></td>
  </tr>
        <td>Orden de Compra :</td>
       <td><input id="txtOrdenCompraE" name="txtOrdenCompraE" type="text" tabindex="3" maxlength="50" size="12" onKeyPress="return acceptNum(event)" value="<%=rsEmp("ORDEN_COMPRA")%>"/></td>
       <td></td>
       <td>Valor :</td>
       <td><input id="txtValorE" name="txtValorE" type="text" tabindex="4" maxlength="50" size="12"  onKeyPress="return acceptNum(event)" value="<%=rsEmp("VALOR_OC")%>"/></td>
     </tr>
     <tr>
       <td>OTIC :</td>
    <td><label id="txtRutOtic" name="txtRutOtic"></label></td>
        <td><input type="hidden" id="txtIdOtic" name="txtIdOtic"/></td>
       <td>Raz&oacute;n Social : </td>
       <td><label id="txtRsocOtic" name="txtRsocOtic"></label></td>
     </tr>
     <tr> 
      <tr>
       <td>Orden de Compra :</td>
       <td><input id="txtOrdenCompraO" name="txtOrdenCompraO" type="text" tabindex="6" maxlength="50" size="12"  onKeyPress="return acceptNum(event)" value="<%=rsEmp("ORDEN_COMPRA_OTIC")%>"/></td>
       <td></td>
       <td>Valor :</td>
       <td><input id="txtValorO" name="txtValorO" type="text" tabindex="7" maxlength="50" size="12"  onKeyPress="return acceptNum(event)" value="<%=rsEmp("VALOR_OCOMPRA_OTIC")%>"/></td>
     </tr>
     <tr> 
       <td>Fecha Vigencia :</td>
       <td><input id="txtFecha" name="txtFecha" type="text" tabindex="8" maxlength="50" size="12" value="<%=rsEmp("FECHA__AUTORIZACION")%>"/></td>
     </tr>
       <tr>
       <td>Asistentes :</td>
       <td><input id="txtValorAutorizador" name="txtValorAutorizador" type="text" tabindex="10" maxlength="50" size="12" onKeyPress="return acceptNum(event)" value="<%=rsEmp("VALOR_AUTORIZADO")%>"/></td>
     </tr>
      <tr>
       <td>Inscritos :</td>
      <td><label id="txtInscritos" name="txtInscritos"></label></td>
     </tr>
     <tr>
       <td>Solicitud :</td>
      <td><select id="solicitud" name="solicitud" tabindex="11"></select><input id="txtSol" name="txtSol" type="hidden" value="<%=rsEmp("SOLICITUD")%>"/></td>
     </tr>
   </table>
</form> 
<%
   else
%> <form name="frmAutorizacion" id="frmAutorizacion" action="autorizacion/insertar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="155">C&oacute;digo Programaci&oacute;n:</td>
       <td width="150"><select id="Curriculo" name="Curriculo" tabindex="1" onchange="cargaDatosInscritos(this.value);"></select></td>
       <td width="20" ></td>
       <td width="120">&nbsp;</td>
       <td width="350">&nbsp;</td>

     </tr>
     <tr>
       <td>Rut Empresa :</td>
       <td><input id="txRut" name="txRut" type="text" tabindex="2" maxlength="11" size="11" onkeyup="lookup(this.value);"/>
       <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:100px;left:150px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList">
              &nbsp;
            </div>
         </div>
       </td>
       <td><input type="hidden" id="Empresa" name="Empresa"/></td>
       <td>Raz&oacute;n Social :</td>
       <td><label id="txtRsoc" name="txtRsoc"></label></td>
     </tr>
      <tr>
       <td>Orden de Compra :</td>
       <td><input id="txtOrdenCompraE" name="txtOrdenCompraE" type="text" tabindex="3" maxlength="50" onKeyPress="return acceptNum(event)" size="12"/></td>
       <td></td>
       <td>Valor :</td>
       <td><input id="txtValorE" name="txtValorE" type="text" tabindex="4" maxlength="50"  onKeyPress="return acceptNum(event)" size="12"/></td>
     </tr>
    <tr>
       <td>OTIC :</td>
     <td><label id="txtRutOtic" name="txtRutOtic"></label></td>
     <td><input type="hidden" id="txtIdOtic" name="txtIdOtic"/></td>
     <td>Raz&oacute;n Social : </td>
     <td><label id="txtRsocOtic" name="txtRsocOtic"></label></td>
     </tr>
      <tr>
      
       <td>Orden de Compra :</td>
       <td><input id="txtOrdenCompraO" name="txtOrdenCompraO" type="text" tabindex="6" maxlength="50"  onKeyPress="return acceptNum(event)" size="12"/></td>
       <td></td>
       <td>Valor :</td>
       <td><input id="txtValorO" name="txtValorO" type="text" tabindex="7" maxlength="50"  onKeyPress="return acceptNum(event)" size="12"/></td>
     </tr>
     <tr>
       <td>Fecha Vigencia :</td>
       <td><input id="txtFecha" name="txtFecha" type="text" tabindex="8" maxlength="50" size="12"/></td>
     </tr>
       <tr>
       <td>Asistentes :</td>
       <td><input id="txtValorAutorizador" name="txtValorAutorizador" type="text" tabindex="10" maxlength="50" size="12"  onKeyPress="return acceptNum(event)"/></td>
     </tr>
      <tr>
       <td>Inscritos :</td>
        <td><label id="txtInscritos" name="txtInscritos">0</label></td>
     </tr>
      <tr>
       <td>Solicitud :</td>
      <td><select id="solicitud" name="solicitud" tabindex="11"></select></td>
     </tr>
   </table>
</form> 
<%
end if
%>
<div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 
