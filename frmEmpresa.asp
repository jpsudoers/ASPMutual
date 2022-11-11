<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query= "select * from EMPRESAS where id_empresa="&vid
set rsEmp = conn.execute (query)

dim query2
query2= "select * from EMPRESA_CONDICION_COMERCIAL where id_empresa="&vid
set rsCond = conn.execute (query2)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmEmpresa" id="frmEmpresa" action="empresa/modificar.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
    	<td width="100">Rut :</td>
      	<td width="166"><%=rsEmp("rut")%><input id="txRut" name="txRut" type="hidden" value="<%=rsEmp("rut")%>"/></td>
        <td width="20" ><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("id_empresa")%>" /></td>
        <td width="100">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="100">&nbsp;</td>
        <td width="131" >&nbsp;</td>
   	</tr>
	<tr>
        <td>Raz&oacute;n Social :</td>
        <td><input id="txtRsoc" name="txtRsoc" type="text" tabindex="2" maxlength="99" value="<%=rsEmp("r_social")%>" size="30"/></td>
        <td></td>
    	<td>Giro:</td>
        <td><input id="txtGiro" name="txtGiro" type="text" tabindex="3" size="30" maxlength="99" value="<%=rsEmp("giro")%>"/></td>
    </tr>
	<tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" type="text" tabindex="4" maxlength="80" value="<%=rsEmp("direccion")%>" size="30"/></td>
        <td></td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="5" size="30" maxlength="40" value="<%=rsEmp("comuna")%>" /></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="6"  size="30"  maxlength="30" value="<%=rsEmp("ciudad")%>" /></td>
    </tr>
	<tr>
        <td>Tel&eacute;fono :</td>
        <td><input id="txtFon" name="txtFon" type="text" tabindex="7" size="12" maxlength="20"  value="<%=rsEmp("fono")%>"/></td>
        <td></td>
        <td>Fax :</td>
        <td><input id="txtFax" name="txtFax" type="text" tabindex="8" maxlength="20" value="<%=rsEmp("fax")%>"  size="12"/></td>
    </tr>
	<tr>
    	<td>Mutual :</td>
        <td><select id="txtMut" name="txtMut" tabindex="9"></select><input id="txtIdMut" name="txtIdMut" type="hidden" value="<%=rsEmp("mutual")%>"/></td>
        <td></td>
        <td>Otic :</td>
        <td colspan="4"><select id="OTIC" name="OTIC" tabindex="10" style="width:45em;"></select><input id="txtIdOtic" name="txtIdOtic" type="hidden" value="<%=rsEmp("ID_OTIC")%>"/></td>

    </tr>
    <tr>
    	<td colspan="8">&nbsp;</td>
    </tr>
    <table id= "list2"> </table>
	<div id="pager2"></div>
     
	 
            <%
            vEDP=""
            vOC=""
            vTransferencia=""
            vDepositoBanco=""
			vConHes=""
			vConMigo=""

            Do Until rsCond.EOF


            if(rsCond("ID_CONDICION_COMERCIAL")="5")then  
                vEDP="checked='checked'" 
            end if
            if(rsCond("ID_CONDICION_COMERCIAL")="0")then  
                vOC="checked='checked'"

				if(rsCond("CON_HES")="1")then  
					vConHes="checked='checked'" 
				end if

				if(rsCond("CON_MIGO")="1")then  
					vConMigo="checked='checked'" 
				end if

				
            end if
            if(rsCond("ID_CONDICION_COMERCIAL")="3")then  
                vTransferencia="checked='checked'" 
            end if
            if(rsCond("ID_CONDICION_COMERCIAL")="2")then  
                vDepositoBanco="checked='checked'" 
            end if

            rsCond.MoveNext

            Loop
            rsCond.Close                
            Set rsCond=Nothing
            %>
		
		<tr>
          <td></td>
          <td  style="vertical-align: top!important;">
            <label>EDP</label><input type="checkbox" name="cond_edp" id="cond_edp"  value="1"   <%=vEDP%>  tabindex="23"/></td>    
        
            <td colspan="2" style="vertical-align: top!important;"><label>OC</label><input type="checkbox" name="cond_oc" id="cond_oc" value="1"  <%=vOC%>   tabindex="24" onchange="MuestraRef();"  /></td>     
         
            <td colspan="2"  style="vertical-align: top!important;"><label>Transferencia</label><input type="checkbox" name="cond_transferencia" id="cond_transferencia"  value="1"  <%=vTransferencia%>  tabindex="25"/></td>     
        
            <td colspan="2" style="vertical-align: top!important;" ><label>Deposito Banco</label><input type="checkbox" name="cond_depositoBanco"  id="cond_depositoBanco"  value="1"  <%=vDepositoBanco%>  tabindex="26"/>
            
            </td>    
          </tr> 

		 
		 <%
		  if(vOC = "")then 
		 %>
		 <div id="filaRef" style="display:none">
			<tr >
			<td id="lbl1"></td>
          <td  style="vertical-align: top!important;" id="lbl2"></td>    
          <td colspan="2" style="vertical-align: top!important; display:none" id="lblHES"><label>HES</label><input type="checkbox" name="cond_ref_hes" id="cond_ref_hes" value="1"  <%=vConHes%>   tabindex="24"/></td>     
          <td colspan="2"  style="vertical-align: top!important; display:none" id="lblMIGO"><label>MIGO</label><input type="checkbox" name="cond_ref_migo" id="cond_ref_migo"  value="1" <%=vConMigo%>  tabindex="25"/></td>     
          <td colspan="2" style="vertical-align: top!important;" id="lbl5"></td>    
         </tr>
		 </div>
		 <%
		else
		 %>
		 ' <div id="filaRef">
		 ' 
		<tr >
		  <td id="lbl1"></td>
          <td  style="vertical-align: top!important;" id="lbl2"></td>    
          <td colspan="2" style="vertical-align: top!important;" id="lblHES"><label>HES</label><input type="checkbox" name="cond_ref_hes" id="cond_ref_hes" value="1"  <%=vConHes%>   tabindex="24"/></td>     
          <td colspan="2"  style="vertical-align: top!important;" id="lblMIGO"><label>MIGO</label><input type="checkbox" name="cond_ref_migo" id="cond_ref_migo"  value="1" <%=vConMigo%>  tabindex="25"/></td>     
          <td colspan="2" style="vertical-align: top!important;" id="lbl5"></td>    
         </tr>
		 </div>
		 <%
		 end if
		 %>
          
		 
		 
     </tr>
</table>
</form> 
<%
   else
%><form name="frmEmpresa" id="frmEmpresa" action="empresa/insertar.asp" method="post">
	<table cellspacing="0" cellpadding="1" border=0 >
	<tr>
    	<td width="100">Rut :</td>
   	  <td width="166"><input id="txRut" name="txRut" type="text" tabindex="1" size="12" maxlength="11" value="" onchange= "validarCamposNuevaEmpresa()"/></td>
        <td width="20">&nbsp;</td>
        <td width="100">&nbsp;</td>
        <td width="172">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="100">&nbsp;</td>
        <td width="131">&nbsp;</td> 
   	</tr>
	<tr>
        <td>Raz&oacute;n Social :</td>
        <td><input id="txtRsoc" name="txtRsoc" type="text" tabindex="2" maxlength="99" value="" size="30" onchange= "validarCamposNuevaEmpresa()"/></td>
        <td> </td>
    	<td>Giro:</td>
        <td><input id="txtGiro" name="txtGir" type="text" tabindex="3" size="30" maxlength="99"  value="" onchange= "validarCamposNuevaEmpresa()"/></td>
    </tr>
	<tr>
        <td>Direcci&oacute;n :</td>
        <td><input id="txtDir" name="txtDir" type="text" tabindex="4" maxlength="50" value="" size="30" onchange= "validarCamposNuevaEmpresa()"/></td>
        <td></td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="5" size="30" maxlength="40" value="" onchange= "validarCamposNuevaEmpresa()"/></td>
        <td></td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="6"  size="30"  maxlength="30" value="" onchange= "validarCamposNuevaEmpresa()"/></td>
    </tr>
	<tr>
        <td>Tel&eacute;fono :</td>
        <td><input id="txtFon" name="txtFon" type="text" tabindex="7" size="12" maxlength="20" onchange= "validarCamposNuevaEmpresa()"/></td>
        <td></td>
        <td>Fax :</td>
        <td><input id="txtFax" name="txtFax" type="text" tabindex="8" maxlength="20" value=""  size="12" onchange= "validarCamposNuevaEmpresa()"/></td>
    </tr>
	<tr>
    	<td>Mutual :</td>
        <td><select id="txtMut" name="txtMut" tabindex="9"></select></td>
        <td></td>
        <td>Otic :</td>
        <td colspan="4"><select id="OTIC" name="OTIC" tabindex="10" style="width:45em;"></select></td>
    </tr>
    <tr>
    	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
     <table id="list3"></table>
	 <div id="pager3"></div>
    <tr>
          <td></td>
          <!-- <td  style="vertical-align: top!important;"> -->
            <label>EDP</label><input type="checkbox" name="cond_edp" id="cond_edp"  value="1" tabindex="23"/></td>    
            <td colspan="2" style="vertical-align: top!important;"><label>OC</label><input type="checkbox" name="cond_oc" id="cond_oc" value="1"  tabindex="24" onchange="MuestraRef();" /></td>     
            <td colspan="2"  style="vertical-align: top!important;"><label>Transferencia</label><input type="checkbox" name="cond_transferencia" id="cond_transferencia"  value="1"  tabindex="25"/></td>     
            <td colspan="2" style="vertical-align: top!important;" ><label>Deposito Banco</label><input type="checkbox" name="cond_depositoBanco"  id="cond_depositoBanco"  value="1" tabindex="26"/></td>    
    </tr>
	<div id="filaRef" style="display:none">
	<tr >
          <td id="lbl1"></td>
          <td  style="vertical-align: top!important;" id="lbl2"></td>    
          <td colspan="2" style="vertical-align: top!important;" id="lblHES" type="hidden"><label>HES</label><input type="checkbox" name="cond_ref_hes" id="cond_ref_hes" value="1"    tabindex="24"/></td>     
          <td colspan="2"  style="vertical-align: top!important;" id="lblMIGO" type="hidden"><label>MIGO</label><input type="checkbox" name="cond_ref_migo" id="cond_ref_migo"  value="1" tabindex="25"/></td>     
          <td colspan="2" style="vertical-align: top!important;" id="lbl5"></td>    
         </tr>
		 </div>
</table>
</form>
<%
   end if
%>
<div id="messageBox1" style="height:60px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 
