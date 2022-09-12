<!--#include file="../conexion.asp"-->
<%
on error resume next
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
vvacantes = Request("vacantes")

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

'dim valor_curso
'valor_curso="CURRICULO.VALOR"
'if(Request("tipo")="1")then
	'valor_curso="PROGRAMA.VALOR_ESPECIAL as VALOR"
'else
	'if(rs("MUTUAL")="807")then
		'valor_curso="CURRICULO.VALOR_AFILIADOS as VALOR"
	'else
		'valor_curso="CURRICULO.VALOR"	
	'end if
'end if

dim query

query="select dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,CURRICULO.SENCE,CURRICULO.CODIGO,"
query= query&"PROGRAMA.VALOR_ESPECIAL as VALOR,CURRICULO.ID_MUTUAL,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&" CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,PROGRAMA.ID_PROGRAMA "
query= query&" from PROGRAMA "
query= query&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" where ID_PROGRAMA="&vid

set rsEmp = conn.execute (query)

dim query3
query3 = "SELECT ID_BANCO, NOMBRE_BANCO FROM BANCOS"
set rsBanco = conn.execute (query3)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmProgEmp" id="frmProgEmp" action="emp_programacion/insertar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
  	 <tr>
       <td>Rut :</td>
       <%
	   if(Request("tipoCurso")="2")then
	   %>
       <td><input id="txRut" name="txRut" type="text" tabindex="1" maxlength="11" size="11" onkeyup="lookup(this.value);"/>
       <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:1;left:150px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList">
              &nbsp;
            </div>
       </div></td>
       <td><input type="hidden" id="Empresa" name="Empresa"/></td>
       <td>Raz&oacute;n Social :</td>
       <td colspan="4"><label id="txtRsoc" name="txtRsoc"></label></td>
       <%
	   else
	   set rs = conn.execute ("select ID_EMPRESA,RUT,R_SOCIAL,TIPO,ID_OTIC,MUTUAL from EMPRESAS where ID_EMPRESA='"&Request("empresa")&"'")
	   %>
       <td><%=replace(FormatNumber(mid(rs("RUT"), 1,len(rs("RUT"))-2),0)&mid(rs("RUT"), len(rs("RUT"))-1,len(rs("RUT"))),",",".")%></td>
       <td><input type="hidden" id="Empresa" name="Empresa" value="<%=rs("ID_EMPRESA")%>"/><input type="hidden" id="rutOticb" name="rutOticb" value="<%=rs("RUT")%>"/></td>
       <td>Raz&oacute;n Social :</td>
	   <td colspan="4"><label id="txtRsoc" name="txtRsoc"><%=rs("R_SOCIAL")%></label></td>
       <%
	   end if
	   %>       
     </tr>
     <tr>
       <td>C&oacute;digo Curso :</td>
       <td><%=rsEmp("CODIGO")%></td>
       <td><input type="hidden" id="programa" name="programa" value="<%=vid%>"/>
       <input type="hidden" id="TipoEmpresa" name="TipoEmpresa"/>       
       <input type="hidden" id="id_curso" name="id_curso" value="<%=rsEmp("ID_MUTUAL")%>"/></td>
       <td>Nombre Curso :</td>
       <td colspan="4"><%=rsEmp("NOMBRE_CURSO")%><%if(rsEmp("ID_MUTUAL")="39" or rsEmp("ID_MUTUAL")="40" or rsEmp("ID_MUTUAL")="41")then%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">Pertenece a Proyecto:</font>&nbsp;<input type="text" id="proPert" name="proPert" maxlength="100" size="12"/>
<%end if%></td>
     </tr>
     <tr>
       <td width="100">Fecha Inicio :</td>
       <td width="150"><label id="f_inicio"><%=rsEmp("FECHA_INICIO_")%></label></td>
       <td width="20">&nbsp;</td>
       <td width="105">Fecha Termino :</td>
 	   <td width="160"><%=rsEmp("FECHA_TERMINO")%></td>
       <td width="20"></td>
       <td width="160"></td>
       <td width="180"><input type="hidden" id="COfran" name="COfran"/></td>       
     </tr>
     <tr>
       <td colspan="2"><font color="#FF0000">Utiliza Franquicia Sence? :</font>&nbsp;
       	   <input type="radio" name="COfranS" id="COfranS1" value="1" onclick="$('#COfran').val(this.value);MuestraFilas(1);llena_tipo_compromiso(0);"/>Si
           <input type="radio" name="COfranS" id="COfranS0" value="0" onclick="$('#COfran').val(this.value);MuestraFilas(0);llena_tipo_compromiso(0);"/>No</td>
       <td><input type="hidden" id="ConOtic" name="ConOtic" value="0"/></td>
       <td id="Oticpreg" style="display:none">Con OTIC :</td>       
       <td id="pregOtic" style="display:none">
       	  <input type="radio" name="COtic" id="COtic1" value="1" onclick="$('#ConOtic').val(this.value);$('#Mandante').show();llena_tipo_compromiso(1);llena_otic(0);$('#EmpMan').val('0');"/>Si
          <input type="radio" name="COtic" id="COtic0" value="0" onclick="$('#ConOtic').val(this.value);$('#Mandante').hide();llena_tipo_compromiso(0);$('#EmpMan').val('0');"/>No</td>
       <td>&nbsp;</td>
       <td id="Regpreg" style="display:none">N&deg; de Registro Sence:</td>
       <td id="RegFran" style="display:none"><input type="text" id="NReg_Fran" name="NReg_Fran" maxlength="7" size="20"/></td>
     </tr>    
     <tr name="Mandante" id="Mandante" style="display:none">
                   <td>OTIC : </td>
                   <td colspan="7"><select id="txtOtic" name="txtOtic" style="width:37em;" onchange="$('#EmpMan').val(this.value);"></select><input type="hidden" id="EmpMan" name="EmpMan" value="0"/></td>
     </tr>
      <tr>
        <td>Tipo Venta : </td>
        <td><select id="tipoVenta" name="tipoVenta" tabindex="10" style="width:15em;">
          <option value="">Seleccione</option>
          <option value="1">Venta Directa</option>
          <option value="2">Venta Broker</option>
          </select>
        </td>
      </tr>
     <tr>
      <td colspan="8">&nbsp;<input id="tabProgId" name="tabProgId" type="hidden" value="<%=fecha%>"/></td> 
     </tr>
   </table>
<div width="200"><table id="list2"></table> 
   		<div id="pager2"></div> 
   </div>
   <table cellspacing="0" cellpadding="1" border=0 class="tabla">   
     <tr>
	   <td width="155">&nbsp;</td>
       <td width="145"></td>
       <td width="20"></td>
       <td width="155"></td>
       <td width="145"></td>
       <td width="20"></td>
       <td width="155"></td>
       <td width="145"></td>
      </tr>
 <tr>
       <td>N&#176; de Participantes : </td>
       <td><label id="txtNParticipantes" name="txtNParticipantes">0</label>
       <input type="hidden" id="NParticipantes" name="NParticipantes" value="0"/></td>
	   <td>&nbsp;</td>
       <td>Valor por Alumno : </td>
       <td><label id="txtValor" name="txtValor"><%="$ "&replace(FormatNumber(rsEmp("VALOR"),0),",",".")%></label>
       <input type="hidden" id="Valor" name="Valor" value="<%=rsEmp("VALOR")%>"/></td>
       <td>&nbsp;</td>
       <td>Valor Total :</td>
       <td><label id="txtValorTotal" name="txtValorTotal">$ 0</label>
       <input type="hidden" id="ValorTotal" name="ValorTotal" value="0"/></td>
     </tr>
      <tr>
     
       <td>Tipo Compromisos : </td>
       <td>
           <select id="Compromiso" name="Compromiso" tabindex="3" onchange="busDocEmpresa();">
          		 <OPTION VALUE="">Seleccione</OPTION>
                 <OPTION VALUE="4">Cartas Compromiso</OPTION>
          		 <OPTION VALUE="0">Orden de Compra</OPTION>
   				 <OPTION VALUE="1">Vale Vista</OPTION>
   				 <OPTION VALUE="2">Dep&oacute;sito Cheque</OPTION>
                 <OPTION VALUE="3">Transferencia</OPTION>
                  <OPTION VALUE="5">EDP</OPTION>
           </select>
           
       </td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><input id="txtNum" name="txtNum" type="text" tabindex="4" maxlength="50" size="12" onkeyup="busDocEmpresa();"/></td>
       <td><input type="hidden" id="Costo" name="Costo"/></td>
       <td>Subir Documento : </td>
       <td><input type="file" name="txtDoc" id="txtDoc" tabindex="5" maxlength="400" size="18"></td>
     </tr>
     <tr class="transferencia" style="display: none">
     <td>Banco : </td>
     <td>
       <select id="Banco" name="Banco" tabindex="6" style="width:10em;"  onchange="llena_combo_cuentaCorriente(this.value,0)">
         <OPTION VALUE="">Seleccione</OPTION>
         <% if not rsBanco.EoF then
      do until rsBanco.EoF
         response.write "<option value=" & rsBanco(0) & ">" & rsBanco(1) & "</option>"
         rsBanco.moveNext
      loop
   end if %>
       </select>
    </td>
    <td>&nbsp;</td>
    <td>N&uacute;mero Cuenta : </td>
      <td>
        <select id="numeroCuenta" name="numeroCuenta" tabindex="7" style="width:10em;" >
        <OPTION VALUE="">Seleccione</OPTION>
        </select>
      </td><td></td>
      <td>Rut Depositante : </td>
      <td><input id="txtRutDepositante" name="txtRutDepositante" type="text" tabindex="8" maxlength="50" size="12" /></td>
    </tr>
    <tr class="transferencia" style="display: none">
      <td>Monto: </td>
      <td><input id="txtMonto" name="txtMonto" type="text" tabindex="9" maxlength="50" size="12" /></td>
      <td>&nbsp;</td>
      <td>Fecha D&eacute;posito: </td>
      <td><input id="txtFechaDeposito" name="txtFechaDeposito" type="text" tabindex="10" maxlength="50" size="12"  /></td>
      <td></td>
      <td>Nombre Depositante : </td>
      <td><input id="txtNombreDepositante" name="txtNombreDepositante" type="text" tabindex="11" maxlength="50" size="12" /></td>
    </tr>
    <tr class="transferencia" style="display: none;">
      <td>Relaci&oacute;n Depositante con Empresa : </td>
      <td><input id="txtRelacion" name="txtRelacion" type="text" tabindex="12" maxlength="50" size="12" /></td>
      <td></td>
      <td>Mail de Depositante : </td>
      <td><input id="txtEmail" name="txtEmail" type="text" tabindex="13" maxlength="50" size="12" /></td>
      <td>&nbsp;</td>
      <td></td>
      <td></td>
    </tr>
     <tr>
     <td colspan="8"><center><img id="loading" src="../images/loader.gif" style="display:none;"/></center></td>
     </tr>
     <tr>
      <td>&nbsp;<input id="datostabla" name="datostabla" type="hidden"/></td>
      <td><input id="datosfran" name="datosfran" type="hidden"/>
       <input type="hidden" id="docExiste" name="docExiste" value="1"/>
       <input type="hidden" id="Vacantes" name="Vacantes" value="<%=vvacantes%>"/>
       <input type="hidden" id="eBloqueo" name="eBloqueo" value="1"/></td>
     </tr>
     <tr>
      <td ><div id="LinkDoc1"></div></td>
     </tr>
     <tr>
      <td colspan="8"><input type="checkbox" name="terminos" id="terminos" tabindex="17"/>
      Aceptamos los <a href="#" onclick="abreTerminos();">T&eacute;rminos y Condiciones</a></label></td>
     </tr>
   </table>
   <div id="messageBox1" style="height:60px;overflow:auto;width:650px;"> 
  	<ul></ul> 
</div> 
</form> 
<%
   end if
%>

