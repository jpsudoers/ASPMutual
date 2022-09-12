<!--#include file="../conexion.asp"-->
<%
on error resume next
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
vvacantes = Request("vacantes")

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

set rs = conn.execute ("select ID_EMPRESA,RUT,R_SOCIAL,TIPO,ID_OTIC,MUTUAL from EMPRESAS where ID_EMPRESA='"&Session("usuario")&"'")

dim valor_curso
'valor_curso="CURRICULO.VALOR"
	'if(rs("MUTUAL")="807")then
		'valor_curso="CURRICULO.VALOR_AFILIADOS as VALOR"
	'else
		valor_curso="PROGRAMA.VALOR_ESPECIAL as VALOR"	
	'end if

dim query
query="select dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,CURRICULO.CODIGO,"&valor_curso&",CURRICULO.ID_MUTUAL,"
'query= query&"(select CASE WHEN CURRICULO.SENCE='0' then 'Si' WHEN CURRICULO.SENCE='1' then 'No' END) as SENCE, " 
query= query&" CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&" CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,PROGRAMA.ID_PROGRAMA, "
query= query&"(select top 1 (CASE WHEN S.ID_SEDE =  27 THEN bq.nom_sede WHEN S.ID_SEDE <>  27 THEN "
query= query&"S.DIRECCION+', '+S.CIUDAD END) from bloque_programacion bq inner join SEDES s on s.ID_SEDE=bq.id_sede "
query= query&" where CONVERT(bigint,bq.id_programa)=PROGRAMA.ID_PROGRAMA order by bq.id_bloque asc) as sede "
query= query&" from PROGRAMA "
query= query&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" where PROGRAMA.ID_PROGRAMA="&vid

set rsEmp = conn.execute (query)

dim query3
query3 = "SELECT ID_BANCO, NOMBRE_BANCO FROM BANCOS"
set rsBanco = conn.execute (query3)

if not rs.eof and not rs.bof then 
%>
 <form name="frmProgEmp" id="frmProgEmp" action="emp_programacion/insertar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
  	 <tr>
       <%
	   if(rs("TIPO")="2")then
	   %>
        <td>Rut Mandante :</td>
       <td><input id="txRut" name="txRut" type="text" tabindex="1" maxlength="11" size="11" onkeyup="lookup(this.value);"/>
       <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:1;left:150px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList">
              &nbsp;
            </div>
       </div></td>
       <td><input type="hidden" id="Empresa" name="Empresa"/></td>
       <%
	   end if
	   if(rs("TIPO")="1")then
	   %>
       <td>Rut :</td>
       <td><%=replace(FormatNumber(mid(rs("RUT"), 1,len(rs("RUT"))-2),0)&mid(rs("RUT"), len(rs("RUT"))-1,len(rs("RUT"))),",",".")%></td>
       <td><input type="hidden" id="Empresa" name="Empresa" value="<%=Session("usuario")%>"/></td>
       <%
	   end if
	   %>
      
       <td>Raz&oacute;n Social :</td>
        <%
	   if(rs("TIPO")="2")then
	   %>
       <td colspan="4"><label id="txtRsoc" name="txtRsoc">&nbsp;</label></td>
       <%
	   else
	   %>
        <td colspan="4"><label id="txtRsoc" name="txtRsoc"><%=rs("R_SOCIAL")%></label></td>
       <%
	   end if
	   %>
     </tr>
     <tr>
       <td>C&oacute;digo Curso :</td>
       <td><%=rsEmp("CODIGO")%></td>
       <td><input type="hidden" id="programa" name="programa" value="<%=vid%>"/>
       <input type="hidden" id="TipoEmpresa" name="TipoEmpresa" value="<%=rs("TIPO")%>"/>
       <input type="hidden" id="id_curso" name="id_curso" value="<%=rsEmp("ID_MUTUAL")%>"/></td>
       <td>Nombre Curso :</td>
       <td colspan="4"><%=rsEmp("NOMBRE_CURSO")%><%if(rsEmp("ID_MUTUAL")="39" or rsEmp("ID_MUTUAL")="40")then%>
&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">Pertenece a Proyecto:</font>&nbsp;<input type="text" id="proPert" name="proPert" maxlength="100" size="5"/>
<%end if%></td>
     </tr>
     <tr>
       <td width="120">Fecha Inicio :</td>
       <td width="125"><label id="f_inicio"><%=rsEmp("FECHA_INICIO_")%></label></td>
       <td width="20">&nbsp;</td>
       <td width="110">Fecha Termino :</td>
 	   <td width="130"><%=rsEmp("FECHA_TERMINO")%></td>
       <td width="20">&nbsp;</td>
       <td width="150">&nbsp;</td>
       <td width="250"><input type="hidden" id="COfran" name="COfran"/></td>
     </tr>
     <tr>
       <td colspan="8">Lugar de Ejecución : <%=rsEmp("sede")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if (rsEmp("ID_MUTUAL")<>"2363") then%>Horario: 08:30 a 13:00 y de 14:30 a 18:00<%end if%>&nbsp;<%if (rsEmp("CODIGO")="12.37.8599-88" or rsEmp("CODIGO")="12.37.8329-61") then%> ambos días<%end if%></td>
     </tr>      
     <tr>
       <td colspan="2" id="RegFranSence"><font color="#FF0000">Utiliza Franquicia Sence? :</font>&nbsp;
       	   <input type="radio" name="COfranS" id="COfranS1" value="1" onclick="$('#COfran').val(this.value);MuestraFilas(1);llena_tipo_compromiso(0);"/>Si
           <input type="radio" name="COfranS" id="COfranS0" value="0" onclick="$('#COfran').val(this.value);MuestraFilas(0);llena_tipo_compromiso(0);"/>No</td>
       <td><input type="hidden" id="ConOtic" name="ConOtic" value="0"/></td>
       <td id="Oticpreg" name="Oticpreg" style="display:none">Con OTIC :</td>       
	   <td id="pregOtic" name="pregOtic" style="display:none">
       <input type="radio" name="COtic" id="COtic1" value="1" onclick="$('#ConOtic').val(this.value);$('#Mandante').show();llena_tipo_compromiso(1);llena_otic(0);$('#EmpMan').val('0');"/>Si
       <input type="radio" name="COtic" id="COtic0" value="0" onclick="$('#ConOtic').val(this.value);$('#Mandante').hide();llena_tipo_compromiso(0);$('#EmpMan').val('0');"/>No</td>
       <td>&nbsp;</td>
       <td id="Regpreg" style="display:none">N&deg; de Registro Sence:</td>
       <td id="RegFran" style="display:none"><input type="text" id="NReg_Fran" name="NReg_Fran" maxlength="7" size="8"/></td>
     </tr> 
     <tr name="Mandante" id="Mandante" style="display:none">
                   <td>OTIC : </td>
                   <td colspan="7"><select id="txtOtic" name="txtOtic" style="width:37em;" onchange="$('#EmpMan').val(this.value);"></select><input type="hidden" id="EmpMan" name="EmpMan" value="0"/></td>
     </tr>
     <tr>
      <td colspan="8">&nbsp;<input id="tabProgId" name="tabProgId" type="hidden" value="<%=fecha%>"/></td> 
     </tr>
   </table>
<div width="200"><table id="list2"></table> 
   		<div id="pager2"></div> 
   </div>
   <table cellspacing="0" cellpadding="1" border=0  class="tabla">   
     <tr>
	   <td width="140">&nbsp;</td>
       <td width="145"></td>
       <td width="20"></td>
       <td width="140"></td>
       <td width="100"></td>
       <td width="20"></td>
       <td width="130"></td>
       <td width="210"></td>
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
     
       <td>Tipo Compromiso : </td>
       <td>
           <select id="Compromiso" name="Compromiso" tabindex="3" onchange="busDocEmpresa();"></select>
       </td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><input id="txtNum" name="txtNum" type="text" tabindex="4" maxlength="50" size="12" onkeyup="busDocEmpresa();"/>  <select id="selectOC" name="selectOC" style="display:none" ><option>Seleccione orden de compra</option></select> </td>
       <td><input type="hidden" id="Costo" name="Costo"/></td> 
       <td>Subir Documento : </td>
       <td><input type="file" name="txtDoc" id="txtDoc" tabindex="5" maxlength="400" size="18"></td>
     </tr>
	 <tr  >
		<td><input id="btnIngresarOC" name="btnIngresarOC" type="submit" onclick="IngresarOC();" style="display:none" value="Ingresar Orden de Compra"/></td>
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
      <td colspan="8">&nbsp;<input id="datostabla" name="datostabla" type="hidden"/>
       <input id="datosfran" name="datosfran" type="hidden"/>
       <input type="hidden" id="docExiste" name="docExiste" value="1"/>      
       <input type="hidden" id="Vacantes" name="Vacantes" value="<%=vvacantes%>"/></td>
     </tr>
     <tr>
      <td colspan="4"><input type="checkbox" name="terminos" id="terminos" tabindex="17"/>
      Aceptamos los <a href="#" onclick="abreTerminos();">T&eacute;rminos y Condiciones</a></label></td>
	  <td colspan="4"><font size="-2" color="#FF0000"><i><strong><u>Datos Transferencias:</u></strong> A nombre de Mutual de Seguridad Capacitación S.A. - Rut 76.410.180-4 <br/> N° Cuenta Corriente 10675612 del Banco BCI / enviar notificación a <b><u>cobranza.masc@mutualasesorias.cl</s></u></b></i></font></td>      
     </tr>
   </table>
   <br />
   <div id="messageBox1" style="height:60px;overflow:auto;width:650px;"> 
  	<ul></ul> 
</div> 
</form> 
<%
   end if
%>

