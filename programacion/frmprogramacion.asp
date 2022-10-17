<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")

vUser=0
if(request("u")<>"")then vUser=request("u") end if
dim uf

  uf=" if EXISTS (select ValorUF from UF_Diaria where  month(FechaValor) = month (getdate()) and year(FechaValor) = year (getdate()) and day(FechaValor) = day (getdate())) "
  uf= uf&"begin "
  uf= uf&"select ValorUF from UF_Diaria "
  uf= uf&"where  month(FechaValor) = month (getdate()) "
  uf= uf&"and year(FechaValor) = year (getdate()) "
  uf= uf&"and day(FechaValor) = day (getdate()) "
  uf= uf&"end "
  uf= uf&"else "
  uf= uf&"begin "
   uf= uf&" select ValorUF from UF_Diaria "
  uf= uf&"where  month(FechaValor) = month (getdate()) "
  uf= uf&"and year(FechaValor) = year (getdate()) "
  uf= uf&"and day(FechaValor) = day (getdate()-1) "
  uf= uf&"end "
  

set rsUF = conn.execute(uf)

dim query
query="select ID_PROGRAMA,ID_MUTUAL,ID_INSTRUCTOR,ID_SEDE,SENCE,TIPO,"
query= query&"CONVERT(VARCHAR(10),FECHA_APERTURA, 105) as FECHA_APERTURA,"
query= query&"CONVERT(VARCHAR(10),FECHA_CIERRE, 105) as FECHA_CIERRE,"
query= query&"CONVERT(VARCHAR(10),FECHA_INICIO_, 105) as FECHA_INICIO_,"
query= query&"CONVERT(VARCHAR(10),FECHA_TERMINO, 105) as FECHA_TERMINO,VALOR,CUPOS,INSCRITOS"
query= query&",VACANTES,ESTADO,ID_EMPRESA,DIR_EJEC,VALOR_ESPECIAL,MONTO_TOTAL, ValorUF, "
query= query&"(CASE WHEN PROGRAMA.TIPO =  1 or PROGRAMA.TIPO =  3 THEN (select RUT from EMPRESAS where ID_EMPRESA=PROGRAMA.ID_EMPRESA) "
query= query&"WHEN PROGRAMA.TIPO =  2 THEN '' END)as rut_Empresa,"
query= query&"(CASE WHEN PROGRAMA.TIPO =  1 or PROGRAMA.TIPO =  3 THEN (select R_SOCIAL from EMPRESAS where ID_EMPRESA=PROGRAMA.ID_EMPRESA)"
query= query&"WHEN PROGRAMA.TIPO =  2 THEN '' END)as Nom_Empresa,BMI,BMF,BTI,BTF,ID_Modalidad,ID_ZOOM, tipoValor, "
query= query&"usr=ISNULL((select 'Coordinado por '+u.NOMBRES+' '+u.A_PATERNO+' '+u.A_MATERNO from usuarios u where u.ID_USUARIO=PROGRAMA.USUARIO_CREA),'') "
query= query&" from PROGRAMA where ID_PROGRAMA="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
fecha = vid
%>
 <form name="frmProgramacion" id="frmProgramacion" action="programacion/modificar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td>C&oacute;digo Curso :</td>
       <td><select id="Curriculo" name="Curriculo" tabindex="1" onchange="cargaDatos(this.value);"></select></td>
       <td><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("ID_PROGRAMA")%>"  /><input type="hidden" id="txtIdCurriculo" name="txtIdCurriculo" value="<%=rsEmp("ID_MUTUAL")%>" /></td>
		<td><input type="hidden" id="txtVlUfActual" name="txtVlUfActual" value="<%=rsUF("ValorUF")%>"  />
	   <td>Nombre :</td>
       <td colspan="4"><label id="txtCurso" name="txtCurso"></label></td>
     </tr>
     <tr>
	   <td width="130">Sence :</td>
       <td width="170"><input type="radio" name="Sence_SI" id="Sence_SI" disabled value="0">Si
		   				<input type="radio" name="Sence_NO" id="Sence_NO" value="1" disabled>No</td> 
       <td width="20"></td>
       <td width="120"></td>
       <td width="110"></td>
       <td width="20"></td>
       <td width="90"></td>
       <td width="248"></td>
      </tr>
      <tr>
       <td>Tipo :</td>
         <td><select id="Tipo" name="Tipo" tabindex="3" onchange="mostrarFilaEmpresa(this.value);"></select></td>
         <td><input type="hidden" id="txtTipo" name="txtTipo" value="<%=rsEmp("TIPO")%>" /></td>
     </tr>
     <tr id="filaEmpresa">
       <td>Rut de Empresa :</td>
       <td><input id="txRut" name="txRut" type="text" tabindex="4" maxlength="11" size="12" onkeyup="lookup(this.value);" value="<%=rsEmp("rut_Empresa")%>"/>
       <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:1;left:150px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList">
              &nbsp;
            </div>
       </div></td>
       <td><input type="hidden" id="id_empresa" name="id_empresa" value="<%=rsEmp("ID_EMPRESA")%>"/></td>
       <td>Raz&oacute;n Social :</td>
       <td colspan="4"><label id="txtRsoc" name="txtRsoc"><%=rsEmp("Nom_Empresa")%></label></td>
     </tr>
     <tr>
       <td>Fecha Apertura :</td>
       <td><input id="txtFechApertura" name="txtFechApertura" type="text" tabindex="7" maxlength="50" value="<%=rsEmp("FECHA_APERTURA")%>" size="12"/></td>
       <td></td>
       <td>Fecha Cierre :</td>
       <td><input id="txtFechCierre" name="txtFechCierre" type="text" tabindex="8" maxlength="50" value="<%=rsEmp("FECHA_CIERRE")%>" size="12"/></td>
     </tr>
     <tr>
       <td>Fecha Inicio :</td>
       <td><input id="txtFechInicio" name="txtFechInicio" type="text" tabindex="9" maxlength="50" value="<%=rsEmp("FECHA_INICIO_")%>" size="12"/></td>
       <td></td>
       <td>Fecha Termino :</td>
       <td><input id="txtFechTermino" name="txtFechTermino" type="text" tabindex="10" maxlength="50" value="<%=rsEmp("FECHA_TERMINO")%>" size="12"/></td>
       <td></td>
       <td>Direcci&oacute;n :</td>
       <td><input id="txLugar" name="txLugar" type="text" tabindex="11" maxlength="50" size="30" value="<%=rsEmp("DIR_EJEC")%>"/></td>
     </tr>
	 <tr>
	 <td id="tdlabTipoVa">Tipo de Cobro:</td>
	 <td id="tdTipoValor">
	   <select id="selectTValor" name="selectTValor" tabindex="12"  size="1" onchange="MostrarValorUF()">
	   <%
	   if rsEmp("tipoValor") = 0 then
	   %>
		<option value="0" selected>Pesos</option>
		<Option value="1">UF </option>
	   <%
		else
	   %>
	   <option value="0" >Pesos</option>
		<Option value="1" selected>UF </option>
	   <%
	   end if
	   %>
	   </select>
	   </td>
	   
	   
	   <%
	   if rsEmp("tipoValor") = 0 then
	   %>
	   <td id="ContentUF" name="ContentUF" > 
	     <td id="tdlabTipoVaUF" style="display:none">Cantidad UF:</td>
		 <td><input id="txtValUF" name="txtValUF" type="text" tabindex="12" maxlength="50" size="12" value="<%=rsEmp("ValorUF")%>" style="display:none" onchange="CalcularPrecioUF()"/></td>
	
	  </td>
	   <%
		else
	   %>
		<td id="ContentUF" name="ContentUF"> 
		  <td id="tdlabTipoVaUF">Cantidad UF:</td>
		  <td><input id="txtValUF" name="txtValUF" type="text" tabindex="12" maxlength="50" size="12" value="<%=rsEmp("ValorUF")%>" onchange="CalcularPrecioUF()"/></td>
	    </td>
	   <%
	    end if
	   %>
		
	 </tr>
     <tr>
       <td><label id="labValor" name="labValor">Valor :</label></td>
       <td><label id="txtValor" name="txtValor">&nbsp;</label></td>
       <td>&nbsp;</td>       
       <td id="tdLabValEsp">Valor Especial:</td>
       <td id="tdValEsp"><input id="txtValEsp" name="txtValEsp" type="text" tabindex="12" maxlength="50" size="12" value="<%=rsEmp("VALOR_ESPECIAL")%>"/></td> 
       <td id="tdLabValTot">Monto Total:</td>
       <td id="tdValTot"><input id="txtValTot" name="txtValTot" type="text" tabindex="12" maxlength="50" size="12" value="<%=rsEmp("MONTO_TOTAL")%>"/></td>           
     </tr>
      <tr>
       <td>Cupos :</td>
       <td><input id="txtCupo" name="txtCupo" type="text" tabindex="13" maxlength="50" value="<%=rsEmp("CUPOS")%>" size="12"  onkeyup="calVacantes()"  onKeyPress="return acceptNum(event)"/></td>
       <td></td>
       <td>Inscritos :</td>
       <td><label id="txtInscritos" name="txtInscritos"></label></td>
       <td></td>
       <td>Vacantes :</td>
       <td><label id="txtVacantes" name="txtVacantes"><%=rsEmp("VACANTES")%></label></td>
     </tr>
     <tr>
         <td>Modalidad : </td>
         <td>
	<input type="radio" name="txtModalidad" id="txtModalidad_2" value="2" <%if(rsEmp("ID_Modalidad")="2")then%>checked="checked"<%end if%>  onclick="modalidad(2);">Asincr&oacute;nico
	<input type="radio" name="txtModalidad" id="txtModalidad_1" value="1" <%if(rsEmp("ID_Modalidad")="1")then%>checked="checked"<%end if%>  onclick="modalidad(1);">Presencial
	 <input type="radio" name="txtModalidad" id="txtModalidad_0" value="0" <%if(rsEmp("ID_Modalidad")="0")then%>checked="checked"<%end if%> onclick="modalidad(0);">Virtual</td> 
	 <td>&nbsp;</td>
         <td colspan="2" id="TdLabelHorario">Horario :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<select id="txtBMInicio" name="txtBMInicio"></select> a <select id="txtBTFin" name="txtBTFin"></select></td>
         <td>&nbsp;</td>
         <td id="TdLabelZoom" style="display:none">ID Zoom :</td>
         <td id="TdZoom" style="display:none"><input id="txtIDZoom" name="txtIDZoom" type="text" maxlength="12" size="12" onKeyPress="return acceptNum(event)" value="<%=rsEmp("ID_ZOOM")%>"/></td>                                
     </tr>
     <tr>
       <td colspan="8">&nbsp;</td>
     </tr>
   </table>
     <table id="list2"></table> 
     <div id="pager2"></div> 
     <table cellspacing="0" cellpadding="0" border=0>
          <tr>
              <td>&nbsp;<input id="tabFecha" name="tabFecha" type="hidden" value="<%=fecha%>"/>
	              <input id="numBloques" name="numBloques" type="hidden" value="1"/>
                  <input id="vValEsp" name="vValEsp" type="hidden" value="<%=rsEmp("VALOR_ESPECIAL")%>"/>
                  <input id="vValTot" name="vValTot" type="hidden" value="<%=rsEmp("MONTO_TOTAL")%>"/>
                  <input id="vmi" name="vmi" type="hidden" value="<%=rsEmp("BMI")%>"/> 
		  <input id="vmf" name="vmf" type="hidden" value="<%=rsEmp("BMF")%>"/>    
                  <input id="vti" name="vti" type="hidden" value="<%=rsEmp("BTI")%>"/> 
		  <input id="vtf" name="vtf" type="hidden" value="<%=rsEmp("BTF")%>"/> 
		  <input id="txtIDModalidad" name="txtIDModalidad" type="hidden" value="<%=rsEmp("ID_Modalidad")%>"/>
              </td> 
          </tr>
    </table>
</form> 
<%
   else
   fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
   fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
%> <form name="frmProgramacion" id="frmProgramacion" action="programacion/insertar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td>C&oacute;digo Curso :</td>
       <td><select id="Curriculo" name="Curriculo" tabindex="1" onchange="cargaDatos(this.value);"></select></td>
	   <td><input type="hidden" id="txtVlUfActual" name="txtVlUfActual" value="<%=rsUF("ValorUF")%>"  />
       <td></td>
       <td>Nombre :</td>
       <td colspan="4"><label id="txtCurso" name="txtCurso"></label></td>
     </tr>
     <tr>
	   <td width="130">Sence :</td>
       <td width="170"><input type="radio" name="Sence_SI" id="Sence_SI" disabled value="0">Si
		   				<input type="radio" name="Sence_NO" id="Sence_NO" value="1" disabled>No</td> 
       <td width="20"></td>
       <td width="120"></td>
       <td width="110"></td>
       <td width="20"></td>
       <td width="90"></td>
       <td width="248"></td>
      </tr>
     <tr>
       <td>Tipo :</td>
         <td><select id="Tipo" name="Tipo" tabindex="3" onchange="mostrarFilaEmpresa(this.value);"></select></td>
     </tr>
     <tr id="filaEmpresa">
       <td>Rut de Empresa :</td>
       <td><input id="txRut" name="txRut" type="text" tabindex="4" maxlength="11" size="12" onkeyup="lookup(this.value);"/>
       <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:1;left:150px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList">
              &nbsp;
            </div>
       </div></td>
       <td><input type="hidden" id="id_empresa" name="id_empresa"/></td>
       <td>Raz&oacute;n Social :</td>
       <td colspan="4"><label id="txtRsoc" name="txtRsoc"></label></td>
     </tr>     
     <tr>
       <td>Fecha Apertura :</td>
       <td><input id="txtFechApertura" name="txtFechApertura" type="text" tabindex="7" maxlength="50" size="12"/></td>
       <td></td>
       <td>Fecha Cierre :</td>
       <td><input id="txtFechCierre" name="txtFechCierre" type="text" tabindex="8" maxlength="50" size="12"/></td>
     </tr>
     <tr>
       <td>Fecha Inicio :</td>
       <td><input id="txtFechInicio" name="txtFechInicio" type="text" tabindex="9" maxlength="50" size="12"/></td>
       <td></td>
       <td>Fecha Termino :</td>
       <td><input id="txtFechTermino" name="txtFechTermino" type="text" tabindex="10" maxlength="50" size="12"/></td>
       <td></td>
       <td>Direcci&oacute;n :</td>
       <td><input id="txLugar" name="txLugar" type="text" tabindex="11" maxlength="50" size="30"/></td>
     </tr>
	 <tr>
	  <td id="tdLabTipoVa">Tipo de Cobro:</td>
		<td id="tdTipoValor">
		<Select id="selectTValor" name="selectTValor" tabindex="12"  size="1" onchange="MostrarValorUF()">
		<option value="0" selected>Pesos</option>
		<Option value="1">UF </option>
		</select> </td>
		<td>
		 <td id="tdlabTipoVaUF" style="display:none">Cantidad UF:</td>
		 <td><input id="txtValUF" name="txtValUF" type="text" tabindex="12" maxlength="50" size="12" value="0" style="display:none" onchange="CalcularPrecioUF()" /></td>
		<td>
	 </tr>
     <tr>
       <td><label id="labValor" name="labValor">Valor :</label></td>
       <td><label id="txtValor" name="txtValor">&nbsp;</label></td>
       <td>&nbsp;</td>       
       <td id="tdLabValEsp">Valor Especial:</td>
       <td id="tdValEsp"><input id="txtValEsp" name="txtValEsp" type="text" tabindex="12" maxlength="50" size="12"/></td> 
	   
	  
       <td id="tdValTot"><input id="txtValTot" name="txtValTot" type="text" tabindex="12" maxlength="50" size="12"/></td>   
     </tr>
      <tr>
       <td>Cupos :</td>
       <td><input id="txtCupo" name="txtCupo" type="text" tabindex="13" maxlength="50" size="12" onkeyup="calcula()"  onKeyPress="return acceptNum(event)"/></td>
       <td></td>
       <td>Inscritos :</td>
       <td><label id="txtInscritos" name="txtInscritos"></label></td>
       <td></td>
       <td>Vacantes :</td>
       <td><label id="txtVacantes" name="txtVacantes"></label></td>
     </tr>
      <tr>
         <td>Modalidad : </td>
         <td>
<input type="radio" name="txtModalidad" id="txtModalidad_2" value="2" onclick="modalidad(2);">Asincr&oacute;nico
<input type="radio" name="txtModalidad" id="txtModalidad_1" value="1" checked="checked" onclick="modalidad(1);">Presencial
	 <input type="radio" name="txtModalidad" id="txtModalidad_0" value="0" onclick="modalidad(0);">Virtual</td>  
         <td>&nbsp;</td>   
         <td colspan="2" id="TdLabelHorario">Horario :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<select id="txtBMInicio" name="txtBMInicio"></select> a <select id="txtBTFin" name="txtBTFin"></select></td>       
         <td>&nbsp;</td> 
         <td id="TdLabelZoom" style="display:none">ID Zoom :</td>
         <td id="TdZoom" style="display:none"><input id="txtIDZoom" name="txtIDZoom" type="text" maxlength="12" size="12" onKeyPress="return acceptNum(event)"/></td>                     </tr>
      <tr>
       <td colspan="8">&nbsp;</td>
     </tr>
   </table>
     <table id="list2"></table> 
     <div id="pager2"></div> 
     <table cellspacing="0" cellpadding="0" border=0>
          <tr>
               <td>&nbsp;<input id="tabFecha" name="tabFecha" type="hidden" value="<%=fecha%>"/>
                   <input id="numBloques" name="numBloques" type="hidden"/>
                   <input id="txtUser" name="txtUser" type="hidden" value="<%=vUser%>"/> 
		   <input id="txtIDModalidad" name="txtIDModalidad" type="hidden" value="1"/>
               </td> 
          </tr>
    </table>
</form> 
<%
   end if
%>
<div id="messageBox2" style="height:50px;overflow:auto;width:500px;"> 
  	<ul></ul> 
</div> 
