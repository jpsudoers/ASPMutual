<!--#include file="../conexion.asp"-->
<%
on error resume next
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
vvacantes = Request("vacantes")

dim query

query="select CURRICULO.NOMBRE_CURSO,CURRICULO.SENCE,CURRICULO.CODIGO,CURRICULO.VALOR, "
query= query&" INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO as instructor, "
query= query&" SEDES.NOMBRE,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&" CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,PROGRAMA.ID_PROGRAMA "
query= query&" from PROGRAMA "
query= query&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=PROGRAMA.ID_INSTRUCTOR "
query= query&" inner join SEDES on SEDES.ID_SEDE=PROGRAMA.ID_SEDE "
query= query&" where ID_PROGRAMA="&vid

set rsEmp = conn.execute (query)

set rs = conn.execute ("select ID_EMPRESA,RUT,R_SOCIAL,TIPO,ID_OTIC from EMPRESAS where EMPRESAS.ID_EMPRESA='"&Session("usuario")&"'")

if not rsEmp.eof and not rsEmp.bof then 
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
       <input type="HIDDEN" id="TipoEmpresa" name="TipoEmpresa" value="<%=rs("TIPO")%>"/></td>
       <td>Nombre Curso :</td>
       <td colspan="4"><%=rsEmp("NOMBRE_CURSO")%></td>
     </tr>
     <tr>
	   <td width="120">Sence :</td>
       <%
	   checkedSi=""
	   checkedNo=""
	   
	   if(rsEmp("SENCE")=0)then
	   checkedSi="checked='checked'"
	   else
	   checkedNo="checked='checked'"
	   end if
	   %>
       
       <td width="100"><input type="radio" name="Sence_SI" id="Sence_SI" disabled value="0" <%=checkedSi%>/>
         Si
           <input type="radio" name="Sence_NO" id="Sence_NO" value="1" disabled <%=checkedNo%>/>
       No</td> 
       <td width="20">&nbsp;</td>
       <td width="120">&nbsp;</td>
       <td width="172">&nbsp;</td>
       <td width="20"></td>
       <td width="90"></td>
       <td width="131"></td>
      </tr>
     <tr>
       <td>Fecha Inicio :</td>
       <td><%=rsEmp("FECHA_INICIO_")%></td>
       <td>&nbsp;</td>
       <td>Fecha Termino :</td>
 	   <td><%=rsEmp("FECHA_TERMINO")%></td>
     </tr>
     <%
	 if(rs("TIPO")="1")then
		 if(rs("ID_OTIC")<>"0")then
			 set rsOtic = conn.execute ("select ID_EMPRESA,RUT,R_SOCIAL from EMPRESAS where EMPRESAS.ID_EMPRESA="&rs("ID_OTIC"))
			 %>
               <tr name="Mandante" id="Mandante" style.display ="">
                   <td>Rut OTIC :</td>
                   <td><%=FormatNumber(mid(rsOtic("RUT"), 1,len(rsOtic("RUT"))-2),0)&mid(rsOtic("RUT"), len(rsOtic("RUT"))-1,len(rsOtic("RUT")))%></td>
                   <td><input type="hidden" id="EmpMan" name="EmpMan" value="<%=rsOtic("ID_EMPRESA")%>"/></td>
                   <td>Raz&oacute;n Social :</td>
                   <td colspan="4"><label id="txtRsocOtic" name="txtRsocOtic"><%=rsOtic("R_SOCIAL")%></label></td>
               </tr>
			 <%
			 else
			 %>
              <tr name="Mandante" id="Mandante" style.display ="">
                   <td><input type="hidden" id="EmpMan" name="EmpMan" value="<%=rs("ID_OTIC")%>"/></td>
              </tr>
             <%
		 end if
	end if
	if(rs("TIPO")="2")then
	%>
          <tr name="Mandante" id="Mandante" style.display ="">
			   <td>Rut OTIC :</td>
			   <td><%=FormatNumber(mid(rs("RUT"), 1,len(rs("RUT"))-2),0)&mid(rs("RUT"), len(rs("RUT"))-1,len(rs("RUT")))%></td>
			   <td><input type="hidden" id="EmpMan" name="EmpMan" value="<%=Session("usuario")%>"/></td>
			   <td>Raz&oacute;n Social :</td>
			   <td colspan="4"><label id="txtRsocOtic" name="txtRsocOtic"><%=rs("R_SOCIAL")%></label></td>
		 </tr>
    <%
	end if
	%>
     <tr>
      <td>&nbsp;</td>
     </tr>
   </table>
<div width="200"><table id="list2"></table> 
   		<div id="pager2"></div> 
   </div>
   <table cellspacing="0" cellpadding="1" border=0>   
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
       <input type="hidden" id="NParticipantes" name="NParticipantes"/></td>
	   <td>&nbsp;</td>
       <td>Valor por Alumno : </td>
       <td><label id="txtValor" name="txtValor"><%="$ "&replace(FormatNumber(rsEmp("VALOR"),0),",",".")%></label>
       <input type="hidden" id="Valor" name="Valor" value="<%=rsEmp("VALOR")%>"/></td>
       <td>&nbsp;</td>
       <td>Valor Total :</td>
       <td><label id="txtValorTotal" name="txtValorTotal">$ 0</label>
       <input type="hidden" id="ValorTotal" name="ValorTotal"/></td>
     </tr>
      <tr>
       <td>Costo Empresa :</td>
       <td><label id="txtCosto" name="txtCosto">$ 0</label>
       <input type="hidden" id="Costo" name="Costo"/></td>
       <td>&nbsp;</td>
       <td>Tipo Compromiso : </td>
       <td>
           <select id="Compromiso" name="Compromiso" tabindex="3">
          		 <OPTION VALUE="">Seleccione</OPTION>
          		 <OPTION VALUE="0">Orden de Compra</OPTION>
   				 <OPTION VALUE="1">Carta Inscripci&oacute;n</OPTION>
   				 <OPTION VALUE="2">Dep&oacute;sito Cheque</OPTION>
           </select>
       </td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><input id="txtNum" name="txtNum" type="text" tabindex="4" maxlength="50" size="12"/></td>
     </tr>
       <tr>
       <td>Subir Documento Pdf : </td>
       <td colspan="4"><input type="file" name="txtDoc" id="txtDoc" tabindex="5" maxlength="200" size="30"></td>
       <td colspan="3"></td>
     </tr>
     <tr>
        <td colspan="8" style="text-align:right"><center>
        <img id="loading" src="../images/loader.gif" style="display:none;"></center></td>
     </tr>
     <tr>
      <td>&nbsp;<input id="datostabla" name="datostabla" type="hidden"/></td>
      <td><input id="datosfran" name="datosfran" type="hidden"/>
       <input type="hidden" id="Vacantes" name="Vacantes" value="<%=vvacantes%>"/></td>
     </tr>
     <tr>
      <td colspan="8"><input type="checkbox" name="terminos" id="terminos" tabindex="17"/>
    	Aceptamos los T&eacute;rminos y Condiciones</label></td>
     </tr>
   </table>
</form> 
<%
   end if
%>

<div id="messageBox1" style="height:60px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div> 
</div>
