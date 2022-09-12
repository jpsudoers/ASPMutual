<!--#include file="../conexion.asp"-->
<%
on error resume next
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
vvacantes = Request("vacantes")

 if(Session("tipoUsuario")="")then
		Session.Abandon
		Response.Redirect("index.asp")
 end if

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

dim query

query="select CURRICULO.NOMBRE_CURSO,CURRICULO.SENCE,CURRICULO.CODIGO,CURRICULO.VALOR, "
query= query&" CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&" CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,PROGRAMA.ID_PROGRAMA "
query= query&" from PROGRAMA "
query= query&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" where ID_PROGRAMA="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmProgEmp" id="frmProgEmp" action="emp_programacion/insertar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
  	 <tr>
       <td>Rut :</td>
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
     </tr>
     <tr>
       <td>C&oacute;digo Curso :</td>
       <td><%=rsEmp("CODIGO")%></td>
       <td><input type="hidden" id="programa" name="programa" value="<%=vid%>"/>
       <input type="hidden" id="TipoEmpresa" name="TipoEmpresa"/></td>
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
 	   <td onclick="MostrarFilas('EmpOtic');"><%=rsEmp("FECHA_TERMINO")%></td>
     </tr>
     <tr name="EmpOtic" id="EmpOtic" style="display:none">
       <td>Rut OTIC :</td>
       <td><label id="txtRutOtic" name="txtRutOtic"></label></td>
       <td><input type="hidden" id="EmpMan" name="EmpMan"/></td>
       <td>Raz&oacute;n Social :</td>
       <td colspan="4"><label id="txtRsocOtic" name="txtRsocOtic"></label></td>
     </tr>
     <tr>
      <td>&nbsp;<input id="tabProgId" name="tabProgId" type="hidden" value="<%=fecha%>"/></td> 
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
           <select id="Compromiso" name="Compromiso" tabindex="3">
          		 <OPTION VALUE="">Seleccione</OPTION>
                 <OPTION VALUE="4">Carta Compromiso</OPTION>
          		 <OPTION VALUE="0">Orden de Compra</OPTION>
   				 <OPTION VALUE="1">Vale Vista</OPTION>
   				 <OPTION VALUE="2">Dep&oacute;sito Cheque</OPTION>
                 <OPTION VALUE="3">Transferencia</OPTION>
           </select>
       </td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><input id="txtNum" name="txtNum" type="text" tabindex="4" maxlength="50" size="12"/></td>
       <td><input type="hidden" id="Costo" name="Costo"/></td>      
       <td>&nbsp;</td>
       <td>&nbsp;</td>
     </tr>
     <tr>
       <td>Subir Documento Pdf : </td>
       <td colspan="4"><input type="file" name="txtDoc" id="txtDoc" tabindex="5" maxlength="200" size="30"></td>
       <td colspan="3"></td>
     </tr>
     <tr>
     <td colspan="8"><center><img id="loading" src="../images/loader.gif" style="display:none;"/></center></td>
     </tr>
     <tr>
      <td>&nbsp;<input id="datostabla" name="datostabla" type="hidden"/></td>
      <td><input id="datosfran" name="datosfran" type="hidden"/>
       <input type="hidden" id="Vacantes" name="Vacantes" value="<%=vvacantes%>"/></td>
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

