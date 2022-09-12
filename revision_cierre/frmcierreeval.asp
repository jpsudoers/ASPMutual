<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vautorizacion = Request("id")

dim query
query= "SELECT CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, "
query= query&"EMPRESAS.R_SOCIAL,EMPRESAS.RUT,AUTORIZACION.ID_OTIC,AUTORIZACION.ID_TIPO_VENTA,"
query= query&"CURRICULO.CODIGO,dbo.MayMinTexto (CURRICULO.NOMBRE_CURSO) as NOMBRE_CURSO,CURRICULO.ID_MUTUAL, "
query= query&"AUTORIZACION.DOCUMENTO_COMPROMISO as doc,(CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra'"
query= query&" WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso' WHEN AUTORIZACION.TIPO_DOC='5' then 'EDP' END) as 'Tipo Documento',"
query= query&"AUTORIZACION.ORDEN_COMPRA,AUTORIZACION.VALOR_OC,AUTORIZACION.N_PARTICIPANTES,AUTORIZACION.VALOR_CURSO, "
query= query&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor,"
query= query&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloque_programacion.nom_sede "
query= query&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as sede,"
query= query&"AUTORIZACION.CON_OTIC,AUTORIZACION.con_franquicia,AUTORIZACION.n_reg_sence,AUTORIZACION.ID_BLOQUE,CON_R=isnull(EMPRESAS.CON_REFERENCIA,0),AUTORIZACION.TIPO_DOC, "
query= query&"AUTORIZACION.N_REFERENCIA,trf.TIPO_REFERENCIA FROM AUTORIZACION " 
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA " 
query= query&" left join TIPO_REFERENCIA trf on trf.ID_TIPO_REFERENCIA=AUTORIZACION.ID_TIPO_REFERENCIA " 
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=AUTORIZACION.id_empresa " 
query= query&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
query= query&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
query= query&" inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede "
query= query&" where AUTORIZACION.ID_AUTORIZACION="&vautorizacion

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmcierreeval" id="frmcierreeval" action="revision_cierre/modificar.asp" method="post">
    <table cellspacing="1" cellpadding="1" border=0>
     <tr>
       <td>Rut Empresa :</td>
       <td><%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td>
       <td>&nbsp;</td>
       <td>Raz&oacute;n Social : </td>
       <td><%=rsEmp("R_SOCIAL")%></td>
    </tr>
    <tr>
       <td>C&oacute;digo Curso :</td>
       <td><%=rsEmp("CODIGO")%></td>
       <td>&nbsp;</td>
       <td>Nombre Curso :</td>
       <td><%=rsEmp("NOMBRE_CURSO")%></td>
    </tr>
    <tr>
       <td width="100">Fecha Inicio :</td>
       <td width="180"><%=rsEmp("FECHA_INICIO_")%></td>
       <td width="20">&nbsp;</td>
       <td width="105">Fecha T&eacute;rmino : </td>
       <td width="495"><%=rsEmp("FECHA_TERMINO")%></td>
    </tr>
    <tr>
       <td colspan="2">Relator :&nbsp;&nbsp;<%=rsEmp("instructor")%></td>
       <td>&nbsp;</td>
       <td colspan="2">Sede :&nbsp;&nbsp;<%=rsEmp("sede")%></td>
    </tr>
    </table>
     <table cellspacing="1" cellpadding="1" border=0>
     <tr>
       <%
	   ckdFranSi=""
	   ckdFranNo=""
	   habfilas=""
   
	   if(rsEmp("con_franquicia")="1")then
		  ckdFranSi="checked='checked'"
		  habfilas=""
	   else
		  ckdFranNo="checked='checked'"
		  habfilas="style='display:none'"
	   end if
	   %> 
       <td width="278">Utiliza Franquicia Sence? :&nbsp;
       	   <input type="radio" name="COfranS" id="COfranS1" value="1" <%=ckdFranSi%> disabled="disabled"/>Si
           <input type="radio" name="COfranS" id="COfranS0" value="0" <%=ckdFranNo%> disabled="disabled"/>No</td>
       <td width="20">&nbsp;</td>
       <td width="310" id="Oticpreg" <%=habfilas%>>Con OTIC :&nbsp;&nbsp;
       <%
	   checkedSi=""
	   checkedNo=""
       habFilaOtic=""
	   
	   if(rsEmp("con_otic")="1")then
		   checkedSi="checked='checked'"
	       habFilaOtic=""		   
	   else
		   checkedNo="checked='checked'"
	       habFilaOtic="style='display:none'"				   
	   end if
	   %>   
       <input type="radio" name="COtic" id="COtic1" value="1" <%=checkedSi%> disabled="disabled"/>Si
       <input type="radio" name="COtic" id="COtic0" value="0" <%=checkedNo%> disabled="disabled"/>No</td>
       <td id="Regpreg" <%=habfilas%>>N&deg; de Registro Sence:&nbsp;&nbsp;<%=rsEmp("n_reg_sence")%></td>
     </tr> 
     <%
      if(rsEmp("ID_OTIC")<>"0")then
		 set rsOtic = conn.execute ("select RUT,R_SOCIAL from EMPRESAS where EMPRESAS.ID_EMPRESA="&rsEmp("ID_OTIC"))
	 %>
     <tr name="Mandante" id="Mandante" <%=habFilaOtic%>>
         <td>Rut OTIC :&nbsp;&nbsp;<%=replace(FormatNumber(mid(rsOtic("RUT"), 1,len(rsOtic("RUT"))-2),0)&mid(rsOtic("RUT"), len(rsOtic("RUT"))-1,len(rsOtic("RUT"))),",",".")%></td>
         <td>&nbsp;</td>
         <td colspan="3">Raz&oacute;n Social :&nbsp;&nbsp;<%=rsOtic("R_SOCIAL")%></td>
     </tr>
     <%end if%>
     <tr>
      <td>Tipo Venta:              
          
        <% if(rsEmp("id_tipo_venta")="1")then %>  Venta Directa 

        <% elseif(rsEmp("id_tipo_venta")="2") then%> Venta Broker
    
     <% else %>
      Sin Asignar
     <% end if %>
      </td>
  
    </tr>
     <tr>
       <td colspan="6">&nbsp;</td>
     </tr>
   </table>
   <div style="width:900px;"> 
      <table id="mytable" border="1" width="900">
      <thead> 
      <tr>
           <th>Rut / Ident.</th> 
           <th>Alumno</th> 
           <th>Asistencia (%)</th> 
           <th>Calificaci&oacute;n (%)</th> 
           <th>&nbsp;Evaluaci&oacute;n&nbsp;</th> 
      </tr>
      </thead> 
      <tbody> 
           <%
			   sol = "select HISTORICO_CURSOS.ID_HISTORICO_CURSO,TRABAJADOR.RUT,TRABAJADOR.NOMBRES, "
			   sol = sol&"HISTORICO_CURSOS.ASISTENCIA,HISTORICO_CURSOS.CALIFICACION,HISTORICO_CURSOS.EVALUACION, "
			   sol = sol&"TRABAJADOR.NACIONALIDAD,TRABAJADOR.ID_EXTRANJERO from HISTORICO_CURSOS "
			   sol = sol&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
			   sol = sol&" where HISTORICO_CURSOS.ID_AUTORIZACION="&vautorizacion
			   'sol = sol&" and HISTORICO_CURSOS.RELATOR="&vrelator
			   sol = sol&" order by TRABAJADOR.NOMBRES asc"
			   
			   set rsSol =  conn.execute(sol)
			   Hist="HiTra"
			   Asis="AsTra"
			   Cal="caTra"
			   Eva="evaluacion"
			   Eva2="eval"
			   i=0
			   tabAsis=1
			   tabCal=2
			   
			   trabRut=""
			   while not rsSol.eof
			   i=i+1
			   tabAsis=tabAsis+2
			   tabCal=tabCal+2
			   
			   if(rsSol("NACIONALIDAD")="0")then
			   		trabRut=replace(FormatNumber(mid(rsSol("rut"), 1,len(rsSol("rut"))-2),0)&mid(rsSol("rut"), len(rsSol("rut"))-1,len(rsSol("rut"))),",",".")
			   else
			   		trabRut=rsSol("ID_EXTRANJERO")
			   end if
			   
	      %>
          <tr>
           <td><input id="<%=Hist&i%>" name="<%=Hist&i%>" type="hidden" value="<%=rsSol("ID_HISTORICO_CURSO")%>"/><%=trabRut%></td>
           <td><%=rsSol("NOMBRES")%></td> 
           <td><input id="<%=Asis&i%>" name="<%=Asis&i%>" type="text" tabindex="<%=tabAsis%>" maxlength="3" size="4" onKeyPress="return acceptNum(event)" onblur="cambiaEstado();" value="<%=rsSol("ASISTENCIA")%>"/>%</td>
           <td><input id="<%=Cal&i%>" name="<%=Cal&i%>" type="text" tabindex="<%=tabCal%>" maxlength="3" size="4" onKeyPress="return acceptNum(event)" onblur="cambiaEstado();" value="<%=rsSol("CALIFICACION")%>"/>%</td>
           <td><input id="<%=Eva2&i%>" name="<%=Eva2&i%>" type="hidden" value="<%=rsSol("EVALUACION")%>"/><label id="<%=Eva&i%>" name="<%=Eva&i%>"><%=rsSol("EVALUACION")%></label></td>
          </tr>
          <%
		  	   rsSol.Movenext
			wend
		  %>
       </tbody> 
     </table>
   </div> 

   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
	   <td width="150">&nbsp;</td>
       <td width="130"></td>
       <td width="20"></td>
       <td width="140"></td>
       <td width="100"></td>
       <td width="20"></td>
       <td width="150"></td>
       <td width="130"></td>
      </tr>
    <tr>
       <td>N&#176; de Participantes : </td>
       <td><%=rsEmp("N_PARTICIPANTES")%></td>
	   <td>&nbsp;</td>
       <td>Valor por Alumno : </td>
       <td><%="$ "&replace(FormatNumber(rsEmp("VALOR_CURSO"),0),",",".")%></td>
       <td>&nbsp;</td>
       <td>Valor Total :</td>
       <td><%="$ "&replace(FormatNumber(rsEmp("VALOR_OC"),0),",",".")%></td>
     </tr>
      <tr>
       <td>Tipo Compromiso : </td>
       <td><%=rsEmp("Tipo Documento")%></td>
	   <td>&nbsp;</td>
       <td>N&uacute;mero OC/CI/DEP : </td>
       <td><%=rsEmp("ORDEN_COMPRA")%></td>
       <td></td>    
       <td>Ver Documento :</td>
       <td><a href="#" onclick="documento('<%=rsEmp("doc")%>');">Ver Documento</a></td>
     </tr>
<%
if(rsEmp("CON_R")="1" and rsEmp("TIPO_DOC")="0") then
%>
      <tr>
       <td>Tipo Referencia : </td>
       <td><%=rsEmp("TIPO_REFERENCIA")%></td>
       <td></td>    
       <td>N&uacute;mero Referencia : </td>
       <td><%=rsEmp("N_REFERENCIA")%></td>
      </tr>

<%
end if
%>
     <tr>
	   <td colspan="8">&nbsp;<input id="countFilas" name="countFilas" type="hidden" value="<%=i%>"/>
       <input id="Bloque" name="Bloque" type="hidden" value="<%=rsEmp("ID_BLOQUE")%>"/>
       <input type="hidden" id="AId" name="AId" value="<%=vautorizacion%>"/>
       <input id="txtIdMutual" name="txtIdMutual" type="hidden" value="<%=rsEmp("ID_MUTUAL")%>"/>
       <input id="txtSoloCert" name="txtSoloCert" type="hidden" value="0"/>
       </td>
      </tr>
     <tr>
       <td colspan="8">&nbsp;</td>
     </tr>
    <tr>
        <td colspan="8"><input type="checkbox" name="liberarsin" id="liberarsin" onclick="eliberarsin();">Liberar Solo Certificados</td>
    </tr>

   </table>
</form> 
<%
end if
%>
<div id="messageBox2" style="height:60px;overflow:auto;width:450px;"> 
  	<ul></ul> 
</div> 
 