<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
vrelator = Request("relator")

dim query
query="select distinct PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,CURRICULO.SENCE, "
query= query&"CURRICULO.NOMBRE_CURSO, CURRICULO.PORCE_ASISTENCIA, CURRICULO.PORCE_CALIFICACION ,CURRICULO.ID_MUTUAL, "
query= query&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor, "
query= query&"SEDES.NOMBRE,INSTRUCTOR_RELATOR.ID_INSTRUCTOR, "
query= query&"CONVERT(VARCHAR(10),FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),FECHA_TERMINO, 105) as FECHA_TERMINO "
query= query&" from PROGRAMA "
query= query&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" inner join HISTORICO_CURSOS on HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA "  
query= query&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=HISTORICO_CURSOS.RELATOR "
query= query&" inner join SEDES on SEDES.ID_SEDE=HISTORICO_CURSOS.SEDE "
query= query&" where PROGRAMA.ID_PROGRAMA="&vid
query= query&" and HISTORICO_CURSOS.RELATOR="&vrelator

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmcierreeval" id="frmcierreeval" action="evaluacioncierre/modificar.asp" method="post">
 <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="155">C&oacute;digo Curso :</td>
       <td width="200"><%=rsEmp("CODIGO")%></td> 
       <td width="20"><input id="txtIdMutual" name="txtIdMutual" type="hidden" value="<%=rsEmp("ID_MUTUAL")%>"/></td>
       <td width="145">Nombre Curso :</td>
       <td width="400"><%=rsEmp("NOMBRE_CURSO")%></td>
     </tr>
     <tr>
       <td>Fecha de Inicio : </td>
       <td><%=rsEmp("FECHA_INICIO_")%></td>
       <td>&nbsp;</td>
       <td>Fecha de T&eacute;rmino : </td>
       <td><%=rsEmp("FECHA_TERMINO")%></td>
      </tr>
      <tr>
	   <td>Instructor :</td>
       <td><%=rsEmp("instructor")%></td> 
       <td>&nbsp;</td>
       <td>Sede :</td>
       <td><%=rsEmp("NOMBRE")%></td>
      </tr>
      <tr>
       <td colspan="5">&nbsp;</td>
	    <INPUT id="AsistenciaNecesaria" TYPE="hidden" value="<%=rsEmp("PORCE_ASISTENCIA")%>" />
	   <INPUT id="CalificacionNecesaria" TYPE="hidden" value="<%=rsEmp("PORCE_CALIFICACION")%>" />
      </tr>
   </table>
        <div style="width:960px;"> 
      <table id="mytable" border="1" width="960">
      <thead> 
      <tr>
           <th>Rut / Ident.</th> 
           <th>Alumno</th> 
           <th>Rut Empresa</th> 
           <th>Raz&oacute;n Social</th> 
           <th>Asistencia (%)</th> 
           <th>Calificaci&oacute;n (%)</th> 
           <th>&nbsp;Evaluaci&oacute;n&nbsp;</th> 
      </tr>
      </thead> 
      <tbody> 
           <%
			   sol = " select HISTORICO_CURSOS.ID_HISTORICO_CURSO,TRABAJADOR.RUT,TRABAJADOR.NOMBRES,"
			   sol = sol&" TRABAJADOR.NACIONALIDAD,TRABAJADOR.ID_EXTRANJERO,"
			   sol = sol&" EMPRESAS.RUT as emp_rut,dbo.MayMinTexto(EMPRESAS.R_SOCIAL) as R_SOCIAL,"
			   sol = sol&" CURRICULO.PORCE_ASISTENCIA,  CURRICULO.PORCE_CALIFICACION  from HISTORICO_CURSOS"
			   sol = sol&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
			   sol = sol&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
			   sol = sol&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA = HISTORICO_CURSOS.ID_PROGRAMA"
			   sol = sol&" inner join CURRICULO on CURRICULO.ID_MUTUAL = PROGRAMA.ID_MUTUAL"
			   sol = sol&" where HISTORICO_CURSOS.ID_PROGRAMA="&vid
			   sol = sol&" and HISTORICO_CURSOS.RELATOR="&vrelator
			   sol = sol&" order by TRABAJADOR.NOMBRES asc"
			   
			   set rsSol =  conn.execute(sol)
			   Hist="H"
			   Asis="A"
			   Cal="C"
			   Eva="evaluacion"
			   Eva2="E"
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
           <td><%=replace(FormatNumber(mid(rsSol("emp_rut"), 1,len(rsSol("emp_rut"))-2),0)&mid(rsSol("emp_rut"), len(rsSol("emp_rut"))-1,len(rsSol("emp_rut"))),",",".")%></td> 
           <td><%=rsSol("R_SOCIAL")%></td> 
		   <td><select id="<%=Asis&i%>" name="<%=Asis&i%>" maxlength="3" size="1" tabindex="<%=tabAsis%>" onclick="cambiaEstado();">
		   <%
		   for number=100 to 0 step -1 
		   %>
			<option><%=number%></option>
		   <%
			next
		   %>
		   </select></td> 
		   <td><select id="<%=Cal&i%>" name="<%=Cal&i%>" maxlength="3" size="1" tabindex="<%=tabCal%>" onclick="cambiaEstado();">
			<%
		   for number=100 to 0 step -1 
		   %>
			<option><%=number%></option>
		   <%
			next
		   %>
		   </select></td> 
           <td><label id="<%=Eva&i%>" name="<%=Eva&i%>">Reprobado</label><input id="<%=Eva2&i%>" name="<%=Eva2&i%>" type="hidden" value="Reprobado"/></td>
          </tr>
          <%
		  	   rsSol.Movenext
			wend
		  %>
      </tbody> 
     </table>
      </div> 
      <table cellspacing="0" cellpadding="0" border=0>
          <tr>
           <td><input id="countFilas" name="countFilas" type="hidden" value="<%=i%>"/></td>
          </tr>
       </table>
</form> 
<%
   end if
%>

