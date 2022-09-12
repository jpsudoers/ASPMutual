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
query= query&"CURRICULO.NOMBRE_CURSO, CURRICULO.PORCE_ASISTENCIA, CURRICULO.PORCE_CALIFICACION,CURRICULO.ID_MUTUAL, "
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
 <form name="frmingevacdn" id="frmingevacdn" action="evaluacioncierre/modificarIngEvaCdn.asp" method="post">
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
        <div style="width:1030px;"> 
      <table id="mytable2" border="1" width="1030">
      <thead> 
      <tr>
           <th>Rut / Ident.</th> 
           <th>Alumno</th> 
           <!--<th>Rut Empresa</th> -->
           <th>Raz&oacute;n Social</th> 
           <th>Asis. CDN(%)</th> 
           <th>Cal. CDN(%)</th> 
           <th>Eval. CDN</th>
           <th>Asis. REL(%)</th> 
           <th>Cal. REL(%)</th> 
           <th>Eval. REL</th>
           <th>PO</th>
           <th>PVMO</th>
           <th>Editar</th>             
      </tr>
      </thead> 
      <tbody> 
           <%
			   sol = "select HISTORICO_CURSOS.ID_HISTORICO_CURSO,TRABAJADOR.RUT,TRABAJADOR.NOMBRES,"
			   sol = sol&"TRABAJADOR.NACIONALIDAD,TRABAJADOR.ID_EXTRANJERO,"
			 sol = sol&"HISTORICO_CURSOS.ASIS_CDN,HISTORICO_CURSOS.CAL_CDN,EVA_CDN=ISNULL(HISTORICO_CURSOS.EVA_CDN,'Reprobado'),"
             sol = sol&"HISTORICO_CURSOS.ASIS_REL,HISTORICO_CURSOS.CAL_REL,EVA_REL=ISNULL(HISTORICO_CURSOS.EVA_REL,'Reprobado'),"	
			 sol= sol&" isnull(HISTORICO_CURSOS.PO,0) as PO,isnull(HISTORICO_CURSOS.PVMO,0) as PVMO,"
			   sol = sol&"EMPRESAS.RUT as emp_rut,dbo.MayMinTexto(EMPRESAS.R_SOCIAL) as R_SOCIAL from HISTORICO_CURSOS "
			   sol = sol&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
			   sol = sol&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
			   sol = sol&" where HISTORICO_CURSOS.ID_PROGRAMA="&vid
			   sol = sol&" and HISTORICO_CURSOS.RELATOR="&vrelator
			   sol = sol&" order by TRABAJADOR.NOMBRES asc"
			   set rsSol =  conn.execute(sol)
			   Hist="H"
			   Asis="A"
			   Cal="C"
			   Eva="evaluacion"
			   Eva2="E"
			   
			   Asis2="Ar"
			   Cal2="Cr"
			   AsiR="ASR"
			   CalR="CAR"
			   Eva2R="evaluacionr"
			   Eva_2="Er"			   
			   
			   Editar="EDT"
			   E_Edt="E_DT"
			   C_Edt="C_DT"
			   i=0
		   
			   trabRut=""
			   while not rsSol.eof
			   i=i+1
			   
			   if(rsSol("NACIONALIDAD")="0")then
			   		trabRut=replace(FormatNumber(mid(rsSol("rut"), 1,len(rsSol("rut"))-2),0)&mid(rsSol("rut"), len(rsSol("rut"))-1,len(rsSol("rut"))),",",".")
			   else
			   		trabRut=rsSol("ID_EXTRANJERO")
			   end if
	      %>
          <tr>
           <td><input id="<%=Hist&i%>" name="<%=Hist&i%>" type="hidden" value="<%=rsSol("ID_HISTORICO_CURSO")%>"/><%=trabRut%></td>
           <td><%=rsSol("NOMBRES")%></td> 
           <!--<td><%=replace(FormatNumber(mid(rsSol("emp_rut"), 1,len(rsSol("emp_rut"))-2),0)&mid(rsSol("emp_rut"), len(rsSol("emp_rut"))-1,len(rsSol("emp_rut"))),",",".")%></td> -->
           <td><%=rsSol("R_SOCIAL")%></td> 
           <td><input id="<%=Asis&i%>" name="<%=Asis&i%>" type="text" maxlength="3" size="4" onKeyPress="return acceptNum(event)" onblur="cambiaEstado();" value="<%=rsSol("ASIS_CDN")%>" <%if IsNull(rsSol("ASIS_CDN")) and IsNull(rsSol("CAL_CDN")) and not IsNull(rsSol("ASIS_REL")) and not IsNull(rsSol("CAL_REL")) then%> style="display:none;" <%end if%>/></td>
           <td><input id="<%=Cal&i%>" name="<%=Cal&i%>" type="text" maxlength="3" size="4" onKeyPress="return acceptNum(event)" onblur="cambiaEstado();" value="<%=rsSol("CAL_CDN")%>" <%if IsNull(rsSol("ASIS_CDN")) and IsNull(rsSol("CAL_CDN")) and not IsNull(rsSol("ASIS_REL")) and not IsNull(rsSol("CAL_REL")) then%> style="display:none;" <%end if%>/></td>
           <td><label id="<%=Eva&i%>" name="<%=Eva&i%>" <%if IsNull(rsSol("ASIS_CDN")) and IsNull(rsSol("CAL_CDN")) and not IsNull(rsSol("ASIS_REL")) and not IsNull(rsSol("CAL_REL")) then%> style="display:none;" <%end if%>><%=rsSol("EVA_CDN")%></label><input id="<%=Eva2&i%>" name="<%=Eva2&i%>" type="hidden" value="<%=mid(rsSol("EVA_CDN"),1,1)%>"/></td>
           <td><label id="<%=AsiR&i%>" name="<%=AsiR&i%>"><%=rsSol("ASIS_REL")%></label><input id="<%=Asis2&i%>" name="<%=Asis2&i%>" type="text" maxlength="3" size="4" onKeyPress="return acceptNum(event)" onblur="cambiaEstadoRel();" value="<%=rsSol("ASIS_REL")%>" style="display:none;"/></td>
           <td><label id="<%=CalR&i%>" name="<%=CalR&i%>"><%=rsSol("CAL_REL")%></label><input id="<%=Cal2&i%>" name="<%=Cal2&i%>" type="text"  maxlength="3" size="4" onKeyPress="return acceptNum(event)" onblur="cambiaEstadoRel();" value="<%=rsSol("CAL_REL")%>" style="display:none;"/></td>
           <td><label id="<%=Eva2R&i%>" name="<%=Eva2R&i%>" <%if IsNull(rsSol("ASIS_REL")) or IsNull(rsSol("CAL_REL")) then%> style="display:none;" <%end if%>><%=rsSol("EVA_REL")%></label><input id="<%=Eva_2&i%>" name="<%=Eva_2&i%>" type="hidden" value="<%=mid(rsSol("EVA_REL"),1,1)%>"/></td>
           <td><%if int(rsSol("PO")) = 0 Then %>No<%else%>Si<%end if%></td>
           <td><%if int(rsSol("PVMO")) = 0 Then %>No<%else%>Si<%end if%></td>
			<td width="95"><a href="#" onclick="edtInfoRel(<%=i%>);" id="<%=Editar&i%>" name="<%=Editar&i%>" <%if IsNull(rsSol("ASIS_REL")) or IsNull(rsSol("CAL_REL")) then%> style="display:none;" <%end if%>>Editar</a>&nbsp;<a href="#" onclick="edtInfoRelCan(<%=i%>);" id="<%=C_Edt&i%>" name="<%=C_Edt&i%>" style="display:none;">Cancelar</a><input id="<%=E_Edt&i%>" name="<%=E_Edt&i%>" type="hidden" value="0"/></td>           
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
           <td><input id="countFilas" name="countFilas" type="hidden" value="<%=i%>"/><input id="u" name="u" type="hidden" value="<%=Request("u")%>"/><input id="b" name="b" type="hidden" value="<%=Request("b")%>"/><input id="t" name="t" type="hidden" value="<%=Request("t")%>"/><input id="ec" name="ec" type="hidden" value="<%=Request("ec")%>"/></td>
          </tr>
       </table>
</form> 
<%
   end if
%>

