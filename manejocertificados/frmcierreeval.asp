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
query= query&"CURRICULO.NOMBRE_CURSO, "
query= query&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor, "
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
 <form name="frmcierreeval" id="frmcierreeval" action="#" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td>C&oacute;digo Programaci&oacute;n :</td>
       <td><label id="txtCodProg" name="txtCodProg"><%=right("000000"&rsEmp("ID_PROGRAMA"),5)%></label></td>
       <td><input type="hidden" id="txtId" name="txtId" value="<%=rsEmp("ID_PROGRAMA")%>"/></td>
       <td>&nbsp;</td>
     </tr>
     <tr>
	   <td width="155">C&oacute;digo Curso :</td>
       <td width="200"><%=rsEmp("CODIGO")%></td> 
       <td width="20">&nbsp;</td>
       <td width="130">Nombre Curso :</td>
       <td width="500"><%=rsEmp("NOMBRE_CURSO")%></td>
      </tr>
       <tr>
       <td>Fecha de Inicio : </td>
       <td><%=rsEmp("FECHA_INICIO_")%></td>
       <td>&nbsp;</td>
       <td>Fecha de T&eacute;rmino : </td>
       <td colspan="4"><%=rsEmp("FECHA_TERMINO")%></td>
      </tr>
      <tr>
	   <td width="155">Instructor :</td>
       <td width="200"><%=rsEmp("instructor")%><input id="ins_relator" name="ins_relator" type="hidden" value="<%=rsEmp("ID_INSTRUCTOR")%>"/></td> 
       <td width="20"></td>
       <td width="120">Sede :</td>
       <td width="172" colspan="4"><%=rsEmp("NOMBRE")%></td>
      </tr>
      <tr>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td colspan="4">&nbsp;</td>
      </tr>
   </table>
      <table cellspacing="0" cellpadding="0" border=1>
      <tr>
           <td width="100"><center><h3><em style="text-transform: capitalize;">Rut</em></h3></center></td>
           <td width="250"><center><h3><em style="text-transform: capitalize;">Alumno</em></h3></center></td> 
           <td width="350"><center><h3><em style="text-transform: capitalize;">Empresa</em></h3></center></td> 
           <td width="140"><center><h3><em style="text-transform: capitalize;">Asistencia (%)</em></h3></center></td>
           <td width="140"><center><h3><em style="text-transform: capitalize;">Calificaci&oacute;n (%)</em></h3></center></td>
           <td width="140"><center><h3><em style="text-transform: capitalize;">Evaluaci&oacute;n</em></h3></center></td>
           <td width="70"><h3><input type="checkbox" id="chkall" name="chkall" onclick="checkear();"></h3></td>
          </tr>
           <%
			   sol = "select HISTORICO_CURSOS.ID_HISTORICO_CURSO,HISTORICO_CURSOS.ID_TRABAJADOR,TRABAJADOR.RUT,TRABAJADOR.NOMBRES,"
			   sol = sol&"EMPRESAS.R_SOCIAL,HISTORICO_CURSOS.ASISTENCIA,HISTORICO_CURSOS.CALIFICACION,HISTORICO_CURSOS.EVALUACION "
			   sol = sol&" from HISTORICO_CURSOS "
			   sol = sol&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
			   sol = sol&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
			   sol = sol&" where HISTORICO_CURSOS.ID_PROGRAMA="&vid
			   sol = sol&" and HISTORICO_CURSOS.RELATOR="&vrelator
			   
			   set rsSol =  conn.execute(sol)
			   i=0
			   check="check"
			   Hist="traId"
			   while not rsSol.eof
			   i=i+1
		  %>
					  <tr>
                      <td><input id="<%=Hist&i%>" name="<%=Hist&i%>" type="hidden" value="<%=rsSol("ID_TRABAJADOR")%>"/><%=rsSol("RUT")%></td>
                           <td><%=rsSol("NOMBRES")%></td> 
                           <td><%=rsSol("R_SOCIAL")%></td> 
                           <td><center><%=rsSol("ASISTENCIA")&" %"%></center></td>
                           <td><center><%=rsSol("CALIFICACION")&" %"%></center></td>
                           <td><center><%=rsSol("EVALUACION")%></center></td>
                           <td align="center"><input type="checkbox" id="<%=check&i%>" name="<%=check&i%>" value="<%=i%>" onclick="desCheckearAll();"></td>
					  </tr>
		  <%
		  	   rsSol.Movenext
			  wend
		  %>
     </table>
     <table cellspacing="0" cellpadding="0" border=0>
          <tr>
           <td width="200"><input id="countFilas" name="countFilas" type="hidden" value="<%=i%>"/></td>
           <td width="200">&nbsp;</td> 
           <td width="200">&nbsp;</td>
           <td width="200">&nbsp;</td>
           <td width="200">&nbsp;</td>
          </tr>
    </table>
</form> 
<%
   end if
%>

