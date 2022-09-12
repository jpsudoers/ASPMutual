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
query="select distinct PROGRAMA.ID_PROGRAMA,CURRICULO.CODIGO,CURRICULO.NOMBRE_CURSO, "
query= query&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor, "
query= query&"(CASE WHEN SEDES.ID_SEDE = 27 THEN bloque_programacion.nom_sede "
query= query&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE END) as NOMBRE, "
query= query&"INSTRUCTOR_RELATOR.ID_INSTRUCTOR, "
query= query&"CONVERT(VARCHAR(10),FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),FECHA_TERMINO, 105) as FECHA_TERMINO "
query= query&" from PROGRAMA "
query= query&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" inner join HISTORICO_CURSOS on HISTORICO_CURSOS.ID_PROGRAMA=PROGRAMA.ID_PROGRAMA " 
query= query&" inner join bloque_programacion on bloque_programacion.id_bloque=HISTORICO_CURSOS.ID_BLOQUE "  
query= query&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=HISTORICO_CURSOS.RELATOR "
query= query&" inner join SEDES on SEDES.ID_SEDE=HISTORICO_CURSOS.SEDE "
query= query&" where PROGRAMA.ID_PROGRAMA="&vid
query= query&" and HISTORICO_CURSOS.RELATOR="&vrelator

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmcierreeval" id="frmcierreeval" action="#" method="post">
   <table cellspacing="0" cellpadding="2" border=0>
     <tr>
	   <td width="115">C&oacute;digo Curso :</td>
       <td width="200"><%=rsEmp("CODIGO")%></td> 
       <td width="20"><input type="hidden" id="IdPrograma" name="IdPrograma" value="<%=rsEmp("ID_PROGRAMA")%>"/></td>
       <td width="150">Nombre Curso :</td>
       <td width="450"><%=rsEmp("NOMBRE_CURSO")%></td>
      </tr>
       <tr>
       <td>Fecha de Inicio : </td>
       <td><%=rsEmp("FECHA_INICIO_")%></td>
       <td>&nbsp;</td>
       <td>Fecha de T&eacute;rmino : </td>
       <td colspan="4"><%=rsEmp("FECHA_TERMINO")%></td>
      </tr>
      <tr>
       <td colspan="2">Instructor :&nbsp;&nbsp;<%=rsEmp("instructor")%><input id="IdRelator" name="IdRelator" type="hidden" value="<%=rsEmp("ID_INSTRUCTOR")%>"/></td>
       <td>&nbsp;</td>
	   <td colspan="5">Sede :&nbsp;&nbsp;<%=rsEmp("NOMBRE")%></td> 
      </tr>
      <tr>
       <td colspan="5">&nbsp;</td>
      </tr>
   </table>
   <table id="listEmpresas"></table> 
   <div id="pageEmpresas"></div> 
</form> 
<%
   end if
%>

