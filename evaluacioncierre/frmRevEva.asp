<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
      <table  cellpadding="0" cellspacing="0" border="1" id="tabResumen">
      <thead> 
      <tr>
           <th>Rut / Ident.</th> 
           <th>Alumno</th> 
           <th>Rut Empresa</th> 
           <th>Raz&oacute;n Social</th> 
           <th>Asis CDN(%)</th> 
           <th>Cal CDN(%)</th> 
           <th>Eva CDN</th> 
           <th>Asis Rel(%)</th> 
           <th>Cal Rel(%)</th> 
           <th>Eva Rel</th>            
      </tr>
      </thead> 
      <tbody> 
           <%
			   sol = "select HISTORICO_CURSOS.ID_HISTORICO_CURSO,TRABAJADOR.RUT,TRABAJADOR.NOMBRES,"
			   sol = sol&"TRABAJADOR.NACIONALIDAD,TRABAJADOR.ID_EXTRANJERO,"
			   sol = sol&"HISTORICO_CURSOS.ASIS_CDN,HISTORICO_CURSOS.CAL_CDN,HISTORICO_CURSOS.EVA_CDN,"
               sol = sol&"HISTORICO_CURSOS.ASIS_REL,HISTORICO_CURSOS.CAL_REL,HISTORICO_CURSOS.EVA_REL,"
			   sol = sol&"EMPRESAS.RUT as emp_rut,dbo.MayMinTexto(EMPRESAS.R_SOCIAL) as R_SOCIAL from HISTORICO_CURSOS "
			   sol = sol&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
			   sol = sol&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
			   sol = sol&" where HISTORICO_CURSOS.ID_BLOQUE="&Request("b")
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
           <td><%=rsSol("CAL_CDN")%></td>
            <td><%=rsSol("ASIS_CDN")%></td>
            <td><%=rsSol("EVA_CDN")%></td>
           <td><%=rsSol("CAL_REL")%></td>
            <td><%=rsSol("ASIS_REL")%></td>
            <td><%=rsSol("EVA_REL")%></td>            
          </tr>
          <%
		  	   rsSol.Movenext
			wend
		  %>
      </tbody> 
     </table>