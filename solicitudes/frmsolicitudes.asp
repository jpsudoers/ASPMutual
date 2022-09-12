<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
dim query
query="SELECT SOLICITUD.id_solicitud,EMPRESAS.RUT, EMPRESAS.R_SOCIAL, MUTUALES.nomb_mutual, OTIC.R_SOCIAL as otic, "
query= query&"CURRICULO.codigo, CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),SOLICITUD.fecha, 105) as fecha, SOLICITUD.participantes "
query= query&" FROM SOLICITUD "
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=SOLICITUD.id_empresa "
query= query&" INNER JOIN MUTUALES ON MUTUALES.Mutual_id =EMPRESAS.MUTUAL "
query= query&" INNER JOIN OTIC ON OTIC.ID_OTIC=EMPRESAS.ID_OTIC "
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=SOLICITUD.id_mutual "
query= query&" where SOLICITUD.id_solicitud ="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmSolicitudes" id="frmSolicitudes" action="" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="155">Rut Empresa :</td>
       <td width="200"><label id="txtRut" name="txtRut"><%=rsEmp("RUT")%></label></td>
       <td width="20" >&nbsp;</td>
       <td width="120">Raz&oacute;n Social :</td>
       <td width="300" colspan="4"><label id="txtR_social" name="txtR_social"><%=rsEmp("R_SOCIAL")%></label></td>
     </tr>
     <tr>
	   <td width="155">Otic :</td>
       <td width="200" colspan="4"><label id="txtOtic" name="txtOtic"><%=rsEmp("otic")%></label></td> 
      </tr>
     <tr>
       <td>Mutual :</td>
       <td  colspan="4"><label id="txtMutual" name="txtMutual"><%=rsEmp("nomb_mutual")%></label></td>
     </tr>
    <tr>
       <td>C&oacute;digo Curso:</td>
         <td><label id="txtCodCurso" name="txtCodCurso"><%=rsEmp("CODIGO")%></label></td>
         <td width="20" >&nbsp;</td>
       <td width="120">Nombre Curso :</td>
       <td width="300" colspan="4" rowspan="2"><label id="txtNombCurso" name="txtNombCurso"><%=rsEmp("NOMBRE_CURSO")%></label></td>
     </tr>
     <tr>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td> </td>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
     </tr>
     <tr>
       <td>Fecha Requerida :</td>
       <td><label id="txtFecha" name="txtFecha"><%=rsEmp("fecha")%></label></td>
       <td></td>
       <td>N&deg; Participantes :</td>
       <td><label id="txtNParticipantes" name="txtNParticipantes"><%=rsEmp("participantes")%></label></td>
     </tr>
     <tr>
       <td>&nbsp;</td>
       <td></td>
       <td></td>
       <td></td>
       <td></td>
     </tr>
   </table>
      <table cellspacing="0" cellpadding="0" border=1>
       <tr>
     
        <td width="90"><center><h3><em style="text-transform: capitalize;">Rut</em></h3></center></td>
        <td width="340"><center><h3><em style="text-transform: capitalize;">Nombre, Apellido Paterno, Apellido Materno</em></h3></center></td>
        <td width="200"><center><h3><em style="text-transform: capitalize;">Cargo En La Empresa</em></h3></center></td>
        <td width="250"><center><h3><em style="text-transform: capitalize;">Escolaridad</em></h3></center></td>
       </tr>
       <%
       	while not rsEmp.eof
			sol = "select SOLICITUD_TRABAJADOR.id_trabajador from SOLICITUD_TRABAJADOR "
			sol = sol&" inner join SOLICITUD on SOLICITUD.id_solicitud=SOLICITUD_TRABAJADOR.id_solicitud "
			sol = sol&" where SOLICITUD.id_solicitud='"&rsEmp("id_solicitud")&"'"
			sol = sol&" order by SOLICITUD_TRABAJADOR.id_trabajador asc "
		   
		   set rsSol =  conn.execute(sol)
		   
		  		while not rsSol.eof
					sql2 = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES,TRABAJADOR.CARGO_EMPRESA,TRABAJADOR.ESCOLARIDAD "
					sql2 = sql2&" from TRABAJADOR "
					sql2 = sql2&" where TRABAJADOR.ID_TRABAJADOR="&rsSol("id_trabajador")
					set rsTrab = conn.execute (sql2)
					
					
					
		if(rsTrab("ESCOLARIDAD")=0)then
		escolaridad="Sin Escolaridad"
		end if
		
		if(rsTrab("ESCOLARIDAD")=1)then
		escolaridad="Básica Incompleta"
		end if
		
		if(rsTrab("ESCOLARIDAD")=2)then
		escolaridad="Básica Completa"
		end if


		if(rsTrab("ESCOLARIDAD")=3)then
		escolaridad="Media Incompleta"
		end if

		if(rsTrab("ESCOLARIDAD")=4)then
		escolaridad="Media Completa"
		end if
		
		if(rsTrab("ESCOLARIDAD")=5)then
		escolaridad="Superior Técnica Incompleta"
		end if

		if(rsTrab("ESCOLARIDAD")=6)then
		escolaridad="Superior Técnica Profesional Completa"
		end if

		if(rsTrab("ESCOLARIDAD")=7)then
		escolaridad="Universitaria Incompleta"
		end if
		
		if(rsTrab("ESCOLARIDAD")=8)then
		escolaridad="Universitaria Completa"
		end if
		
					
					
			   %>
			   <tr>
				<td><%=rsTrab("RUT")%></td>
				<td><%=rsTrab("NOMBRES")%></td>
				<td><%=rsTrab("CARGO_EMPRESA")%></td>
				<td><%=escolaridad%></td>
			   </tr>
			   <%
					rsSol.Movenext
					wend
			rsEmp.Movenext
			wend
	   %>
      </table>
</form> 
<%
   end if
%>

