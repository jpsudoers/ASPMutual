<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vautorizacion = Request("id")

dim query
query= "SELECT CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, "
query= query&"EMPRESAS.R_SOCIAL,EMPRESAS.RUT,PROGRAMA.ID_PROGRAMA,EMPRESAS.ID_EMPRESA, "
query= query&"CURRICULO.CODIGO,CURRICULO.SENCE, "
query= query&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor,"
query= query&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloque_programacion.nom_sede "
query= query&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+' - 3er Piso, '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as sede,"
query= query&"AUTORIZACION.CON_OTIC,AUTORIZACION.CON_FRANQUICIA,AUTORIZACION.N_REG_SENCE,AUTORIZACION.ID_OTIC,"
query= query&"(case AUTORIZACION.CON_OTIC when 1 then 'Si' when 0 then 'No' end) as 'estConOtic',"
query= query&"(case AUTORIZACION.CON_FRANQUICIA when 1 then 'Si' when 0 then 'No' end) as 'estConFran'"
query= query&" FROM AUTORIZACION " 
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA " 
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=AUTORIZACION .id_empresa " 
query= query&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
query= query&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
query= query&" inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede "
query= query&" where AUTORIZACION.ID_AUTORIZACION="&vautorizacion

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frminscripcion" id="frminscripcion" method="post">
    <table cellspacing="0" cellpadding="1" border=0>
      <tr>
       <td>Rut Empresa : <%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td>
       <td><input type="hidden" id="Empresa" name="Empresa" value="<%=rsEmp("ID_EMPRESA")%>"/></td>
       <td>Raz&oacute;n Social : </td>
       <td colspan="4"><%=rsEmp("R_SOCIAL")%></td>
    </tr>
    <tr>
       <td>C&oacute;digo Curso : <%=rsEmp("CODIGO")%></td>
       <td><input type="hidden" id="Programa" name="Programa" value="<%=rsEmp("ID_PROGRAMA")%>"/>
       				  <input type="hidden" id="autorizacionId" name="autorizacionId" value="<%=vautorizacion%>"/></td>
       <td>Nombre Curso :</td>
       <td colspan="4"><%=rsEmp("NOMBRE_CURSO")%></td>
    </tr>
    <tr>
     <td>Fecha Inicio : <%=rsEmp("FECHA_INICIO_")%></td>
       <td>&nbsp;</td>
       <td>Fecha T&eacute;rmino : </td>
       <td><%=rsEmp("FECHA_TERMINO")%></td>
    </tr>
     <tr>
       <td width="220">Utiliza Franquicia Sence? : <%=rsEmp("estConFran")%></td>
       <td width="20">&nbsp;</td>
       <td width="120">Con OTIC : </td>
       <td width="172"><%=rsEmp("estConOtic")%></td>
       <td width="20">&nbsp;</td>
       <td width="150">N&deg; de Registro Sence : </td>
       <td width="243"><%=rsEmp("n_reg_sence")%></td>
    </tr>      
	<%
		 if(rsEmp("id_otic")<>"0" and rsEmp("con_otic")="1" and rsEmp("con_franquicia")="1")then
			 set rsOtic = conn.execute ("select ID_EMPRESA,RUT,R_SOCIAL from EMPRESAS where EMPRESAS.ID_EMPRESA="&rsEmp("id_otic"))
			 %>
               <tr name="Mandante" id="Mandante" style.display ="">
                   <td>Rut OTIC : <%=replace(FormatNumber(mid(rsOtic("RUT"), 1,len(rsOtic("RUT"))-2),0)&mid(rsOtic("RUT"), len(rsOtic("RUT"))-1,len(rsOtic("RUT"))),",",".")%></td>
                   <td></td>
                   <td>Raz&oacute;n Social :</td>
                   <td colspan="4"><label id="txtRsocOtic" name="txtRsocOtic"><%=rsOtic("R_SOCIAL")%></label></td>
               </tr>
    <%
	    end if
	%>    
     <tr>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
     </tr>
   </table>
</form> 
<div width="200">
<table id="list2"></table> 
<div id="pager2"></div> 
</div>
 <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td width="70">&nbsp;</td>
       <td width="320">&nbsp;</td>
       <td width="50">&nbsp;</td>
       <td width="340">&nbsp;</td>
     </tr>
      <tr>
       <td>Relator : </td>
       <td><%=rsEmp("instructor")%></td>
       <td>Sede : </td>
       <td><%=rsEmp("sede")%></td>
     </tr>
     <tr>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
     </tr>
   </table>
<%
end if
%>
