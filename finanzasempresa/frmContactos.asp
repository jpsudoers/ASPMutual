<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim sql
sql="SELECT E.RUT,UPPER(E.R_SOCIAL) as R_SOCIAL,UPPER(E.GIRO) as GIRO,UPPER(E.DIRECCION) as DIRECCION,"
sql=sql&"UPPER(E.COMUNA) as COMUNA,UPPER(E.CIUDAD) as CIUDAD,E.FONO,E.FAX,UPPER(E.NOMBRES) as NOMBRES,"
sql=sql&"UPPER(E.EMAIL) as EMAIL,UPPER(E.CARGO) as CARGO,UPPER(E.NOMBRE_CONTA) as NOMBRE_CONTA,"
sql=sql&"UPPER(E.CARGO_CONTA) as CARGO_CONTA,UPPER(E.EMAIL_CONTA) as EMAIL_CONTA,E.FONO_CONTACTO,E.FONO_CONTABILIDAD,"
sql=sql&"(select UPPER(E2.R_SOCIAL) from EMPRESAS E2 where E2.ID_EMPRESA=E.MUTUAL) as Nom_Mutual,"  
sql=sql&"(select UPPER(E3.R_SOCIAL) from EMPRESAS E3 where E3.ID_EMPRESA=E.ID_OTIC) as Nom_OTIC" 
sql=sql&" FROM EMPRESAS E where E.ID_EMPRESA='"&Request("id")&"'"

set rsEmp = conn.execute (sql)

%>
<form name="frmEmpresa" id="frmEmpresa" action="#" method="post">
<table cellspacing="3" cellpadding="1" border=0 width="960">
	<tr>
    	<td>Rut :</td>
      	<td><%=rsEmp("rut")%></td>
        <td>&nbsp;</td>
        <td>Raz&oacute;n Social :</td>
        <td colspan="4"><%=rsEmp("r_social")%></td>
   	</tr>
	<tr>
    	<td>Giro:</td>
        <td colspan="7"><%=rsEmp("giro")%></td>
    </tr>
	<tr>
        <td width="70">Direcci&oacute;n :</td>
        <td width="300"><%=rsEmp("direccion")%></td>
        <td width="20">&nbsp;</td>
        <td width="110">Comuna :</td>
        <td width="200"><%=rsEmp("comuna")%></td>
        <td width="20">&nbsp;</td>
        <td width="80">Ciudad :</td>
        <td width="160"><%=rsEmp("ciudad")%></td>
    </tr>
	<tr>
        <td>Tel&eacute;fono :</td>
        <td><%=rsEmp("fono")%></td>
        <td></td>
        <td>Fax :</td>
        <td colspan="4"><%=rsEmp("fax")%></td>
    </tr>
	<tr>
    	<td>Mutual :</td>
        <%if isnull(rsEmp("Nom_Mutual"))THEN%>
        	<td>SIN MUTUAL</td>
		<%else%>
        	<td><%=rsEmp("Nom_Mutual")%></td>      
        <%end if%>
        <td></td>
        <td>Otic :</td>
        <%if isnull(rsEmp("Nom_OTIC"))THEN%>
        	<td colspan="4">SIN OTIC</td>
		<%else%>
        	<td colspan="4"><%=rsEmp("Nom_OTIC")%></td>        
        <%end if%>
    </tr>
    <tr>
    	<td colspan="8">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="8">
      		<table id="list2"></table> 
	      	<div id="pager2"></div>         
      </td> 
     </tr>    
</table>
</form> 