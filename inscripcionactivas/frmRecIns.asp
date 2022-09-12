<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<div> 
	<table id="tablaRechIns" border="1"> 
    <thead> 
    	<tr> 
            <th width="140">Fecha</th> 
            <th width="150">Usuario</th> 
            <th>Rechazos</th> 
        </tr> 
    </thead> 
    <tbody> 
    	  <%
				dim query
				query="select FECHA_INGRESO=convert(datetime, VI.FECHA_INGRESO),"
				query=query&"NOMBRES=(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO),RAZONES=UPPER(VI.RAZONES) "
				query=query&" from verifica_inscripcion VI " 
				query=query&" inner join usuarios u on u.ID_USUARIO=VI.ID_USUARIO where VI.ID_AUTORIZACION="&Request("Id")
				query=query&" order by VI.FECHA_INGRESO desc"
				
				set rsEmp = conn.execute (query)
			   
			   while not rsEmp.eof
		       %>
					  <tr>
                   		   <td><%=rsEmp("FECHA_INGRESO")%></td>
                           <td><%=rsEmp("NOMBRES")%></td> 
                           <td><%=rsEmp("RAZONES")%></td> 
					  </tr>
		      <%
		  	   rsEmp.Movenext
			  wend
		      %>
    </tbody> 
    </table> 
</div> 