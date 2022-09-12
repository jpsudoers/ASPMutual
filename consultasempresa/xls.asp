<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=Bloqueos_"&fecha&".xls"
		
	qsHc="select DISTINCT e.RUT,Empresa=upper(e.R_SOCIAL),USR=(u.NOMBRES+' '+u.A_PATERNO+' '+u.A_MATERNO), "
	qsHc=qsHc&"case when EB.ESTADO=0 then 'Bloqueada' else 'Desbloqueada' end 'ESTADO_DESC',EB.FECHA_ESTADO,EB.DESCRIPCION "
	qsHc=qsHc&"from EMPRESAS_BLOQUEOS eb inner join EMPRESAS e on e.ID_EMPRESA=eb.ID_EMPRESA "
	qsHc=qsHc&"inner join USUARIOS u on u.ID_USUARIO=eb.ID_USUARIO WHERE EB.ID_BLOQUEO_EMPRESA IN (SELECT MAX(EB2.ID_BLOQUEO_EMPRESA) "
	qsHc=qsHc&"FROM EMPRESAS_BLOQUEOS eb2 WHERE EB2.ID_EMPRESA=e.ID_EMPRESA) and e.ESTADO=1 and e.TIPO=1  "
	qsHc=qsHc&"order by eb.FECHA_ESTADO "

	set rsHc =  conn.execute(qsHc)%>

    <table border="1">
    <tr>
      <td align="center"><b><font size="2">Rut</font></b></td>
      <td align="center"><b><font size="2">Raz&oacute;n Social</font></b></td>      
      <td align="center"><b><font size="2">Usuario</font></b></td>   
      <td align="center"><b><font size="2">Estado Empresa</font></b></td>
      <td align="center"><b><font size="2">Fecha Estado</font></b></td>
      <td align="center"><b><font size="2">Descripci&oacute;n</font></b></td>      
    </tr>
	<%
    while not rsHc.eof
    %>
    <tr>
      <td><font size="2"><%=rsHc("RUT")%></font></td>
      <td><font size="2"><%=rsHc("Empresa")%></font></td>     
      <td><font size="2"><%=rsHc("USR")%></font></td>                  
      <td><font size="2"><%=rsHc("ESTADO_DESC")%></font></td> 
      <td><font size="2"><%=rsHc("FECHA_ESTADO")%></font></td>
      <td><font size="2"><%=rsHc("DESCRIPCION")%></font></td>
    </tr>	
    <%
        rsHc.MoveNext
    wend
    %>
    </table>
