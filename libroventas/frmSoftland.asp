<%
fecha_soft=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)
%>
<form name="frmsoftland" id="frmsoftland" action="#" method="post">
    <table cellspacing="1" cellpadding="1" border=0>
    <tr>
    	<td colspan="5"><p>Seleccione rango de fechas:</p></td>
    </tr>
    <tr>
       <td width="60">Desde :</td>
       <td width="100"><input id="f_inicio_soft" name="f_inicio_soft" type="text" tabindex="1" size="10" value="<%=fecha_soft%>"/></td>
       <td width="20">&nbsp;</td>
       <td width="60">Hasta :</td>
       <td width="100"><input id="f_hasta_soft" name="f_hasta_soft" type="text" tabindex="2" size="10" value="<%=fecha_soft%>"/></td>
    </tr>
   </table>
</form> 