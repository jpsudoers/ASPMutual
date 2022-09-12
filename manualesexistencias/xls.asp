<!--#include file="../conexion.asp"-->
<%
	dim mes(12) 
	mes(0)="Enero"
	mes(1)="Febrero"
	mes(2)="Marzo"
	mes(3)="Abril"
	mes(4)="Mayo"
	mes(5)="Junio"
	mes(6)="Julio"
	mes(7)="Agosto"
	mes(8)="Septiembre"
	mes(9)="Octubre"
	mes(10)="Noviembre"
	mes(11)="Diciembre"

	fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
	fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
	Response.ContentType ="application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "inline; filename=Informes_Movimientos_"&fecha&".xls"

	qsMan="select a.DESC_ARTICULO as 'Nombre Articulo',isnull((SELECT SUM(CANTIDAD) "
	qsMan=qsMan&" FROM MOVIMIENTOS m "
	qsMan=qsMan&" inner join bloque_programacion bq on bq.id_bloque=m.ID_PROG_BLOQUE "
	qsMan=qsMan&" inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "
	qsMan=qsMan&" WHERE (m.TIPO_MOVIMIENTO = 3) and m.ID_ARTICULO=a.ID_ARTICULO "
	qsMan=qsMan&" and CONVERT(date,p.FECHA_INICIO_)>=CONVERT(date,'"&request("f_ini")&"') "
	qsMan=qsMan&" and CONVERT(date,p.FECHA_INICIO_)<=CONVERT(date,'"&request("f_fin")&"')),0) as 'total_hasta', "
	qsMan=qsMan&" isnull((SELECT SUM(CANTIDAD) FROM MOVIMIENTOS "
	qsMan=qsMan&" WHERE (TIPO_MOVIMIENTO = 1) and ID_ARTICULO=a.ID_ARTICULO),0)+ "
	qsMan=qsMan&" isnull((SELECT SUM(CANTIDAD) FROM MOVIMIENTOS "
	qsMan=qsMan&" WHERE (TIPO_MOVIMIENTO = 3) and ID_ARTICULO=a.ID_ARTICULO),0) as 'Stock Actual' from ARTICULOS a WHERE A.ESTADO_ARTICULO=1"

	set rsMan =  conn.execute(qsMan)%>

    <table width="800" border="0">
        <tr>
     	   <td colspan="3">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="3"><b><font size="5">Informe de Existencias</font></b></td>
        </tr>
        <tr>
          <td colspan="3">&nbsp;</td>
        </tr>
        <tr>
          <td width="500" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2">Nombre Articulo</font></b></td>
          <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%="Salidas desde el "&mid(request("f_ini"),1,2)&" de "&mes(mid(request("f_ini"),4,2)-1)&" de "&mid(request("f_ini"),7,4)&" Hasta el "&mid(request("f_fin"),1,2)&" de "&mes(mid(request("f_fin"),4,2)-1)&" de "&mid(request("f_fin"),7,4)%></font></b></td>
          <td width="150" align="center" style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%="Stock Actual (Hasta el "&right("0"&day(now()),2)&" de "&mes(month(now())-1)&" de "&year(now)&")"%></font></b></td>
        </tr>
        <%
        while not rsMan.eof
        %>
        <tr>
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=rsMan("Nombre Articulo")%></font></b></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=replace(FormatNumber(rsMan("total_hasta"),0),",",".")%></font></b></td>
          <td style="border-width: 1px;border: solid; border-color:#000;"><b><font size="2"><%=replace(FormatNumber(rsMan("Stock Actual"),0),",",".")%></font></b></td>
        </tr>	
        <%
            rsMan.MoveNext
        wend
        %>
    </table>
