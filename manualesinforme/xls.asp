<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=Informes_Movimientos_"&fecha&".xls"
	
	dim totIngresos
	dim totEgresos
	dim totAjustesNeg
	dim totAjustesPos
	Dim totGeneral
	
	totIngresos=0
	totEgresos=0
	totAjustesNeg=0
	totAjustesPos=0
	totGeneral=0
	
	if(Request("tipo")="0")then	
	
	qsIng="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Mov_Articulo,CANTIDAD,"
	qsIng=qsIng&" dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_User "
	qsIng=qsIng&" FROM MOVIMIENTOS M "
	qsIng=qsIng&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO " 
	qsIng=qsIng&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsIng=qsIng&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO " 
	qsIng=qsIng&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsIng=qsIng&" WHERE M.TIPO_MOVIMIENTO IN (1,2) AND M.MODULO=1 AND M.ESTADO=1" 
	qsIng=qsIng&" ORDER BY M.FECHA DESC "

	set rsIng =  conn.execute(qsIng)%>

    <table width="1300" border="0">
        <tr>
     		  <td width="200" align="center">&nbsp;</td>
              <td width="130" align="center">&nbsp;</td>
              <td width="300" align="center">&nbsp;</td>
              <td width="50" align="center">&nbsp;</td>
              <td width="200" align="center">&nbsp;</td>
              <td width="200" align="center">&nbsp;</td>
              <td width="400" align="center">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="7" align="center"><b><font size="5">Informe de Movimientos</font></b></td>
        </tr>
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="7"><b><font size="4">Ingresos</font></b></td>
        </tr>
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
	</table>
    <table width="1300" border="1">
    <tr>
      <td width="200" align="center"><b><font size="2">Fecha</font></b></td>
      <td width="130" align="center"><b><font size="2">Tipo de Movimiento</font></b></td>
      <td width="300" align="center"><b><font size="2">Articulo</font></b></td>
      <td width="50" align="center"><b><font size="2">Cant.</font></b></td>
      <td width="200" align="center"><b><font size="2">Razones</font></b></td> 
      <td width="200"align="center"><b><font size="2">Fecha / Relator</font></b></td>      
      <td width="400" align="center"><b><font size="2">Usuario</font></b></td>
    </tr>
	<%
    while not rsIng.eof
	totIngresos=totIngresos+cdbl(rsIng("CANTIDAD"))
    %>
    <tr>
      <td  align="center"><b><font size="2"><%=rsIng("FECHA_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsIng("NOMBRE_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsIng("Mov_Articulo")%></font></b></td>
      <td align="right"><b><font size="2"><%=FormatNumber(rsIng("CANTIDAD"),0)%></font></b></td>
      <td>&nbsp;</td>  
      <td>&nbsp;</td>            
      <td><b><font size="2"><%=rsIng("Mov_User")%></font></b></td>
    </tr>	
    <%
        rsIng.MoveNext
    wend
    %>
    </table>
  
    <table width="1300" border="0">
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
        <tr>
			<td width="200">&nbsp;</td>
      		<td width="130">&nbsp;</td>
      		<td width="300">&nbsp;</td>
      		<td width="50">&nbsp;</td>
     		<td width="200">&nbsp;</td>
      		<td width="200">&nbsp;</td>     
	        <td width="400"><b><font size="3">Total Ingresos : <%=FormatNumber(totIngresos,0)%></font></b></td>
        </tr> 
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>       
        <tr>
          <td colspan="7"><b><font size="4">Salidas</font></b></td>
        </tr>
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
	</table>
    
    <%
 	qsEgres="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Nom_Articulo,CANTIDAD, "
	qsEgres=qsEgres&"dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_Nom, "
	qsEgres=qsEgres&"CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105)+', '+"
	qsEgres=qsEgres&"dbo.MayMinTexto(IR.NOMBRES+' '+IR.A_PATERNO+' '+IR.A_MATERNO) AS FECHA_NOM_RELATOR "
	qsEgres=qsEgres&" FROM MOVIMIENTOS M " 
	qsEgres=qsEgres&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO "
	qsEgres=qsEgres&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsEgres=qsEgres&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO " 
	qsEgres=qsEgres&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsEgres=qsEgres&" inner join bloque_programacion BP ON BP.id_bloque=M.ID_PROG_BLOQUE "
	qsEgres=qsEgres&" INNER JOIN PROGRAMA P ON P.ID_PROGRAMA=BP.id_programa "
	qsEgres=qsEgres&" INNER JOIN INSTRUCTOR_RELATOR IR ON IR.ID_INSTRUCTOR=BP.id_relator "
	qsEgres=qsEgres&" WHERE M.TIPO_MOVIMIENTO IN (3) AND M.MODULO=2 AND M.ESTADO=1" 
	qsEgres=qsEgres&" ORDER BY M.FECHA DESC "

	set rsEgres =  conn.execute(qsEgres)%>
    
    <table width="1300" border="1">
    <tr>
      <td width="200" align="center"><b><font size="2">Fecha</font></b></td>
      <td width="130" align="center"><b><font size="2">Tipo de Movimiento</font></b></td>
      <td width="300" align="center"><b><font size="2">Articulo</font></b></td>
      <td width="50" align="center"><b><font size="2">Cant.</font></b></td>
      <td width="200" align="center"><b><font size="2">Razones</font></b></td> 
      <td width="200"align="center"><b><font size="2">Fecha / Relator</font></b></td>      
      <td width="400" align="center"><b><font size="2">Usuario</font></b></td>
    </tr>
	<%
    while not rsEgres.eof
	totEgresos=totEgresos+cdbl(rsEgres("CANTIDAD"))
    %>
    <tr>
      <td align="center"><b><font size="2"><%=rsEgres("FECHA_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsEgres("NOMBRE_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsEgres("Nom_Articulo")%></font></b></td>      
      <td align="right"><b><font size="2"><%=FormatNumber(rsEgres("CANTIDAD"),0)%></font></b></td>
      <td>&nbsp;</td>
      <td><b><font size="2"><%=rsEgres("FECHA_NOM_RELATOR")%></font></b></td>
      <td><b><font size="2"><%=rsEgres("Mov_Nom")%></font></b></td>      
    </tr>	
    <%
        rsEgres.MoveNext
    wend
    %>
    </table>
   
    <table width="1300" border="0">
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
        <tr>
			<td width="200">&nbsp;</td>
      		<td width="130">&nbsp;</td>
      		<td width="300">&nbsp;</td>
      		<td width="50">&nbsp;</td>
     		<td width="200">&nbsp;</td>
      		<td width="200">&nbsp;</td> 
            <td width="400"><b><font size="3">Total Salidas: <%=FormatNumber(totEgresos,0)%></font></b></td>
        </tr> 
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>  
        <tr>
          <td colspan="7"><b><font size="4">Ajustes</font></b></td>
        </tr>
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
	</table>
    
    <%
	qsAjuste="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Nom_Articulo,CANTIDAD, "
	qsAjuste=qsAjuste&"M.EXPLICACION,dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_Nom,M.TIPO_AJUSTE " 
	qsAjuste=qsAjuste&" FROM MOVIMIENTOS M " 
	qsAjuste=qsAjuste&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO " 
	qsAjuste=qsAjuste&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsAjuste=qsAjuste&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO "
	qsAjuste=qsAjuste&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsAjuste=qsAjuste&" WHERE M.TIPO_MOVIMIENTO IN (4,5,6) AND M.MODULO=3 AND M.ESTADO=1" 
	qsAjuste=qsAjuste&" ORDER BY M.FECHA DESC " 

	set rsAjuste =  conn.execute(qsAjuste)%>
    
    <table width="1300" border="1">
    <tr>
      <td width="200" align="center"><b><font size="2">Fecha</font></b></td>
      <td width="130" align="center"><b><font size="2">Tipo de Movimiento</font></b></td>
      <td width="300" align="center"><b><font size="2">Articulo</font></b></td>
      <td width="50" align="center"><b><font size="2">Cant.</font></b></td>
      <td width="200" align="center"><b><font size="2">Razones</font></b></td> 
      <td width="200"align="center"><b><font size="2">Fecha / Relator</font></b></td>      
      <td width="400" align="center"><b><font size="2">Usuario</font></b></td>
    </tr>
	<%
    while not rsAjuste.eof
		if(rsAjuste("TIPO_AJUSTE")="2")then
			totAjustesNeg=totAjustesNeg+cdbl(rsAjuste("CANTIDAD"))
		else
			totAjustesPos=totAjustesPos+cdbl(rsAjuste("CANTIDAD"))
		end if
    %>
    <tr>
      <td align="center"><b><font size="2"><%=rsAjuste("FECHA_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsAjuste("NOMBRE_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsAjuste("Nom_Articulo")%></font></b></td>
      <td align="right"><b><font size="2"><%=FormatNumber(rsAjuste("CANTIDAD"),0)%></font></b></td>
	  <td><b><font size="2"><%=rsAjuste("EXPLICACION")%></font></b></td>      
	  <td>&nbsp;</td>        
      <td><b><font size="2"><%=rsAjuste("Mov_Nom")%></font></b></td>
    </tr>	
    <%
        rsAjuste.MoveNext
    wend
    %>
</table>
    <table width="1300" border="0">
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
        <tr>
			<td width="200">&nbsp;</td>
      		<td width="130">&nbsp;</td>
      		<td width="300">&nbsp;</td>
      		<td width="50">&nbsp;</td>
     		<td width="200">&nbsp;</td>
      		<td width="200">&nbsp;</td> 
            <td width="400"><b><font size="3">Total Ajustes Negativos : <%=FormatNumber(totAjustesNeg,0)%></font></b></td>
        </tr> 
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>          
          <td><b><font size="3">Total Ajustes Positivos : <%=FormatNumber(totAjustesPos,0)%></font></b></td>
        </tr> 
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>        
        <tr>
          <td colspan="7">&nbsp;</td>
          <%
			'totGeneral=cdbl(cdbl(cdbl(totIngresos)+cdbl(totAjustesPos))-cdbl(cdbl(totEgresos)+cdbl(totAjustesNeg)))
			'<td width="400"><b><font size="3">Total General : <%=FormatNumber(totGeneral,0)</font></b></td>
		  %>
        </tr>         
	</table>
    <%else%>
		
    <table width="1300" border="0">
        <tr>
     		  <td width="200" align="center">&nbsp;</td>
              <td width="130" align="center">&nbsp;</td>
              <td width="300" align="center">&nbsp;</td>
              <td width="50" align="center">&nbsp;</td>
              <td width="200" align="center">&nbsp;</td>
              <td width="200" align="center">&nbsp;</td>
              <td width="400" align="center">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="7" align="center"><b><font size="5">Informe de Movimientos</font></b></td>
        </tr>
    </table>
    <%
	if(Request("tipo")="1")then	
	qsIng="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Mov_Articulo,CANTIDAD,"
	qsIng=qsIng&" dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_User "
	qsIng=qsIng&" FROM MOVIMIENTOS M "
	qsIng=qsIng&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO " 
	qsIng=qsIng&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsIng=qsIng&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO " 
	qsIng=qsIng&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsIng=qsIng&" WHERE M.TIPO_MOVIMIENTO IN (1,2) AND M.MODULO=1 " 
	qsIng=qsIng&" ORDER BY M.FECHA DESC "

	set rsIng =  conn.execute(qsIng)%>
    
    <table width="1300" border="0">
        <tr>
          <td colspan="5">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="5"><b><font size="4">Ingresos</font></b></td>
        </tr>
        <tr>
          <td colspan="5">&nbsp;</td>
        </tr>
	</table>
   
    <table width="1300" border="1">
    <tr>
      <td width="200" align="center"><b><font size="2">Fecha</font></b></td>
      <td width="130" align="center"><b><font size="2">Tipo de Movimiento</font></b></td>
      <td width="300" align="center"><b><font size="2">Articulo</font></b></td>
      <td width="50" align="center"><b><font size="2">Cant.</font></b></td>
      <td width="200" align="center"><b><font size="2">Razones</font></b></td> 
      <td width="200"align="center"><b><font size="2">Fecha / Relator</font></b></td>      
      <td width="400" align="center"><b><font size="2">Usuario</font></b></td>
    </tr>
	<%
    while not rsIng.eof
	totIngresos=totIngresos+cdbl(rsIng("CANTIDAD"))
    %>
    <tr>
      <td  align="center"><b><font size="2"><%=rsIng("FECHA_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsIng("NOMBRE_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsIng("Mov_Articulo")%></font></b></td>
      <td align="right"><b><font size="2"><%=FormatNumber(rsIng("CANTIDAD"),0)%></font></b></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>            
      <td><b><font size="2"><%=rsIng("Mov_User")%></font></b></td>
    </tr>	
    <%
        rsIng.MoveNext
    wend
    %>
    </table>
      <table width="1300" border="0">
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
        <tr>
			<td width="200">&nbsp;</td>
      		<td width="130">&nbsp;</td>
      		<td width="300">&nbsp;</td>
      		<td width="50">&nbsp;</td>
     		<td width="200">&nbsp;</td>
      		<td width="200">&nbsp;</td> 
            <td width="400"><b><font size="3">Total Ingresos : <%=FormatNumber(totIngresos,0)%></font></b></td>
        </tr> 
       </table>
        
        <%
		end if
		
		if(Request("tipo")="2")then	
		%>
        <table width="1300" border="0">
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>       
        <tr>
          <td colspan="7"><b><font size="4">Salidas</font></b></td>
        </tr>
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
	  </table>
    
    <%
 	qsEgres="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Nom_Articulo,CANTIDAD, "
	qsEgres=qsEgres&"dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_Nom, "
	qsEgres=qsEgres&"CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105)+', '+"
	qsEgres=qsEgres&"dbo.MayMinTexto(IR.NOMBRES+' '+IR.A_PATERNO+' '+IR.A_MATERNO) AS FECHA_NOM_RELATOR "
	qsEgres=qsEgres&" FROM MOVIMIENTOS M " 
	qsEgres=qsEgres&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO "
	qsEgres=qsEgres&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsEgres=qsEgres&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO " 
	qsEgres=qsEgres&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsEgres=qsEgres&" inner join bloque_programacion BP ON BP.id_bloque=M.ID_PROG_BLOQUE "
	qsEgres=qsEgres&" INNER JOIN PROGRAMA P ON P.ID_PROGRAMA=BP.id_programa "
	qsEgres=qsEgres&" INNER JOIN INSTRUCTOR_RELATOR IR ON IR.ID_INSTRUCTOR=BP.id_relator "
	qsEgres=qsEgres&" WHERE M.TIPO_MOVIMIENTO IN (3) AND M.MODULO=2 " 
	qsEgres=qsEgres&" ORDER BY M.FECHA DESC "

	set rsEgres =  conn.execute(qsEgres)%>
    
    <table width="1300" border="1">
    <tr>
      <td width="200" align="center"><b><font size="2">Fecha</font></b></td>
      <td width="130" align="center"><b><font size="2">Tipo de Movimiento</font></b></td>
      <td width="300" align="center"><b><font size="2">Articulo</font></b></td>
      <td width="50" align="center"><b><font size="2">Cant.</font></b></td>
      <td width="200" align="center"><b><font size="2">Razones</font></b></td> 
      <td width="200"align="center"><b><font size="2">Fecha / Relator</font></b></td>      
      <td width="400" align="center"><b><font size="2">Usuario</font></b></td>
    </tr>
	<%
    while not rsEgres.eof
	totEgresos=totEgresos+cdbl(rsEgres("CANTIDAD"))
    %>
    <tr>
      <td align="center"><b><font size="2"><%=rsEgres("FECHA_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsEgres("NOMBRE_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsEgres("Nom_Articulo")%></font></b></td>
      <td align="right"><b><font size="2"><%=FormatNumber(rsEgres("CANTIDAD"),0)%></font></b></td>
      <td>&nbsp;</td>      
      <td><b><font size="2"><%=rsEgres("FECHA_NOM_RELATOR")%></font></b></td>
      <td><b><font size="2"><%=rsEgres("Mov_Nom")%></font></b></td>      
    </tr>	
    <%
        rsEgres.MoveNext
    wend
    %>
    </table>
   
    <table width="1300" border="0">
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
        <tr>
			<td width="200">&nbsp;</td>
      		<td width="130">&nbsp;</td>
      		<td width="300">&nbsp;</td>
      		<td width="50">&nbsp;</td>
     		<td width="200">&nbsp;</td>
      		<td width="200">&nbsp;</td> 
            <td width="400"><b><font size="3">Total Salidas: <%=FormatNumber(totEgresos,0)%></font></b></td>
        </tr> 
     </table>
      <%
		end if
	
		if(Request("tipo")="3")then	
      %>
        <table width="1300" border="0">
        <tr>
          <td colspan="5">&nbsp;</td>
        </tr>  
        <tr>
          <td colspan="5"><b><font size="4">Ajustes</font></b></td>
        </tr>
        <tr>
          <td colspan="5">&nbsp;</td>
        </tr>
	</table>
    
    <%
	qsAjuste="SELECT CONVERT(VARCHAR(10),M.FECHA, 105) as FECHA_MOV,TP.NOMBRE_MOV,A.DESC_ARTICULO as Nom_Articulo,CANTIDAD, "
	qsAjuste=qsAjuste&"M.EXPLICACION,dbo.MayMinTexto(U.NOMBRES+' '+U.A_PATERNO+' '+U.A_MATERNO) as Mov_Nom,M.TIPO_AJUSTE " 
	qsAjuste=qsAjuste&" FROM MOVIMIENTOS M " 
	qsAjuste=qsAjuste&" inner join ARTICULOS A on A.ID_ARTICULO=M.ID_ARTICULO " 
	qsAjuste=qsAjuste&" inner join BODEGAS B on B.ID_BODEGA=M.ID_BODEGA " 
	qsAjuste=qsAjuste&" inner JOIN TIPO_MOVIMIENTO TP on TP.ID_TIPO_MOV=M.TIPO_MOVIMIENTO "
	qsAjuste=qsAjuste&" inner join USUARIOS U ON U.ID_USUARIO=M.ID_USUARIO "
	qsAjuste=qsAjuste&" WHERE M.TIPO_MOVIMIENTO IN (4,5,6) AND M.MODULO=3 " 
	qsAjuste=qsAjuste&" ORDER BY M.FECHA DESC " 

	set rsAjuste =  conn.execute(qsAjuste)%>
    
    <table width="1300" border="1">
    <tr>
      <td width="200" align="center"><b><font size="2">Fecha</font></b></td>
      <td width="130" align="center"><b><font size="2">Tipo de Movimiento</font></b></td>
      <td width="300" align="center"><b><font size="2">Articulo</font></b></td>
      <td width="50" align="center"><b><font size="2">Cant.</font></b></td>
      <td width="200" align="center"><b><font size="2">Razones</font></b></td> 
      <td width="200"align="center"><b><font size="2">Fecha / Relator</font></b></td>      
      <td width="400" align="center"><b><font size="2">Usuario</font></b></td>
    </tr>
	<%
    while not rsAjuste.eof
		if(rsAjuste("TIPO_AJUSTE")="2")then
			totAjustesNeg=totAjustesNeg+cdbl(rsAjuste("CANTIDAD"))
		else
			totAjustesPos=totAjustesPos+cdbl(rsAjuste("CANTIDAD"))
		end if
    %>
    <tr>
      <td align="center"><b><font size="2"><%=rsAjuste("FECHA_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsAjuste("NOMBRE_MOV")%></font></b></td>
      <td><b><font size="2"><%=rsAjuste("Nom_Articulo")%></font></b></td>
      <td align="right"><b><font size="2"><%=FormatNumber(rsAjuste("CANTIDAD"),0)%></font></b></td>
	  <td><b><font size="2"><%=rsAjuste("EXPLICACION")%></font></b></td>     
      <td>&nbsp;</td>       
      <td><b><font size="2"><%=rsAjuste("Mov_Nom")%></font></b></td>
    </tr>	
    <%
        rsAjuste.MoveNext
    wend
    %>
</table>
    <table width="1300" border="0">
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>
        <tr>
			<td width="200">&nbsp;</td>
      		<td width="130">&nbsp;</td>
      		<td width="300">&nbsp;</td>
      		<td width="50">&nbsp;</td>
     		<td width="200">&nbsp;</td>
      		<td width="200">&nbsp;</td> 
            <td width="400"><b><font size="3">Total Ajustes Negativos : <%=FormatNumber(totAjustesNeg,0)%></font></b></td>
        </tr> 
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td><b><font size="3">Total Ajustes Positivos : <%=FormatNumber(totAjustesPos,0)%></font></b></td>
        </tr> 
        <tr>
          <td colspan="7">&nbsp;</td>
        </tr>   
        </table>
		<%
        end if
     end if%>