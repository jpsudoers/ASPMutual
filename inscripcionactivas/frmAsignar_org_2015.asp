<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim codSelec

if(Request("IdProg")="")then
codSelec=""
else
codSelec="and bloqProg.id_programa='"&Request("IdProg")&"'"
end if

dim query
query= "SELECT CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,PROGRAMA.ID_PROGRAMA, "
query= query&"CURRICULO.CODIGO,CURRICULO.SENCE,AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.N_PARTICIPANTES,"
query= query&"PROGRAMA.ID_MUTUAL as id_curso FROM AUTORIZACION "
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
query= query&" where AUTORIZACION.ID_AUTORIZACION="&Request("IdAuto")

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmAsignar" id="frmAsignar" action="#" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td width="120">C&oacute;digo Curso :</td>
       <td width="100"><%=rsEmp("CODIGO")%></td>
       <td width="20">&nbsp;</td>
       <td width="120">Nombre Curso :</td>
       <td width="467"><%=rsEmp("NOMBRE_CURSO")%></td>
    </tr>
    <tr>
       <td>Sence :</td>
       <%
	   checkedSi=""
	   checkedNo=""
	   
	   if(rsEmp("SENCE")=0)then
	   checkedSi="checked='checked'"
	   else
	   checkedNo="checked='checked'"
	   end if
	   %>
       <td colspan="4"><input type="radio" name="Sence_SI" id="Sence_SI" disabled value="0" <%=checkedSi%>/>
         Si
           <input type="radio" name="Sence_NO" id="Sence_NO" value="1" disabled <%=checkedNo%>/>
       No</td> 
    </tr>
     <tr>
       <td colspan="8">&nbsp;</td>
    </tr>
   </table>
	<div style="width:835px;"> 
	<table id="mytable" border="1" width="835"> 
    <thead> 
    	<tr> 
            <th>Fechas del Curso</th> 
            <th>Relator</th> 
            <th>Sala</th> 
            <th>Cupos</th>
            <th>Inscritos</th>  
            <th>&nbsp;Ver</th> 
            <th>&nbsp;Sel</th> 
        </tr> 
    </thead> 
    <tbody> 
    	  <%
				sql = "select bloqProg.id_relator,bloqProg.id_sede, "
	            sql = sql&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor, "
				sql = sql&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloqProg.nom_sede "
	  		 	sql = sql&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as sede," 
			    sql = sql&"bloqProg.id_bloque,bloqProg.cupos,bloqProg.estado,(select COUNT(*) from "
				sql = sql&"HISTORICO_CURSOS where HISTORICO_CURSOS.ID_BLOQUE=bloqProg.id_bloque) as incritos,"
				sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105)as FINICIO, "
				sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105)as FTERMINO,PROGRAMA.ID_PROGRAMA as programa "
				sql = sql&" from bloque_programacion bloqProg"
				sql = sql&" inner join SEDES on SEDES.ID_SEDE=bloqProg.id_sede "
				sql = sql&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloqProg.id_relator "
				sql = sql&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=bloqProg.id_programa "
				sql = sql&" where bloqProg.estado=1 and PROGRAMA.FECHA_INICIO_>=CONVERT(date,GETDATE(), 105) "&codSelec
				sql = sql&" and PROGRAMA.ID_MUTUAL='"&rsEmp("id_curso")&"'"
				sql = sql&" order by PROGRAMA.FECHA_INICIO_ asc "

			   set rsSol =  conn.execute(sql)
			   i=0

			   check="check"
			   totInsBloqProg=0
			   
			   while not rsSol.eof
			   i=i+1
			   
			   totInsBloqProg=cdbl(cdbl(rsSol("cupos"))-cdbl(rsSol("incritos")))
			   
			   		if(rsSol("estado")="1" and totInsBloqProg>0 and cdbl(rsEmp("N_PARTICIPANTES"))<=totInsBloqProg)then
		  %>
					  <tr>
                   		   <td><%=rsSol("FINICIO")&" - "&rsSol("FTERMINO")%></td>
                           <td><%=rsSol("instructor")%></td> 
                           <td><%=rsSol("sede")%></td> 
                           <td><%=rsSol("cupos")%></td>
                           <td><%=rsSol("incritos")%></td>
                           <td><span class="ui-state-valid" ><a href="#" title="Modificar Registro" class="ui-icon ui-icon-pencil" name="aContrato" onclick="MostrarInscritos(<%=rsSol("id_bloque")%>);"></a></span></td>
                 		   <td><input type="checkbox" id="<%=check&i%>" name="<%=check&i%>" value="<%=i%>" 
                           onclick="checkear(<%=i%>,<%=rsSol("id_bloque")%>,<%=rsSol("id_relator")%>,<%=rsSol("id_sede")%>,'<%=rsSol("instructor")%>','<%=rsSol("sede")%>',<%=rsSol("programa")%>,'<%=rsSol("FINICIO")%>','<%=rsSol("FTERMINO")%>');"></td>
					  </tr>
		  <%
		  			end if
		  	   rsSol.Movenext
			  wend
		  %>
    </tbody> 
    </table> 
     </div> 
      <table cellspacing="0" cellpadding="0" border=0>
          <tr>
           <td><input id="countFilas" name="countFilas" type="hidden" value="<%=i%>"/>
           <input id="SelBloque" name="SelBloque" type="hidden" value=""/>
           <input id="SelRelator" name="SelRelator" type="hidden"/>
           <input id="SelSala" name="SelSala" type="hidden"/>
           <input id="SelRelNom" name="SelRelNom" type="hidden"/>
           <input id="SelSalaNom" name="SelSalaNom" type="hidden"/>
           <input id="SelProg" name="SelProg" type="hidden"/>
           <input id="SelFechI" name="SelFechI" type="hidden"/>
           <input id="SelFechT" name="SelFechT" type="hidden"/>
           </td>
          </tr>
    </table>
</form> 
<%
end if
%>
<div id="messageBox1" style="height:50px;overflow:auto;width:200px;"> 
  	<ul></ul> 
</div> 