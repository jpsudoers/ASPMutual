<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query= "SELECT CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO,PROGRAMA.ID_PROGRAMA, "
query= query&"CURRICULO.CODIGO,CURRICULO.SENCE,AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.N_PARTICIPANTES,"
query= query&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor,"
query= query&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloque_programacion.nom_sede "
query= query&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as sede,"
query= query&"PROGRAMA.ID_MUTUAL as id_curso,EMPRESAS.R_SOCIAL,EMPRESAS.RUT,(CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra'"
query= query&" WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista ' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque ' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia ' "
query= query&" WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso ' " 
query= query&" END) as 'Tipo Documento',AUTORIZACION.ORDEN_COMPRA FROM AUTORIZACION "
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=AUTORIZACION.id_empresa "
query= query&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "
query= query&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
query= query&" inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede " 
query= query&" where AUTORIZACION.ID_AUTORIZACION="&Request("IdAuto")

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmSepIns" id="frmSepIns" action="inscripcionactivas/USSepUnir.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td colspan="5"><b>Inscripción Mandante :</b></td>
    </tr>    
    <tr>
       <td colspan="5">&nbsp;</td>
    </tr>  
    <tr>
       <td>Rut Empresa :</td>
       <td><%=rsEmp("RUT")%></td> 
       <td>&nbsp;</td> 
       <td>Razón Social :</td> 
       <td><%=rsEmp("R_SOCIAL")%></td> 
    </tr>       
    <tr>
       <td width="95">C&oacute;digo Curso :</td>
       <td width="140"><%=rsEmp("CODIGO")%></td>
       <td width="20">&nbsp;</td>
       <td width="110">Nombre Curso :</td>
       <td width="462"><%=rsEmp("NOMBRE_CURSO")%></td>
    </tr>
    <tr>
       <td>Fecha Inicio :</td>
       <td><%=rsEmp("FECHA_INICIO_")%></td> 
       <td>&nbsp;</td> 
       <td>Fecha Termino :</td> 
       <td><%=rsEmp("FECHA_TERMINO")%></td> 
    </tr>    
    <tr>
       <td colspan="2">Relator :&nbsp;&nbsp; <%=rsEmp("instructor")%></td> 
       <td>&nbsp;</td> 
       <td colspan="2">
       	  <table width="467"  cellspacing="0" cellpadding="1" border=0>
          <tr>
            <td width="337">Sede :&nbsp;&nbsp; <%=rsEmp("sede")%></td>
            <td width="120" align="right"><!--<button id="btVer" onclick="alert('ver');" style="height: 20px;background:#0077A7 url(../images/img3.gif) repeat-x left bottom;font-size:11px;color:#FFFFFF;border:1px solid #999;-moz-border-radius:5px;-ms-border-radius: 5px;-webkit-border-radius:5px;border-radius:5px">Ver Inscripción</button>-->&nbsp;</td>
          </tr>
		  </table>
	</td> 
    </tr>   
    <tr>
       <td colspan="2">N° Participantes :&nbsp;&nbsp;<%=rsEmp("N_PARTICIPANTES")%></td> 
       <td>&nbsp;</td> 
       <td colspan="2">Documento de Compromiso :&nbsp;&nbsp;<%=rsEmp("Tipo Documento")%> N° <%=rsEmp("ORDEN_COMPRA")%></td> 
    </tr>         
    <tr>
       <td colspan="8">&nbsp;</td>
    </tr>
    <tr>
       <td colspan="8"><b>Seleccione los Participantes a Separar de la Inscripción :</b></td>
    </tr>
    <tr>
       <td colspan="8">&nbsp;</td>
    </tr>    
   </table>
	<div style="width:835px;"> 
	<table id="UStableSepIns" border="1" width="835"> 
    <thead> 
    	<tr> 
            <th>Rut</th> 
            <th>Nombre Participante</th> 
            <th>Cargo</th> 
            <th>Escolaridad</th>
            <th>Sel</th> 
        </tr> 
    </thead> 
    <tbody> 
    	  <%

				sql = "select T.RUT,T.NOMBRES,T.CARGO_EMPRESA, "
				sql = sql&"(case T.ESCOLARIDAD when 0 then 'Sin Escolaridad' "
				sql = sql&" when 1 then 'Básica Incompleta' "
				sql = sql&" when 2 then 'Básica Completa' "
				sql = sql&" when 3 then 'Media Incompleta' "
				sql = sql&" when 4 then 'Media Completa' "
				sql = sql&" when 5 then 'Superior Técnica Incompleta' "
				sql = sql&" when 6 then 'Superior Técnica Profesional Completa' "
				sql = sql&" when 7 then 'Universitaria Incompleta' "
				sql = sql&" when 8 then 'Universitaria Completa' end) as ESCOLARIDAD,HC.ID_HISTORICO_CURSO "
				sql = sql&" from HISTORICO_CURSOS HC "
				sql = sql&" inner join TRABAJADOR T on T.ID_TRABAJADOR=HC.ID_TRABAJADOR " 
				sql = sql&" where HC.ID_AUTORIZACION="&Request("IdAuto")

			   set rsSol =  conn.execute(sql)
			   i=0
			   check="chkIdHst"
			   
			   while not rsSol.eof
			   i=i+1
		  %>
					  <tr>
                   		   <td><%=rsSol("RUT")%></td>
                           <td><%=rsSol("NOMBRES")%></td> 
                           <td><%=rsSol("CARGO_EMPRESA")%></td> 
                           <td><%=rsSol("ESCOLARIDAD")%></td>
                 		   <td><input type="checkbox" id="<%=check&i%>" name="<%=check&i%>" value="<%=rsSol("ID_HISTORICO_CURSO")%>"></td>
					  </tr>
		  <%
		  	   rsSol.Movenext
			  wend
		  %>
    </tbody> 
    </table> 
     </div> 
      <table cellspacing="0" cellpadding="0" border=0>
          <tr>
           <td><input id="conSepFilas" name="conSepFilas" type="hidden" value="<%=i%>"/>
           <input id="SepSelIdHst" name="SepSelIdHst" type="hidden" value=""/>
           <input id="SepSelTotal" name="SepSelTotal" type="hidden" value="1"/>
           <input id="IdAuSep" name="IdAuSep" type="hidden" value="<%=Request("IdAuto")%>"/>
           </td>
          </tr>
    </table>
</form> 
<%
end if
%>
<div id="messageBox1" style="height:50px;overflow:auto;width:500px;"> 
  	<ul></ul> 
</div> 