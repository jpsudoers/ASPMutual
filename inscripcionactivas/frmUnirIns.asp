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
query= query&" END) as 'Tipo Documento',AUTORIZACION.ORDEN_COMPRA,AUTORIZACION.id_empresa,"
query= query&"bloque_programacion.id_relator,bloque_programacion.id_sede,AUTORIZACION.ID_BLOQUE FROM AUTORIZACION "
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
 <form name="frmUnirIns" id="frmUnirIns" action="inscripcionactivas/USInsUnir.asp" method="post">
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
       <td colspan="8"><b>Seleccione Inscripciones a Unir :</b></td>
    </tr>
    <tr>
       <td colspan="8">&nbsp;</td>
    </tr>    
   </table>
	<div style="width:835px;"> 
	<table id="UStableIns" border="1" width="835"> 
    <thead> 
    	<tr> 
            <th>Fecha</th> 
            <th>Rut</th> 
            <th>Nombre de Empresa</th> 
            <th>Nombre de Curso</th>
            <th>Relator</th> 
            <th>N° Part</th> 
            <th>Ver</th> 
            <th>Sel</th> 
        </tr> 
    </thead> 
    <tbody> 
    	  <%
				sql = "select EMPRESAS.RUT as rut,dbo.MayMinTexto (EMPRESAS.R_SOCIAL) as empresa, "
				sql = sql&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO, " 
				sql = sql&"dbo.MayMinTexto (CURRICULO.NOMBRE_CURSO) as nombre, " 
				sql = sql&"AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.DOCUMENTO_COMPROMISO as doc,AUTORIZACION.N_PARTICIPANTES,"
				sql = sql&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) as instructor "
				sql = sql&" from AUTORIZACION " 
				sql = sql&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA " 
				sql = sql&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA " 
				sql = sql&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
				sql = sql&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE " 
				sql = sql&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator " 
				sql = sql&" WHERE AUTORIZACION.ESTADO=1 and AUTORIZACION.ID_PROGRAMA='"&rsEmp("ID_PROGRAMA")&"'"
				sql = sql&" and AUTORIZACION.ID_AUTORIZACION not in ('"&Request("IdAuto")&"')"
				sql = sql&" and AUTORIZACION.ID_EMPRESA='"&rsEmp("id_empresa")&"'"

				'sql = sql&" where a.ID_AUTORIZACION="&Request("IdAuto")&") "

			   set rsSol =  conn.execute(sql)
			   i=0

			   check="chkIdUS"
			   
			   while not rsSol.eof
			   i=i+1
		  %>
					  <tr>
                   		   <td><%=rsSol("FECHA_INICIO")%></td>
                           <td><%=rsSol("rut")%></td> 
                           <td><%=rsSol("empresa")%></td> 
                           <td><%=rsSol("nombre")%></td>
                           <td><%=rsSol("instructor")%></td>
                           <td><%=rsSol("N_PARTICIPANTES")%></td>
                           <td><span class="ui-state-valid"><a href="#" title="Ver Registro" class="ui-icon ui-icon-pencil" name="aContrato" onclick="update(<%=rsSol("ID_AUTORIZACION")%>);"></a></span></td>
                 		   <td><input type="checkbox" id="<%=check&i%>" name="<%=check&i%>" value="<%=rsSol("ID_AUTORIZACION")%>"></td>
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
           <td><input id="conUSFilas" name="conUSFilas" type="hidden" value="<%=i%>"/>
           <input id="USSelIdAu" name="USSelIdAu" type="hidden" value=""/>
           <input id="IdAuMan" name="IdAuMan" type="hidden" value="<%=Request("IdAuto")%>"/>
           <input id="USIdbloque" name="USIdbloque" type="hidden" value="<%=rsEmp("ID_BLOQUE")%>"/>
           <input id="USIdRelator" name="USIdRelator" type="hidden" value="<%=rsEmp("id_relator")%>"/>
           <input id="USIdSede" name="USIdSede" type="hidden" value="<%=rsEmp("id_sede")%>"/>
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