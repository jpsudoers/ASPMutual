<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

vid = Request("id")
vprog=Request("prog")
vcon_otic=Request("con_otic")
vcon_fran=Request("con_fran")
vreg_sence=Request("reg_sence")
vidproyecto=Request("id_proyecto")

dim query
query= "SELECT CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
query= query&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, "
query= query&"EMPRESAS.R_SOCIAL,EMPRESAS.RUT,preinscripciones.id_otic, PROGRAMA.ID_PROGRAMA,EMPRESAS.ID_EMPRESA,"
query= query&"CURRICULO.CODIGO,CURRICULO.SENCE,preinscripciones.tipo_compromiso,preinscripciones.numero_compromiso, "
query= query&"preinscripciones.id_preinscripcion,preinscripciones.doc_compromiso as doc,preinscripciones.participantes,"
query= query&"preinscripciones.valor,preinscripciones.valor_total,preinscripciones.costo,preinscripciones.tipo_empresa, "
query= query&"preinscripciones.id_otic, preinscripciones.id_tipo_venta, preinscripciones.id_banco, preinscripciones.id_cuenta_corriente, preinscripciones.monto_transferencia, FORMAT (preinscripciones.fecha_transferencia, 'dd/MM/yyyy ') AS fecha_transferencia, preinscripciones.rut_depositante, preinscripciones.nombre_depositante, preinscripciones.email_depositante, preinscripciones.relacion_depositante "
query= query&" FROM preinscripciones "
query= query&" INNER JOIN PROGRAMA ON PROGRAMA.ID_PROGRAMA=preinscripciones.id_programacion "
query= query&" INNER JOIN CURRICULO ON CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
query= query&" INNER JOIN EMPRESAS ON EMPRESAS.ID_EMPRESA=preinscripciones.id_empresa "
query= query&" where preinscripciones.id_preinscripcion="&vid

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
 <form name="frmAsignar" id="frmAsignar" action="inspendientes/modificar.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <%
       if(rsEmp("tipo_empresa")="1")then
	   %>
       <td>Rut Empresa :</td>
       <%
	   end if
	   if(rsEmp("tipo_empresa")="2")then
	   %>
       <td>Rut Mandante :</td>
       <%
	   end if
	   %>
       <td><%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td>
       <td><input type="HIDDEN" id="Empresa" name="Empresa" value="<%=rsEmp("ID_EMPRESA")%>"/></td>
       <td>Raz&oacute;n Social : </td>
       <td colspan="4"><%=rsEmp("R_SOCIAL")%></td>
    </tr>
    <tr>
       <td>C&oacute;digo Curso :</td>
       <td><%=rsEmp("CODIGO")%></td>
       <td><input type="hidden" id="Programa" name="Programa" value="<%=rsEmp("ID_PROGRAMA")%>"/></td>
       <td>Nombre Curso :</td>
       <td colspan="4"><%=rsEmp("NOMBRE_CURSO")%></td>
    </tr>
     <tr style="display:none">
       <td width="120">Sence :</td>
      <%
	   checkedSi=""
	   checkedNo=""
	   
	   if(rsEmp("SENCE")=0)then
	   checkedSi="checked='checked'"
	   else
	   checkedNo="checked='checked'"
	   end if
	   %>
       
       <td width="100"><input type="radio" name="Sence_SI" id="Sence_SI" disabled value="0" <%=checkedSi%>/>
         Si
           <input type="radio" name="Sence_NO" id="Sence_NO" value="1" disabled <%=checkedNo%>/>
       No</td> 
       <td width="20">&nbsp;</td>
       <td width="120">&nbsp;</td>
       <td width="172">&nbsp;</td>
       <td width="20">&nbsp;</td>
       <td width="90">&nbsp;</td>
       <td width="303">&nbsp;</td>
    </tr>
    <tr>
       <td width="120">Fecha Inicio :</td>
       <td width="100"><%=rsEmp("FECHA_INICIO_")%></td>
       <td width="20">&nbsp;</td>
       <td width="120">Fecha T&eacute;rmino : </td>
       <td width="172"><%=rsEmp("FECHA_TERMINO")%></td>
       <td width="20">&nbsp;</td>
       <td width="90">&nbsp;</td>
       <td width="303">&nbsp;</td>       
    </tr>
    <tr>
       <td>&nbsp;</td>
       <td>&nbsp;</td>
       <td><input type="hidden" id="txtpreinscripcion" name="txtpreinscripcion" value="<%=rsEmp("id_preinscripcion")%>"/></td>
       <td><input type="hidden" id="txtIdOtic" name="txtIdOtic" value="<%=rsEmp("id_otic")%>"/>
       <input type="hidden" id="txtValorE" name="txtValorE" value="<%=rsEmp("costo")%>"/>
       <input type="hidden" id="txtOrdenCompraE" name="txtOrdenCompraE" value="<%=rsEmp("numero_compromiso")%>"/></td>
       <td></td>
    </tr>
   </table>
 
   <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td><input type="hidden" id="progId" name="progId" value="<%=vprog%>"/>
       <input type="hidden" id="EmpresaAsig" name="EmpresaAsig" value="<%=rsEmp("ID_EMPRESA")%>"/>
       <input type="hidden" id="ProgramaAsig" name="ProgramaAsig" value="<%=rsEmp("ID_PROGRAMA")%>"/>
       <input type="hidden" id="txtpreinscripcionAsig" name="txtpreinscripcionAsig" value="<%=rsEmp("id_preinscripcion")%>"/>
       <input type="hidden" id="txtIdOticAsig" name="txtIdOticAsig" value="<%=rsEmp("id_otic")%>"/>
       <input type="hidden" id="txtValorEAsig" name="txtValorEAsig" value="<%=rsEmp("costo")%>"/>
       <input type="hidden" id="txtOrdenCompraEAsig" name="txtOrdenCompraEAsig" value="<%=rsEmp("numero_compromiso")%>"/>
       <input type="hidden" id="documentosAsig" name="documentosAsig" value="<%=rsEmp("doc")%>"/>
       <input type="hidden" id="totalTrabAsig" name="totalTrabAsig" value="0"/>
       <input type="hidden" id="totalInsAsig" name="totalInsAsig" value="0"/>
       <input type="hidden" id="relatorSedeAsig" name="relatorSedeAsig" value=""/>
       <input type="hidden" id="bloqueSel" name="bloqueSel" value=""/>
       <input type="hidden" id="Relator" name="Relator" value=""/>
       <input type="hidden" id="Sede" name="Sede" value=""/>
       <input type="hidden" id="tipoDocAsig" name="tipoDocAsig" value="<%=rsEmp("tipo_compromiso")%>"/>
       <input type="hidden" id="NPartInsc" name="NPartInsc" value="<%=rsEmp("participantes")%>"/>
       <input type="hidden" id="ValorCurso" name="ValorCurso" value="<%=rsEmp("valor")%>"/>
       <input type="hidden" id="C_Otic" name="C_Otic" value="<%=vcon_otic%>"/>
       <input type="hidden" id="C_Fran" name="C_Fran" value="<%=vcon_fran%>"/> 
       <input type="hidden" id="R_Sence" name="R_Sence" value="<%=vreg_sence%>"/>     
       <input type="hidden" id="id_Proy" name="id_Proy" value="<%=vidproyecto%>"/> 
       <input type="hidden" id="tipoVenta" name="tipoVenta" value="<%=rsEmp("id_tipo_venta")%>"/>    

       <input type="hidden" id="Banco" name="Banco" value="<%=rsEmp("id_banco")%>"/>  
       <input type="hidden" id="numeroCuenta" name="numeroCuenta" value="<%=rsEmp("id_cuenta_corriente")%>"/>  
       <input type="hidden" id="txtMonto" name="txtMonto" value="<%=rsEmp("monto_transferencia")%>"/>  
       <input type="hidden" id="txtFechaDeposito" name="txtFechaDeposito" value="<%=rsEmp("fecha_transferencia")%>"/>  
       <input type="hidden" id="txtRutDepositante" name="txtRutDepositante" value="<%=rsEmp("rut_depositante")%>"/>  
       <input type="hidden" id="txtNombreDepositante" name="txtNombreDepositante" value="<%=rsEmp("nombre_depositante")%>"/>  
       <input type="hidden" id="txtEmail" name="txtEmail" value="<%=rsEmp("email_depositante")%>"/>  
       <input type="hidden" id="txtRelacion" name="txtRelacion" value="<%=rsEmp("relacion_depositante")%>"/>                 
       </td>
     </tr>
   </table>
    <div style="width:835px;"> 
	<table id="mytable" border="1" width="835"> 
    <thead> 
    	<tr> 
            <th>Rut</th> 
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
sql = "select INSTRUCTOR_RELATOR.RUT,bloqProg.id_relator,bloqProg.id_sede,dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+"
sql = sql&"INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor, "
sql = sql&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloqProg.nom_sede "
sql = sql&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as sede,"
sql = sql&"bloqProg.id_bloque, bloqProg.cupos, "
sql = sql&"bloqProg.estado,(select COUNT(*) from HISTORICO_CURSOS where HISTORICO_CURSOS.ID_BLOQUE=bloqProg.id_bloque) as incritos"
sql = sql&" from bloque_programacion bloqProg"
sql = sql&" inner join SEDES on SEDES.ID_SEDE=bloqProg.id_sede "
sql = sql&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloqProg.id_relator "
sql = sql&" where bloqProg.id_programa='"&rsEmp("ID_PROGRAMA")&"'"

			   set rsSol =  conn.execute(sql)
			   i=0
			   
			   relId="relId"
			   Sala="Sala"
			   Bloque="Bloque"
			   check="check"
			   estadoBloqueo=""
			   totInsBloqProg=0
			   
			   while not rsSol.eof
			   i=i+1
			   
			   totInsBloqProg=cdbl(cdbl(rsSol("cupos"))-cdbl(rsSol("incritos")))
			   
			   if(rsSol("estado")="1" and totInsBloqProg>0 and cdbl(rsEmp("participantes"))<=totInsBloqProg)then
			     estadoBloqueo=""
			   else
			     estadoBloqueo="disabled"
			   end if
			   
		  %>
					  <tr>
                      <td><input id="<%=relId&i%>" name="<%=relId&i%>" type="hidden" value="<%=rsSol("id_relator")%>"/><input id="<%=Sala&i%>" name="<%=Sala&i%>" type="hidden" value="<%=rsSol("id_sede")%>"/><input id="<%=Bloque&i%>" name="<%=Bloque&i%>" type="hidden" value="<%=rsSol("id_bloque")%>"/><%=replace(FormatNumber(mid(rsSol("RUT"), 1,len(rsSol("RUT"))-2),0)&mid(rsSol("RUT"), len(rsSol("RUT"))-1,len(rsSol("RUT"))),",",".")%></td>
                           <td><%=rsSol("instructor")%></td> 
                           <td><%=rsSol("sede")%></td> 
                           <td align="center"><%=rsSol("cupos")%></td>
                           <td align="center"><%=rsSol("incritos")%></td>
                           <td align="center"><span class="ui-state-valid" ><a href="#" title="Modificar Registro" class="ui-icon ui-icon-pencil" name="aContrato" onclick="MostrarInscritos(<%=rsSol("id_bloque")%>);"></a></span></td>
                 <td align="center"><input type="checkbox" id="<%=check&i%>" name="<%=check&i%>" <%=estadoBloqueo%> value="<%=i%>" onclick="checkear(<%=i%>);"></td>
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
           <td width="200"><input id="countFilas" name="countFilas" type="hidden" value="<%=i%>"/></td>
           <td width="200">&nbsp;</td> 
           <td width="200">&nbsp;</td>
           <td width="200">&nbsp;</td>
           <td width="200">&nbsp;</td>
          </tr>
    </table>
</form> 
<%
end if
%>
<div id="messageBox1" style="height:45px;overflow:auto;width:200px;"> 
  	<ul></ul> 
</div> 