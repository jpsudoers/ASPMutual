<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query="select g.ID_GASTO,g.GASTO,gp.DETALLE,gp.MONTO,gp.ESTADO,gp.ID_GASTOS_PROGRAMA,gp.ID_PROVEEDORES,"&_
	  "p.PROVEEDOR,p.rut from GASTOS_PROGRAMA gp "&_
      " inner join gastos g on g.ID_GASTO=gp.ID_GASTO "&_
	  " inner join proveedores p on p.ID_PROVEEDORES=gp.ID_PROVEEDORES "&_
      " where g.estado=1 and gp.ID_BLOQUE="&Request("bloque")

set rsCostos = conn.execute (query)

if not rsCostos.eof and not rsCostos.bof then %>
<form name="frmCostos" id="frmCostos" action="programacion/modificarcostos.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td colspan="2">Seleccione e Ingrese Costos Asociados.</td>
    </tr>
    <tr><td>&nbsp;</td></tr>
	<%
	codId=0
	
	WHILE NOT rsCostos.EOF
		codId=codId+1
    %> 
    <tr>
    <td width="300"><input id="txtIDCosto<%=codId%>" name="txtIDCosto<%=codId%>" type="hidden" value="<%=rsCostos("ID_GASTOS_PROGRAMA")%>"/><input id="IDCosto<%=codId%>" name="IDCosto<%=codId%>" type="hidden" value="<%=rsCostos("ID_GASTO")%>"/><%=rsCostos("GASTO")%>&nbsp;:<%if(rsCostos("ID_GASTO")="10")then%>&nbsp;&nbsp;&nbsp;<input id="txtDetItem<%=codId%>" name="txtDetItem<%=codId%>" type="text" <%if(rsCostos("ESTADO")="0")then%>style="display:none"<%end if%> size="20" value="<%=rsCostos("DETALLE")%>"/><%end if%></td> 
    <td width="550"><input type="radio" name="txtGasto<%=codId%>" id="txtGasto<%=codId%>_1" value="1" onclick="$('#txtMont<%=codId%>').show();if(<%=rsCostos("ID_GASTO")%>==10){$('#txtDetItem<%=codId%>').show();};$('#lbRutProv<%=codId%>').show();$('#txtRutProv<%=codId%>').show();$('#lbRazProv<%=codId%>').show();" <%if(rsCostos("ESTADO")="1")then%>checked="checked"<%end if%>>Si
	<input type="radio" name="txtGasto<%=codId%>" id="txtGasto<%=codId%>_0" value="0" onclick="$('#txtMont<%=codId%>').hide();$('#txtMont<%=codId%>').val('0');if(<%=rsCostos("ID_GASTO")%>==10){$('#txtDetItem<%=codId%>').hide();$('#txtDetItem<%=codId%>').val('');};$('#lbRutProv<%=codId%>').hide();$('#txtRutProv<%=codId%>').hide();$('#lbRazProv<%=codId%>').hide();$('#id_prov<%=codId%>').val('');" <%if(rsCostos("ESTADO")="0")then%>checked="checked"<%end if%>>No
    &nbsp;&nbsp;&nbsp;<input id="txtMont<%=codId%>" name="txtMont<%=codId%>" type="text" <%if(rsCostos("ESTADO")="0")then%>style="display:none"<%end if%> value="<%=rsCostos("MONTO")%>" onKeyPress="return acceptNum(event)" size="10" maxlength="9" />&nbsp;&nbsp;&nbsp;<label id="lbRutProv<%=codId%>" name="lbRutProv<%=codId%>" <%if(rsCostos("ESTADO")="0")then%>style="display:none"<%end if%>>Rut : </label><input id="txtRutProv<%=codId%>" name="txtRutProv<%=codId%>" type="text" <%if(rsCostos("ESTADO")="0")then%>style="display:none"<%end if%> size="11" maxlength="11" onkeyup="lookupProv(this.value,<%=codId%>);" value="<%if(rsCostos("ID_PROVEEDORES")<>"16")then%><%=rsCostos("rut")%><%end if%>"/><div class="suggestionsBox" id="suggestions<%=codId%>" style="display: none;position:absolute;z-index:1;left:430px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList<%=codId%>">
              &nbsp;
            </div>
       </div><label id="lbRazProv<%=codId%>" name="lbRazProv<%=codId%>" <%if(rsCostos("ESTADO")="0")then%>style="display:none"<%end if%>><%if(rsCostos("ID_PROVEEDORES")<>"16")then%>&nbsp;&nbsp;&nbsp;Proveedor : <%=rsCostos("PROVEEDOR")%><%end if%></label><input id="id_prov<%=codId%>" name="id_prov<%=codId%>" type="hidden" value="<%=rsCostos("ID_PROVEEDORES")%>"/>
    </td>  
    </tr>
    <%
    	rsCostos.MoveNext
	WEND
	%>
    <tr>
    	<td colspan="2">&nbsp;<input id="idCusr" name="idCusr" type="hidden" value="<%=Request("usr")%>"/>
    			  <input id="idCbloque" name="idCbloque" type="hidden" value="<%=Request("bloque")%>"/>
                  <input id="totitemCostos" name="totitemCostos" type="hidden" value="<%=codId%>"/>
                  <input id="valItems" name="valItems" type="hidden" value=""/>
        </td>
    </tr>    
   </table>
</form> 
<%
   else
		dim Costos
	    Costos="select g.ID_GASTO,g.GASTO from gastos g where g.estado=1"

		set rsDetCostos = conn.execute (Costos)   
%> 
<form name="frmCostos" id="frmCostos" action="programacion/insertarcostos.asp" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
       <td colspan="2">Seleccione e Ingrese Costos Asociados.</td>
    </tr>
    <tr><td>&nbsp;</td></tr>
	<%
	codId=0
	
	WHILE NOT rsDetCostos.EOF
		codId=codId+1
    %> 
    <tr>
    <td width="300"><input id="txtIDCosto<%=codId%>" name="txtIDCosto<%=codId%>" type="hidden" value="<%=rsDetCostos("ID_GASTO")%>"/><%=rsDetCostos("GASTO")%>&nbsp;:&nbsp;&nbsp;&nbsp;<input id="txtDetItem<%=codId%>" name="txtDetItem<%=codId%>" type="text" style="display:none" size="20"/></td> 
    <td width="550"><input type="radio" name="txtGasto<%=codId%>" id="txtGasto<%=codId%>_1" value="1" onclick="$('#txtMont<%=codId%>').show();if(<%=rsDetCostos("ID_GASTO")%>==10){$('#txtDetItem<%=codId%>').show();};$('#lbRutProv<%=codId%>').show();$('#txtRutProv<%=codId%>').show();$('#lbRazProv<%=codId%>').show();">Si
	<input type="radio" name="txtGasto<%=codId%>" id="txtGasto<%=codId%>_0" value="0" onclick="$('#txtMont<%=codId%>').hide();$('#txtMont<%=codId%>').val('0');if(<%=rsDetCostos("ID_GASTO")%>==10){$('#txtDetItem<%=codId%>').hide();$('#txtDetItem<%=codId%>').val('');};$('#lbRutProv<%=codId%>').hide();$('#txtRutProv<%=codId%>').hide();$('#lbRazProv<%=codId%>').hide();$('#id_prov<%=codId%>').val('');" checked="checked">No
    &nbsp;&nbsp;&nbsp;<input id="txtMont<%=codId%>" name="txtMont<%=codId%>" type="text" style="display:none" onKeyPress="return acceptNum(event)" size="10" maxlength="9"/>&nbsp;&nbsp;&nbsp;<label id="lbRutProv<%=codId%>" name="lbRutProv<%=codId%>" style="display:none">Rut : </label><input id="txtRutProv<%=codId%>" name="txtRutProv<%=codId%>" type="text" style="display:none" size="11" maxlength="11" onkeyup="lookupProv(this.value,<%=codId%>);"/><div class="suggestionsBox" id="suggestions<%=codId%>" style="display: none;position:absolute;z-index:1;left:430px">
            <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
            <div class="suggestionList" id="autoSuggestionsList<%=codId%>">
              &nbsp;
            </div>
       </div><label id="lbRazProv<%=codId%>" name="lbRazProv<%=codId%>" style="display:none"></label><input id="id_prov<%=codId%>" name="id_prov<%=codId%>" type="hidden"/>
    </td>  
    </tr>
    <%
    	rsDetCostos.MoveNext
	WEND
	%>
    <tr>
    	<td colspan="2">&nbsp;<input id="idCusr" name="idCusr" type="hidden" value="<%=Request("usr")%>"/>
    			  <input id="idCbloque" name="idCbloque" type="hidden" value="<%=Request("bloque")%>"/>
                  <input id="totitemCostos" name="totitemCostos" type="hidden" value="<%=codId%>"/>
                  <input id="valItems" name="valItems" type="hidden" value=""/>
        </td>
    </tr>    
   </table>
</form> 
<%
   end if
%>
<div id="messageBox2" style="height:70px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 