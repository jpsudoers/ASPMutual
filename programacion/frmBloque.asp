<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

if(Request("estado")="1")then
vid = Request("id")
else
vid="0"
end if
dim query

query="select bloProg.id_relator,bloProg.id_sede,bloProg.cupos,bloProg.id_rel_seg,bloProg.nom_sede,bloProg.n_relatores "
query= query&" from bloque_programacion bloProg "
query= query&" where bloProg.id_programa="&vid&" and bloProg.id_bloque='"&Request("bloque")&"'"

set rsEmp = conn.execute (query)

set rsTotInsc = conn.execute ("select isnull(SUM(bloqProg.cupos),0)as totCupos from bloque_programacion bloqProg where bloqProg.id_programa='"&Request("id")&"'")

dim vacantesFrm
if(Request("totalVacantes")<>"")then
vacantesFrm=Request("totalVacantes")
else
vacantesFrm="0"
end if

dim totDisponible
totDisponible=cdbl(cdbl(vacantesFrm)-cdbl(rsTotInsc("totCupos")))

if not rsEmp.eof and not rsEmp.bof then 
totDisponible=cdbl(cdbl(totDisponible)+cdbl(rsEmp("cupos")))
%>
 <form name="frmBloque" id="frmBloque" action="programacion/modificar_tab.asp" method="post">
 <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td width="180">Primer Relator:</td>
       <td width="200"><select id="relator_frm" name="relator_frm" tabindex="13" onchange="llena_seg_rel(this.value);"></select>&nbsp;&nbsp;<img id="AgregaRel" src="images/Agregar2.png" onclick="agrega_relator();" title="Agregar Segundo Relator" width="11" height="11"/></td>
     </tr>
     <tr id="filaRelSeg">
       <td>Segundo Relator:</td>
       <td><select id="relator_frm_seg" name="relator_frm_seg" tabindex="14"></select>&nbsp;&nbsp;<img src="images/remove.png" onclick="elimina_relator();" title="Eliminar Segundo Relator"/></td>
     </tr>          
     <tr>
       <td>Sala/Sede</td>
       <td><select id="salaSede_frm" name="salaSede_frm" tabindex="15" onchange="mostrarDirSede(this.value);" style="width:32em;"></select></td>
     </tr>
     <tr id="filaDir">
       <td>Direcci&oacute;n:</td>
       <td><input id="txtDir_frm" name="txtDir_frm" type="text" tabindex="16" maxlength="60" size="60" value="<%=rsEmp("nom_sede")%>"/></td>
     </tr>  
     <tr>
       <td>Cupos</td>
       <td><input id="txtCupo_frm" name="txtCupo_frm" type="text" tabindex="17" maxlength="3" size="4" value="<%=rsEmp("cupos")%>" onKeyPress="return acceptNum(event)"/></td>
     </tr>
   </table>
     <table cellspacing="0" cellpadding="0" border=0>
          <tr>
           <td><input id="tabPrograma" name="tabPrograma" type="hidden" value="<%=Request("id")%>"/>
           <input id="txtBloque" name="txtBloque" type="hidden" value="<%=Request("bloque")%>"/></td>
           <td><input id="tabRelator" name="tabRelator" type="hidden" value="<%=rsEmp("id_relator")%>"/>
            <input id="tabSala" name="tabSala" type="hidden" value="<%=rsEmp("id_sede")%>"/>
            <input id="totProgDisp" name="totProgDisp" type="hidden" value="<%=totDisponible%>"/>
            <input id="tabRelSeg" name="tabRelSeg" type="hidden" value="<%=rsEmp("id_rel_seg")%>"/>
            <input id="txtNRelator" name="txtNRelator" type="hidden" value="<%=rsEmp("n_relatores")%>"/>
            <input id="selRelSeg" name="selRelSeg" type="hidden" value=""/>
            </td>
          </tr>
    </table>
</form> 
<%
   else
%> <form name="frmBloque" id="frmBloque" action="programacion/insertar_tab.asp" method="post">
   <table cellspacing="1" cellpadding="1" border=0>
     <tr>
       <td width="180">Primer Relator:</td>
       <td width="200"><select id="relator_frm" name="relator_frm" tabindex="13" onchange="llena_seg_rel(this.value);"></select>&nbsp;&nbsp;<img id="AgregaRel" src="images/Agregar2.png" onclick="agrega_relator();" title="Agregar Segundo Relator" width="11" height="11"/></td>
     </tr>
     <tr id="filaRelSeg">
       <td>Segundo Relator:</td>
       <td><select id="relator_frm_seg" name="relator_frm_seg" tabindex="14"></select>&nbsp;&nbsp;<img src="images/remove.png" onclick="elimina_relator();" title="Eliminar Segundo Relator"/></td>
     </tr>     
     <tr>
       <td>Sala/Sede:</td>
       <td><select id="salaSede_frm" name="salaSede_frm" tabindex="15" onchange="mostrarDirSede(this.value);" style="width:32em;"></select></td>
     </tr>
     <tr id="filaDir">
       <td>Direcci&oacute;n:</td>
       <td><input id="txtDir_frm" name="txtDir_frm" type="text" tabindex="16" maxlength="60" size="60"/></td>
     </tr>     
     <tr>
       <td>Cupos:</td>
       <td><input id="txtCupo_frm" name="txtCupo_frm" type="text" tabindex="17" maxlength="3" size="4" onKeyPress="return acceptNum(event)"/></td>
     </tr>
   </table>
     <table cellspacing="0" cellpadding="0" border=0>
          <tr>
           <td><input id="tabPrograma" name="tabPrograma" type="hidden" value="<%=Request("id")%>"/> 
           <input id="txtBloque" name="txtBloque" type="hidden" value="<%=Request("bloque")%>"/>
           <input id="totProgDisp" name="totProgDisp" type="hidden" value="<%=totDisponible%>"/>
           <input id="txtNRelator" name="txtNRelator" type="hidden" value="1"/>
           <input id="tabRelSeg" name="tabRelSeg" type="hidden" value="0"/>
           <input id="tabRelator" name="tabRelator" type="hidden" value="0"/>
           <input id="selRelSeg" name="selRelSeg" type="hidden" value=""/>
           </td> 
           <td>&nbsp;</td>
          </tr>
    </table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:90px;overflow:auto;width:350px;"> 
  	<ul></ul> 
</div> 
