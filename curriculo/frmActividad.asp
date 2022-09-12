<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

query="select cb.ID_MUTUAL,cb.NOMBRE_BLOQUE,cb.horas,cb.DIA,ca.TEMA,ca.ACTIVIDAD,ca.ID_ACTIVIDAD_CURSO,ca.ID_BLOQUE_CURSO"&_
		  " from CURRICULO_BLOQUE cb "&_
		  " inner join CURRICULO_ACTIVIDADES ca on ca.ID_BLOQUE_CURSO=cb.ID_BLOQUE_CURSO "&_
		  " where ca.ID_ACTIVIDAD_CURSO="&Request("id")
		  
'response.Write(query)
'response.End()		  
		  

set rsEmp = conn.execute (query)

if not rsEmp.eof and not rsEmp.bof then 
%>
<form name="frmActividad" id="frmActividad" action="curriculo/guardaAct.asp" method="post">
<table cellspacing="0" cellpadding="1" border=0>
	<tr>
        <td>Dia :</td>
        <td><select name="actDia" id="actDia"></select>
        <input id="txtactId" name="txtactId" type="hidden" value="<%=rsEmp("ID_ACTIVIDAD_CURSO")%>" />
        <input id="txtactDiaId" name="txtactDiaId" type="hidden" value="<%=rsEmp("DIA")%>" />        
        <input id="txtactBloqueId" name="txtactBloqueId" type="hidden" value="<%=rsEmp("NOMBRE_BLOQUE")%>" />
        <input id="txtactHorasId" name="txtactHorasId" type="hidden" value="<%=rsEmp("horas")%>" />  
        <input id="txtIdMutual" name="txtIdMutual" type="hidden" value="<%=rsEmp("ID_MUTUAL")%>" />   
        <input id="txtBloqueId" name="txtBloqueId" type="hidden" value="<%=rsEmp("ID_BLOQUE_CURSO")%>" />               
        </td>
    </tr>
    <tr>
    	<td width="110">Bloque</td>
      	<td width="600"><select name="actBloque" id="actBloque"></select></td>
   	</tr>
  <tr>
        <td>Tema :</td>
        <td><textarea type="text" rows="3" cols="100" id="actTema" tabindex="3" name="actTema"><%=rsEmp("TEMA")%></textarea></td>
    </tr>
     <tr>
    	<td>Actividad :</td>
        <td><textarea type="text" rows="3" cols="100" id="actAct"  tabindex="4" name="actAct"><%=rsEmp("ACTIVIDAD")%></textarea></td>
    </tr>
      <tr>
        <td>Horas :</td>
        <td>
        	<select name="actHoras" id="actHoras"></select></td>
    </tr>    
</table>
</form> 
<%
   else
%> <form name="frmActividad" id="frmActividad" action="curriculo/guardaAct.asp" method="post">
<table cellspacing="1" cellpadding="1" border=0>
	<tr>
        <td>Dia :</td>
        <td><select name="actDia" id="actDia"></select>
        <input id="txtactId" name="txtactId" type="hidden" value="0" />
        <input id="txtIdMutual" name="txtIdMutual" type="hidden" value="" />
        <input id="txtBloqueId" name="txtBloqueId" type="hidden" value="0" />         
        </td>
    </tr>
    <tr>
    	<td width="110">Bloque</td>
      	<td width="600"><select name="actBloque" id="actBloque"></select></td>
   	</tr>
  <tr>
        <td>Tema :</td>
        <td><textarea type="text" rows="3" cols="100" id="actTema"  tabindex="3" name="actTema"/></textarea></td>
    </tr>
     <tr>
    	<td>Actividad :</td>
        <td><textarea type="text" rows="3" cols="100" id="actAct"  tabindex="4" name="actAct"/></textarea></td>
    </tr>
      <tr>
        <td>Horas :</td>
        <td><select name="actHoras" id="actHoras"></select></td>
    </tr>
</table>
</form> 
<%
   end if
%>
<div id="messageBox1" style="height:100px;overflow:auto;width:600px;"> 
  	<ul></ul> 
</div> 
