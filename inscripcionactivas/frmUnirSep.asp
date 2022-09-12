 <form name="frmunirsep" id="frmunirsep" action="#" method="post">
   <table cellspacing="0" cellpadding="1" border=0>
    <tr>
    	<td>&nbsp;</td>
    </tr>
    <tr>
    	<td>Seleccione Acción a realizar con la Inscripción:</td>
    </tr>
    <tr>
    	<td>&nbsp;</td>
    </tr>
    <tr>
       <td align="center"><input type="radio" name="selopins" id="selopinsUnir" value="0"/>Unir&nbsp;&nbsp;&nbsp;<input type="radio" name="selopins" id="selopinsSep" value="1"/>Separar</td> 
    </tr>
     <tr>
       <td><input id="USIdAuto" name="USIdAuto" type="hidden" value="<%=request("IdAuto")%>"/>&nbsp;</td>
    </tr>
   </table>
</form> 
<div id="messageBox1" style="height:50px;overflow:auto;width:200px;"> 
  	<ul></ul> 
</div> 