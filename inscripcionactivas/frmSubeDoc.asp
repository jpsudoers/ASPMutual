<%
on error resume next
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
 <form name="frmSubDoc" id="frmSubDoc" action="#" method="post">
 <table cellspacing="0" cellpadding="1" border=0>
     <tr>
       <td>Tipo Compromiso  :</td>
       <td>
           <select id="SubCom" name="SubCom" tabindex="1">
          		 <OPTION VALUE="">Seleccione</OPTION>
                 <OPTION VALUE="4">Carta Compromiso</OPTION>
          		 <OPTION VALUE="0">Orden de Compra</OPTION>
   				 <OPTION VALUE="1">Vale Vista</OPTION>
   				 <OPTION VALUE="2">Dep&oacute;sito Cheque</OPTION>
                 <OPTION VALUE="3">Transferencia</OPTION>
           </select>
       </td>
     </tr>
     <tr>
       <td width="160">N&uacute;mero OC/CI/DEP : </td>
       <td width="250"><input id="SubNum" name="SubNum" type="text" tabindex="2" maxlength="49" size="20"/></td>
     </tr>
     <tr>
       <td>Subir Documento :</td>
       <td><input type="file" name="SubDoc" id="SubDoc" tabindex="3" maxlength="200" size="20" onchange="changeFile(this.value)"></td>
     </tr>
     <tr>
       <td colspan="8"><center><img id="loading" src="../images/loader.gif" style="display:none;"/></center></td>
     </tr>
   </table>
</form> 
<div id="messageBox2" style="height:90px;overflow:auto;width:350px;"> 
  	<ul></ul> 
</div> 