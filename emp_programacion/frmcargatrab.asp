<!--#include file="../conexion.asp"-->
<%
on error resume next
Response.CodePage = 65001
Response.CharSet = "utf-8"

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
%>
 <form name="frmCargaTrab" id="frmCargaTrab" action="#" method="post">
   <table cellspacing="0" cellpadding="1" border=0>   
      <tr>
	   <td colspan="2">Para ingresar la Carga Masiva de Trabajadores necesita realizar los siguientes pasos:</td>
      </tr>   
      <tr>
	   <td colspan="2">&nbsp;</td>
      </tr>        
      <tr>
	   <td height="31" colspan="2">1.- Descargar planilla Excel e ingresar la información solicitada:</td>
     </tr>   
      <tr>
	   <td height="119" colspan="2" align="center"><a href="documentos/Carga_Masiva_MC.xls" title="Descargar Aqui"><img src="images/logoexcel.jpg" width="70" height="65" border="0" /><br /><b>Descargar Aqui</b></a></td>
      </tr>       
      <tr>
	   <td colspan="2">2.- Seleccione planilla Excel a Cargar:</td>
      </tr>
      <tr>
	   <td colspan="2">&nbsp;</td>
      </tr>      
      <tr>
       <td width="140">Seleccione Archivo: </td>
       <td width="345"><input type="file" name="txtXLS" id="txtXLS" tabindex="5" maxlength="400" size="18">
       <input id="txtXLSPreins" name="txtXLSPreins" type="hidden" value="<%=Request("preinsIdxls")%>"/></td>
      </tr>
      <tr>
    	 <td colspan="2">&nbsp;</td>
      </tr>      
      <tr>
    	 <td colspan="2"><center><img id="loadingExcel" src="../images/loader.gif" style="display:none;"/></center></td>
      </tr>
   </table>
   <br />
   <div id="messageBox2" style="height:60px;overflow:auto;width:400px;"> 
  	<ul></ul> 
   </div> 
</form> 