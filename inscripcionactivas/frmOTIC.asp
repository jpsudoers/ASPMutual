<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

%>
 <form name="frmOTIC" id="frmOTIC" action="#" method="post">
	<div style="width:835px;"> 
	<table id="tableOTIC" border="1" width="835"> 
    <thead> 
    	<tr> 
            <th>Rut</th> 
            <th>Nombre OTIC</th> 
            <th>&nbsp;Sel</th> 
        </tr> 
    </thead> 
    <tbody> 
    	  <%
				sql = "SELECT E.RUT, E.R_SOCIAL, E.ID_EMPRESA FROM EMPRESAS E "
				sql = sql&"WHERE E.TIPO=2 AND E.ESTADO=1 AND E.RUT NOT IN ('95207000-2')"

			   set rsSol =  conn.execute(sql)
			   i=0
		   
		       check="checkOTIC"
		   
			   while not rsSol.eof
			   i=i+1
		  %>
					  <tr>
                   		   <td><%=rsSol("RUT")%></td>
                           <td><%=rsSol("R_SOCIAL")%></td> 
                 		   <td><input type="checkbox" id="<%=check&i%>" name="<%=check&i%>" value="<%=i%>" 
                           onclick="checkearOTIC(<%=i%>,<%=rsSol("ID_EMPRESA")%>,'<%=rsSol("RUT")%>','<%=rsSol("R_SOCIAL")%>');"></td>
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
           <td><input id="SelcountFilas" name="SelcountFilas" type="hidden" value="<%=i%>"/>
           <input id="SelOTICRut" name="SelOTICRut" type="hidden"/>
           <input id="SelOTICNom" name="SelOTICNom" type="hidden"/>
           <input id="SelOTICId" name="SelOTICId" type="hidden"/>
           </td>
          </tr>
    </table>
</form> 
<div id="messageBox1" style="height:50px;overflow:auto;width:200px;"> 
  	<ul></ul> 
</div> 