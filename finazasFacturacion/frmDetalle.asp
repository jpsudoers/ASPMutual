<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/NumeroaLetra.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query="select AU.ID_AUTORIZACION,E.RUT,UPPER(E.R_SOCIAL) as 'Empresa',UPPER(E.DIRECCION) AS DIRECCION,"
query= query&"UPPER(E.DIRECCION+', '+E.CIUDAD) AS DIR, E.ID_EMPRESA,"
query= query&"DBO.MayMinTexto(E.COMUNA) AS COMUNA, DBO.MayMinTexto(E.CIUDAD) AS CIUDAD, E.FONO, "
query= query&"UPPER(E.GIRO) AS GIRO,AU.N_PARTICIPANTES,AU.VALOR_OC,AU.VALOR_CURSO, "
query= query&"(CASE WHEN AU.TIPO_DOC='0' then 'O/C '+AU.ORDEN_COMPRA " 
query= query&" WHEN AU.TIPO_DOC='1' then 'Vale Vista N° '+AU.ORDEN_COMPRA "
query= query&" WHEN AU.TIPO_DOC='2' then 'DV - '+AU.ORDEN_COMPRA " 
query= query&" WHEN AU.TIPO_DOC='3' then 'Transferencia N° '+AU.ORDEN_COMPRA " 
query= query&" WHEN AU.TIPO_DOC='4' then 'CONTRA FACTURA' END) as 'DocFactura', "
query= query&"(CASE WHEN AU.TIPO_DOC='0' then 'Orden de Compra N° ' "
query= query&" WHEN AU.TIPO_DOC='1' then 'Vale Vista N° ' " 
query= query&" WHEN AU.TIPO_DOC='2' then 'Depósito Cheque N° ' " 
query= query&" WHEN AU.TIPO_DOC='3' then 'Transferencia N° ' " 
query= query&" WHEN AU.TIPO_DOC='4' then 'Carta Compromiso N° ' " 
query= query&" END) as 'Tipo Documento', AU.TIPO_DOC, AU.ORDEN_COMPRA, AU.DOCUMENTO_COMPROMISO, "
query= query&" C.DESCRIPCION,C.CODIGO,C.HORAS, "
query= query&" CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"
query= query&" CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO, "
query= query&" ISNULL(convert(varchar,AU.N_REG_SENCE),'') as N_REG_SENCE "
query= query&" from AUTORIZACION AU "
query= query&" INNER JOIN EMPRESAS E ON E.ID_EMPRESA=AU.ID_EMPRESA "
query= query&" inner join PROGRAMA P on P.ID_PROGRAMA=AU.ID_PROGRAMA "
query= query&" INNER JOIN CURRICULO C ON C.ID_MUTUAL=P.ID_MUTUAL "
query= query&" where AU.ID_AUTORIZACION='"&Request("id")&"'"

set rsEmp = conn.execute (query)

%>
 <form name="frmDetalle" id="frmDetalle" action="finazasFacturacion/modificar.asp" method="post">
  <table cellspacing="0" cellpadding="3" border=0>
    <tr>
	   <td>Total a Cancelar :</td>
       <td><%="$"&replace(FormatNumber(rsEmp("VALOR_OC"),0),",",".")%>
        <% if(rsEmp("TIPO_DOC")<>"0" and rsEmp("TIPO_DOC")<>"4")then %>
       		<input type="hidden" id="Pagada" name="Pagada" value="1"/>
       <% else %>
       		<input type="hidden" id="Pagada" name="Pagada" value="0"/>
       <% end if %>
       </td> 
     </tr>
    <tr>
	   <td width="190">Documento de Compromiso :</td>
       <td width="400"><%=rsEmp("Tipo Documento")&rsEmp("ORDEN_COMPRA")%> - <a href="#" onclick="docOC('<%=rsEmp("DOCUMENTO_COMPROMISO")%>');">Ver Documento</a></td> 
     </tr>
     <% if(rsEmp("TIPO_DOC")<>"0" and rsEmp("TIPO_DOC")<>"4")then %>
     <tr>
	   <td>Curso Cancelado :</td>
       <td><input type="radio" name="InsPagada" id="InsPagada" value="1" checked='checked' onclick="$('#Pagada').val(this.value); MostrarFechPago('filaFechPago');$('#txtFechIngreso').val('');"/>Si
           <input type="radio" name="InsPagada" id="InsPagada" value="0" onclick="$('#Pagada').val(this.value); OcultarFechPago('filaFechPago');$('#txtFechIngreso').val('01-01-1990');"/>No
           </td> 
     </tr>
     <% end if %>
     <tr name="filaFechPago" id="filaFechPago">
        <td>Fecha de Pago :</td>
        <td><input id="txtFechIngreso" name="txtFechIngreso" type="text" tabindex="4"  size="15"  maxlength="20"/></td>
     </tr>
     <tr>
	   <td colspan="2">&nbsp;<input type="hidden" id="idAutorizacion" name="idAutorizacion" value="<%=Request("id")%>"/>
						     <input type="hidden" id="montoFact" name="montoFact" value="<%=rsEmp("VALOR_OC")%>"/>
                             <input type="hidden" id="idEmpresa" name="idEmpresa" value="<%=rsEmp("ID_EMPRESA")%>"/>
                             <input type="hidden" id="NDocFact" name="NDocFact" value="<%=rsEmp("ORDEN_COMPRA")%>"/>
                             <input type="hidden" id="FormaPago" name="FormaPago" value="<%=rsEmp("TIPO_DOC")%>"/>
                             <input type="hidden" id="DOC_COMP" name="DOC_COMP" value="<%=rsEmp("DOCUMENTO_COMPROMISO")%>"/>
       </td> 
     </tr>
   </table>
   <table cellspacing="0" cellpadding="3" border=0>
     <tr>
	   <td width="170">Rut :</td>
       <td width="550"><%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td> 
     </tr>
     <tr>
       <td>Empresa : </td>
       <td><%=rsEmp("Empresa")%></td>
     </tr>
     <tr>
       <td>Direcci&oacute;n : </td>
       <td><%=rsEmp("DIR")%></td>
     </tr>
     <tr>
       <td>Giro : </td>
       <td><%=rsEmp("GIRO")%></td>
     </tr>
     <tr>
       <td>Fecha de Facturaci&oacute;n : </td>
       <td><%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%></td>
     </tr>
          <tr>
       <td>N&deg; de Participantes : </td>
       <td><%=rsEmp("N_PARTICIPANTES")%></td>
     </tr>
     <tr>
       <td>Valor Empresa : </td>
       <td><%="$" &replace(FormatNumber(rsEmp("VALOR_OC"),0),",",".")%></td>
     </tr>
     <tr>
       <td>N&deg; de Factura: </td>
       <td><input type="text" id="NfacturaEmp" name="NfacturaEmp" tabindex="1" onkeyup="$('#facEmp').html(this.value);"/></td>
     </tr>
     <tr>
       <td colspan="2">&nbsp;</td>
      </tr>
   </table>
</form> 
<div id="messageBox2" style="height:33px;overflow:auto;width:300px;"> 
  	<ul></ul> 
</div> 
<table width="500" border="0">
  <tr>
     <td><br/><br/><font color="#FF0000"><b>&bull; Verifique la posici&oacute;n de inicio de impresi&oacute;n de la impresora.</b></font> <a href="#" onclick="verManual();">Ver Manual</a></td>
  </tr>
</table>
<div style="text-align:center;display:none" id="imprimeFactura">
<table style="margin:0px;padding:0px;width:740px;background:#FFFFFF;font:Courier,10px, normal;color:#000000;" border="0">
  <tr>
  	<td style="text-align:center;font-size:8px;width:500px;height:185px">&nbsp;</td>
    <td style="text-align:center;font-size:12px;width:240px;height:185px">
        <table width="200" border="0">
          <tr>
            <td height="90">&nbsp;</td>
          </tr>
          <tr>
            <td style="text-align:center;font-family:Arial;font-size:12px"><label id="facEmp"></label></td>
          </tr>
        </table>
    </td>
  </tr>
  <tr>
  	<td colspan="2">
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
        <tr>
        <td colspan="4">
            <table border="0" cellpadding="0" cellspacing="0" style="width:100%;">
            <tr>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:465px;height:30px;vertical-align:middle;"><%=rsEmp("Empresa")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:125px;height:30px;vertical-align:middle;"><%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;height:30px;vertical-align:middle;">
                        <%=right("0"&day(now()),2)&"/"&right("0"&month(now()),2)&"/"&year(now)%></td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%;">
           <tr>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:355px;height:30px;vertical-align:middle;">
                        <%=rsEmp("DIRECCION")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:160px;height:30px;vertical-align:middle;">
                        <%=rsEmp("COMUNA")%></td>
               <td style="font-family:Arial;font-size:12px;text-align:left;width:110px;height:30px;vertical-align:middle;">
                        <%=rsEmp("FONO")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:100px;height:30px;">
                        <%=rsEmp("CIUDAD")%></td>
            </tr>
             <tr>
              <td colspan="5" style="font-family:Arial;font-size:12px;text-align:left;width:100%;height:30px;vertical-align:middle;"><%=rsEmp("GIRO")%></td>
            </tr>
            <tr>
              <td colspan="7" style="width:60px;height:55px;">&nbsp;</td>
            </tr>
          </table>
      	<div style="height:365px;">
    	<table border="0" cellpadding="0" cellspacing="0" style="width:100%;text-align:left;max-height:525px">
          <tr style="height:15px;vertical-align:top">
            <td colspan="2" style="text-align:left;font-family:Arial;font-size:12px;width:590px;height:20px;">Corresponde a Curso de:</td>
          </tr>
          <tr style="height:15px;vertical-align:top">
            <td style="text-align:left;font-family:Arial;font-size:12px;width:590px;height:20px;">
            <table width="590" border="0">
              <tr>
                <td colspan="2" style="text-align:left;font-family:Arial;font-size:12px;"><%=rsEmp("DESCRIPCION")%></td>
              </tr>
              <tr>
                <td colspan="2" style="text-align:left;font-family:Arial;font-size:12px;">Dictado por Mutual de Seguridad Capacitaci&oacute;n S.A.</td>
              </tr>
              <tr>
                <td width="200" style="text-align:left;font-family:Arial;font-size:12px;">Direcci&oacute;n : </td>
                <td width="290" style="text-align:left;font-family:Arial;font-size:12px;">Latorre 2631 Antofagasta</td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">C&oacute;digo Sence : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsEmp("CODIGO")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">N&uacute;mero de Horas : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsEmp("HORAS")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">Fecha de Inicio :</td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=replace(rsEmp("FECHA_INICIO"),"-","/")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">Fecha de T&eacute;rmino : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=replace(rsEmp("FECHA_TERMINO"),"-","/")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">N&uacute;mero de Participantes : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsEmp("N_PARTICIPANTES")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">Valor por Participantes : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=FormatNumber(rsEmp("VALOR_CURSO"),0)&".-"%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">N&uacute;mero Orden de Compra : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsEmp("DocFactura")&" "%><label id="fechPago"></label></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">Registro Sence : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsEmp("N_REG_SENCE")%></td>
              </tr>
              <tr>
                <td colspan="2">&nbsp;</td>
              </tr>
              <tr>
             <td colspan="2" style="text-align:left;font-family:Arial;font-size:12px;">INFORMACI&Oacute;N SOLICITADA POR SENCE</td>
              </tr>
              <tr>
                <td colspan="2" style="text-align:left;font-family:Arial;font-size:12px;">&nbsp;&nbsp;&nbsp;EN CERTIFICADO DE ASISTENCIA</td>
              </tr>
            </table>
            </td>
            <td style="text-align:right;font-family:Arial;font-size:12px;"><%=FormatNumber(rsEmp("VALOR_OC"),0)%></td>
          </tr>
        </table>
        </div>
		</td>
	</tr>
    <tr>
  	<td>
    	<table width="100%" border="0">
            <tr>
                <td style="width:50px">&nbsp;</td>
                <td style="font-family:Arial;font-size:12px;text-align:left;text-transform:uppercase;">
                        <%=CONVERTIR(cdbl(FormatNumber(rsEmp("VALOR_OC"),0)))%>
                </td>
            </tr>
        </table>
    </td>
	</tr>
    <tr>
   		<td>
        <table border="0" width="100%">
            <tr>
              <td colspan="2" style="font-family:Arial;font-size:12px;text-align:right;width:590px;">&nbsp;</td>
            </tr>
            <tr>
              <td style="font-family:Arial;font-size:12px;text-align:right;width:590px;">&nbsp;</td>
              <td style="font-family:Arial;font-size:12px;text-align:right"><%=FormatNumber(rsEmp("VALOR_OC"),0)%></td>
            </tr>
            <tr>
              <td colspan="2" style="width:590px;height:182px;" valign="bottom">&nbsp;</td>
            </tr>
        </table>
        </td>
   </tr>
</table></td>
	</tr>
</table>
</div>
