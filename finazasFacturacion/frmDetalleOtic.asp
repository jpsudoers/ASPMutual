<%
on error resume next
%>
<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/NumeroaLetra.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

dim query
query="select AU.ID_AUTORIZACION,E.ID_OTIC,E.RUT,UPPER(E.R_SOCIAL) as 'Empresa',UPPER(E.DIRECCION) AS DIRECCION,"
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
query= query&" CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO "
query= query&" from AUTORIZACION AU "
query= query&" INNER JOIN EMPRESAS E ON E.ID_EMPRESA=AU.ID_EMPRESA "
query= query&" inner join PROGRAMA P on P.ID_PROGRAMA=AU.ID_PROGRAMA "
query= query&" INNER JOIN CURRICULO C ON C.ID_MUTUAL=P.ID_MUTUAL "
query= query&" where AU.ID_AUTORIZACION='"&Request("id")&"'"


set rsEmp = conn.execute (query)

%>
 <form name="frmDetalleOtic" id="frmDetalleOtic" action="finazasFacturacion/modificarOtic.asp" method="post">
  <table cellspacing="0" cellpadding="3" border=0>
    <tr>
	   <td>Total a Cancelar :</td>
       <td><%="$"&replace(FormatNumber(rsEmp("VALOR_OC"),0),",",".")%></td> 
     </tr>
     <tr>
	   <td width="200">Documento de Compromiso :</td>
       <td width="600"><%=rsEmp("Tipo Documento")&rsEmp("ORDEN_COMPRA")%> - <a href="#" onclick="docOC('<%=rsEmp("DOCUMENTO_COMPROMISO")%>');">Ver Documento</a></td> 
     </tr>
     <tr>
	   <td colspan="2">&nbsp;<input type="hidden" id="idAutorizacionOtic" name="idAutorizacionOtic" value="<%=Request("id")%>"/>
                             <input type="hidden" id="idEmpresaOtic" name="idEmpresaOtic" value="<%=rsEmp("ID_EMPRESA")%>"/>
                             <input type="hidden" id="NDocFactOtic" name="NDocFactOtic" value="<%=rsEmp("ORDEN_COMPRA")%>"/>
                             <input type="hidden" id="FormaPagoOtic" name="FormaPagoOtic" value="<%=rsEmp("TIPO_DOC")%>"/>
                          <input type="hidden" id="DOC_COMP_Otic" name="DOC_COMP_Otic" value="<%=rsEmp("DOCUMENTO_COMPROMISO")%>"/>
                             <input type="hidden" id="IdOtic" name="IdOtic" value="<%=rsEmp("ID_OTIC")%>"/>
      </td> 
     </tr>
   </table>
   <table cellspacing="0" cellpadding="3" border=0>
     <tr>
	   <td width="40">Rut :</td>
       <td width="250"><%=replace(FormatNumber(mid(rsEmp("RUT"), 1,len(rsEmp("RUT"))-2),0)&mid(rsEmp("RUT"), len(rsEmp("RUT"))-1,len(rsEmp("RUT"))),",",".")%></td> 
       <td width="70">Empresa :</td>
       <td width="430"><%=rsEmp("Empresa")%></td>
     </tr>
     <tr>
       <td>Giro : </td>
       <td><%=rsEmp("GIRO")%></td>
       <td>Direcci&oacute;n : </td>
       <td><%=rsEmp("DIR")%></td>
     </tr>
     </table>
     <table cellspacing="0" cellpadding="3" border=0>
     <tr>
       <td width="150">Fecha de Facturaci&oacute;n : </td>
       <td width="650"><%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%></td>
     </tr>
          <tr>
       <td>N&deg; de Participantes : </td>
       <td><%=rsEmp("N_PARTICIPANTES")%></td>
     </tr>
     <tr>
       <td>Valor Empresa : </td>
       <td><input type="hidden" id="V_Inscripcion" name="V_Inscripcion" value="<%=rsEmp("VALOR_OC")%>"/><input type="text" id="ValorEmpresa" name="ValorEmpresa" tabindex="1" onkeyup="calcularOtic();"/></td>
     </tr>
     <tr id="filaNFactConN">
       <td>N&deg; de Factura: </td>
       <td><input type="text" id="Nfactura" name="Nfactura" tabindex="2" onkeyup="calFacOtic();"/></td>
     </tr>
     <tr id="filaNFactSinN">
       <td>N&deg; de Factura: </td>
       <td><input type="hidden" id="NfacturaSin" name="NfacturaSin" value="0"/>No Aplica</td>
     </tr>
     <tr>
       <td colspan="2">&nbsp;</td>
     </tr>
 </table>
     <%
	    dim queryOtic
		queryOtic ="select E.RUT,UPPER(E.R_SOCIAL) as 'Otic',UPPER(E.DIRECCION+', '+E.CIUDAD) AS DIR,"
		queryOtic = queryOtic&"UPPER(E.DIRECCION+', '+E.CIUDAD) AS DIRECCION,DBO.MayMinTexto(E.COMUNA) AS COMUNA,"
		queryOtic = queryOtic&"DBO.MayMinTexto(E.CIUDAD) AS CIUDAD, E.FONO,"
		queryOtic = queryOtic&"UPPER(E.GIRO) AS GIRO from EMPRESAS E where E.ID_EMPRESA='"&rsEmp("ID_OTIC")&"'"

		set rsOtic = conn.execute (queryOtic)
	 %>
     <table cellspacing="0" cellpadding="3" border=0>
     <tr>
	   <td width="55">Rut Otic:</td>
       <td width="235"><%=replace(FormatNumber(mid(rsOtic("RUT"), 1,len(rsOtic("RUT"))-2),0)&mid(rsOtic("RUT"), len(rsOtic("RUT"))-1,len(rsOtic("RUT"))),",",".")%></td> 
       <td width="70">Otic : </td>
       <td width="440"><%=rsOtic("Otic")%></td>
     </tr>
     <tr>
       <td>Giro : </td>
       <td><%=rsOtic("GIRO")%></td>
       <td>Direcci&oacute;n : </td>
       <td><%=rsOtic("DIRECCION")%></td>
     </tr>
     </table>
     <table cellspacing="0" cellpadding="3" border=0>
     <tr>
       <td width="150">Fecha de Facturaci&oacute;n : </td>
       <td width="650"><%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%></td>
     </tr>
          <tr>
       <td>N&deg; de Participantes : </td>
       <td><%=rsEmp("N_PARTICIPANTES")%></td>
     </tr>
     <tr>
       <td>Valor Otic : </td>
       <td><input type="hidden" id="ValOtic" name="ValOtic"/><label id="ValorOtic">$ 0</label></td>
     </tr>
     <tr id="filaNFactConNOtic">
       <td>N&deg; de Factura: </td>
       <td><input type="text" id="NfacturaOticCon" name="NfacturaOticCon" tabindex="2" onkeyup="calFacOticCon();"/><label id="factOticNoAplica"></label></td>
     </tr>
      <tr id="filaNFactSinNOtic" style="display:none">
       <td>N&deg; de Factura: </td>
       <td><input type="text" id="NfacturaOtic" name="NfacturaOtic"/><label id="NfactOtic"></label><br/></td>
     </tr>
   </table>
</form> 
<div id="messageBox1" style="height:75px;overflow:auto;width:500px;"> 
  	<ul></ul> 
</div> 
<table width="600" border="0">
  <tr>
     <td><br/><br/><font color="#FF0000"><b>&bull; Verifique la posici&oacute;n de inicio de impresi&oacute;n de la impresora.</b></font> <a href="#" onclick="verManual();">Ver Manual</a></td>
  </tr>
</table>
<div style="text-align:center;display:none" id="imprimeFacturaOtic">
<table id="factEmpTab" style="margin:0px;padding:0px;width:740px;background:#FFFFFF;font:Courier,10px, normal;color:#000000;" border="0">
  <tr>
  	<td style="text-align:center;font-size:8px;width:500px;height:185px">&nbsp;</td>
    <td style="text-align:center;font-size:12px;width:240px;height:185px">
        <table width="200" border="0">
          <tr>
            <td height="90">&nbsp;</td>
          </tr>
          <tr>
            <td style="text-align:center;font-family:Arial;font-size:12px"><label id="facEmpresa"></label></td>
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
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsEmp("DocFactura")%></td>
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
              <tr>
                <td colspan="2">&nbsp;</td>
              </tr>
            </table>
            </td>
            <td style="text-align:right;font-family:Arial;font-size:12px;"><label id="txtValFacEmp"></label></td>
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
                        	<label id="txtFacEmp"></label>
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
              <td style="font-family:Arial;font-size:12px;text-align:right"><label id="txtTotFacEmp"></label></td>
            </tr>
            <tr>
              <td colspan="2" style="width:590px;height:186px;" valign="bottom">&nbsp;</td>
            </tr>
        </table>
        </td>
   </tr>
</table></td>
	</tr>
</table>
<table id="factOticTab"  style="margin:0px;padding:0px;width:740px;background:#FFFFFF;font:Courier,10px, normal;color:#000000;" border="0">
  <tr id="imprimeFila1">
  	<td style="text-align:center;font-size:8px;width:500px;height:165px">&nbsp;</td>
    <td style="text-align:center;font-size:12px;width:240px;height:165px">
        <table width="200" border="0">
          <tr>
            <td height="62">&nbsp;</td>
          </tr>
          <tr>
            <td style="text-align:center;font-family:Arial;font-size:12px"><label id="facOtic"></label></td>
          </tr>
        </table>
    </td>
  </tr>
  <tr id="imprimeFila2">
  	<td style="text-align:center;font-size:8px;width:500px;height:185px">&nbsp;</td>
    <td style="text-align:center;font-size:12px;width:240px;height:185px">
        <table width="200" border="0">
          <tr>
            <td height="90">&nbsp;</td>
          </tr>
          <tr>
            <td style="text-align:center;font-family:Arial;font-size:12px"><label id="facOtic2"></label></td>
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
              <td style="font-family:Arial;font-size:12px;text-align:left;width:465px;height:30px;vertical-align:middle;"><%=rsOtic("Otic")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:125px;height:30px;vertical-align:middle;"><%=replace(FormatNumber(mid(rsOtic("RUT"), 1,len(rsOtic("RUT"))-2),0)&mid(rsOtic("RUT"), len(rsOtic("RUT"))-1,len(rsOtic("RUT"))),",",".")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;height:30px;vertical-align:middle;">
                        <%=right("0"&day(now()),2)&"/"&right("0"&month(now()),2)&"/"&year(now)%></td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%;">
           <tr>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:355px;height:30px;vertical-align:middle;">
                        <%=rsOtic("DIRECCION")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:160px;height:30px;vertical-align:middle;">
                        <%=rsOtic("COMUNA")%></td>
               <td style="font-family:Arial;font-size:12px;text-align:left;width:110px;height:30px;vertical-align:middle;">
                        <%=rsOtic("FONO")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:100px;height:30px;">
                        <%=rsOtic("CIUDAD")%></td>
            </tr>
             <tr>
              <td colspan="5" style="font-family:Arial;font-size:12px;text-align:left;width:100%;height:30px;vertical-align:middle;"><%=rsOtic("GIRO")%></td>
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
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsEmp("DocFactura")%></td>
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
              <tr>
                <td colspan="2">&nbsp;</td>
              </tr>
            </table>
            </td>
            <td style="text-align:right;font-family:Arial;font-size:12px;"><label id="txtValFacOtic"></label></td>
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
											<label id="txtFacOtic"></label>
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
              <td style="width:590px;">&nbsp;</td>
              <td style="font-family:Arial;font-size:12px;text-align:right"><label id="txtTotFacOtic"></label></td>
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
