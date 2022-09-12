<!--#include file="../conexion.asp"-->
<!--#include file="../funciones/NumeroaLetra.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

on error goto 0

dim SQL
SQL="select AU.ID_AUTORIZACION,E.RUT,UPPER(E.R_SOCIAL) as 'Empresa',UPPER(E.DIRECCION) AS DIRECCION,"
SQL= SQL&"DBO.MayMinTexto(E.COMUNA) AS COMUNA, DBO.MayMinTexto(E.CIUDAD) AS CIUDAD, E.FONO, "
SQL= SQL&"UPPER(E.GIRO) AS GIRO,AU.N_PARTICIPANTES,AU.VALOR_OC,AU.VALOR_CURSO, "
SQL= SQL&"(CASE WHEN AU.TIPO_DOC='0' then 'Orden de Compra ' "
SQL= SQL&" WHEN AU.TIPO_DOC='1' then 'Vale Vista ' " 
SQL= SQL&" WHEN AU.TIPO_DOC='2' then 'Depósito Cheque ' " 
SQL= SQL&" WHEN AU.TIPO_DOC='3' then 'Transferencia ' " 
SQL= SQL&" WHEN AU.TIPO_DOC='4' then 'Carta Compromiso ' " 
SQL= SQL&" END) as 'Tipo Documento', AU.ORDEN_COMPRA, "
SQL= SQL&" C.DESCRIPCION,C.CODIGO,C.HORAS, "
SQL= SQL&" CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"
SQL= SQL&" CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO "
SQL= SQL&" from AUTORIZACION AU "
SQL= SQL&" INNER JOIN EMPRESAS E ON E.ID_EMPRESA=AU.ID_EMPRESA "
SQL= SQL&" inner join PROGRAMA P on P.ID_PROGRAMA=AU.ID_PROGRAMA "
SQL= SQL&" INNER JOIN CURRICULO C ON C.ID_MUTUAL=P.ID_MUTUAL "
SQL= SQL&" where AU.ID_AUTORIZACION=1351"

set rsSQL = conn.execute (SQL)

%>
<div style="text-align:center">
<table style="margin:0px;padding:0px;width:740px;background:#FFFFFF;font:Courier,10px, normal;color:#000000;" border="0">
  <tr>
  	<td style="text-align:center;font-size:8px;width:500px;height:140px">&nbsp;</td>
    <td style="text-align:center;font-size:12px;width:240px;height:140px">
        <table width="200" border="0">
          <tr>
            <td height="56">&nbsp;</td>
          </tr>
          <tr>
            <td  style="text-align:center;font-family:Arial;font-size:12px">12343</td>
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
              <td style="font-family:Arial;font-size:12px;text-align:left;width:465px;height:30px;vertical-align:middle;"><%=rsSQL("Empresa")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:125px;height:30px;vertical-align:middle;"><%=replace(FormatNumber(mid(rsSQL("RUT"), 1,len(rsSQL("RUT"))-2),0)&mid(rsSQL("RUT"), len(rsSQL("RUT"))-1,len(rsSQL("RUT"))),",",".")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;height:30px;vertical-align:middle;">
                        <%=right("0"&day(now()),2)&"/"&right("0"&month(now()),2)&"/"&year(now)%></td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="width:100%;">
           <tr>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:355px;height:30px;vertical-align:middle;">
                        <%=rsSQL("DIRECCION")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:160px;height:30px;vertical-align:middle;">
                        <%=rsSQL("COMUNA")%></td>
               <td style="font-family:Arial;font-size:12px;text-align:left;width:110px;height:30px;vertical-align:middle;">
                        <%=rsSQL("FONO")%></td>
              <td style="font-family:Arial;font-size:12px;text-align:left;width:100px;height:30px;">
                        <%=rsSQL("CIUDAD")%></td>
            </tr>
             <tr>
              <td colspan="5" style="font-family:Arial;font-size:12px;text-align:left;width:100%;height:30px;vertical-align:middle;"><%=rsSQL("GIRO")%></td>
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
                <td colspan="2" style="text-align:left;font-family:Arial;font-size:12px;"><%=rsSQL("DESCRIPCION")%></td>
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
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsSQL("CODIGO")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">N&uacute;mero de Horas : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsSQL("HORAS")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">Fecha de Inicio :</td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsSQL("FECHA_INICIO")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">Fecha de T&eacute;rmino : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsSQL("FECHA_TERMINO")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">N&uacute;mero de Participantes : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=rsSQL("N_PARTICIPANTES")%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">Valor por Participantes : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;"><%=FormatNumber(rsSQL("VALOR_CURSO"),0)%></td>
              </tr>
              <tr>
                <td style="text-align:left;font-family:Arial;font-size:12px;">N&uacute;mero Orden de Compra : </td>
                <td style="text-align:left;font-family:Arial;font-size:12px;">&nbsp;</td>
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
            <td style="text-align:right;font-family:Arial;font-size:12px;"><%=FormatNumber(rsSQL("VALOR_OC"),0)%></td>
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
                        <%=CONVERTIR(cdbl(FormatNumber("140000",0)))&" pesos.-"%>
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
              <td style="font-family:Arial;font-size:12px;text-align:right"><%=FormatNumber("140000",0)%></td>
            </tr>
        </table>
        </td>
   </tr>
</table></td>
	</tr>
</table>
</div>