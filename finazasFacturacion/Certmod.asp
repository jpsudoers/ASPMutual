<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/msword"
Response.AddHeader "content-disposition", "inline; filename=cert_"&fecha&".doc"

	IdAuto=Request("IdAuto")

	dim datosProg
datosProg="select E.RUT,UPPER(E.R_SOCIAL) as R_SOCIAL,C.DESCRIPCION,C.CODIGO,CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"
datosProg=datosProg&"CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO,C.HORAS,A.N_REG_SENCE from AUTORIZACION A "
datosProg=datosProg&" INNER JOIN PROGRAMA P ON P.ID_PROGRAMA=A.ID_PROGRAMA "
datosProg=datosProg&" inner join CURRICULO C on C.ID_MUTUAL=P.ID_MUTUAL "
datosProg=datosProg&" INNER JOIN EMPRESAS E ON E.ID_EMPRESA=A.ID_EMPRESA "
datosProg=datosProg&" where A.ID_AUTORIZACION='"&IdAuto&"'"
	
	set rsDatosProg = conn.execute (datosProg)

	set rsMutual = conn.execute ("SELECT RUT,R_SOCIAL,CONTACTO,RUT_CONTACTO FROM INSTITUCION")	
		%>
        
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:"Century Gothic";
	panose-1:2 11 5 2 2 2 2 2 2 4;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0cm;
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman","serif";}
p.MsoFootnoteText, li.MsoFootnoteText, div.MsoFootnoteText
	{mso-style-link:"Texto nota pie Car";
	margin:0cm;
	margin-bottom:.0001pt;
	font-size:10.0pt;
	font-family:"Times New Roman","serif";}
p.MsoHeader, li.MsoHeader, div.MsoHeader
	{margin:0cm;
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman","serif";}
p.MsoFooter, li.MsoFooter, div.MsoFooter
	{margin:0cm;
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman","serif";}
span.MsoFootnoteReference
	{vertical-align:super;}
span.TextonotapieCar
	{mso-style-name:"Texto nota pie Car";
	mso-style-link:"Texto nota pie";}
 /* Page Definitions */
 @page WordSection1
	{size:595.3pt 841.9pt;
	margin:70.9pt 3.0cm 70.9pt 72.0pt;}
div.WordSection1
	{page:WordSection1;}
-->
</style>        
<div class=WordSection1>       
<table border=0 cellspacing=0 cellpadding=0 width=700 style='width:496.15pt;margin-left:-37.15pt;border-collapse:collapse;
 border:none'>
<tr>
	<td><img width=170 height=50 src="http://norte.otecmutual.cl/images/logoMutual.jpg"></td>
</tr>
<tr>
	<td>&nbsp;</td>
</tr>
<tr>
<td>        
<table border=0 cellspacing=0 cellpadding=0 width=700 style='width:496.15pt;margin-left:-37.15pt;border-collapse:collapse;
 border:none'>
 <tr>
 <td><font FACE="Calibri" size="+1"><b><center><span>&nbsp;&nbsp;&nbsp;&nbsp;CERTIFICADO DE ASISTENCIA ACTIVIDAD OTEC, CFT O ENTIDAD
NIVELADORA DE ESTUDIOS, IMPUTADA EN FORMA TOTAL O PARCIAL A
FRANQUICIA TRIBUTARIA DE CAPACITACIÓN</span></center></b></font></td>
</tr>
</table>
</td>
</tr>
 
<tr>
<td>  
<table border=0 cellspacing=0 cellpadding=0 width=700>
<tr>
 <td colspan="2">&nbsp;</td> 
</tr> 
<tr>
 <td  height="30" style='border:solid windowtext 1.0pt;padding:0cm 0.5pt 0cm 0.5pt'><CENTER><B><font FACE="Calibri" size="-1">X</font></B></CENTER></td> 
 <td><font FACE="Calibri" size="-1">&nbsp;&nbsp;Actividad dentro del año calendario</font></td>
</tr> 
<tr>
 <td colspan="2">&nbsp;</td> 
</tr> 
<tr>
 <td  height="30" style='border:solid windowtext 1.0pt;padding:0cm 0.5pt 0cm 0.5pt'>&nbsp;</td> 
 <td><font FACE="Calibri" size="-1">&nbsp;&nbsp;Actividad parcial</font></td>
</tr>
<tr>
 <td colspan="2">&nbsp;</td> 
</tr> 
<tr>
 <td  height="30" style='border:solid windowtext 1.0pt;padding:0cm 0.5pt 0cm 0.5pt'>&nbsp;</td> 
 <td><font FACE="Calibri" size="-1">&nbsp;&nbsp;Actividad complementaria</font></td>
</tr>
<tr>
 <td>&nbsp;</td> 
 <td>&nbsp;</td> 
</tr>
<tr>
 <td colspan="2"><p><font FACE="Calibri" size="-1">Se extiende el presente certificado de asistencia correspondiente a la actividad de capacitación que a continuación se señala:</font></p></td> 
</tr>
<tr>
 <td colspan="2">&nbsp;</td> 
</tr>
</table>
</td>
</tr>

<tr>
<td>  
<table width="700"  style='width:496.15pt;margin-left:-37.15pt;border-collapse:collapse;
 border:none'>
 <tr>
  <td style='border:solid windowtext 1.0pt;width:30pt'><b><font FACE="Calibri" size="-1">Razón social OTEC, CFT o entidad niveladora</font></b></td>
  <td style='border:solid windowtext 1.0pt;width:670pt'><font FACE="Calibri" size="-1"><%=rsMutual("R_SOCIAL")%></font></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">RUT OTEC, CFT o entidad niveladora</font></b></td>
  <td style='border:solid windowtext 1.0pt'><font FACE="Calibri" size="-1"><%=rsMutual("RUT")%></font></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">Razón social empresa</font></b></td>
  <td style='border:solid windowtext 1.0pt'><font FACE="Calibri" size="-1"><%=rsDatosProg("R_SOCIAL")%></font></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">RUT empresa</font></b></td>
  <td style='border:solid windowtext 1.0pt'><font FACE="Calibri" size="-1"><%=replace(FormatNumber(mid(rsDatosProg("RUT"), 1,len(rsDatosProg("RUT"))-2),0)&mid(rsDatosProg("RUT"), len(rsDatosProg("RUT"))-1,len(rsDatosProg("RUT"))),",",".")%></font></td>
 </tr>
 <tr>
  <td height="30" style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">Nombre de la actividad [1]</font></b></td>
  <td height="30" style='border:solid windowtext 1.0pt'><p><font FACE="Calibri" size="-1"><%=rsDatosProg("DESCRIPCION")%></font></p></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">Código Sence</font></b></td>
  <td style='border:solid windowtext 1.0pt'><font FACE="Calibri" size="-1"><%=rsDatosProg("CODIGO")%></font></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">Fecha inicio</font></b></td>
  <td style='border:solid windowtext 1.0pt'><font FACE="Calibri" size="-1"><%=rsDatosProg("FECHA_INICIO")%></font></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">Fecha de término</font></b></td>
  <td style='border:solid windowtext 1.0pt'><font FACE="Calibri" size="-1"><%=rsDatosProg("FECHA_TERMINO")%></font></td>
 </tr>
 <tr>
  <td height="50" style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1"><p>Nº de horas (para actividades parciales o complementarias
   indicar número efectivo de horas realizadas en el año 
   correspondiente)</p></font></b></td>
  <td height="50" style='border:solid windowtext 1.0pt'><font FACE="Calibri" size="-1"><%=rsDatosProg("HORAS")%></font></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">Nº de factura</font></b></td>
  <td style='border:solid windowtext 1.0pt;padding:0cm 0.5pt 0cm 0.5pt'><font FACE="Calibri" size="-1">&nbsp;</font></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt'><b><font FACE="Calibri" size="-1">Nº registro de acción Sence</font></b></td>
  <td style='border:solid windowtext 1.0pt;padding:0cm 0.5pt 0cm 0.5pt'><font FACE="Calibri" size="-1"><%=rsDatosProg("N_REG_SENCE")%></font></td>
 </tr>
</table>
</td>
</tr>
<tr>
  <td>&nbsp;</td>
 </tr>
 <tr>
  <td><font FACE="Calibri" size="-1">Participantes:</font></td>
 </tr>
 <tr>
  <td>&nbsp;</td>
</tr>
<tr>
<td>  
<table cellspacing=0 cellpadding=0 width=700 style='width:496.15pt;margin-left:-37.15pt;border-collapse:collapse;
 border:none'>
  <tr>
  <td width=30 style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-1">Nº</font></b></td>
  <td width=90 style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-1">RUT</font></b></td>
  <td width=160 style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-1">Apellido paterno</font></b></td>
  <td width=160 style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-1">Apellido materno</font></b></td>
  <td width=160 style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-1">Nombres</font></b></td>
  <td width=100 style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-1">Porcentaje de asistencia</font></b></td>
 </tr>
 
 <%
 	dim datosTrab
	datosTrab="select (CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',"
	datosTrab=datosTrab&"T.APATERTRAB,T.AMATERTRAB,T.NOM_TRAB,H.ASISTENCIA,NACIONALIDAD from HISTORICO_CURSOS H "
	datosTrab=datosTrab&" inner join TRABAJADOR T on T.ID_TRABAJADOR=H.ID_TRABAJADOR " 
	datosTrab=datosTrab&" WHERE H.ID_AUTORIZACION='"&IdAuto&"'"

	set rsTrab = conn.execute (datosTrab)
	cont=0
		while not rsTrab.eof
		cont=cont+1
 %>

 <tr>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><font FACE="Calibri" size="-1"><%=cont%></font></td>  
  
  				<%if(rsTrab("NACIONALIDAD")="0")then%>
                     <td align="right" style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><font FACE="Calibri" size="-1"><%=replace(FormatNumber(mid(rsTrab("TrabId"), 1,len(rsTrab("TrabId"))-2),0)&mid(rsTrab("TrabId"), len(rsTrab("TrabId"))-1,len(rsTrab("TrabId"))),",",".")%></font></td> 	
				<%else%>
                     <td align="right" style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><font FACE="Calibri" size="-1"><%=rsTrab("TrabId")%></font></td> 
				<%end if%>	
  
 
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><font FACE="Calibri" size="-1"><%=""&rsTrab("APATERTRAB")%></font></td> 
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><font FACE="Calibri" size="-1"><%=""&rsTrab("AMATERTRAB")%></font></td> 
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><font FACE="Calibri" size="-1"><%=""&rsTrab("NOM_TRAB")%></font></td> 
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><center><font FACE="Calibri" size="-1"><%=""&rsTrab("ASISTENCIA")%></font></center></td>             
 </tr>
 <%
 			rsTrab.Movenext
	wend	
 %>
 
</table>
</td>
</tr>
 <tr>
  <td>&nbsp;</td>
 </tr> 
<tr>
<td>  
<table cellspacing=0 cellpadding=0 width="500">
 <tr>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-2">Firma representante legal OTEC, CFT o entidad niveladora [2]</font></b></td>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt' width=300>&nbsp;</td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-2">Nombre representante legal OTEC, CFT o entidad niveladora</font></b></td>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-2">Héctor Garay Carvajal</font></b></td>
 </tr>
 <tr>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-2">RUT representante legal OTEC, CFT o entidad niveladora</font></b></td>
  <td style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'><b><font FACE="Calibri" size="-2">6.164.330-3</font></b></td>
 </tr>
</table>
</td>
</tr>

 <tr>
  <td>&nbsp;</td>
 </tr>

<tr>
<td>  
<table cellspacing=0 cellpadding=0 align=right width=180>
 <tr>
  <td height="60" style='border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt'>
<CENTER><font FACE="Calibri" size="-1">&nbsp;Fecha emisión: <%=right("0"&day(now()),2)&"/"&right("0"&month(now()),2)&"/"&year(now)%></font>
</CENTER></td>
 </tr>
</table>
</td>
</tr>

 <tr>
  <td>&nbsp;</td>
 </tr>
 
  <tr>
  <td><center><img width=600 height=80 src="http://norte.otecmutual.cl/images/logoCertSence.jpg"></center></td>
 </tr>
</table>
</div>
