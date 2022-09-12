<!--#include file="../conexion.asp"-->
<%
fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)
Response.ContentType ="application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=CREDENCIAL_"&fecha&".xls"

	
	qs="select UPPER (EMPRESAS.R_SOCIAL)as 'RSOCIAL',EMPRESAS.RUT AS RUTE,nomPart=UPPER (TRABAJADOR.NOM_TRAB),apPart=UPPER (TRABAJADOR.APATERTRAB + ' ' + TRABAJADOR.AMATERTRAB),"
	qs=qs&"UPPER (TRABAJADOR.NOMBRES)as 'PARTICIPANTE', TRABAJADOR.RUT, TRABAJADOR.NACIONALIDAD, TRABAJADOR.ID_EXTRANJERO,"
	qs=qs&"PROGRAMA.FECHA_INICIO_,PROGRAMA.FECHA_TERMINO, sedes.NOMBRE, "
	qs=qs&"CF=(select CASE WHEN A2.CON_FRANQUICIA=1 THEN 'SI' ELSE 'NO' END  from autorizacion a2 "
	qs=qs&" where a2.ID_AUTORIZACION=HISTORICO_CURSOS.ID_AUTORIZACION), TRABAJADOR.CORREO, TRABAJADOR.EMAIL, "
qs=qs&"COD_AUTENFICACION=(select CASE WHEN T.NACIONALIDAD='0' then " 
qs=qs&" right('00000' + CONVERT(VARCHAR,HC.ID_PROGRAMA), 5)+ "
qs=qs&" right('0000000' + CONVERT(VARCHAR,HC.ID_HISTORICO_CURSO), 7)+'-'+ "
qs=qs&" right('0000000000' + CONVERT(VARCHAR,T.RUT), 10) "
qs=qs&" WHEN T.NACIONALIDAD='1' then "
qs=qs&" right('00000' + CONVERT(VARCHAR,HC.ID_PROGRAMA), 5)+ "
qs=qs&" right('0000000' + CONVERT(VARCHAR,HC.ID_HISTORICO_CURSO), 7)+'-'+ "
qs=qs&" right('0000000000' + CONVERT(VARCHAR,T.ID_EXTRANJERO), 10) "
qs=qs&" END from HISTORICO_CURSOS HC "
qs=qs&" inner join TRABAJADOR T on T.ID_TRABAJADOR=HC.ID_TRABAJADOR "
qs=qs&" where HC.ID_HISTORICO_CURSO=HISTORICO_CURSOS.ID_HISTORICO_CURSO),'idHistorico'=HISTORICO_CURSOS.ID_HISTORICO_CURSO "
	qs=qs&" from HISTORICO_CURSOS "
	qs=qs&" inner join trabajador on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
	qs=qs&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=HISTORICO_CURSOS.ID_EMPRESA "
	qs=qs&" inner join programa on programa.ID_PROGRAMA=HISTORICO_CURSOS.ID_PROGRAMA "
	qs=qs&" inner join sedes on sedes.ID_SEDE=HISTORICO_CURSOS.SEDE "
	qs=qs&" where HISTORICO_CURSOS.ID_PROGRAMA='"&Request("prog")&"' and HISTORICO_CURSOS.RELATOR='"&Request("relator")&"'"
	qs=qs&" ORDER BY R_SOCIAL ASC "

	set rs =  conn.execute(qs)%>
<table width="1200" border="1">
<tr>
  <td width="150" align="center"><b><font size="2">RUT</font></b></td>
  <td width="400" align="center"><b><font size="2">RAZ&Oacute;N SOCIAL</font></b></td>
  <td width="40">&nbsp;</td>
  <td width="300" align="center"><b><font size="2">NOMBRES PARTICIPANTE</font></b></td>
  <td width="300" align="center"><b><font size="2">APELLIDOS PARTICIPANTE</font></b></td>
  <td width="150" align="center"><b><font size="2">RUT/ID EXTRANJERO</font></b></td>
  <td width="130" align="center"><b><font size="2">TEL&Eacute;FONO</font></b></td>
  <td width="230" align="center"><b><font size="2">EMAIL</font></b></td>
  <td width="130" align="center"><b><font size="2">NOTA</font></b></td>
  <td width="130" align="center"><b><font size="2">CON FRANQUICIA</font></b></td>
  <td width="300" align="center"><b><font size="2">COD_AUTENFICACION</font></b></td>
  <td width="300" align="center"><b><font size="2">ID HISTORICO</font></b></td>
</tr>
<%i=1
while not rs.eof
fechaInicio=rs("FECHA_INICIO_")
fechaTermino=rs("FECHA_TERMINO")
Sala=rs("NOMBRE")
%>
<tr>
  <td><b><font size="2"><%=rs("RUTE")%></font></b></td>
  <td  align="left"><b><font size="2"><%=rs("RSOCIAL")%></font></b></td>
  <td><b><font size="2"><%=i%></font></b></td>
  <td><b><font size="2"><%=rs("nomPart")%></font></b></td>
    <td><b><font size="2"><%=rs("apPart")%></font></b></td>
  <%if(rs("NACIONALIDAD")="0")then%>
  <td align="right"><b><font size="2"><%=rs("rut")%></font></b></td>
  <%else%>
  <td align="right"><b><font size="2" color="#FF0000"><%=rs("ID_EXTRANJERO")%></font></b></td>
  <%end if%>
  <td><b><font size="2"><%=rs("CORREO")%></font></b></td>
  <td><b><font size="2"><%=rs("EMAIL")%></font></b></td>
  <td><b><font size="2">&nbsp;</font></b></td>
  <td align="center"><b><font size="2"><%=rs("CF")%></font></b></td>
  <td align="center"><b><font size="2"><%=rs("COD_AUTENFICACION")%></font></b></td>
  <td align="center"><b><font size="2"><%=rs("idHistorico")%></font></b></td>
</tr>	
<%
 i=i+1
	rs.MoveNext
wend

fechaSala=""
mesInicio=""
mesTermino=""
if(fechaInicio<>"")then

if(month(fechaInicio)="1")then
	mesInicio="Enero"
end if

if(month(fechaInicio)="2")then
	mesInicio="Febrero"
end if

if(month(fechaInicio)="3")then
	mesInicio="Marzo"
end if

if(month(fechaInicio)="4")then
	mesInicio="Abril"
end if

if(month(fechaInicio)="5")then
	mesInicio="Mayo"
end if

if(month(fechaInicio)="6")then
	mesInicio="Junio"
end if

if(month(fechaInicio)="7")then
	mesInicio="Julio"
end if

if(month(fechaInicio)="8")then
	mesInicio="Agosto"
end if

if(month(fechaInicio)="9")then
	mesInicio="Septiembre"
end if

if(month(fechaInicio)="10")then
	mesInicio="Octubre"
end if

if(month(fechaInicio)="11")then
	mesInicio="Noviembre"
end if

if(month(fechaInicio)="12")then
	mesInicio="Diciembre"
end if



if(month(fechaTermino)="1")then
	mesTermino="Enero"
end if

if(month(fechaTermino)="2")then
	mesTermino="Febrero"
end if

if(month(fechaTermino)="3")then
	mesTermino="Marzo"
end if

if(month(fechaTermino)="4")then
	mesTermino="Abril"
end if

if(month(fechaTermino)="5")then
	mesTermino="Mayo"
end if

if(month(fechaTermino)="6")then
	mesTermino="Junio"
end if

if(month(fechaTermino)="7")then
	mesTermino="Julio"
end if

if(month(fechaTermino)="8")then
	mesTermino="Agosto"
end if

if(month(fechaTermino)="9")then
	mesTermino="Septiembre"
end if

if(month(fechaTermino)="10")then
	mesTermino="Octubre"
end if

if(month(fechaTermino)="11")then
	mesTermino="Noviembre"
end if

if(month(fechaTermino)="12")then
	mesTermino="Diciembre"
end if

fechaSala="desde el "&right("0"&day(fechaInicio),2)&" de "&mesInicio&" hasta el "&right("0"&day(fechaTermino),2)&" de "&mesTermino&", "&Sala&"."
end if
%>
<tr>
  <td colspan="5" align="center"><b><font size="3"><%=fechaSala%></font></b></td>
</tr>
</table>