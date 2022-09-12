<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vcurriculo = Request("Curriculo")
vsence=Request("Sence")
vtipo=Request("Tipo")
vfechapertura=Request("txtFechApertura")
vfechcierre=Request("txtFechCierre")
vfechinicio=Request("txtFechInicio")
vfechtermino=Request("txtFechTermino")
vcupo=Request("txtCupo")
vinscritos=Request("txtInscritos")
vvacantes=Request("txtVacantes")
vprogId=Request("tabFecha")

if(Request("id_empresa")<>"")then
vid_empresa=Request("id_empresa")
else
vid_empresa="NULL"
end if

if(trim(Request("txLugar"))<>"")then
vtxlugar="dbo.MayMinTexto('"&trim(Request("txLugar"))&"')"
else
vtxlugar="NULL"
end if

if(Request("txtValEsp")<>"")then
vvalesp="'"&Request("txtValEsp")&"'"
else
vvalesp="NULL"
end if

if(Request("txtValTot")<>"")then
vvaltot="'"&Request("txtValTot")&"'"
else
vvaltot="NULL"
end if

vtxtBMInicio="0"
if(Request("txtBMInicio")<>"")then
	vtxtBMInicio=Request("txtBMInicio")
end if

vtxtBTFin="0"
if(Request("txtBTFin")<>"")then
	vtxtBTFin=Request("txtBTFin")
end if

vtxtIDZoom="NULL"
if(Request("txtIDZoom")<>"")then
	vtxtIDZoom=Request("txtIDZoom")
end if

vtxtIDModalidad=Request("txtIDModalidad")

dim query
query = "IF EXISTS (select * from bloque_programacion bp where bp.id_programa='"&vprogId&"') BEGIN "
query = query&"insert into PROGRAMA (ID_MUTUAL,SENCE,TIPO,FECHA_APERTURA,FECHA_CIERRE"
query = query&",FECHA_INICIO_,FECHA_TERMINO,CUPOS,INSCRITOS,VACANTES,ESTADO,VIGENCIA,ID_EMPRESA,DIR_EJEC,VALOR_ESPECIAL,MONTO_TOTAL,BMI,BMF,BTI,BTF,ID_Modalidad,ID_ZOOM) "
query = query&" values('"&vcurriculo&"','"&vsence&"','"&vtipo&"',CONVERT(datetime,'"&vfechapertura&"',105)"
query = query&",CONVERT(datetime,'"&vfechcierre&"',105),CONVERT(datetime,'"&vfechinicio&"',105),"
query = query&"CONVERT(datetime,'"&vfechtermino&"',105),'"&vcupo&"',0,'"&vvacantes&"',1,1"
query = query&","&vid_empresa&","&vtxlugar&","&vvalesp&","&vvaltot&","&vtxtBMInicio&",Null,Null,"&vtxtBTFin&","&vtxtIDModalidad&","&vtxtIDZoom&") END "

'response.Write(query)
'response.End()

conn.execute (query)

conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(query,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Programacion de Cursos');")

'response.Write("<sql2>insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),n'"&Replace(query,"'","`")&"');</sql2>")
'response.End()


set rsProg = conn.execute ("select IDENT_CURRENT('PROGRAMA')AS UltProg")

query2="update bloque_programacion set id_programa='"&rsProg("UltProg")&"' where id_programa='"&vprogId&"'"

conn.execute (query2)

conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(query2,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Programacion de Cursos');")

Response.Write("<sql>"&query&"</sql>")

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("<mensaje>"&vmensaje&"</mensaje>")
Response.Write("</insertar>") 
%>