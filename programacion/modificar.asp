<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
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

query = "UPDATE PROGRAMA SET ID_MUTUAL='"&vcurriculo&"', SENCE='"&vsence&"',TIPO='"&vtipo&"',"
query = query&"FECHA_APERTURA=CONVERT(datetime,'"&vfechapertura&"',105),FECHA_CIERRE=CONVERT(datetime,'"&vfechcierre&"',105), "
query = query&"FECHA_INICIO_=CONVERT(datetime,'"&vfechinicio&"',105),FECHA_TERMINO=CONVERT(datetime,'"&vfechtermino&"',105), "
query = query&"CUPOS='"&vcupo&"', INSCRITOS='"&vinscritos&"', VACANTES='"&vvacantes&"', ID_EMPRESA="&vid_empresa&", "
query = query&"DIR_EJEC="&vtxlugar&", VALOR_ESPECIAL="&vvalesp&", MONTO_TOTAL="&vvaltot&", BMI="&vtxtBMInicio&", BTF="&vtxtBTFin&", ID_Modalidad="&vtxtIDModalidad&", ID_ZOOM="&vtxtIDZoom&"  WHERE ID_PROGRAMA = '"&vid&"'"

'response.Write(query)
'response.End()

conn.execute (query)

conn.execute ("insert into Log_Update_Insert values("&Session("usuarioMutual")&",getdate(),'"&Replace(query,"'","`")&"','"&Request.ServerVariables("PATH_INFO")&"','Programacion de Cursos');")

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>