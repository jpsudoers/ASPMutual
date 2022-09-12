<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
'Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
vId  = Request("txtId")
vrSocial=Request("txtRsoc")
vgiro=Request("txtGiro")
vdir=Request("txtDir")
vcom=Request("txtCom")
vciu=Request("txtCiu")
vfon=Request("txtFon")

vEDP = Request("cond_edp")
vOC = Request("cond_oc")
vTransferencia = Request("cond_transferencia")
vDepositoBanco = Request("cond_depositoBanco")

vCond_ref_hes = Request("cond_ref_hes")
vCond_ref_migo = Request("cond_ref_migo")

vCon_ref = 0
if (vOC = "1" and (vCond_ref_hes="1" or vCond_ref_migo="1")) then
	vCon_ref = 1
end if

if (vCond_ref_hes<>1) then
	vCond_ref_hes=0
end if

if (vCond_ref_migo<>1) then
	vCond_ref_migo=0
end if	

if(vEDP<>1)then
	vEDP=0
end if
if(vOC<>1)then
	vOC=0
	vCond_ref_hes=0
	vCond_ref_migo=0
end if
if(vTransferencia<>1)then
	vTransferencia=0
end if
if(vDepositoBanco<>1)then
	vDepositoBanco=0
end if

vfax="NULL"
if(Request("txtFax")<>"")then
	vfax="'"&Request("txtFax")&"'"
end if

vmut=Request("txtMut")
votic=Request("OTIC") 

vnomb=Request("txtNomb")
vmail=Request("txtMail") 
vcargo=Request("txtCargo")  

vnombconta=Request("txtNombConta")
vmailconta=Request("txtMailConta") 
vcargoconta=Request("txtCargoConta") 

vfonocont=Request("txtFonoCont") 
vcargocontafono=Request("txtContaFono") 

vpasscord=Request("txtPassCord") 

vpassconta="NULL"
if(Request("txtPassConta")<>"")then
	vpassconta="LOWER('" & Request("txtPassConta") &"')"
end if

vtipocontacto=Request("contactoIgual") 

query = "UPDATE EMPRESAS SET R_SOCIAL=dbo.MayMinTexto('"&vrSocial&"'), GIRO=dbo.MayMinTexto('"&vgiro&"'), "
query = query&"DIRECCION=dbo.MayMinTexto('"&vdir&"'), COMUNA=dbo.MayMinTexto('"&vcom&"'), "
query = query&"CIUDAD=dbo.MayMinTexto('"&vciu&"'), FONO='"&vfon&"', FAX="&vfax&", MUTUAL='"&vmut&"', "
query = query&"NOMBRES=dbo.MayMinTexto('"&vnomb&"'), EMAIL=LOWER('"&vmail&"'), CARGO=dbo.MayMinTexto('"&vcargo&"'), ID_OTIC='"&votic&"', "
query = query&"NOMBRE_CONTA=dbo.MayMinTexto('"&vnombconta&"'),EMAIL_CONTA=LOWER('"&vmailconta&"'),"
query = query&"CARGO_CONTA=dbo.MayMinTexto('"&vcargoconta&"'),FONO_CONTACTO='"&vfonocont&"',FONO_CONTABILIDAD='"&vcargocontafono&"', "
query = query&"PASSWORD_COORDINACION=LOWER('"&vpasscord&"'), PASSWORD_CONTA="&vpassconta&", TIPO_CONTACTO='"&vtipocontacto&"', "
query = query&"CON_REFERENCIA=" & vCon_ref & " "
query = query&" WHERE id_empresa = '"&vid&"'"

'response.Write(query)
'response.End()

conn.execute (query)

conn.execute ("DELETE  EMPRESA_CONDICION_COMERCIAL WHERE  id_empresa = '"&vid&"'")

if vEDP = 1 then
query3 = "INSERT INTO EMPRESA_CONDICION_COMERCIAL (ID_EMPRESA, ID_CONDICION_COMERCIAL, FECHA_REGISTRO) "
query3 = query3 &" values('"&vid&"',5,getdate() ) "
conn.execute (query3)
end if
if vOC = 1 then
query4 = "INSERT INTO EMPRESA_CONDICION_COMERCIAL (ID_EMPRESA, ID_CONDICION_COMERCIAL, FECHA_REGISTRO, CON_HES, CON_MIGO) "
query4 = query4 &" values('"&vid&"',0,getdate()," & vCond_ref_hes & "," & vCond_ref_migo & ") "
conn.execute (query4)
end if

if vTransferencia = 1 then
query5 = "INSERT INTO EMPRESA_CONDICION_COMERCIAL (ID_EMPRESA, ID_CONDICION_COMERCIAL, FECHA_REGISTRO) "
query5 = query5 &" values('"&vid&"',3,getdate() ) "
conn.execute (query5)
end if

if vDepositoBanco = 1 then
query6 = "INSERT INTO EMPRESA_CONDICION_COMERCIAL (ID_EMPRESA, ID_CONDICION_COMERCIAL, FECHA_REGISTRO) "
query6 = query6 &" values('"&vid&"',2,getdate() ) "
conn.execute (query6)
end if


Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>