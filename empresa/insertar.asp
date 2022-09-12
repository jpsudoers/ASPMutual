<%Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<insertar>") %>
<!--#include file="../conexion.asp"-->
<%
on error resume next
vrut = Request("txRut")
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
if (vCond_ref_hes=1 or vCond_ref_migo=1) then
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

'dim query
'query = "IF NOT EXISTS (select * from EMPRESAS e where e.rut='"&vrut&"' and e.ESTADO=1 and e.PREINSCRITA=1) BEGIN "
'query = query&"insert into EMPRESAS (RUT,R_SOCIAL ,GIRO ,DIRECCION,COMUNA,CIUDAD ,FONO ,FAX ,MUTUAL, ESTADO "
'query = query&",NOMBRES, EMAIL,CARGO,ID_OTIC,NOMBRE_CONTA,EMAIL_CONTA,CARGO_CONTA,FONO_CONTACTO,FONO_CONTABILIDAD,TIPO,"
'query = query&"PASSWORD_COORDINACION,PASSWORD_CONTA,PREINSCRITA,TIPO_CONTACTO,ACTIVA,CON_REFERENCIA) "
'query = query&" values('"&vrut&"',dbo.MayMinTexto('"&vrSocial&"'),dbo.MayMinTexto('"&vgiro&"'),"
'query = query&"dbo.MayMinTexto('"&vdir&"'),dbo.MayMinTexto('"&vcom&"'),dbo.MayMinTexto('"&vciu&"'),'"&vfon&"',"&vfax&", "
'query = query&"'"&vmut&"',1,dbo.MayMinTexto('"&vnomb&"'),LOWER('"&vmail&"'),dbo.MayMinTexto('"&vcargo&"'),'"&votic&"',"
'query = query&"dbo.MayMinTexto('"&vnombconta&"'),LOWER('"&vmailconta&"'),dbo.MayMinTexto('"&vcargoconta&"'),"
'query = query&"'"&vfonocont&"',dbo.MayMinTexto('"&vcargocontafono&"'),1,LOWER('"&vpasscord&"'),"&vpassconta&",1,'"&vtipocontacto&"'," & vCon_ref & ");"
'query = query&" INSERT INTO EMPRESA_TIPO_COMPROMISO(ID_EMPRESA, ID_COMPROMISO_PAGO, ESTADO_EMPRESA_COMPROMISO, FECHA_EMPRESA_COMPROMISO)"
'query = query&" VALUES ((select top 1 e.id_empresa from EMPRESAS e where e.rut='"&vrut&"' and e.ESTADO=1 and e.PREINSCRITA=1 order by e.id_empresa desc), 1, 0, GETDATE());"
'query = query&" END"


dim query
query = "IF NOT EXISTS (select * from EMPRESAS e where e.rut='"&vrut&"' and e.ESTADO=1 and e.PREINSCRITA=1) BEGIN "
query = query&"insert into EMPRESAS (RUT,R_SOCIAL ,GIRO ,DIRECCION,COMUNA,CIUDAD ,FONO ,FAX ,MUTUAL, ESTADO, ID_OTIC "
query = query&",PREINSCRITA, ACTIVA, TIPO, TIPO_CONTACTO, CON_REFERENCIA)"
query = query&" values('"&vrut&"',dbo.MayMinTexto('"&vrSocial&"'),dbo.MayMinTexto('"&vgiro&"'),"
query = query&"dbo.MayMinTexto('"&vdir&"'),dbo.MayMinTexto('"&vcom&"'),dbo.MayMinTexto('"&vciu&"'),'"&vfon&"',"&vfax&", "
query = query&"'"&vmut&"',1,'"&votic&"',1,1,1,1," & vCon_ref & ");"
query = query&" INSERT INTO EMPRESA_TIPO_COMPROMISO(ID_EMPRESA, ID_COMPROMISO_PAGO, ESTADO_EMPRESA_COMPROMISO, FECHA_EMPRESA_COMPROMISO)"
query = query&" VALUES ((select top 1 e.id_empresa from EMPRESAS e where e.rut='"&vrut&"' and e.ESTADO=1 and e.PREINSCRITA=1 order by e.id_empresa desc), 1, 0, GETDATE());"
query = query&" END"
'response.Write(query)
'response.End()

'Response.Write("<sql>"&query&"</sql>")
conn.execute (query)

set rsCond = conn.execute ("select top 1 e.id_empresa from EMPRESAS e where e.rut='"&vrut&"' and e.ESTADO=1 and e.PREINSCRITA=1 order by e.id_empresa desc")
vid=rsCond("id_empresa")
 
 
queryUP = "update EMPRESAS_USUARIOS set ID_EMPRESA="&vid&" where ID_EMPRESA = 0 and Rut_Empresa='"&vrut&"'"
conn.execute(queryUP)

conn.execute ("DELETE  EMPRESA_CONDICION_COMERCIAL WHERE  id_empresa = '"&vid&"'")

if vEDP = 1 then
query3 = "INSERT INTO EMPRESA_CONDICION_COMERCIAL (ID_EMPRESA, ID_CONDICION_COMERCIAL, FECHA_REGISTRO) "
query3 = query3 &" values('"&vid&"',5,getdate() ) "
conn.execute (query3)
end if
if vOC = 1 then
query4 = "INSERT INTO EMPRESA_CONDICION_COMERCIAL (ID_EMPRESA, ID_CONDICION_COMERCIAL, FECHA_REGISTRO) "
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


if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
'Response.Write("<mensaje>"&vmensaje&"</mensaje>")
Response.Write("</insertar>") 
%>