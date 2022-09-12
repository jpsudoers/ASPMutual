<!--#include file="../funciones/xelupload.asp"-->
<!--#include file="../conexion.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload
objUpload.Upload()

vempresa = objUpload.Form("Empresa")
vprograma = objUpload.Form("programa")
vempman = objUpload.Form("EmpMan")
vcompromiso = objUpload.Form("Compromiso")
vnum = objUpload.Form("txtNum")
vparticipantes = objUpload.Form("NParticipantes")
vvalor = objUpload.Form("Valor")
vvalortotal = objUpload.Form("ValorTotal")
vSelect = objUpload.Form("selectOC")

vcosto = objUpload.Form("Costo")
vtipoempresa = objUpload.Form("TipoEmpresa")
vconotic = objUpload.Form("ConOtic")

vcofran = objUpload.Form("COfran")
nom_arch = ""
if(vnum="") then 
	vnum = vSelect
end if

vreg_sence = ""
if(objUpload.Form("NReg_Fran")<>"")then
	vreg_sence = objUpload.Form("NReg_Fran")
else
	vreg_sence = "NULL"
end if

vpropert = ""
if(objUpload.Form("proPert")<>"")then
	vpropert = "'"&objUpload.Form("proPert")&"'"
else
	vpropert = "NULL"
end if

set objFich = objUpload.Ficheros("txtDoc")

aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

if(ext <> "") then
nom_arch="Documento_"&fecha&"_"&vempresa&"."&ext
end if
vtabProgId = objUpload.Form("tabProgId")

vTipoVenta = "1"

if(objUpload.Form("tipoVenta")<>"")then
	vTipoVenta = objUpload.Form("tipoVenta")
end if

if(objUpload.Form("Banco")<>"")then
	vid_banco = objUpload.Form("Banco")
else
	vid_banco = "NULL"
end if

if(objUpload.Form("numeroCuenta")<>"")then
		vid_cuenta_corriente = objUpload.Form("numeroCuenta")
else
		vid_cuenta_corriente = "NULL"
end if

if(objUpload.Form("txtMonto")<>"")then
	vmonto_transferencia = objUpload.Form("txtMonto")
else
	vmonto_transferencia = "NULL"
end if


if(objUpload.Form("txtFechaDeposito")<>"")then
	vfecha_transferencia = " CONVERT(datetime,'"&objUpload.Form("txtFechaDeposito")&"',105) "
else
	vfecha_transferencia = "NULL"
end if

if(objUpload.Form("txtRutDepositante")<>"")then
	vrut_depositante = "'"&objUpload.Form("txtRutDepositante")&"'"
else
	vrut_depositante = "NULL"
end if

if(objUpload.Form("txtNombreDepositante")<>"")then
	vnombre_depositante = "'"&objUpload.Form("txtNombreDepositante")&"'"
else
	vnombre_depositante = "NULL"
end if

if(objUpload.Form("txtEmail")<>"")then
	vemail_depositante = "'"&objUpload.Form("txtEmail")&"'"
else
	vemail_depositante = "NULL"
end if

if(objUpload.Form("txtRelacion")<>"")then
	vrelacion_depositante = "'"&objUpload.Form("txtRelacion")&"'"
else
	vrelacion_depositante = "NULL"
end if

if (vSelect <> "") then 


dim empresaoc
empresaoc = "update EMPRESAS_OC"
empresaoc = empresaoc&" set MONTOUTILIZADO = (MONTOUTILIZADO + try_cast("&vvalortotal 
empresaoc= empresaoc&" as int)) where NRO_OC ='"&vSelect&"' and id_empresa="&vempresa

conn.execute(empresaoc)
dim Narchivo
Narchivo = "select NOMBREARCHIVO from EMPRESAS_OC where NRO_OC ='"&vSelect&"' and ID_EMPRESA ="&vempresa


set narchi = conn.execute(Narchivo)

end if



if(nom_arch = "") then 
	nom_arch = narchi("NOMBREARCHIVO")
end if


dim preins
preins = "IF NOT EXISTS (select p.* from PREINSCRIPCIONES p where p.cod_preinscripcion='"&vtabProgId&"') BEGIN "
preins = preins&"insert into PREINSCRIPCIONES(id_programacion,id_empresa,id_otic,ESTADO,tipo_compromiso,numero_compromiso,"
preins = preins&"doc_compromiso,fecha_preinscripcion,participantes,valor,valor_total,costo,tipo_empresa,con_otic,"
preins = preins&"con_franquicia,n_reg_sence,id_proyecto,cod_preinscripcion,id_tipo_venta,id_banco, id_cuenta_corriente, monto_transferencia, fecha_transferencia, rut_depositante, nombre_depositante, email_depositante, relacion_depositante) "
preins = preins&" values('"&vprograma&"','"&vempresa&"','"&vempman&"',1,'"&vcompromiso&"','"&vnum&"','"&nom_arch&"',GETDATE (),"
preins = preins&"'"&vparticipantes&"','"&vvalor&"','"&vvalortotal&"','"&vcosto&"','"&vtipoempresa&"','"&vconotic&"',"
preins = preins&"'"&vcofran&"',"&vreg_sence&","&vpropert&",'"&vtabProgId&"','"&vTipoVenta&"',"&vid_banco&","&vid_cuenta_corriente&","&vmonto_transferencia&","&vfecha_transferencia&","&vrut_depositante&", "&vnombre_depositante&", "&vemail_depositante&" , "&vrelacion_depositante&" ) END"
conn.execute (preins)

set rs = conn.execute ("select IDENT_CURRENT('PREINSCRIPCIONES')AS UltPreIns")

dim preIns_Trab
preIns_Trab = "update PREINSCRIPCION_TRABAJADOR set id_preinscripcion='"&rs("UltPreIns")&"'"
preIns_Trab = preIns_Trab&",preinscripcionTemp=null where preinscripcionTemp='"&vtabProgId&"'"



conn.execute (preIns_Trab)

conn.execute ("exec [dbo].[ACTUALIZA_INSCRIPCIONES_AUTORIZADAS] 1")

objFich.GuardarComo nom_arch,Server.MapPath("../ordenes")

set oFich = nothing
set objUpload = nothing


Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

' send one copy with Google SMTP server (with autentication)
schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com" 
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "notificaciones@otecmutual.cl"
Flds.Item(schema & "sendpassword") =  "admin_2019"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

'creamos el nombre del archivo 
archivo= Server.MapPath("mail.html")
'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 
'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 
'cerramos el fichero 
fich.close() 

set rsEmpresa = conn.execute ("select R_SOCIAL from EMPRESAS where ID_EMPRESA='"&vempresa&"'")

texto_fichero = Replace(texto_fichero,"<!--#empresa-->",rsEmpresa("R_SOCIAL"))

dim queryProg
queryProg = "select CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_ from PROGRAMA "
queryProg = queryProg&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL where PROGRAMA.ID_PROGRAMA='"&vprograma&"'"

set rsQueryProg = conn.execute (queryProg)

texto_fichero = Replace(texto_fichero,"<!--#curso-->",rsQueryProg("NOMBRE_CURSO"))
texto_fichero = Replace(texto_fichero,"<!--#fecha-->",rsQueryProg("FECHA_INICIO_"))

texto_fichero = Replace(texto_fichero,"<!--#asistentes-->",vparticipantes)  
texto_fichero = Replace(texto_fichero,"<!--#valor-->","$ "&replace(FormatNumber(vcosto,0),",","."))

texto_fichero = Replace(texto_fichero,"<!--#doc_tipo-->","http://norte.otecmutual.cl/ordenes/"&nom_arch)

'response.End()

dim queryTabla
queryTabla="select TRABAJADOR.RUT,dbo.MayMinTexto(TRABAJADOR.NOMBRES) as NOMBRES from PREINSCRIPCION_TRABAJADOR "
queryTabla=queryTabla&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=PREINSCRIPCION_TRABAJADOR.id_trabajador "
queryTabla=queryTabla&" where PREINSCRIPCION_TRABAJADOR.id_preinscripcion='"&rs("UltPreIns")&"' order by TRABAJADOR.NOMBRES asc "

set rsQueryTabla = conn.execute (queryTabla)

dim textoTabla
textoTabla=""
while not rsQueryTabla.eof
	textoTabla=textoTabla&"<tr>"
	textoTabla=textoTabla&"<td>"&replace(FormatNumber(mid(rsQueryTabla("RUT"), 1,len(rsQueryTabla("RUT"))-2),0)&mid(rsQueryTabla("RUT"), len(rsQueryTabla("RUT"))-1,len(rsQueryTabla("RUT"))),",",".")&"</td>"
	textoTabla=textoTabla&"<td colspan='2'>"&rsQueryTabla("NOMBRES")&"</td>"
	textoTabla=textoTabla&"</tr>"
rsQueryTabla.Movenext
wend
texto_fichero = Replace(texto_fichero,"<!--#tabla-->",textoTabla)  

if(vcompromiso=0)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Orden de Compra")  
end if

if(vcompromiso=1)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Vale Vista")  
end if

if(vcompromiso=2)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Depósito Cheque")  
end if

if(vcompromiso=3)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Transferencia")  
end if

if(vcompromiso=4)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Carta Compromiso")  
end if

'.To = "notificaciones@otecmutual.cl"

With iMsg
.To = "notificaciones@otecmutual.cl"
.From = "Solicitud de Inscripción a Curso <notificaciones@otecmutual.cl>"
.Subject = "Solicitud de Inscripción a Curso"
.HTMLBody = texto_fichero
.Sender = "Mutual Capacitación"
.Organization = "Mutual Capacitación"
Set .Configuration = iConf
SendEmailGmail = .Send
End With

set iMsg = nothing
set iConf = nothing
set Flds = nothing
%>