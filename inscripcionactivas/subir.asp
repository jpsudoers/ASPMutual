<!--#include file="../funciones/xelupload.asp"-->
<!--#include file="../conexion.asp"-->
<%
on error resume next
Dim objUpload, objFich

set objUpload = new xelUpload
objUpload.Upload()

vcompromiso = objUpload.Form("Compromiso")
vorden = objUpload.Form("txtNOrden")
votic = objUpload.Form("txtIdOtic")
vvalortotal = objUpload.Form("ValorTotal")
vValor = objUpload.Form("Valor")
vnparticipantes = objUpload.Form("NParticipantes")
vbloque = objUpload.Form("txtBloque")
vsala = objUpload.Form("txtSala")
vrelator = objUpload.Form("txtRelator")
vempresa = objUpload.Form("Empresa")
vprograma = objUpload.Form("Programa")
vautorizacionid = objUpload.Form("autorizacionId")
vtabpart = objUpload.Form("tabFechaPart")
vconotic = objUpload.Form("ConOtic")
vcofran = objUpload.Form("COfran")

vTpReferencia = objUpload.Form("TpReferencia")
vtxtReferencia = objUpload.Form("txtReferencia")

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


if(objUpload.Form("NReg_Fran")<>"")then
	vreg_fran = objUpload.Form("NReg_Fran")
else
	vreg_fran = "NULL"
end if

vTipoVenta = objUpload.Form("tipoVenta")

dim regAutoHist
regAutoHist="select AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.ORDEN_COMPRA, "
regAutoHist=regAutoHist&"AUTORIZACION.VALOR_CURSO,AUTORIZACION.VALOR_OC,AUTORIZACION.DOCUMENTO_COMPROMISO, "
regAutoHist=regAutoHist&"AUTORIZACION.TIPO_DOC,AUTORIZACION.N_PARTICIPANTES,AUTORIZACION.ID_PROGRAMA, "
regAutoHist=regAutoHist&"HISTORICO_CURSOS.ID_HISTORICO_CURSO,HISTORICO_CURSOS.ID_TRABAJADOR, "
regAutoHist=regAutoHist&"HISTORICO_CURSOS.RELATOR,HISTORICO_CURSOS.SEDE,HISTORICO_CURSOS.ID_BLOQUE,AUTORIZACION.CON_OTIC,"
regAutoHist=regAutoHist&"AUTORIZACION.CON_FRANQUICIA,AUTORIZACION.N_REG_SENCE "
regAutoHist=regAutoHist&" from AUTORIZACION "
regAutoHist=regAutoHist&" inner join HISTORICO_CURSOS on HISTORICO_CURSOS.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION "
regAutoHist=regAutoHist&" where AUTORIZACION.ID_AUTORIZACION='"&vautorizacionid&"'"

set rsAutoHist = conn.execute (regAutoHist)

dim regInsc

	while not rsAutoHist.eof
		regInsc = "INSERT INTO REG_INSCRIPCION(ID_AUTORIZACION,ORDEN_COMPRA,VALOR_CURSO,VALOR_OC,DOCUMENTO_COMPROMISO,"
		regInsc = regInsc&"TIPO_DOC,N_PARTICIPANTES,ID_PROGRAMA,ID_HISTORICO_CURSO,ID_TRABAJADOR,RELATOR,SEDE,ID_BLOQUE,FECHA,"
		regInsc = regInsc&"USUARIO,CON_OTIC,CON_FRANQUICIA,N_REG_SENCE) VALUES('"&rsAutoHist("ID_AUTORIZACION")&"'"
		regInsc = regInsc&",'"&rsAutoHist("ORDEN_COMPRA")&"','"&rsAutoHist("VALOR_CURSO")&"','"&rsAutoHist("VALOR_OC")&"',"
		regInsc = regInsc&"'"&rsAutoHist("DOCUMENTO_COMPROMISO")&"','"&rsAutoHist("TIPO_DOC")&"',"
		regInsc = regInsc&"'"&rsAutoHist("N_PARTICIPANTES")&"','"&rsAutoHist("ID_PROGRAMA")&"',"
		regInsc = regInsc&"'"&rsAutoHist("ID_HISTORICO_CURSO")&"','"&rsAutoHist("ID_TRABAJADOR")&"',"
		regInsc = regInsc&"'"&rsAutoHist("RELATOR")&"','"&rsAutoHist("SEDE")&"','"&rsAutoHist("ID_BLOQUE")&"',"
		regInsc = regInsc&"GETDATE(),'"&Session("usuarioMutual")&"','"&rsAutoHist("CON_OTIC")&"',"
		regInsc = regInsc&"'"&rsAutoHist("CON_FRANQUICIA")&"',"&rsAutoHist("N_REG_SENCE")&")"

		conn.execute (regInsc)

		rsAutoHist.Movenext
	wend

dim delHistTrab
delHistTrab = "delete from HISTORICO_CURSOS "
delHistTrab = delHistTrab&" where ID_AUTORIZACION='"&vautorizacionid&"' and TRABDEL='"&vtabpart&"'"

conn.execute (delHistTrab)

dim HistUp
HistUp="select HISTORICO_CURSOS.ID_HISTORICO_CURSO,HISTORICO_CURSOS.TRABIDUP from HISTORICO_CURSOS "
HistUp=HistUp&" where HISTORICO_CURSOS.TRABUP='"&vtabpart&"' and HISTORICO_CURSOS.ID_AUTORIZACION='"&vautorizacionid&"'"

set rsHistUp = conn.execute (HistUp)

dim regHistUp
	while not rsHistUp.eof
		regHistUp = "update HISTORICO_CURSOS set ID_TRABAJADOR='"&rsHistUp("TRABIDUP")&"'"
		regHistUp = regHistUp&" where HISTORICO_CURSOS.TRABUP='"&vtabpart&"'"
		regHistUp = regHistUp&" and HISTORICO_CURSOS.ID_AUTORIZACION='"&vautorizacionid&"'"
		regHistUp = regHistUp&" and HISTORICO_CURSOS.ID_HISTORICO_CURSO='"&rsHistUp("ID_HISTORICO_CURSO")&"'"

		conn.execute (regHistUp)

		rsHistUp.Movenext
	wend

set objFich = objUpload.Ficheros("SubDocSelec")

aux=split(objFich.Nombre,".")
ext=aux(ubound(aux))

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

nom_arch="Documento_"&fecha&"."&ext

objFich.GuardarComo nom_arch,Server.MapPath("../ordenes")

set oFich = nothing
set objUpload = nothing

dim UpAutoriza
UpAutoriza = "update AUTORIZACION set ID_BANCO="&vid_banco&", ID_CUENTA_CORRIENTE="&vid_cuenta_corriente&", MONTO_TRANSFERENCIA="&vmonto_transferencia&", FECHA_TRANSFERENCIA="&vfecha_transferencia&", RUT_DEPOSITANTE="&vrut_depositante&", NOMBRE_DEPOSITANTE="&vnombre_depositante&", EMAIL_DEPOSITANTE="&vemail_depositante&", RELACION_DEPOSITANTE="&vrelacion_depositante&", ID_PROGRAMA='"&vprograma&"', ORDEN_COMPRA='"&vorden&"', ID_OTIC='"&votic&"', "
UpAutoriza = UpAutoriza&"VALOR_OC='"&vvalortotal&"', VALOR_CURSO='"&vValor&"', DOCUMENTO_COMPROMISO='"&nom_arch&"', CON_OTIC='"&vconotic&"', "  
UpAutoriza = UpAutoriza&"CON_FRANQUICIA='"&vcofran&"', N_REG_SENCE="&vreg_fran&", ID_EMPRESA='"&vempresa&"',"
UpAutoriza = UpAutoriza&"ID_BLOQUE='"&vbloque&"', TIPO_DOC='"&vcompromiso&"',N_PARTICIPANTES='"&vnparticipantes&"',"
UpAutoriza = UpAutoriza&"ID_TIPO_REFERENCIA='"&vTpReferencia&"', N_REFERENCIA='"&vtxtReferencia&"', ID_TIPO_VENTA= "&vTipoVenta
UpAutoriza = UpAutoriza&" where ID_AUTORIZACION='"&vautorizacionid&"'"

conn.execute (UpAutoriza)

dim UpHistorico
UpHistorico = "update HISTORICO_CURSOS set ID_PROGRAMA='"&vprograma&"', ID_EMPRESA='"&vempresa&"'," 
UpHistorico = UpHistorico&" RELATOR='"&vsala&"', SEDE='"&vrelator&"', ID_BLOQUE='"&vbloque&"' "
UpHistorico = UpHistorico&" where ID_AUTORIZACION='"&vautorizacionid&"'"

conn.execute (UpHistorico)

if(vautorizacionid="0")then

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
Flds.Item(schema & "sendpassword") =  "admin_2015"
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

set rsEmpresa = conn.execute ("select R_SOCIAL,EMAIL from EMPRESAS where ID_EMPRESA='"&vempresa&"'")

texto_fichero = Replace(texto_fichero,"<!--#empresa-->",rsEmpresa("R_SOCIAL"))

dim queryProg
queryProg = "select CURRICULO.NOMBRE_CURSO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
queryProg = queryProg&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO from PROGRAMA "
queryProg = queryProg&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL where PROGRAMA.ID_PROGRAMA='"&vprograma&"'"

set rsQueryProg = conn.execute (queryProg)

texto_fichero = Replace(texto_fichero,"<!--#curso-->",rsQueryProg("NOMBRE_CURSO"))
texto_fichero = Replace(texto_fichero,"<!--#fecha-->",rsQueryProg("FECHA_INICIO_"))
texto_fichero = Replace(texto_fichero,"<!--#fechat-->",rsQueryProg("FECHA_TERMINO"))

texto_fichero = Replace(texto_fichero,"<!--#asistentes-->",vnparticipantes)  
texto_fichero = Replace(texto_fichero,"<!--#valor-->","$ "&replace(FormatNumber(vvalortotal,0),",","."))

texto_fichero = Replace(texto_fichero,"<!--#doc_tipo-->","http://norte.otecmutual.cl/ordenes/"&nom_arch)

dim queryTabla
queryTabla="select TRABAJADOR.RUT,dbo.MayMinTexto(TRABAJADOR.NOMBRES) as NOMBRES from HISTORICO_CURSOS "
queryTabla=queryTabla&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.id_trabajador "
queryTabla=queryTabla&" where HISTORICO_CURSOS.ID_AUTORIZACION='"&vautorizacionid&"' order by TRABAJADOR.NOMBRES asc "

set rsQueryTabla = conn.execute (queryTabla)

dim textoTabla
textoTabla=""
while not rsQueryTabla.eof
	textoTabla=textoTabla&"<tr>"
	textoTabla=textoTabla&"<td>"&replace(FormatNumber(mid(rsQueryTabla("RUT"), 1,len(rsQueryTabla("RUT"))-2),0)&mid(rsQueryTabla("RUT"), len(rsQueryTabla("RUT"))-1,len(rsQueryTabla("RUT"))),",",".")&"</td>"
	textoTabla=textoTabla&"<td>"&rsQueryTabla("NOMBRES")&"</td>"
	textoTabla=textoTabla&"</tr>"
rsQueryTabla.Movenext
wend
texto_fichero = Replace(texto_fichero,"<!--#tabla-->",textoTabla)  

if(vcompromiso=0)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Orden de Compra N° "&vorden)  
end if

if(vcompromiso=1)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Vale Vista N° "&vorden)  
end if

if(vcompromiso=2)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Depósito Cheque N° "&vorden)  
end if

if(vcompromiso=3)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Transferencia N° "&vorden)  
end if

if(vcompromiso=4)then
texto_fichero = Replace(texto_fichero,"<!--#TipoDocumento-->","Carta Compromiso N° "&vorden)  
end if

dim relSala

relSala="select dbo.MayMinTexto(INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+"
relSala=relSala&"INSTRUCTOR_RELATOR.A_MATERNO) AS instructor, "
relSala=relSala&"dbo.MayMinTexto(SEDES.NOMBRE+' - 3er Piso, '+SEDES.DIRECCION+', '+SEDES.CIUDAD) as sede "
relSala=relSala&" from bloque_programacion "
relSala=relSala&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
relSala=relSala&" inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede "
relSala=relSala&" where bloque_programacion.id_bloque='"&vbloque&"'"

set rsRelSala = conn.execute (relSala)

texto_fichero = Replace(texto_fichero,"<!--#relator-->",rsRelSala("instructor"))
texto_fichero = Replace(texto_fichero,"<!--#sala-->",rsRelSala("sede"))

With iMsg
.To = rsEmpresa("EMAIL")
.BCC = "notificaciones@otecmutual.cl"
.From = "Mutual Capacitación <notificaciones@otecmutual.cl>"
.Subject = "Inscripción a Curso Modificada"
.HTMLBody = texto_fichero
.Sender = "Mutual Capacitación"
.Organization = "Mutual Capacitación"
Set .Configuration = iConf
SendEmailGmail = .Send
End With

set iMsg = nothing
set iConf = nothing
set Flds = nothing

end if
%>
