<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next

vprograma = Request("ProgramaAsig")
vempresa = Request("EmpresaAsig")
vordenemp = Request("txtOrdenCompraEAsig")
vvalor=Request("txtValorEAsig")
votic = Request("txtIdOticAsig")
vordenotic = Request("txtOrdenCompraOAsig")
vvalorotic=Request("txtValorOAsig")
vfecha=Request("txtFechaAsig")
vvalorautorizador=Request("txtValorAutorizadorAsig")
vinscrito=Request("txtInscritoAsig")
vpreins=Request("txtpreinscripcionAsig")
vdocumento=Request("documentosAsig")
vrelator=Request("Relator")
vsede=Request("Sede")
vbloque=Request("bloqueSel")
cadena = split(Request("relatorSedeAsig"), "/")
vtipodocasig = Request("tipoDocAsig")
vnpartinsc = Request("NPartInsc")
vvalorcurso = Request("ValorCurso")

dim query
query = "insert into AUTORIZACION (ID_PROGRAMA,ID_EMPRESA,ID_OTIC,ORDEN_COMPRA,VALOR_OC,FECHA__AUTORIZACION"
query = query&",INSCRITOS,ESTADO,ORDEN_COMPRA_OTIC,VALOR_OCOMPRA_OTIC,DOCUMENTO_COMPROMISO,ID_BLOQUE,TIPO_DOC,N_PARTICIPANTES,"
query = query&"VALOR_CURSO)"
query = query&" values('"&vprograma&"','"&vempresa&"','"&votic&"','"&vordenemp&"','"&vvalor&"',GETDATE ()"
query = query&",'"&vinscrito&"',1,'"&vordenotic&"','"&vvalorotic&"','"&vdocumento&"','"&vbloque&"','"&vtipodocasig&"',"
query = query&"'"&vnpartinsc&"','"&vvalorcurso&"')"

conn.execute (query)

set rsPreins = conn.execute ("select id_trabajador from PREINSCRIPCION_TRABAJADOR where id_preinscripcion="&vpreins)

set rsAutorizacion = conn.execute ("select IDENT_CURRENT('AUTORIZACION')AS UltAuto")

while not rsPreins.eof
	dim query2
	query2 = "insert into HISTORICO_CURSOS (ID_EMPRESA,ID_PROGRAMA,ID_TRABAJADOR,ESTADO,RELATOR,SEDE,ID_AUTORIZACION,ID_BLOQUE)"
	query2 = query2&" values('"&vempresa&"','"&vprograma&"','"&rsPreins("id_trabajador")&"',1,'"&vrelator&"','"&vsede&"',"
	query2 = query2&"'"&rsAutorizacion("UltAuto")&"','"&vbloque&"') "
	
	conn.execute (query2)
	rsPreins.Movenext
wend

'conn.execute ("UPDATE preinscripciones set estado=0 where id_preinscripcion="&vpreins)

dim upPreins
upPreins="UPDATE preinscripciones set estado=0, situacion=1, situacion_fecha=GETDATE (), "
upPreins=upPreins&"situacion_encargado='"&Session("usuarioMutual")&"',"
upPreins=upPreins&"situacion_razon=NULL where id_preinscripcion="&vpreins

conn.execute (upPreins)


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
archivo= Server.MapPath("mailEmpresa.html")
'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 
'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 
'cerramos el fichero 
fich.close() 

dim queryProg
queryProg = "select CURRICULO.NOMBRE_CURSO,CURRICULO.CODIGO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
queryProg = queryProg&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO "
queryProg = queryProg&" from PROGRAMA "
queryProg = queryProg&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
queryProg = queryProg&" where PROGRAMA.ID_PROGRAMA='"&vprograma&"'"

set rsQueryProg = conn.execute (queryProg)

texto_fichero = Replace(texto_fichero,"<!--#curso-->",rsQueryProg("NOMBRE_CURSO"))
texto_fichero = Replace(texto_fichero,"<!--#inicio-->",rsQueryProg("FECHA_INICIO_"))
texto_fichero = Replace(texto_fichero,"<!--#termino-->",rsQueryProg("FECHA_TERMINO"))


dim TrabAuto
TrabAuto = "select TRABAJADOR.RUT,TRABAJADOR.NOMBRES, "
TrabAuto = TrabAuto&"dbo.MayMinTexto (INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO+' '+INSTRUCTOR_RELATOR.A_MATERNO) as instructor,"
TrabAuto = TrabAuto&"SEDES.NOMBRE,SEDES.DIRECCION,SEDES.CIUDAD,SEDES.COMUNA "
TrabAuto = TrabAuto&" from HISTORICO_CURSOS "
TrabAuto = TrabAuto&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=HISTORICO_CURSOS.RELATOR "
TrabAuto = TrabAuto&" inner join SEDES on SEDES.ID_SEDE=HISTORICO_CURSOS.SEDE "
TrabAuto = TrabAuto&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
TrabAuto = TrabAuto&" inner join AUTORIZACION on AUTORIZACION.ID_AUTORIZACION=HISTORICO_CURSOS.ID_AUTORIZACION "
TrabAuto = TrabAuto&" where HISTORICO_CURSOS.ID_AUTORIZACION='"&rsAutorizacion("UltAuto")&"'"

set rsQueryTabla = conn.execute (TrabAuto)

dim textoTabla
dim salDirCiu
textoTabla=""
salDirCiu=""
while not rsQueryTabla.eof
salDirCiu=rsQueryTabla("NOMBRE")&" - 3er Piso, "&rsQueryTabla("DIRECCION")&", "&rsQueryTabla("CIUDAD")
	textoTabla=textoTabla&"<tr>"
	textoTabla=textoTabla&"<td>"&replace(FormatNumber(mid(rsQueryTabla("RUT"), 1,len(rsQueryTabla("RUT"))-2),0)&mid(rsQueryTabla("RUT"), len(rsQueryTabla("RUT"))-1,len(rsQueryTabla("RUT"))),",",".")&"</td>"
	textoTabla=textoTabla&"<td>"&rsQueryTabla("NOMBRES")&"</td>"
	textoTabla=textoTabla&"<td>"&rsQueryTabla("NOMBRE")&" - 3er Piso, "&rsQueryTabla("DIRECCION")&", "&rsQueryTabla("CIUDAD")&"</td>"
	textoTabla=textoTabla&"</tr>"
rsQueryTabla.Movenext
wend

texto_fichero = Replace(texto_fichero,"<!--#tabla-->",textoTabla) 

'response.End()

set rsEmpresa = conn.execute ("select RUT,R_SOCIAL,EMAIL from EMPRESAS where ID_EMPRESA='"&vempresa&"'")

With iMsg
.To = rsEmpresa("EMAIL")
.From = "Inscripción a Curso <notificaciones@otecmutual.cl>"
.Subject = "Inscripción a Curso"
.HTMLBody = texto_fichero
.Sender = "Mutual Capacitación"
.Organization = "Mutual Capacitación"
Set .Configuration = iConf
SendEmailGmail = .Send
End With

set iMsg = nothing
set iConf = nothing
set Flds = nothing

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

' send one copy with Google SMTP server (with autentication)
schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com" 
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "respaldos.mcfacturas2015@gmail.com"
Flds.Item(schema & "sendpassword") =  "mcfacturas2015"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

'creamos el nombre del archivo 
archivo= Server.MapPath("mailMutual.html")
'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 
'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 
'cerramos el fichero 
fich.close() 

texto_fichero = Replace(texto_fichero,"<!--#rut-->",rsEmpresa("RUT"))
texto_fichero = Replace(texto_fichero,"<!--#empresa-->",rsEmpresa("R_SOCIAL"))
texto_fichero = Replace(texto_fichero,"<!--#curso-->",rsQueryProg("NOMBRE_CURSO"))
texto_fichero = Replace(texto_fichero,"<!--#sence-->",rsQueryProg("CODIGO"))
texto_fichero = Replace(texto_fichero,"<!--#direccion-->",salDirCiu)
texto_fichero = Replace(texto_fichero,"<!--#inicio-->",rsQueryProg("FECHA_INICIO_"))
texto_fichero = Replace(texto_fichero,"<!--#termino-->",rsQueryProg("FECHA_TERMINO"))

dim queryIns
queryIns="select valor,valor_total,participantes,costo,numero_compromiso,doc_compromiso,tipo_compromiso "
queryIns=queryIns&" from preinscripciones where id_preinscripcion='"&vpreins&"'"
set rsQueryIns = conn.execute (queryIns)

texto_fichero = Replace(texto_fichero,"<!--#participantes-->",rsQueryIns("participantes"))
texto_fichero = Replace(texto_fichero,"<!--#valor-->","$ "&replace(FormatNumber(rsQueryIns("valor"),0),",","."))
texto_fichero = Replace(texto_fichero,"<!--#monto-->","$ "&replace(FormatNumber(rsQueryIns("valor_total"),0),",","."))

if(rsQueryIns("tipo_compromiso")="0")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Orden de Compra N° "&rsQueryIns("numero_compromiso"))
end if

if(rsQueryIns("tipo_compromiso")="1")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Vale Vista N° "&rsQueryIns("numero_compromiso")) 
end if

if(rsQueryIns("tipo_compromiso")="2")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Depósito Cheque N° "&rsQueryIns("numero_compromiso")) 
end if

if(rsQueryIns("tipo_compromiso")="3")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Transferencia N° "&rsQueryIns("numero_compromiso")) 
end if

if(rsQueryIns("tipo_compromiso")="4")then
texto_fichero = Replace(texto_fichero,"<!--#orden-->","Carta Compromiso N° "&rsQueryIns("numero_compromiso"))
end if

texto_fichero = Replace(texto_fichero,"<!--#doc_tipo-->","http://norte.otecmutual.cl/ordenes/"&rsQueryIns("doc_compromiso"))

With iMsg
.To = "respaldos.mcfacturas2015@gmail.com"
.From = "Facturar a Empresa <respaldos.mcfacturas2015@gmail.com>"
.Subject = "Facturar a Empresa"
.HTMLBody = texto_fichero
.Sender = "Mutual Capacitación"
.Organization = "Mutual Capacitación"
Set .Configuration = iConf
SendEmailGmail = .Send
End With

set iMsg = nothing
set iConf = nothing
set Flds = nothing

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>