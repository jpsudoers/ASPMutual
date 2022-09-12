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
vc_otic = Request("C_Otic")
vc_fran = Request("C_Fran")

if(Request("R_Sence")<>"")then
	vr_sence = Request("R_Sence")
else
	vr_sence = "NULL"
end if

if(Request("id_Proy")<>"")then
	vid_proyecto = "'"&Request("id_Proy")&"'"
else
	vid_proyecto = "NULL"
end if

dim query
query = "insert into AUTORIZACION (ID_PROGRAMA,ID_EMPRESA,ID_OTIC,ORDEN_COMPRA,VALOR_OC,FECHA__AUTORIZACION"
query = query&",INSCRITOS,ESTADO,ORDEN_COMPRA_OTIC,VALOR_OCOMPRA_OTIC,DOCUMENTO_COMPROMISO,ID_BLOQUE,TIPO_DOC,N_PARTICIPANTES,"
query = query&"VALOR_CURSO,CON_OTIC,FACTURADO,CON_FRANQUICIA,N_REG_SENCE,id_proyecto)"
query = query&" values('"&vprograma&"','"&vempresa&"','"&votic&"','"&vordenemp&"','"&vvalor&"',GETDATE ()"
query = query&",'"&vinscrito&"',1,'"&vordenotic&"','"&vvalorotic&"','"&vdocumento&"','"&vbloque&"','"&vtipodocasig&"',"
query = query&"'"&vnpartinsc&"','"&vvalorcurso&"','"&vc_otic&"',1,'"&vc_fran&"',"&vr_sence&","&vid_proyecto&")"

conn.execute (query)

set rsPreins = conn.execute ("select id_trabajador from PREINSCRIPCION_TRABAJADOR where id_preinscripcion="&vpreins)

set rsAutorizacion = conn.execute ("select IDENT_CURRENT('AUTORIZACION')AS UltAuto")

while not rsPreins.eof
	dim query2
	query2 = "insert into HISTORICO_CURSOS (ID_EMPRESA,ID_PROGRAMA,ID_TRABAJADOR,ESTADO,RELATOR,SEDE,ID_AUTORIZACION,ID_BLOQUE,"
	query2 = query2&"TRABDEL, TRABIDUP, TRABUP) values('"&vempresa&"','"&vprograma&"','"&rsPreins("id_trabajador")&"',1,"
	query2 = query2&"'"&vrelator&"','"&vsede&"','"&rsAutorizacion("UltAuto")&"','"&vbloque&"',1,1,1) "
	
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
TrabAuto = TrabAuto&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloque_programacion.nom_sede "
TrabAuto = TrabAuto&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+', '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as DIR "
TrabAuto = TrabAuto&" from HISTORICO_CURSOS "
TrabAuto = TrabAuto&" inner join TRABAJADOR on TRABAJADOR.ID_TRABAJADOR=HISTORICO_CURSOS.ID_TRABAJADOR "
TrabAuto = TrabAuto&" inner join AUTORIZACION on AUTORIZACION.ID_AUTORIZACION=HISTORICO_CURSOS.ID_AUTORIZACION "
TrabAuto = TrabAuto&" inner join bloque_programacion on bloque_programacion.id_bloque=HISTORICO_CURSOS.ID_BLOQUE " 
TrabAuto = TrabAuto&" inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede "  
TrabAuto = TrabAuto&" where HISTORICO_CURSOS.ID_AUTORIZACION='"&rsAutorizacion("UltAuto")&"'"

set rsQueryTabla = conn.execute (TrabAuto)

dim textoTabla
dim salDirCiu
textoTabla=""
salDirCiu=""
while not rsQueryTabla.eof
salDirCiu=rsQueryTabla("DIR")
	textoTabla=textoTabla&"<tr>"
	textoTabla=textoTabla&"<td>"&replace(FormatNumber(mid(rsQueryTabla("RUT"), 1,len(rsQueryTabla("RUT"))-2),0)&mid(rsQueryTabla("RUT"), len(rsQueryTabla("RUT"))-1,len(rsQueryTabla("RUT"))),",",".")&"</td>"
	textoTabla=textoTabla&"<td>"&rsQueryTabla("NOMBRES")&"</td>"
	textoTabla=textoTabla&"<td>"&rsQueryTabla("DIR")&"</td>"
	textoTabla=textoTabla&"</tr>"
rsQueryTabla.Movenext
wend

texto_fichero = Replace(texto_fichero,"<!--#tabla-->",textoTabla) 

set rsEmpresa = conn.execute ("select RUT,R_SOCIAL,EMAIL from EMPRESAS where ID_EMPRESA='"&vempresa&"'")

With iMsg
.To = rsEmpresa("EMAIL")
.BCC = "notificaciones@otecmutual.cl"
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

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>