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

vTipoVenta = Request("tipoVenta")
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

if(Request("Banco")<>"")then
	vid_banco = Request("Banco")
else
	vid_banco = "NULL"
end if

if(Request("numeroCuenta")<>"")then
		vid_cuenta_corriente = Request("numeroCuenta")
else
		vid_cuenta_corriente = "NULL"
end if

if(Request("txtMonto")<>"")then
	vmonto_transferencia = Request("txtMonto")
else
	vmonto_transferencia = "NULL"
end if


if(Request("txtFechaDeposito")<>"")then
	vfecha_transferencia = " CONVERT(datetime,'"&Request("txtFechaDeposito")&"',105) "
else
	vfecha_transferencia = "NULL"
end if

if(Request("txtRutDepositante")<>"")then
	vrut_depositante = "'"&Request("txtRutDepositante")&"'"
else
	vrut_depositante = "NULL"
end if

if(Request("txtNombreDepositante")<>"")then
	vnombre_depositante = "'"&Request("txtNombreDepositante")&"'"
else
	vnombre_depositante = "NULL"
end if

if(Request("txtEmail")<>"")then
	vemail_depositante = "'"&Request("txtEmail")&"'"
else
	vemail_depositante = "NULL"
end if

if(Request("txtRelacion")<>"")then
	vrelacion_depositante = "'"&Request("txtRelacion")&"'"
else
	vrelacion_depositante = "NULL"
end if


dim query
query = "IF EXISTS (select * from preinscripciones p where p.id_preinscripcion='"&vpreins&"' and p.estado=1) BEGIN "
query = query&"insert into AUTORIZACION (ID_PROGRAMA,ID_EMPRESA,ID_OTIC,ORDEN_COMPRA,VALOR_OC,FECHA__AUTORIZACION"
query = query&",INSCRITOS,ESTADO,ORDEN_COMPRA_OTIC,VALOR_OCOMPRA_OTIC,DOCUMENTO_COMPROMISO,ID_BLOQUE,TIPO_DOC,N_PARTICIPANTES,"
query = query&"VALOR_CURSO,CON_OTIC,FACTURADO,CON_FRANQUICIA,N_REG_SENCE,id_proyecto,ID_TIPO_VENTA, id_banco, id_cuenta_corriente, monto_transferencia, fecha_transferencia, rut_depositante, nombre_depositante, email_depositante, relacion_depositante)"
query = query&" values('"&vprograma&"','"&vempresa&"','"&votic&"','"&vordenemp&"','"&vvalor&"',GETDATE ()"
query = query&",'"&vinscrito&"',1,'"&vordenotic&"','"&vvalorotic&"','"&vdocumento&"','"&vbloque&"','"&vtipodocasig&"',"
query = query&"'"&vnpartinsc&"','"&vvalorcurso&"','"&vc_otic&"',1,'"&vc_fran&"',"&vr_sence&","&vid_proyecto&","&vTipoVenta&","&vid_banco&","&vid_cuenta_corriente&","&vmonto_transferencia&","&vfecha_transferencia&","&vrut_depositante&", "&vnombre_depositante&", "&vemail_depositante&" , "&vrelacion_depositante&" ) END"

conn.execute (query)

set rsPreins = conn.execute ("select id_trabajador from PREINSCRIPCION_TRABAJADOR where id_preinscripcion="&vpreins)

set rsAutorizacion = conn.execute ("select IDENT_CURRENT('AUTORIZACION')AS UltAuto")

while not rsPreins.eof
	dim query2
	query2 = "IF NOT EXISTS (select * from HISTORICO_CURSOS HC where HC.ID_AUTORIZACION='"&rsAutorizacion("UltAuto")&"'"
	query2 = query2&" and HC.ID_TRABAJADOR='"&rsPreins("id_trabajador")&"') BEGIN "
	query2 = query2&"insert into HISTORICO_CURSOS (ID_EMPRESA,ID_PROGRAMA,ID_TRABAJADOR,ESTADO,RELATOR,SEDE,ID_AUTORIZACION,ID_BLOQUE,"
	query2 = query2&"TRABDEL, TRABIDUP, TRABUP) values('"&vempresa&"','"&vprograma&"','"&rsPreins("id_trabajador")&"',1,"
	query2 = query2&"'"&vrelator&"','"&vsede&"','"&rsAutorizacion("UltAuto")&"','"&vbloque&"',1,1,1) END"
	
	conn.execute (query2)
	rsPreins.Movenext
wend

'conn.execute ("UPDATE preinscripciones set estado=0 where id_preinscripcion="&vpreins)

dim upPreins
upPreins="UPDATE preinscripciones set estado=0, situacion=1, situacion_fecha=GETDATE (), "
upPreins=upPreins&"situacion_encargado='"&Session("usuarioMutual")&"',"
upPreins=upPreins&"situacion_razon=NULL where id_preinscripcion="&vpreins

conn.execute (upPreins)

dim queryProg
queryProg = "select CURRICULO.NOMBRE_CURSO,CURRICULO.CODIGO,CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, "
queryProg = queryProg&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, CURRICULO.ID_MUTUAL, "
queryProg = queryProg&"BMI=hb1.HORARIO,BTF=hb2.HORARIO,PROGRAMA.ID_Modalidad,PROGRAMA.ID_ZOOM "
queryProg = queryProg&" from PROGRAMA "
queryProg = queryProg&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "
queryProg = queryProg&" left join HORARIO_BLOQUES hb1 on hb1.ID_HORARIO=PROGRAMA.BMI "
queryProg = queryProg&" left join HORARIO_BLOQUES hb2 on hb2.ID_HORARIO=PROGRAMA.BTF "
queryProg = queryProg&" where PROGRAMA.ID_PROGRAMA='"&vprograma&"'"

set rsQueryProg = conn.execute (queryProg)

if(rsQueryProg("ID_MUTUAL")<>"52")then
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
archivo= Server.MapPath("mailEmpresa.html")
'conectamos con el FSO 
set confile = createObject("scripting.filesystemobject") 
'volvemos a abrir el fichero para lectura 
set fich = confile.OpenTextFile(archivo) 

'leemos el contenido del fichero 
texto_fichero = fich.readAll() 
'cerramos el fichero 
fich.close() 

texto_fichero = Replace(texto_fichero,"<!--#curso-->",rsQueryProg("NOMBRE_CURSO"))
texto_fichero = Replace(texto_fichero,"<!--#inicio-->",rsQueryProg("FECHA_INICIO_"))
texto_fichero = Replace(texto_fichero,"<!--#termino-->",rsQueryProg("FECHA_TERMINO"))

if(rsQueryProg("ID_Modalidad")="1")then
	texto_fichero = Replace(texto_fichero,"<!--#lbHorario-->","<b>Horario de curso :</b>")
	texto_fichero = Replace(texto_fichero,"<!--#horario-->",rsQueryProg("BMI")&" a "&rsQueryProg("BTF"))
	texto_fichero = Replace(texto_fichero,"<!--#modalidad-->","Presencial")
	texto_fichero = Replace(texto_fichero,"<!--#lbZoom-->","")
     	texto_fichero = Replace(texto_fichero,"<!--#idZoom-->","")
     	texto_fichero = Replace(texto_fichero,"<!--#parrafoparte1-->","Los participantes deben presentarse con su cedula de identidad. De lo contrario no podrán ingresar al curso.")
     	texto_fichero = Replace(texto_fichero,"<!--#parrafoparte2-->","<em><b>Es responsabilidad de cada empresa informar a su trabajador que se encuentra haciendo uso de la franquicia SENCE, por lo cual se requiere de manera obligatoria que el trabajador firme el libro electrónico con su huella digital, de lo contrario Mutual Capacitación no se hace responsable de la perdida de la franquicia y el costo del curso pasara directamente a la empresa solicitante de la inducción.</b></em>")
elseif(rsQueryProg("ID_Modalidad")="2")then
	texto_fichero = Replace(texto_fichero,"<!--#lbHorario-->","")
	'texto_fichero = Replace(texto_fichero,"<!--#horario-->","")
	texto_fichero = Replace(texto_fichero,"<!--#horario-->",rsQueryProg("BMI")&" a "&rsQueryProg("BTF"))
	texto_fichero = Replace(texto_fichero,"<!--#modalidad-->","Asincrónico")
	texto_fichero = Replace(texto_fichero,"<!--#lbZoom-->","")
     	texto_fichero = Replace(texto_fichero,"<!--#idZoom-->","")
     	texto_fichero = Replace(texto_fichero,"<!--#parrafoparte1-->","<p><font color='#FF0000'><i><b>INFORMACIÓN IMPORTANTE</b></i></font><br/><br/><b><u>MODALIDAD:</u></b><br/>El curso es impartido e-learning y estará disponible 24 horas después de realizada la inscripción y durante 7 días de corrido para que el participante lo realice.<br/><br/>El día 8 ya no estará disponible y se deberá volver a inscribir con costo empresa.<br/><br/>Las inscripciones se aceptarán hasta las 16:00 horas de lunes a jueves y a las 12:00 horas los viernes. La programación será de lunes a lunes, pero las personas que inscribieron fuera de horario serán autorizadas al día siguiente hábil.<br/><br/>Las empresas responsables de las inscripciones de sus trabajadores deberán informar a estos los datos de usuario y clave, link de sitio, como acceder a plataforma, condiciones de aprobación indicados en el siguiente <a href='http://norte.otecmutual.cl/documentos/Ingreso_LMS.pdf' target='_blank'>Link</a>.<br/><br/><b><u>SOPORTE TÉCNICO:</u></b><br/>Contamos con el área de Soporte Técnico, que podrán contactar en caso de que los participantes tengan problemas con el ingreso (usuario y clave) horario de atención es lunes a jueves de 8:30 a 17:00 horas y los viernes de 8:30 a 15:00 horas. Solicitar asistencia a <a href='mailto:contacto@mutualcapacitacion.cl' target='_parent'>contacto@mutualcapacitacion.cl</a><br/><br/><font color='#FF0000'><i><b>INGRESO PARA EJECUTAR CURSO:</b></i></font><br/><br/>Portal <a href='https://formacion.mcap.cl' target='_blank'>https://formacion.mcap.cl</a><br/><br/><b><u>CONDICIONES PARA APROBAR:</u></b><br/>  -	Realizar el 100% de los ejercicios en cada módulo</b><br/>  - Obtener 80% mínimo en evaluación final<br/><br/>Una vez terminado el curso se dispondrá del certificado de aprobación para el participante en <a href='https://formacion.mcap.cl' target='_blank'>https://formacion.mcap.cl</a> y para la empresa se dispondrá el certificado al día siguiente hábil en <a href='http://norte.otecmutual.cl' target='_blank'>http://norte.otecmutual.cl</a></p>")
     	texto_fichero = Replace(texto_fichero,"<!--#parrafoparte2-->","")
else
	texto_fichero = Replace(texto_fichero,"<!--#lbHorario-->","<b>Horario de curso :</b>")
	texto_fichero = Replace(texto_fichero,"<!--#horario-->",rsQueryProg("BMI")&" a "&rsQueryProg("BTF"))
	texto_fichero = Replace(texto_fichero,"<!--#modalidad-->","Virtual")
	texto_fichero = Replace(texto_fichero,"<!--#lbZoom-->","<b>Unirse a Zoom : </b>")
     	texto_fichero = Replace(texto_fichero,"<!--#idZoom-->","https://zoom.us/j/"&rsQueryProg("ID_ZOOM")&"  -  <a href='http://norte.otecmutual.cl/documentos/Instructivo_Uso_ZOOM_Destinatario.pdf' target='_blank'>Ver Instructivo</a>")
     	texto_fichero = Replace(texto_fichero,"<!--#parrafoparte1-->","<p><font color='#FF0000'><i><b>INFORMACIÓN IMPORTANTE</b></i></font><br/><br/><b><u>MONITOR EN SALA VIRTUAL:</u></b><br/>En los cursos virtuales siempre estará el facilitador acompañado de un coordinador que cumple la función de monitor, esto quiere decir que en sala si hay cualquier problema de conexión, ingresos tardes, desactivación de cámaras o cualquier actividad que realicen los participantes fuera del desarrollo normal del curso, el monitor se encargará de contactarse con los participantes y aplicar medidas según requiera el caso.<br/>La toma de asistencia será permanente en el curso, su proceso será pasar lista y que los participantes muestren su Cedula de identidad para dejar constancia de asistencia en la grabación.<br/><br/><b><u>SOPORTE TÉCNICO:</u></b><br/>Contamos con el área de Soporte Técnico TI quienes podremos contactar en caso de que los participantes requieran de apoyo más técnico para conectar sus cámaras o audio.<br/><br/><b><u>DESCRIPCIÓN INSCRIPCIÓN:</u></b><br/>El participante necesitará tener computador, micrófono, cámara e internet (conexión estable)<br/>Puede descargar la aplicación o ingresar directamente al siguiente link <a href='www.zoom.us' target='_blank'>www.zoom.us</a> (la descarga es gratuita)<br/><br/><b><u>PREVIO AL CURSO:</u></b><br/>Probar sistema de audio micrófonos y parlantes o audífonos<br/>Revisar habilitación de cámara<br/><br/><b><u>CONDICIONES PARA APROBAR:</u></b><br/>Participante debe estar conectado el tiempo que dure la capacitación.<br/>La cámara encendida, no pueden tener una foto de perfil.<br/>Micrófono activo para realizar preguntas.<br/>Permanecer en su puesto mientras se realice el curso, no se pueden trasladar de un lugar a otro (metro, caminando,conduciendo u operando equipos entre otros) Hay considerados tiempos de descanso.<br/>Colocar como usuario nombre y apellido del participante.<br/><br/>Si los participantes comparten un solo equipo para realizar esta actividad deben contar con protección mascarilla o protección facial durante todo el curso y mantener una distancia prudente mínima de dos metros. Para ello pueden proyectar el cursos de lo contrario se solicitará que se retiren de la sala virtual y empresa deberá reprogramar inscripción.<br/>Si no lo pueden proyectar sería máximo 3 participantes para poder controlar su asistencia y visualizarlos en pantalla.<br/>Es importante que cada participante tiene que hacer de forma individual  la prueba.<br/>Se esperará a los participantes para su conexión 15 minutos máximos después del inicio del curso. Si no se conectan en este periodo de espera se considerarán reprobado y no podrán ingresar. Además se realizará el cobro de la inscripción.<br/><br/><b>Evaluación se realizará en línea enviando un link a cada participante para desarrollar prueba. El facilitador en conjunto con monitor podrán revisar en línea las evaluaciones antes de terminar el curso.<br/><br/>Se grabarán todos los cursos para evidenciar la asistencia de los participantes.<br/><br/><u>Se sugiere conectarse a la inducción 20 minutos antes de su inicio para resolver problemas técnicos de conexión o algún otro que se pueda presentar.</u></b></p>")
     	texto_fichero = Replace(texto_fichero,"<!--#parrafoparte2-->","")
end if

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

end if

Response.Write("<sql>"&query&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>