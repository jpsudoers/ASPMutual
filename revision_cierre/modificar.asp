<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next

For i = 1 To Request("countFilas") Step 1
query = "update HISTORICO_CURSOS set ASISTENCIA='"&Request("AsTra"&i)&"', "
query = query&"CALIFICACION='"&Request("caTra"&i)&"', EVALUACION='"&Request("eval"&i)&"', ESTADO=0 "
if(Request("eval"&i)="Aprobado")then
query = query&",COD_AUTENFICACION=(select CASE WHEN T.NACIONALIDAD='0' then " 
query = query&" right('00000' + CONVERT(VARCHAR,HC.ID_PROGRAMA), 5)+ "
query = query&" right('0000000' + CONVERT(VARCHAR,HC.ID_HISTORICO_CURSO), 7)+'-'+ "
query = query&" right('0000000000' + CONVERT(VARCHAR,T.RUT), 10) "
query = query&" WHEN T.NACIONALIDAD='1' then "
query = query&" right('00000' + CONVERT(VARCHAR,HC.ID_PROGRAMA), 5)+ "
query = query&" right('0000000' + CONVERT(VARCHAR,HC.ID_HISTORICO_CURSO), 7)+'-'+ "
query = query&" right('0000000000' + CONVERT(VARCHAR,T.ID_EXTRANJERO), 10) "
query = query&" END from HISTORICO_CURSOS HC "
query = query&" inner join TRABAJADOR T on T.ID_TRABAJADOR=HC.ID_TRABAJADOR "
query = query&" where HC.ID_HISTORICO_CURSO='"&Request("HiTra"&i)&"')"
end if
query = query&" where ID_HISTORICO_CURSO='"&Request("HiTra"&i)&"'"
conn.execute (query)
Next
'response.Write(query)
'response.End()

'conn.execute ("UPDATE AUTORIZACION SET ESTADO=0, FECHA_REV_CIERRE=GETDATE () WHERE AUTORIZACION.ID_AUTORIZACION='"&Request("AId")&"'")

dim cerrarProg
cerrarProg="UPDATE PROGRAMA SET VIGENCIA=(select CASE WHEN COUNT(*)=0 then 0 "
cerrarProg=cerrarProg&" WHEN COUNT(*)>0 then 1 END from HISTORICO_CURSOS hc where hc.ESTADO IN (1,2) " 
cerrarProg=cerrarProg&" and hc.ID_PROGRAMA=(SELECT BQ.id_programa FROM bloque_programacion BQ "
cerrarProg=cerrarProg&" WHERE BQ.id_bloque='"&Request("Bloque")&"')) " 
cerrarProg=cerrarProg&" WHERE ID_PROGRAMA=(SELECT BQ.id_programa FROM bloque_programacion BQ "
cerrarProg=cerrarProg&" WHERE BQ.id_bloque='"&Request("Bloque")&"')"

conn.execute (cerrarProg)

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

dim correo
correo="select EMPRESAS.ID_EMPRESA,EMPRESAS.RUT,UPPER (EMPRESAS.R_SOCIAL) as R_SOCIAL,CURRICULO.CODIGO,CURRICULO.NOMBRE_CURSO, "
correo= correo&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_INICIO_, 105) as FECHA_INICIO_, " 
correo= correo&"CONVERT(VARCHAR(10),PROGRAMA.FECHA_TERMINO, 105) as FECHA_TERMINO, " 
correo= correo&"(CASE WHEN bloque_programacion.id_rel_seg IS NOT NULL THEN " 
correo= correo&"dbo.MayMinTexto(INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO)+ "
correo= correo&"' / '+(select dbo.MayMinTexto (ri.NOMBRES+' '+ri.A_PATERNO) " 
correo= correo&" from INSTRUCTOR_RELATOR ri where ri.ID_INSTRUCTOR=bloque_programacion.id_rel_seg) " 
correo= correo&" WHEN bloque_programacion.id_rel_seg IS NULL THEN " 
correo= correo&" dbo.MayMinTexto(INSTRUCTOR_RELATOR.NOMBRES+' '+INSTRUCTOR_RELATOR.A_PATERNO) END) as RELATORES, "
correo= correo&"(CASE WHEN SEDES.ID_SEDE =  27 THEN bloque_programacion.nom_sede " 
correo= correo&" WHEN SEDES.ID_SEDE <>  27 THEN SEDES.NOMBRE+' - 3er Piso, '+SEDES.DIRECCION+', '+SEDES.CIUDAD END) as NOM_SEDE, " 
correo= correo&"(CASE WHEN AUTORIZACION.CON_OTIC='1' then 'SI' " 
correo= correo&"WHEN AUTORIZACION.TIPO_DOC<>'1' then 'NO' END) as C_OTIC, " 
correo= correo&"(CASE WHEN AUTORIZACION.CON_FRANQUICIA='1' then 'SI' " 
correo= correo&"WHEN AUTORIZACION.CON_FRANQUICIA<>'1' then 'NO' END) as C_FRANQUICIA, " 
correo= correo&"AUTORIZACION.N_REG_SENCE, "
correo= correo&"(CASE WHEN AUTORIZACION.CON_OTIC='1' and AUTORIZACION.ID_OTIC<>'0' then " 
correo= correo&"(select E.RUT from EMPRESAS E where ID_EMPRESA=AUTORIZACION.ID_OTIC) "
correo= correo&"else '' END) as RUT_OTIC, " 
correo= correo&"(CASE WHEN AUTORIZACION.CON_OTIC='1' and AUTORIZACION.ID_OTIC<>'0' then " 
correo= correo&"(select E.R_SOCIAL from EMPRESAS E where ID_EMPRESA=AUTORIZACION.ID_OTIC) "
correo= correo&"else '' END) as NOM_OTIC,AUTORIZACION.VALOR_OC,AUTORIZACION.VALOR_CURSO,AUTORIZACION.N_PARTICIPANTES, " 
correo= correo&"AUTORIZACION.ORDEN_COMPRA,(select CASE WHEN AUTORIZACION.TIPO_DOC='0' then 'Orden de Compra N° ' " 
correo= correo&"WHEN AUTORIZACION.TIPO_DOC='1' then 'Vale Vista N° ' " 
correo= correo&"WHEN AUTORIZACION.TIPO_DOC='2' then 'Depósito Cheque N° ' " 
correo= correo&"WHEN AUTORIZACION.TIPO_DOC='3' then 'Transferencia N° ' " 
correo= correo&"WHEN AUTORIZACION.TIPO_DOC='4' then 'Carta Compromiso N° ' END) as DOC, " 
correo= correo&"('http://norte.otecmutual.cl/ordenes/'+AUTORIZACION.DOCUMENTO_COMPROMISO) as DOCUMENTO_COMPROMISO, " 
correo= correo&"('http://norte.otecmutual.cl/finazasFacturacion/pdf.asp?prog=' "
correo= correo&"+CONVERT(VARCHAR,bloque_programacion.id_programa)+ " 
correo= correo&"'&relator='+CONVERT(VARCHAR,bloque_programacion.id_relator)+ " 
correo= correo&"'&empresa='+CONVERT(VARCHAR,EMPRESAS.ID_EMPRESA)+ " 
correo= correo&"'&autorizacion='+CONVERT(VARCHAR,AUTORIZACION.ID_AUTORIZACION)) as Link from AUTORIZACION "  
correo= correo&" inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA " 
correo= correo&" inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "
correo= correo&" inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL " 
correo= correo&" inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE " 
correo= correo&" inner join INSTRUCTOR_RELATOR on INSTRUCTOR_RELATOR.ID_INSTRUCTOR=bloque_programacion.id_relator "
correo= correo&" inner join SEDES on SEDES.ID_SEDE=bloque_programacion.id_sede "
correo= correo&" where AUTORIZACION.ID_AUTORIZACION='"&Request("AId")&"'"
correo= correo&" ORDER BY EMPRESAS.R_SOCIAL asc "

set rsCurso = conn.execute (correo)

facturadoIns=""
'if(rsCurso("ID_EMPRESA")="1365")then
	'facturadoIns=", FACTURADO=0 "
'end if

'conn.execute ("UPDATE AUTORIZACION SET ESTADO=0, FECHA_REV_CIERRE=GETDATE ()"&facturadoIns&" WHERE AUTORIZACION.ID_AUTORIZACION='"&Request("AId")&"'")

if(Request("txtSoloCert")="1")then
	conn.execute ("UPDATE AUTORIZACION SET ESTADO=1, FECHA_REV_CIERRE=null, SOLO_CERTIFICADOS=1 WHERE ID_AUTORIZACION='"&Request("AId")&"'")
else
	conn.execute ("UPDATE AUTORIZACION SET ESTADO=0, FECHA_REV_CIERRE=GETDATE ()"&facturadoIns&" WHERE AUTORIZACION.ID_AUTORIZACION='"&Request("AId")&"'")
end if

texto_fichero = Replace(texto_fichero,"<!--#rut_empresa-->",replace(FormatNumber(mid(rsCurso("RUT"), 1,len(rsCurso("RUT"))-2),0)&mid(rsCurso("RUT"), len(rsCurso("RUT"))-1,len(rsCurso("RUT"))),",","."))
texto_fichero = Replace(texto_fichero,"<!--#nom_empresa-->",rsCurso("R_SOCIAL"))
texto_fichero = Replace(texto_fichero,"<!--#codigocurso-->",rsCurso("CODIGO"))
texto_fichero = Replace(texto_fichero,"<!--#nombrecurso-->",rsCurso("NOMBRE_CURSO"))
texto_fichero = Replace(texto_fichero,"<!--#fechai-->",rsCurso("FECHA_INICIO_"))
texto_fichero = Replace(texto_fichero,"<!--#fechat-->",rsCurso("FECHA_TERMINO"))
texto_fichero = Replace(texto_fichero,"<!--#sala-->",rsCurso("NOM_SEDE"))
texto_fichero = Replace(texto_fichero,"<!--#relator-->",rsCurso("RELATORES"))

texto_fichero = Replace(texto_fichero,"<!--#franquicia-->",rsCurso("C_FRANQUICIA"))

if(rsCurso("C_FRANQUICIA")="SI")then
	texto_fichero = Replace(texto_fichero,"<!--#con_otic-->","<b>Con OTIC :</b>&nbsp;&nbsp;"&rsCurso("C_OTIC"))
else
	texto_fichero = Replace(texto_fichero,"<!--#con_otic-->","")
end if

if(rsCurso("C_FRANQUICIA")="SI")then
texto_fichero = Replace(texto_fichero,"<!--#reg_sence-->","<b>N° Registro Sence :</b>&nbsp;&nbsp;"&rsCurso("N_REG_SENCE"))
else
texto_fichero = Replace(texto_fichero,"<!--#reg_sence-->","")
end if

dim textoOtic
textoOtic=""

if(rsCurso("C_OTIC")="SI")then
	textoOtic="<tr><td><b>Rut OTIC :</b>&nbsp;&nbsp;"&replace(FormatNumber(mid(rsCurso("RUT_OTIC"), 1,len(rsCurso("RUT_OTIC"))-2),0)&mid(rsCurso("RUT_OTIC"), len(rsCurso("RUT_OTIC"))-1,len(rsCurso("RUT_OTIC"))),",",".")&"</td>"
	textoOtic=textoOtic&"<td><b>Razón Social OTIC : </b>&nbsp;&nbsp;"&rsCurso("NOM_OTIC")&"</td></tr>"
	texto_fichero = Replace(texto_fichero,"<!--#td_otic-->",textoOtic)
else
	texto_fichero = Replace(texto_fichero,"<!--#td_otic-->","")
end if

texto_fichero = Replace(texto_fichero,"<!--#n_part-->",rsCurso("N_PARTICIPANTES"))
texto_fichero = Replace(texto_fichero,"<!--#val_alum-->","$"&replace(FormatNumber(rsCurso("VALOR_CURSO"),0),",","."))
texto_fichero = Replace(texto_fichero,"<!--#val_total-->","$"&replace(FormatNumber(rsCurso("VALOR_OC"),0),",","."))
texto_fichero = Replace(texto_fichero,"<!--#tipo_doc-->",rsCurso("DOC")&" "&rsCurso("ORDEN_COMPRA")&" - "&"<a href="&rsCurso("DOCUMENTO_COMPROMISO")&" target=""_blank"">Ver</a>")

if(rsCurso("C_FRANQUICIA")="SI")then
	texto_fichero = Replace(texto_fichero,"<!--#cert_sence-->","<b>Certificado Sence :</b>&nbsp;&nbsp;<a href="&rsCurso("Link")&" target=""_blank"">Ver</a>")
else
	texto_fichero = Replace(texto_fichero,"<!--#cert_sence-->","")
end if


With iMsg
.To = "respaldos.mcfacturas2015@gmail.com"
.From = "Facturar a Empresas <notificaciones@otecmutual.cl>"
.Subject = "Facturar a "&rsCurso("R_SOCIAL")&" - "&rsCurso("FECHA_INICIO_")&" / "&rsCurso("FECHA_TERMINO")&" - "&rsCurso("NOMBRE_CURSO")
.HTMLBody = texto_fichero
.Sender = "Mutual Capacitación"
.Organization = "Mutual Capacitación"
Set .Configuration = iConf
SendEmailGmail = .Send
End With

set iMsg = nothing
set iConf = nothing
set Flds = nothing

if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>