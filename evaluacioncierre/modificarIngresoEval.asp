<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.ContentType="text/xml"
Response.Write("<?xml version='1.0' encoding='UTF-8'?>")
Response.Write("<modificar>") 

on error resume next
dim estado_eval

'upBloque="UPDATE bloque_programacion SET estado_eva_rel=1, id_usuario_rel="&Request("u")&", fecha_rev_rel=GETDATE() WHERE ID_BLOQUE="&Request("b")&";"

if(Request("t")="1")then
	'upBloque="UPDATE bloque_programacion SET estado_eva_cdn=1, id_usuario_cdn="&Request("u")&", fecha_rev_cdn=GETDATE() WHERE ID_BLOQUE="&Request("b")&";"
'else
	upBloque="UPDATE bloque_programacion SET estado_eva_rel=1, id_usuario_rel="&Request("u")&", fecha_rev_rel=GETDATE() WHERE ID_BLOQUE="&Request("b")&";"
end if

conn.execute (upBloque)

For i = 1 To Request("countFilas") Step 1
	if(Request("E"&i)="A")then
		estado_eval="Aprobado"
	elseif(Request("E"&i)="R")then
		estado_eval="Reprobado"
	elseif(Request("E"&i)="C")then
		estado_eval="Con Obs."
	end if
	
	if(Request("A"&i)<>"" AND Request("C"&i)<>"")then	
		'upHistorico = "update HISTORICO_CURSOS set ASIS_REL='"&Request("A"&i)&"', "&_
				     ' "CAL_REL='"&Request("C"&i)&"', EVA_REL='"&estado_eval&"' where ID_HISTORICO_CURSO='"&Request("H"&i)&"'"
					  
		if(Request("t")="0")then
		upHistorico = "update HISTORICO_CURSOS set ASIS_CDN='"&Request("A"&i)&"', "&_
              "CAL_CDN='"&Request("C"&i)&"', EVA_CDN='"&estado_eval&"', PO=0"&Request("po"&i)&", PVMO=0"&Request("pvmo"&i)&"  where ID_HISTORICO_CURSO='"&Request("H"&i)&"'"
		else
		upHistorico = "update HISTORICO_CURSOS set ASIS_REL='"&Request("A"&i)&"', "&_
              "CAL_REL='"&Request("C"&i)&"', EVA_REL='"&estado_eval&"', PO=0"&Request("po"&i)&", PVMO=0"&Request("pvmo"&i)&" where ID_HISTORICO_CURSO='"&Request("H"&i)&"'"
		end if			  
	
		conn.execute (upHistorico)
	end if
Next
'on error resume next
'Response.Write("<sql>"&err.Description&"</sql>")
if err.number <> 0 then
	Response.Write("<commit>false</commit>")
else
	Response.Write("<commit>true</commit>")
end if
Response.Write("</modificar>") 
%>