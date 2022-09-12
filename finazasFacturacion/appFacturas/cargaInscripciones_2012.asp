<!--#include file="../../cnn_string.asp"-->
<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%
Dim DATOS
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select EMPRESAS.RUT,UPPER (EMPRESAS.R_SOCIAL) as R_SOCIAL,"&_
            "EMPRESAS.ID_EMPRESA,AUTORIZACION.VALOR_OC,"&_
            "EMPRESAS.ID_OTIC AS OTIC,AUTORIZACION.N_PARTICIPANTES,AUTORIZACION.CON_OTIC,"&_
            "(select COUNT(*) from FACTURAS "&_
            "where FACTURAS.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION) as 'Facturado',"&_
            "AUTORIZACION.DOCUMENTO_COMPROMISO,AUTORIZACION.ID_AUTORIZACION,AUTORIZACION.CON_FRANQUICIA,"&_
            "bloque_programacion.id_programa,bloque_programacion.id_relator,"&_
            "CONVERT(VARCHAR(10),AUTORIZACION.FECHA_REV_CIERRE, 105) AS FECHA_REV_CIERRE,"&_
            " case (select Datediff(""d"", Min(CONVERT(date,AU.FECHA_REV_CIERRE,105)), Max(CONVERT(date,GETDATE(), 105))) "&_
            " from AUTORIZACION AU where AU.ID_AUTORIZACION=AUTORIZACION.ID_AUTORIZACION) "&_
            " when 0 then '#66ff00' "&_
            " when 1 then '#ffff00' "&_
            " when 2 then '#ff0000' "&_
            " else '#ff0000' end as color from AUTORIZACION "&_
            " inner join EMPRESAS on EMPRESAS.ID_EMPRESA=AUTORIZACION.ID_EMPRESA "&_
            " inner join bloque_programacion on bloque_programacion.id_bloque=AUTORIZACION.ID_BLOQUE "&_
            " inner join PROGRAMA on PROGRAMA.ID_PROGRAMA=AUTORIZACION.ID_PROGRAMA "&_
            " inner join CURRICULO on CURRICULO.ID_MUTUAL=PROGRAMA.ID_MUTUAL "&_
            " where AUTORIZACION.ESTADO=0 AND AUTORIZACION.FACTURADO=1 "& Request("id_curso")&_
            " ORDER BY AUTORIZACION.FECHA_REV_CIERRE asc"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=0
WHILE NOT DATOS.EOF
	Response.Write("<row id="""&fila&""">"&chr(13))
	Response.Write("<FECHA_REV_CIERRE>"&DATOS("FECHA_REV_CIERRE")&"</FECHA_REV_CIERRE>"&chr(13))
	Response.Write("<RUT>"&DATOS("RUT")&"</RUT>"&chr(13))
	Response.Write("<R_SOCIAL>"&replace(DATOS("R_SOCIAL"),"&","y")&"</R_SOCIAL>"&chr(13))
	Response.Write("<N_PARTICIPANTES>"&DATOS("N_PARTICIPANTES")&"</N_PARTICIPANTES>"&chr(13))
	Response.Write("<VALOR_OC>"&DATOS("VALOR_OC")&"</VALOR_OC>"&chr(13))
	Response.Write("<CON_OTIC>"&DATOS("CON_OTIC")&"</CON_OTIC>"&chr(13))
	Response.Write("<ID_AUTORIZACION>"&DATOS("ID_AUTORIZACION")&"</ID_AUTORIZACION>"&chr(13))
	Response.Write("<DOCUMENTO_COMPROMISO>"&DATOS("DOCUMENTO_COMPROMISO")&"</DOCUMENTO_COMPROMISO>"&chr(13))
	Response.Write("<CON_OTIC2>"&DATOS("CON_OTIC")&"</CON_OTIC2>"&chr(13))
	Response.Write("<OTIC>"&DATOS("OTIC")&"</OTIC>"&chr(13))		
	Response.Write("</row>"&chr(13))

	fila=fila+1
	DATOS.MoveNext
WEND
Response.Write("</rows>") 
%>