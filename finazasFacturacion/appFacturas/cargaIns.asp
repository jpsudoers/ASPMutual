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
Dim DATOSOTIC
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOS = Server.CreateObject("ADODB.RecordSet")
DATOS.CursorType=3

sql = "select AU.ID_AUTORIZACION,E.RUT,UPPER(E.R_SOCIAL) as 'Empresa',UPPER(E.DIRECCION) AS DIRECCION,"&_
            "UPPER(E.DIRECCION+', '+E.CIUDAD) AS DIR, E.ID_EMPRESA,"&_
            "DBO.MayMinTexto(E.COMUNA) AS COMUNA, DBO.MayMinTexto(E.CIUDAD) AS CIUDAD, E.FONO, "&_
            "UPPER(E.GIRO) AS GIRO,AU.N_PARTICIPANTES,AU.VALOR_OC,AU.VALOR_CURSO, "&_
            "(CASE WHEN AU.TIPO_DOC='0' then 'O/C ' "&_
            " WHEN AU.TIPO_DOC='1' then 'Vale Vista N° ' "&_
            " WHEN AU.TIPO_DOC='2' then 'DV - ' "&_
            " WHEN AU.TIPO_DOC='3' then 'Transferencia N° ' "&_
            " WHEN AU.TIPO_DOC='4' then 'CONTRA FACTURA' END) as 'DocFactura', "&_
            "(CASE WHEN AU.TIPO_DOC='0' then 'Orden de Compra N° ' "&_
            " WHEN AU.TIPO_DOC='1' then 'Vale Vista N° ' "&_
            " WHEN AU.TIPO_DOC='2' then 'Depósito Cheque N° ' "&_
            " WHEN AU.TIPO_DOC='3' then 'Transferencia N° ' "&_
            " WHEN AU.TIPO_DOC='4' then 'Carta Compromiso N° ' "&_
            " END) as 'Tipo Documento', AU.TIPO_DOC, ORDEN_COMPRA=AU.ORDEN_COMPRA + (case when isnull(e.CON_REFERENCIA,0)=1 and AU.TIPO_DOC=0  "&_
            " and isnull(AU.ID_TIPO_REFERENCIA,0)>0 and isnull(AU.N_REFERENCIA,'')<>'' then "&_
            " ' - ' + (select trf.TIPO_REFERENCIA from TIPO_REFERENCIA trf  "&_
            " where trf.ID_TIPO_REFERENCIA=AU.ID_TIPO_REFERENCIA) + ' - ' + isnull(AU.N_REFERENCIA,'') else '' END), AU.DOCUMENTO_COMPROMISO, "&_
            " C.DESCRIPCION,C.CODIGO,C.HORAS, "&_
            " CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"&_
            " CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO, "&_
            " ISNULL(convert(varchar,AU.N_REG_SENCE),'') as N_REG_SENCE, "&_
			"(select (CASE WHEN bq.ID_SEDE =  27 THEN bq.nom_sede "&_
			" WHEN bq.ID_SEDE <>  27 THEN S.DIRECCION+', '+S.CIUDAD END) from bloque_programacion bq "&_
            " inner join SEDES s on s.ID_SEDE=bq.id_sede "&_
            " inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "&_
            " where bq.id_bloque=AU.ID_BLOQUE) as SEDE "&_
            " from AUTORIZACION AU "&_
            " INNER JOIN EMPRESAS E ON E.ID_EMPRESA=AU.ID_EMPRESA "&_
            " inner join PROGRAMA P on P.ID_PROGRAMA=AU.ID_PROGRAMA "&_
            " INNER JOIN CURRICULO C ON C.ID_MUTUAL=P.ID_MUTUAL "&_
            " where AU.ID_AUTORIZACION='"&Request("idAuto")&"'"

DATOS.Open sql,oConn

Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 
Response.Write("<records>"&DATOS.RecordCount&"</records>"&chr(13))

fila=0
	Response.Write("<row id="""&fila&""">"&chr(13))
	Response.Write("<Empresa>"&replace(DATOS("Empresa"),"&","Y")&"</Empresa>"&chr(13))
	Response.Write("<RUT>"&DATOS("RUT")&"</RUT>"&chr(13))
	Response.Write("<DIR>"&DATOS("DIR")&"</DIR>"&chr(13))
	Response.Write("<GIRO>"&DATOS("GIRO")&"</GIRO>"&chr(13))
	Response.Write("<COMUNA>"&DATOS("COMUNA")&"</COMUNA>"&chr(13))
	Response.Write("<FONO>"&DATOS("FONO")&"</FONO>"&chr(13))
	Response.Write("<CIUDAD>"&DATOS("CIUDAD")&"</CIUDAD>"&chr(13))
	Response.Write("<DESCRIPCION>"&DATOS("DESCRIPCION")&"</DESCRIPCION>"&chr(13))
	Response.Write("<CODIGO>"&DATOS("CODIGO")&"</CODIGO>"&chr(13))
	Response.Write("<HORAS>"&DATOS("HORAS")&"</HORAS>"&chr(13))	
	Response.Write("<FECHA_INICIO>"&DATOS("FECHA_INICIO")&"</FECHA_INICIO>"&chr(13))
	Response.Write("<FECHA_TERMINO>"&DATOS("FECHA_TERMINO")&"</FECHA_TERMINO>"&chr(13))
	Response.Write("<N_PARTICIPANTES>"&DATOS("N_PARTICIPANTES")&"</N_PARTICIPANTES>"&chr(13))
	Response.Write("<VALOR_OC>"&DATOS("VALOR_OC")&"</VALOR_OC>"&chr(13))
	Response.Write("<VALOR_CURSO>"&DATOS("VALOR_CURSO")&"</VALOR_CURSO>"&chr(13))
	Response.Write("<DocFactura>"&DATOS("DocFactura")&"</DocFactura>"&chr(13))
	Response.Write("<N_REG_SENCE>"&""&DATOS("N_REG_SENCE")&"</N_REG_SENCE>"&chr(13))	
	Response.Write("<DOCUMENTO_COMPROMISO>"&DATOS("DOCUMENTO_COMPROMISO")&"</DOCUMENTO_COMPROMISO>"&chr(13))	
	Response.Write("<DOCLABEL>"&DATOS("Tipo Documento")&"</DOCLABEL>"&chr(13))
	Response.Write("<ID_EMPRESA>"&DATOS("ID_EMPRESA")&"</ID_EMPRESA>"&chr(13))	
	Response.Write("<ORDEN_COMPRA>"&DATOS("ORDEN_COMPRA")&"</ORDEN_COMPRA>"&chr(13))
	Response.Write("<TIPO_DOC>"&DATOS("TIPO_DOC")&"</TIPO_DOC>"&chr(13))
	Response.Write("<SEDE>"&DATOS("SEDE")&"</SEDE>"&chr(13))
		
	Response.Write("</row>"&chr(13))
Response.Write("</rows>") 
%>