<!--#include file="../../cnn_string.asp"-->
<%
Response.ContentType = "text/xml"
Response.AddHeader "Cache-control", "private"
Response.AddHeader "Expires", "-1"
Response.CodePage = 65001
Response.CharSet = "utf-8"
%>
<%
Dim DATOSFACT
Dim oConn
SET oConn = Server.CreateObject("ADODB.Connection")
'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
oConn.Open(MM_cnn_STRING)
Set DATOSFACT = Server.CreateObject("ADODB.RecordSet")
DATOSFACT.CursorType=3
				
Response.Write("<?xml version='1.0' encoding='utf-8' ?>"&chr(13))
Response.Write("<rows>"&chr(13)) 

fila=1
			if(Request("tipoEmp")="1")then
			
				sqlfact = "select TOP 1 AU.ID_AUTORIZACION,E.RUT,UPPER(E.R_SOCIAL) as 'Empresa',"&_
						  "UPPER(E.DIRECCION) AS DIRECCION,"&_
						  "UPPER(E.DIRECCION+', '+E.CIUDAD) AS DIR, E.ID_EMPRESA,"&_
						  "DBO.MayMinTexto(E.COMUNA) AS COMUNA, DBO.MayMinTexto(E.CIUDAD) AS CIUDAD, E.FONO,"&_ 
						  "UPPER(E.GIRO) AS GIRO,AU.N_PARTICIPANTES,AU.VALOR_OC,AU.VALOR_CURSO,"&_
						  "(CASE WHEN AU.TIPO_DOC='0' then 'O/C '"&_ 
						  "WHEN AU.TIPO_DOC='1' then 'Vale Vista N° ' "&_
						  "WHEN AU.TIPO_DOC='2' then 'DV - ' "&_
						  "WHEN AU.TIPO_DOC='3' then 'Transferencia N° '"&_
						  "WHEN AU.TIPO_DOC='4' then 'CONTRA FACTURA' END) as 'DocFactura', "&_
						  "(CASE WHEN AU.TIPO_DOC='0' then 'Orden de Compra N° ' "&_
						  "WHEN AU.TIPO_DOC='1' then 'Vale Vista N° ' "&_
						  "WHEN AU.TIPO_DOC='2' then 'Depósito Cheque N° ' "&_
						  "WHEN AU.TIPO_DOC='3' then 'Transferencia N° ' "&_
						  "WHEN AU.TIPO_DOC='4' then 'Carta Compromiso N° ' "&_
						  " END) as 'Tipo Documento', AU.TIPO_DOC, AU.ORDEN_COMPRA, AU.DOCUMENTO_COMPROMISO, "&_
						  "C.DESCRIPCION,C.CODIGO,C.HORAS, "&_
						  "CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"&_
						  "CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO, "&_
						  "ISNULL(convert(varchar,AU.N_REG_SENCE),'') as N_REG_SENCE, "&_
						  "(select (CASE WHEN bq.ID_SEDE =  27 THEN bq.nom_sede "&_
						  " WHEN bq.ID_SEDE <>  27 THEN S.DIRECCION+', '+S.CIUDAD END) from bloque_programacion bq "&_
						  " inner join SEDES s on s.ID_SEDE=bq.id_sede "&_
						  " inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "&_
						  " where bq.id_bloque=AU.ID_BLOQUE) as SEDE,F.VALOR_OC AS VALOR_OC_FACT,F.VALOR_C_DESC,F.DESCUENTO,F.MONTO,"&_
						  "F.VALOR_CURSO AS VALOR_CURSO_FACT from FACTURAS F "&_
						  " inner join AUTORIZACION AU ON AU.ID_AUTORIZACION=F.ID_AUTORIZACION "&_
						  " inner join PROGRAMA P on P.ID_PROGRAMA=AU.ID_PROGRAMA "&_
						  " INNER JOIN CURRICULO C ON C.ID_MUTUAL=P.ID_MUTUAL "&_
						  " INNER JOIN EMPRESAS E ON E.ID_EMPRESA=F.ID_EMPRESA "&_
						  " WHERE F.ID_AUTORIZACION='"&Request("idAuto")&"' AND F.ID_EMPRESA='"&Request("idEmpresa")&"' AND F.ESTADO=0 "&_
						  " ORDER BY F.ID_FACTURA DESC "
			
			
				DATOSFACT.Open sqlfact,oConn
				
				Response.Write("<records>"&DATOSFACT.RecordCount&"</records>"&chr(13))
				Response.Write("<row id="""&fila&""">"&chr(13))
					
				WHILE NOT DATOSFACT.EOF
							Response.Write("<Empresa>"&replace(DATOSFACT("Empresa"),"&","Y")&"</Empresa>"&chr(13))
							Response.Write("<RUT>"&DATOSFACT("RUT")&"</RUT>"&chr(13))
							Response.Write("<DIR>"&DATOSFACT("DIR")&"</DIR>"&chr(13))
							Response.Write("<GIRO>"&DATOSFACT("GIRO")&"</GIRO>"&chr(13))
							Response.Write("<COMUNA>"&DATOSFACT("COMUNA")&"</COMUNA>"&chr(13))
							Response.Write("<FONO>"&DATOSFACT("FONO")&"</FONO>"&chr(13))
							Response.Write("<CIUDAD>"&DATOSFACT("CIUDAD")&"</CIUDAD>"&chr(13))
							Response.Write("<DESCRIPCION>"&DATOSFACT("DESCRIPCION")&"</DESCRIPCION>"&chr(13))
							Response.Write("<CODIGO>"&DATOSFACT("CODIGO")&"</CODIGO>"&chr(13))
							Response.Write("<HORAS>"&DATOSFACT("HORAS")&"</HORAS>"&chr(13))	
							Response.Write("<FECHA_INICIO>"&DATOSFACT("FECHA_INICIO")&"</FECHA_INICIO>"&chr(13))
							Response.Write("<FECHA_TERMINO>"&DATOSFACT("FECHA_TERMINO")&"</FECHA_TERMINO>"&chr(13))
							Response.Write("<N_PARTICIPANTES>"&DATOSFACT("N_PARTICIPANTES")&"</N_PARTICIPANTES>"&chr(13))
							Response.Write("<VALOR_OC>"&DATOSFACT("VALOR_OC")&"</VALOR_OC>"&chr(13))
							Response.Write("<VALOR_CURSO>"&DATOSFACT("VALOR_CURSO")&"</VALOR_CURSO>"&chr(13))
							Response.Write("<DocFactura>"&DATOSFACT("DocFactura")&"</DocFactura>"&chr(13))
							Response.Write("<N_REG_SENCE>"&""&DATOSFACT("N_REG_SENCE")&"</N_REG_SENCE>"&chr(13))	
							Response.Write("<DOCUMENTO_COMPROMISO>"&DATOSFACT("DOCUMENTO_COMPROMISO")&"</DOCUMENTO_COMPROMISO>"&chr(13))	
							Response.Write("<DOCLABEL>"&DATOSFACT("Tipo Documento")&"</DOCLABEL>"&chr(13))
							Response.Write("<ID_EMPRESA>"&DATOSFACT("ID_EMPRESA")&"</ID_EMPRESA>"&chr(13))	
							Response.Write("<ORDEN_COMPRA>"&DATOSFACT("ORDEN_COMPRA")&"</ORDEN_COMPRA>"&chr(13))
							Response.Write("<TIPO_DOC>"&DATOSFACT("TIPO_DOC")&"</TIPO_DOC>"&chr(13))
							Response.Write("<SEDE>"&DATOSFACT("SEDE")&"</SEDE>"&chr(13))
							
							if not isNull(DATOSFACT("VALOR_OC_FACT"))then
								Response.Write("<VALOR_OC_FACT>"&DATOSFACT("VALOR_OC_FACT")&"</VALOR_OC_FACT>"&chr(13))
							else
							 	Response.Write("<VALOR_OC_FACT>"&DATOSFACT("VALOR_OC")&"</VALOR_OC_FACT>"&chr(13))
							end if
							
							if not isNull(DATOSFACT("VALOR_C_DESC"))then
								Response.Write("<VALOR_C_DESC>"&DATOSFACT("VALOR_C_DESC")&"</VALOR_C_DESC>"&chr(13))
							else
								Response.Write("<VALOR_C_DESC>"&DATOSFACT("VALOR_OC")&"</VALOR_C_DESC>"&chr(13))							
							end if
							
							if not isNull(DATOSFACT("DESCUENTO"))then							
								Response.Write("<DESCUENTO>"&DATOSFACT("DESCUENTO")&"</DESCUENTO>"&chr(13))
							else
								Response.Write("<DESCUENTO>0</DESCUENTO>"&chr(13))							
							end if
							
							if not isNull(DATOSFACT("VALOR_CURSO_FACT"))then								
								Response.Write("<VALOR_CURSO_FACT>"&DATOSFACT("VALOR_CURSO_FACT")&"</VALOR_CURSO_FACT>"&chr(13))
							else
								Response.Write("<VALOR_CURSO_FACT>"&DATOSFACT("VALOR_CURSO")&"</VALOR_CURSO_FACT>"&chr(13))							
							end if
								
							Response.Write("<MONTO>"&DATOSFACT("MONTO")&"</MONTO>"&chr(13))								
					DATOSFACT.MoveNext
				WEND

				else 
			
				sqlfact = "select TOP 1 AU.ID_AUTORIZACION,E.RUT,UPPER(E.R_SOCIAL) as 'Empresa',"&_
						  "UPPER(E.DIRECCION) AS DIRECCION,"&_
						  "UPPER(E.DIRECCION+', '+E.CIUDAD) AS DIR, E.ID_EMPRESA,"&_
						  "DBO.MayMinTexto(E.COMUNA) AS COMUNA, DBO.MayMinTexto(E.CIUDAD) AS CIUDAD, E.FONO,"&_ 
						  "UPPER(E.GIRO) AS GIRO,AU.N_PARTICIPANTES,AU.VALOR_OC,AU.VALOR_CURSO,"&_
						  "(CASE WHEN AU.TIPO_DOC='0' then 'O/C '"&_ 
						  "WHEN AU.TIPO_DOC='1' then 'Vale Vista N° ' "&_
						  "WHEN AU.TIPO_DOC='2' then 'DV - ' "&_
						  "WHEN AU.TIPO_DOC='3' then 'Transferencia N° '"&_
						  "WHEN AU.TIPO_DOC='4' then 'CONTRA FACTURA' END) as 'DocFactura', "&_
						  "(CASE WHEN AU.TIPO_DOC='0' then 'Orden de Compra N° ' "&_
						  "WHEN AU.TIPO_DOC='1' then 'Vale Vista N° ' "&_
						  "WHEN AU.TIPO_DOC='2' then 'Depósito Cheque N° ' "&_
						  "WHEN AU.TIPO_DOC='3' then 'Transferencia N° ' "&_
						  "WHEN AU.TIPO_DOC='4' then 'Carta Compromiso N° ' "&_
						  " END) as 'Tipo Documento', AU.TIPO_DOC, AU.ORDEN_COMPRA, AU.DOCUMENTO_COMPROMISO, "&_
						  "C.DESCRIPCION,C.CODIGO,C.HORAS, "&_
						  "CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO,"&_
						  "CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO, "&_
						  "ISNULL(convert(varchar,AU.N_REG_SENCE),'') as N_REG_SENCE, "&_
						  "(select (CASE WHEN bq.ID_SEDE =  27 THEN bq.nom_sede "&_
						  " WHEN bq.ID_SEDE <>  27 THEN S.DIRECCION+', '+S.CIUDAD END) from bloque_programacion bq "&_
						  " inner join SEDES s on s.ID_SEDE=bq.id_sede "&_
						  " inner join PROGRAMA p on p.ID_PROGRAMA=bq.id_programa "&_
						  " where bq.id_bloque=AU.ID_BLOQUE) as SEDE,F.VALOR_OC AS VALOR_OC_FACT,F.VALOR_C_DESC,F.DESCUENTO,F.MONTO,"&_
						  "F.VALOR_CURSO AS VALOR_CURSO_FACT from FACTURAS F "&_
						  " inner join AUTORIZACION AU ON AU.ID_AUTORIZACION=F.ID_AUTORIZACION "&_
						  " inner join PROGRAMA P on P.ID_PROGRAMA=AU.ID_PROGRAMA "&_
						  " INNER JOIN CURRICULO C ON C.ID_MUTUAL=P.ID_MUTUAL "&_
						  " INNER JOIN EMPRESAS E ON E.ID_EMPRESA=F.ID_EMPRESA "&_
						  " WHERE F.ID_AUTORIZACION='"&Request("idAuto")&"' AND F.ID_EMPRESA='"&Request("idEmpresa")&"' AND F.ESTADO=0 "&_
						  " ORDER BY F.ID_FACTURA DESC "
			
			
				DATOSFACT.Open sqlfact,oConn
				
				Response.Write("<records>"&DATOSFACT.RecordCount&"</records>"&chr(13))
				Response.Write("<row id="""&fila&""">"&chr(13))				
				
				WHILE NOT DATOSFACT.EOF
							Response.Write("<Empresa>"&replace(DATOSFACT("Empresa"),"&","Y")&"</Empresa>"&chr(13))
							Response.Write("<RUT>"&DATOSFACT("RUT")&"</RUT>"&chr(13))
							Response.Write("<DIR>"&DATOSFACT("DIR")&"</DIR>"&chr(13))
							Response.Write("<GIRO>"&DATOSFACT("GIRO")&"</GIRO>"&chr(13))
							Response.Write("<COMUNA>"&DATOSFACT("COMUNA")&"</COMUNA>"&chr(13))
							Response.Write("<FONO>"&DATOSFACT("FONO")&"</FONO>"&chr(13))
							Response.Write("<CIUDAD>"&DATOSFACT("CIUDAD")&"</CIUDAD>"&chr(13))
							Response.Write("<DESCRIPCION>"&DATOSFACT("DESCRIPCION")&"</DESCRIPCION>"&chr(13))
							Response.Write("<CODIGO>"&DATOSFACT("CODIGO")&"</CODIGO>"&chr(13))
							Response.Write("<HORAS>"&DATOSFACT("HORAS")&"</HORAS>"&chr(13))	
							Response.Write("<FECHA_INICIO>"&DATOSFACT("FECHA_INICIO")&"</FECHA_INICIO>"&chr(13))
							Response.Write("<FECHA_TERMINO>"&DATOSFACT("FECHA_TERMINO")&"</FECHA_TERMINO>"&chr(13))
							Response.Write("<N_PARTICIPANTES>"&DATOSFACT("N_PARTICIPANTES")&"</N_PARTICIPANTES>"&chr(13))
							Response.Write("<VALOR_OC>"&DATOSFACT("VALOR_OC")&"</VALOR_OC>"&chr(13))
							Response.Write("<VALOR_CURSO>"&DATOSFACT("VALOR_CURSO")&"</VALOR_CURSO>"&chr(13))
							Response.Write("<DocFactura>"&DATOSFACT("DocFactura")&"</DocFactura>"&chr(13))
							Response.Write("<N_REG_SENCE>"&""&DATOSFACT("N_REG_SENCE")&"</N_REG_SENCE>"&chr(13))	
							Response.Write("<DOCUMENTO_COMPROMISO>"&DATOSFACT("DOCUMENTO_COMPROMISO")&"</DOCUMENTO_COMPROMISO>"&chr(13))	
							Response.Write("<DOCLABEL>"&DATOSFACT("Tipo Documento")&"</DOCLABEL>"&chr(13))
							Response.Write("<ID_EMPRESA>"&DATOSFACT("ID_EMPRESA")&"</ID_EMPRESA>"&chr(13))	
							Response.Write("<ORDEN_COMPRA>"&DATOSFACT("ORDEN_COMPRA")&"</ORDEN_COMPRA>"&chr(13))
							Response.Write("<TIPO_DOC>"&DATOSFACT("TIPO_DOC")&"</TIPO_DOC>"&chr(13))
							Response.Write("<SEDE>"&DATOSFACT("SEDE")&"</SEDE>"&chr(13))
							
							if not isNull(DATOSFACT("VALOR_OC_FACT"))then
								Response.Write("<VALOR_OC_FACT>"&DATOSFACT("VALOR_OC_FACT")&"</VALOR_OC_FACT>"&chr(13))
							else
							 	Response.Write("<VALOR_OC_FACT>"&DATOSFACT("VALOR_OC")&"</VALOR_OC_FACT>"&chr(13))
							end if
							
							if not isNull(DATOSFACT("VALOR_C_DESC"))then
								Response.Write("<VALOR_C_DESC>"&DATOSFACT("VALOR_C_DESC")&"</VALOR_C_DESC>"&chr(13))
							else
								Response.Write("<VALOR_C_DESC>"&DATOSFACT("VALOR_OC")&"</VALOR_C_DESC>"&chr(13))							
							end if
							
							if not isNull(DATOSFACT("DESCUENTO"))then							
								Response.Write("<DESCUENTO>"&DATOSFACT("DESCUENTO")&"</DESCUENTO>"&chr(13))
							else
								Response.Write("<DESCUENTO>0</DESCUENTO>"&chr(13))							
							end if
							
							if not isNull(DATOSFACT("VALOR_CURSO_FACT"))then								
								Response.Write("<VALOR_CURSO_FACT>"&DATOSFACT("VALOR_CURSO_FACT")&"</VALOR_CURSO_FACT>"&chr(13))
							else
								Response.Write("<VALOR_CURSO_FACT>"&DATOSFACT("VALOR_CURSO")&"</VALOR_CURSO_FACT>"&chr(13))							
							end if
								
							Response.Write("<MONTO>"&DATOSFACT("MONTO")&"</MONTO>"&chr(13))							
					DATOSFACT.MoveNext
				WEND	
			end if

	Response.Write("</row>"&chr(13))
Response.Write("</rows>") 
%>