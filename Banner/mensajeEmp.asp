<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
vid = Request("id")
msn=""

		dim SQL
		SQL = "SELECT MENSAJE=dbo.ListaBanner(e.TITULO,e.INFORMACION,CONVERT(varchar(30), e.FECHA_PUBLICACION, 0)) "&_
      		      " from BANNER_EMPRESAS e "&_
      		      " where e.ID_EMPRESA="&vid&" and convert(date, e.FECHA_VIGENCIA)>=CONVERT(date, GETDATE()) "&_
      		      " order by e.FECHA_VIGENCIA"

			set rsSQL = conn.execute (SQL)

		  	while not rsSQL.eof
				msn = msn&rsSQL("MENSAJE")

		  	   rsSQL.Movenext
			wend

	      	%>
			<%=msn%>
