<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Set theDoc = Server.CreateObject("ABCpdf7.Doc")

fecha = year(now)&right("0"&month(now()),2)&right("0"&day(now()),2)
fecha = fecha&right("0"&hour(now()),2)&right("0"&minute(now()),2)&right("0"&second(now()),2)

		w = theDoc.MediaBox.Width
		h = theDoc.MediaBox.Height
		l = theDoc.MediaBox.Left
		b = theDoc.MediaBox.Bottom 
		theDoc.Transform.Rotate 90, l, b
		theDoc.Transform.Translate w, 0
		
		theDoc.Rect.Width = h
		theDoc.Rect.Height = w

		theDoc.Rect = "30 50 760 570"
		Tabla(theDoc)
		
		theID = theDoc.GetInfo(theDoc.Root, "Pages")
		theDoc.SetInfo theID, "/Rotate", "90"
		
		sArchivo = "../pdf/Reporte_Cierre_Curso_"&fecha&".pdf"
		theDoc.Save Server.MapPath(sArchivo)


function Tabla(theDoc)
		  sqlProg = "select (case when P.FECHA_INICIO_<P.FECHA_TERMINO then "&_
					"CONVERT(varchar(2), DAY(P.FECHA_INICIO_))+' '+CONVERT(varchar(3), (CASE MONTH(P.FECHA_INICIO_) "&_
					"WHEN 1 THEN 'Ene' "&_
					"WHEN 2 THEN 'Feb' "&_
					"WHEN 3 THEN 'Mar' "&_
					"WHEN 4 THEN 'Abr' "&_
					"WHEN 5 THEN 'May' "&_
					"WHEN 6 THEN 'Jun' "&_
					"WHEN 7 THEN 'Jul' "&_
					"WHEN 8 THEN 'Ago' "&_
					"WHEN 9 THEN 'Sep' "&_
					"WHEN 10 THEN 'Oct' "&_
					"WHEN 11 THEN 'Nov' "&_
					"WHEN 12 THEN 'Dic' END))+' '+ CONVERT(varchar(4), Year(P.FECHA_INICIO_))+' y '+  "&_
					"CONVERT(varchar(2), DAY(P.FECHA_TERMINO))+' '+CONVERT(varchar(3), (CASE MONTH(P.FECHA_TERMINO) "&_
					"WHEN 1 THEN 'Ene' "&_
					"WHEN 2 THEN 'Feb' "&_
					"WHEN 3 THEN 'Mar' "&_
					"WHEN 4 THEN 'Abr' "&_
					"WHEN 5 THEN 'May' "&_
					"WHEN 6 THEN 'Jun' "&_
					"WHEN 7 THEN 'Jul' "&_
					"WHEN 8 THEN 'Ago' "&_
					"WHEN 9 THEN 'Sep' "&_
					"WHEN 10 THEN 'Oct' "&_
					"WHEN 11 THEN 'Nov' "&_
					"WHEN 12 THEN 'Dic' "&_
					"END))+' '+ CONVERT(varchar(4), Year(P.FECHA_TERMINO) ) else "&_
					"CONVERT(varchar(2), DAY(P.FECHA_INICIO_))+' '+CONVERT(varchar(3), (CASE MONTH(P.FECHA_INICIO_) "&_
					"WHEN 1 THEN 'Ene' "&_
					"WHEN 2 THEN 'Feb' "&_
					"WHEN 3 THEN 'Mar' "&_
					"WHEN 4 THEN 'Abr' "&_
					"WHEN 5 THEN 'May' "&_
					"WHEN 6 THEN 'Jun' "&_
					"WHEN 7 THEN 'Jul' "&_
					"WHEN 8 THEN 'Ago' "&_
					"WHEN 9 THEN 'Sep' "&_
					"WHEN 10 THEN 'Oct' "&_
					"WHEN 11 THEN 'Nov' "&_
					"WHEN 12 THEN 'Dic' END))+' '+ CONVERT(varchar(4), Year(P.FECHA_INICIO_)) "&_
					"end) as fecha,p.DIR_EJEC,s.CIUDAD from bloque_programacion bq  "&_
					"inner join PROGRAMA p on P.ID_PROGRAMA=bq.id_programa   "&_
					"inner join sedes s on s.ID_SEDE=bq.id_sede "&_
					"where bq.ID_BLOQUE="&Request("b")

					set rsProg = conn.execute (sqlProg)
	
						theDoc.Font = theDoc.EmbedFont("Arial Black")
						theDoc.FontSize = 11
						Set theTable = New Table
						theTable.Focus theDoc, 1
						theTable.Width(0) = 8
						theTable.Padding = 4	
											
						theTable.NextRow
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "REPORTE DE INDUCCIÓN"

						theTable.NextRow
			
						theDoc.FontSize = 8
						Set theTable = New Table
						theTable.Focus theDoc, 9
						theTable.Width(0) = 0.3
						theTable.Width(1) = 1.4
						theTable.Width(2) = 0.8
						theTable.Width(3) = 1.7				
					    theTable.Width(4) = 0.6
						theTable.Width(5) = 1
						theTable.Width(6) = 0.9					
					    theTable.Width(7) = 0.9	
					    theTable.Width(8) = 0.6																		
						theTable.Padding = 4	
						
						theTable.NextRow
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "N°"
						theTable.SelectCell(0)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "EMPRESA"
						theTable.SelectCell(1)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "PROYECTO"
						theTable.SelectCell(2)
						theTable.Frame True, True, true, true													
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "NOMBRE"
						theTable.SelectCell(3)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "RUT"
						theTable.SelectCell(4)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "FECHA REALIZACIÓN"
						theTable.SelectCell(5)
						theTable.Frame True, True, true, true	
												
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "ESTATUS DÍA I"
						theTable.SelectCell(6)
						theTable.Frame True, True, true, true		
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "ESTATUS DÍA II"
						theTable.SelectCell(7)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "CIUDAD"
						theTable.SelectCell(8)
						theTable.Frame True, True, true, true																	

	sql = "select dbo.MayMinTexto(T.NOMBRES) as NOMBRES,"&_
	         "(CASE WHEN T.NACIONALIDAD='1' then T.ID_EXTRANJERO WHEN T.NACIONALIDAD='0' then T.RUT END) as 'TrabId',"&_
			 " hc.ASISTENCIA,hc.CALIFICACION,hc.EVALUACION, "&_
			 " hc.ASIS_CDN,hc.CAL_CDN,hc.EVA_CDN, "&_
			 " hc.ASIS_REL,hc.CAL_REL,hc.EVA_REL,dbo.MayMinTexto(e.R_SOCIAL) as R_SOCIAL,"&_
    		 " CONVERT(VARCHAR(10),P.FECHA_INICIO_, 105) as FECHA_INICIO_,"&_
    		 " CONVERT(VARCHAR(10),P.FECHA_TERMINO, 105) as FECHA_TERMINO "&_
			 " from HISTORICO_CURSOS hc "&_
			 " inner join trabajador t on t.ID_TRABAJADOR=hc.ID_TRABAJADOR "&_
			 " inner join bloque_programacion bq on bq.id_bloque=hc.ID_BLOQUE "&_
             " inner join PROGRAMA p on P.ID_PROGRAMA=bq.id_programa "&_ 
     	     " inner join empresas e on e.ID_EMPRESA=hc.ID_EMPRESA "&_		 
			 " where bq.ID_BLOQUE="&Request("b")&_
			 " ORDER BY R_SOCIAL, NOMBRES ASC"

	set rsTrab = conn.execute (sql)

	theDoc.Font = theDoc.EmbedFont("Arial")

	cont=0
	n_trab=1
	
	while not rsTrab.eof
				cont=cont+1
					if(cont=31)then
						theDoc.Page = theDoc.AddPage()
						theDoc.Rect = "30 50 760 570"
						theDoc.Font = theDoc.EmbedFont("Arial Black")
						theDoc.FontSize = 8
						Set theTable = New Table
						theTable.Focus theDoc, 9
						theTable.Width(0) = 0.3
						theTable.Width(1) = 0.8
						theTable.Width(2) = 1
						theTable.Width(3) = 1.3					
					    theTable.Width(4) = 0.7	
						theTable.Width(5) = 1
						theTable.Width(6) = 1					
					    theTable.Width(7) = 1	
					    theTable.Width(8) = 1																		
						theTable.Padding = 4	
						
						theTable.NextRow
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "Nº"
						theTable.SelectCell(0)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "EMPRESA"
						theTable.SelectCell(1)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "PROYECTO"
						theTable.SelectCell(2)
						theTable.Frame True, True, true, true													
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "NOMBRE"
						theTable.SelectCell(3)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "RUT"
						theTable.SelectCell(4)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "FECHA REALIZACIÓN"
						theTable.SelectCell(5)
						theTable.Frame True, True, true, true	
												
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "ESTATUS DÍA I"
						theTable.SelectCell(6)
						theTable.Frame True, True, true, true		
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "ESTATUS DÍA II"
						theTable.SelectCell(7)
						theTable.Frame True, True, true, true	
						
						theTable.NextCell
						theDoc.HPos = 0.5
						theTable.AddText "CIUDAD"
						theTable.SelectCell(8)
						theTable.Frame True, True, true, true							
						
						theDoc.Font = theDoc.EmbedFont("Arial")					
					cont=1
				end if			
				
				theDoc.FontSize = 7	
			    theTable.NextRow
				theTable.NextCell
				theDoc.HPos = 0.5
				theTable.AddText n_trab
				theTable.SelectCell(0)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsTrab("R_SOCIAL")
				theTable.SelectCell(1)
				theTable.Frame True, True, true, true	
			
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText ""
				theTable.SelectCell(2)
				theTable.Frame True, True, true, true	

				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsTrab("NOMBRES")
				theTable.SelectCell(3)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsTrab("TrabId")
				theTable.SelectCell(4)
				theTable.Frame True, True, true, true		
				
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsProg("fecha")
				theTable.SelectCell(5)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText ""
				theTable.SelectCell(6)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText ""
				theTable.SelectCell(7)
				theTable.Frame True, True, true, true	
				
				theTable.NextCell
				theDoc.HPos = 0
				theTable.AddText rsProg("CIUDAD")
				theTable.SelectCell(8)
				theTable.Frame True, True, true, true												
								
				n_trab=n_trab+1
			rsTrab.Movenext
	wend	
end function
%>