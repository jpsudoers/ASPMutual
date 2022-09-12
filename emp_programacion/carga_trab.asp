<!--#include file="../conexion.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

'Asigna a la variable Path, la ruta del archivo *.xls
Path=Server.MapPath("../cargas/Excel_"&request("archivo")&".xls")

'Establece una conexión entre el servidor asp y una base de datos
Dim oConexion
Set oConexion = Server.CreateObject("ADODB.Connection")

'Abrimos el objeto con el driver específico para Microsoft Excel
'oConexion.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&Path&";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
oConexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Path&";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1""" 

'Crea un objeto de tipo recordset para retornar la consulta sql
Dim oRS
Set oRS = Server.CreateObject("ADODB.Recordset")

'Se abre el recordset, señalando como tabla el rango de celdas Excel.
oRS.Open "Select * From [Trabajadores$A1:H60]",  oConexion, 3, 2, adCmdText 

'Primero los nombres de las columnas (el encabezado de la tabla)
response.Write "<table id=""listado"" border=""1""><thead><tr>"
For iCont = 0 to oRs.Fields.Count - 1
  Response.Write "<th>"&oRs.Fields(iCont).name&"</th>"
Next
Response.Write "<th>Mensaje</th>"
response.Write "</tr></thead><tbody>" 

' Mostramos todos los campos
'Do While NOT oRs.EOF
   ' response.Write "<tr>"
   ' For iCont = 0 to oRs.Fields.Count - 1
    '    Response.Write "<td>"&oRs.Fields(iCont)&"<td/>"
    'Next
    'response.Write "</tr>"
    'oRs.MoveNext
'Loop 
Mensaje=""
buscar=""
totTrab=0
While not oRs.EoF
				Mensaje=""
                vcheckExtran = "0"
				if ((trim(oRs("Rut"))="") or (isnull(oRs("Rut")))) then 
                	vruttrab = oRs("Rut")
				else
					vruttrab = UCase(replace(replace(replace(oRs("Rut"),",",""),".","")," ",""))
				end if
				
				
                vcargotrab = trim(UCase(oRs("Cargo")))
                'vescolaridadtrab = oRs("Escolaridad")
				
				if(UCase(oRs("Escolaridad"))=UCase("Sin Escolaridad"))then
					vescolaridadtrab = "0"
				end if
				
				if(UCase(oRs("Escolaridad"))=UCase("Básica Incompleta"))then
					vescolaridadtrab = "1"
				end if
				
				if(UCase(oRs("Escolaridad"))=UCase("Básica Completa"))then
					vescolaridadtrab = "2"
				end if
		
				if(UCase(oRs("Escolaridad"))=UCase("Media Incompleta"))then
					vescolaridadtrab = "3"
				end if
		
				if(UCase(oRs("Escolaridad"))=UCase("Media Completa"))then
					vescolaridadtrab = "4"
				end if
				
				if(UCase(oRs("Escolaridad"))=UCase("Superior Técnica Incompleta"))then
					vescolaridadtrab = "5"
				end if
		
				if(UCase(oRs("Escolaridad"))=UCase("Superior Técnica Profesional Completa"))then
					vescolaridadtrab = "6"
				end if
		
				if(UCase(oRs("Escolaridad"))=UCase("Universitaria Incompleta"))then
					vescolaridadtrab = "7"
				end if
				
				if(UCase(oRs("Escolaridad"))=UCase("Universitaria Completa"))then
					vescolaridadtrab = "8"
				end if
				
				
                vnomtrab = trim(oRs("Nombre"))
                vapatertrab = trim(oRs("Ap_Paterno"))
                vamatertrab = trim(oRs("Ap_Materno"))
                vtrabpreins = Request("archivo")

		vfonotrab = trim(UCase(oRs("Teléfono")))
		vemailtrab = trim(UCase(oRs("Email")))
				
                if(vRut(vruttrab)=true and vNombre(vnomtrab)=true and vPaterno(vapatertrab)=true and vMaterno(vamatertrab)=true  and vCargo(vamatertrab)=true and vEscolaridad(oRs("Escolaridad"))=true)then 
                   set rsExtTrab = conn.execute ("select count(*) as existe from TRABAJADOR where RUT='"&vruttrab&"'")
                               
                   vtrabid = rsExtTrab("existe")
                
                if(vtrabid="0")then
                            vIdNacion=""
							
                             if(vcheckExtran="1")then
                                vIdNacion=Request("txtPasTrab")
                             end if
                
                             dim trab
							
        					 trab = "insert into TRABAJADOR (RUT,NOMBRES,CARGO_EMPRESA,ESCOLARIDAD,ESTADO,NOM_TRAB,APATERTRAB,"
                             trab = trab&"AMATERTRAB,NACIONALIDAD,ID_EXTRANJERO,CORREO,EMAIL) "
                             trab = trab&" values('"&vruttrab&"',dbo.MayMinTexto('"&vnomtrab&" "&vapatertrab&" "&vamatertrab&"'),"
                             trab = trab&"dbo.MayMinTexto('"&vcargotrab&"'),'"&vescolaridadtrab&"',1,"
                             trab = trab&"dbo.MayMinTexto('"&vnomtrab&"'),dbo.MayMinTexto('"&vapatertrab&"'),"
                             trab = trab&"dbo.MayMinTexto('"&vamatertrab&"'),'"&vcheckExtran&"','"&vIdNacion&"','"&vfonotrab&"','"&vemailtrab&"') "
							 
                             conn.execute (trab)
                               
                             set rsTrab = conn.execute ("select ID_TRABAJADOR from TRABAJADOR where RUT='"&vruttrab&"'")
                             trabajador_id=rsTrab("ID_TRABAJADOR")
                else
                               'dim trab_up
                               'trab_up = "update TRABAJADOR set NOMBRES=dbo.MayMinTexto('"&vnomtrab&" "&vapatertrab&" "&vamatertrab&"'),"
                               'trab_up = trab_up&"CARGO_EMPRESA=dbo.MayMinTexto('"&vcargotrab&"'),ESCOLARIDAD='"&vescolaridadtrab&"',"
                               'trab_up = trab_up&"NOM_TRAB=dbo.MayMinTexto('"&vnomtrab&"'),APATERTRAB=dbo.MayMinTexto('"&vapatertrab&"'),"
                               'trab_up = trab_up&"AMATERTRAB=dbo.MayMinTexto('"&vamatertrab&"') where RUT='"&vruttrab&"' and ID_TRABAJADOR='"&vtrabid&"'"
                               'conn.execute (trab_up)
                               
                               set rsRegTrab = conn.execute ("select ID_TRABAJADOR from TRABAJADOR where RUT='"&vruttrab&"'")
                               trabajador_id=rsRegTrab("ID_TRABAJADOR")
                end if

                set rsExtInsp = conn.execute ("select count(*) as existe from PREINSCRIPCION_TRABAJADOR where id_trabajador='"&trabajador_id&"' and preinscripcionTemp='"&vtrabpreins&"'")
                
                 dim preIns_Trab
					 if(rsExtInsp("existe")="0")then
						preIns_Trab = "insert into PREINSCRIPCION_TRABAJADOR(id_trabajador,franquicia,preinscripcionTemp) "
						preIns_Trab = preIns_Trab&" values('"&trabajador_id&"',100,'"&vtrabpreins&"')"
						conn.execute (preIns_Trab)
						
						Mensaje=Mensaje&" &bull; Trabajador Ingresado Exitosamente."
						totTrab=totTrab+1
					 else
						Mensaje=Mensaje&" &bull; Trabajador ya Registrado en Curso."
					 end if
				else
					 Mensaje=""
				end if
				
				if(vRut(vruttrab)=true or vNombre(vnomtrab)=true or vPaterno(vapatertrab)=true or vMaterno(vamatertrab)=true or vCargo(vamatertrab)=true or vEscolaridad(oRs("Escolaridad"))=true)then 
								Response.Write("<tr>")
								'Response.Write("<td nowrap>"&UCase(oRs("Nacionalidad"))&"</td>")
								Response.Write("<td width=""30"">"&UCase(vruttrab)&"</td>")
								Response.Write("<td>"&UCase(oRs("Nombre"))&"</td>")
								Response.Write("<td>"&UCase(oRs("Ap_Paterno"))&"</td>")
								Response.Write("<td>"&UCase(oRs("Ap_Materno"))&"</td>")
								Response.Write("<td>"&UCase(oRs("Cargo"))&"</td>")  
								Response.Write("<td>"&UCase(oRs("Escolaridad"))&"</td>")
								Response.Write("<td>"&UCase(oRs("Teléfono"))&"</td>")  
								Response.Write("<td>"&UCase(oRs("Email"))&"</td>")
								Response.Write("<td height=""100%"" width=""22%"">"&Mensaje&"</td>")       
								Response.Write("</tr>")
				end if
      oRs.MoveNext
Wend

response.Write "</tbody></table>" 
response.Write "<input id=""totTrabCargMasiva"" name=""totTrabCargMasiva"" type=""hidden"" value="&totTrab&" />" 

'Se cierra y se destruye el objeto recordset
oRs.Close
Set rsVac = Nothing

'se verifica que no haya escolaridad 0

conn.execute ("exec dbo.setEscolaridadTrabajadores;")


'Se cierra y se destruye el objeto connection
oConexion.Close
Set oConexion = Nothing 

function vRut(rut)
	bRetorno = true
	if ((trim(rut)="") or (isnull(rut))) then 
		bRetorno = false
		Mensaje = Mensaje&" &bull; RUT vac&iacute;o.<br>"
	else
		arr = Split(rut,"-")
		if(ubound(arr)<>1)then  
			bRetorno=false
			Mensaje = Mensaje&" &bull; Formato RUT inv&aacute;lido.<br>"
		elseif(UCase(codigo_veri(arr(0)))<>UCase(arr(1)))then
			bRetorno=false
			Mensaje = Mensaje&" &bull; RUT inv&aacute;lido.<br>"
		end if
	end if
	
	vRut = bRetorno
end Function

Function codigo_veri(rut)
	tur=strreverse(trim(rut))
	mult = 2

	for i = 1 to len(tur)
		if mult > 7 then mult = 2 end if

		suma = mult * mid(tur,i,1) + suma
		mult = mult +1
	next

	valor = 11 - (suma mod 11)

	if valor = 11 then
		codigo_veri = "0"
	elseif valor = 10 then
		codigo_veri = "k"
	else
		codigo_veri = valor
	end if
end function

function vNombre(nom)
	bRetorno = true
	if ((trim(nom)="") or (isnull(nom))) then 
		Mensaje = Mensaje&" &bull; Nombre vac&iacute;o.<br>"
		bRetorno = false
	end if
	vNombre = bRetorno
end Function

function vPaterno(pat)
	bRetorno = true
	if ((trim(pat)="") or (isnull(pat))) then 
		Mensaje = Mensaje&" &bull; Apellido Paterno vac&iacute;o.<br>"
		bRetorno = false
	end if
	vPaterno = bRetorno
end Function

function vMaterno(mat)
	bRetorno = true
	if ((trim(mat)="") or (isnull(mat))) then 
		Mensaje = Mensaje&" &bull; Apellido Materno vac&iacute;o.<br>"
		bRetorno = false
	end if
	vMaterno = bRetorno
end Function

function vCargo(mat)
	bRetorno = true
	if ((trim(mat)="") or (isnull(mat))) then 
		Mensaje = Mensaje&" &bull; Cargo vac&iacute;o.<br>"
		bRetorno = false
	end if
	vCargo = bRetorno
end Function

function vEscolaridad(mat)
	bRetorno = true
	if ((trim(mat)="") or (isnull(mat))) then 
		Mensaje = Mensaje&" &bull; Escolaridad vac&iacute;o.<br>"
		bRetorno = false
	end if
	vEscolaridad = bRetorno
end Function
%>