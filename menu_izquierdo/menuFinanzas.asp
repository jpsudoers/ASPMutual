<ul>
                <%
					dim totOpciones				
					Dim OPCIONES
					OPCIONES=0
	
					if(DATOS("PERMISO7")<>"0")then
					%>
						<li class="first"><a href="consultasRptLF.asp">Facturación</a></li>
                        <li class="first"><a href="finanzashisfacturas.asp">Histórico de Facturas</a></li>  
                    <%else
						OPCIONES=OPCIONES+1
					end if
					if(DATOS("PERMISO8")<>"0")then
					%>
						<li><a href="finanzaspagos.asp">Registro Pagos</a></li>
                    <%else
						OPCIONES=OPCIONES+1
					end if
					if(DATOS("PERMISO9")<>"0")then
					%>
						<li><a href="finanzascuentas.asp">Cuenta Corriente</a></li>
                    <%else
						OPCIONES=OPCIONES+1
					end if
					if(DATOS("PERMISO10")<>"0")then
					%>
						<li><a href="finanzasvencpagos.asp">Informe de Vencimiento</a></li>
                    <%else
						OPCIONES=OPCIONES+1
					end if
					if(DATOS("PERMISO11")<>"0")then
					%>
						<li><a href="finanzasEmpresa.asp">Empresas</a></li>
                        <li><a href="finanzasconsfacturas.asp">Buscar Facturas</a></li>
                    <%else
						OPCIONES=OPCIONES+1
					end if
					
					if(DATOS("PERMISO11")<>"0")then
					%>
						<li><a href="finanzashisttrab.asp">Historico de Cursos</a></li>
                    <%else
						OPCIONES=OPCIONES+1
					end if					
					
					if(DATOS("PERMISO11")<>"0")then
					%>
						<li><a href="finanzashistins.asp">Historico de Inscripciones</a></li>
						<li><a href="finanzasEDP.asp">Listado Estado de Pagos</a></li>
                    <%else
						OPCIONES=OPCIONES+1
					end if							
					
					if(OPCIONES>0)then
						For totOpciones = 1 To OPCIONES Step 1
					%>
						<p>&nbsp;</p>
					<%
						Next
					end if
					%>
			</ul>