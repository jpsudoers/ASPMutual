<!--#include file="cnn_string.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title></title>
<meta name="Keywords" content="" />
<meta name="Description" content="" />
<link href="css/default.css" rel="stylesheet" type="text/css" />
<link href="css/jquery-ui.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
<script type="text/javascript">
$(document).ready(function(){
						   
		$("#pantContrasena").dialog({
			autoOpen: false,
			bgiframe: true,
			height:250,
			width: 430,
			modal: true,
			buttons: {
				'Guardar': function() {
					if($("#frmContrasena").valid())
					{
						if($("#txtContrasena").val()==$("#ConAntigua").val())
						{
							$.post($('#frmContrasena').attr('action')+'?'+$('#frmContrasena').serialize(),function(d){});
							$(this).dialog('close');
							$("#txtmContrasena").html("Contraseña Modificada Exitosamente.");
							$("#mContrasena").dialog('open');
						}
						else
						{
							$("#txtmContrasena").html("La Contraseña Antigua Ingresada es Incorrecta.");
							$("#mContrasena").dialog('open');
						}
					}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
		
		$("#mContrasena").dialog({
			autoOpen: false,
			bgiframe: true,
			height:80,
			width: 450,
			modal: true,
			buttons: {
				Aceptar: function() {
					$(this).dialog('close');
				}
			}
		});
	 });function validarFrm(){
		$("#frmContrasena").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			ConAntigua:{
				required:true
			},
			NuevaCont:{
				required:true,
				rangelength: [5, 50]
			},
			RepConNueva:{
				 required:true,
				 equalTo: "#NuevaCont"
			}
		},
		messages:{
			ConAntigua:{
				required:"&bull; Ingrese la Contraseña Antigua."
			},
			NuevaCont:{
				required:"&bull; Ingrese la Nueva Contraseña.",
				rangelength:"&bull; La Nueva Contraseña debe Contener al Menos 5 Caracteres."
			},
			RepConNueva:{
				required:"&bull; Vuelva a Ingresar la Contraseña Nueva.",
				equalTo:"&bull; Las Contraseñas Nuevas no Coinciden."
			}
		}
	});
    }

	function passChange(){
		$.post("cambioContrasena/frmContrasena.asp",
			   function(f){
				   validarFrm();
				   $('#pantContrasena').html(f);
				   validarFrm();
		});
		$('#pantContrasena').dialog('open')
	}

	function VerInforme()
	{
		window.open("consultascomercial/xls.asp?f_ini="+$('#txtFInicio').val()+"&f_fin="+$('#txtFTermino').val()+"&mes="+$('#SelMes').val()+"&ano="+$('#selAno').val()+"&p=0&c=1&v="+$('#SelVer').val(),'Informe')
		if($.trim($('#txtFInicio').val()) != ""){
			$('#txtFInicio').css("border","1px solid #000");
			fecha_ini = fecha_mysql($('#txtFInicio').val());
				if($.trim($('#txtFTermino').val()) != ""){
					$('#txtFTermino').css("border","1px solid #000");
					fecha_fin = fecha_mysql($('#txtFTermino').val());	
					if(fecha_ini <= fecha_fin){
						//if ($("input[@name='optVer']:checked").val()=="1")	
							//documento("manualesexistencias/pdf.asp?f_ini="+$('#txtFInicio').val()+"&f_fin="+$('#txtFTermino').val());
						//else
							window.open("consultascomercial/xls.asp?f_ini="+$('#txtFInicio').val()+"&f_fin="+$('#txtFTermino').val()+"&mes="+$('#SelMes').val()+"&ano="+$('#selAno').val()+"&p="+$('#SelTipoProy').val()+"&c=<%=clienteID%>",'Informe')
					}
					else{
						$('#txtFInicio').css("border","1px solid #F00");
						$('#txtFTermino').css("border","1px solid #F00");
					}
				}else{
					$('#txtFTermino').css("border","1px solid #F00");
					}
		}else{
			$('#txtFInicio').css("border","1px solid #F00");
		}
	}
	</script>
</head>
<body>
<div id="header">
	<h1><img src="images/logo.png"  /></h1>
	<ul>
    <%if(Session("tipoUsuario")="1")then
	Dim DATOS
	Dim oConn
	SET oConn = Server.CreateObject("ADODB.Connection")
	'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbarauco")
	oConn.Open(MM_cnn_STRING)
	Set DATOS = Server.CreateObject("ADODB.RecordSet")
	DATOS.CursorType=3
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO5,PERMISO6,PERMISO12 from USUARIOS "
	sql = sql&" where USUARIOS.ID_USUARIO='"&Session("usuarioMutual")&"'"

   DATOS.Open sql,oConn
   WHILE NOT DATOS.EOF
		if(DATOS("PERMISO1")<>"0")then
		%>
		<li><a href="administracion.asp" accesskey="1" title="">Administraci&oacute;n</a></li>
        <%end if
		if(DATOS("PERMISO2")<>"0")then
		%>
		<li><a href="operacion.asp" accesskey="2" title="">Operaci&oacute;n</a></li>
        <%end if
		if(DATOS("PERMISO3")<>"0")then
		%>
        <li><a href="manejocursos.asp" accesskey="3" title="">Manejo de Cursos</a></li>
        <%end if
		if(DATOS("PERMISO4")<>"0")then
		%>
		<li><a href="finanzas.asp" accesskey="4" title="">Finanzas</a></li>
        <%end if
		if(DATOS("PERMISO6")<>"0")then
		%>
		<li><a href="consultas.asp" accesskey="5" title="" class="selItem">Consultas</a></li>
        <%end if
		if(DATOS("PERMISO12")<>"0")then
		%>
		<li><a href="manuales.asp" accesskey="6" title="">Manuales</a></li>
        <%end if	
	 DATOS.MoveNext
   WEND
   end if
   if(Session("tipoUsuario")="0")then%>
		<li><a href="empresas.asp" accesskey="5" title="">Empresas</a></li>
    <%end if
	 if(Session("tipoUsuario")="")then
		Session.Abandon
		Response.Redirect("index.asp")
	 end if%>
	</ul>
</div>
<div id="content">
	<div id="colOne">
  	<h3>Login</h3>
	<div class="bg1" align="left">
				Usuario : <strong><%=Session("nombre")%></strong>
      <br />
      <a href="#" onclick="passChange();"><b>Cambiar Contraseña</b></a>
      <br />
      <br />
      <button OnClick="document.location.href='index.asp';">Cerrar Sesión</button>
		</div>
		<h3>Opciones</h3>
		<div class="bg1">
			<ul>
                <li class="first"><a href="consultashistorico.asp">Historico de Cursos</a></li>
                <li><a href="consultasfinanzas.asp">Finanzas</a></li>
                <li><a href="consultashistval.asp">Historica Valorizada</a></li>                
                <li><a href="consultascertificados.asp">Certificados</a></li>
                <li><a href="consultas_Comercial.asp">Informe Comercial</a></li> 
		<li><a href="consultas_cert_hist.asp">Certificados Emitidos</a></li>                  
                <li><a href="ConsultasEmpresas.asp">Empresas</a></li>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
		    	<h2><em style="text-transform: capitalize;">Informe</em><em style="text-transform: capitalize;"> Comercial</em></h2>
        <br />
        <div>
						<table cellpadding="0" cellspacing="0" style="border:0px;width:100%;" class="cabezaMsj">
							<tr class="buscarTitulo">
								<td><h3><em style="text-transform: capitalize;">Informes</em></h3></td>
							</tr>
						</table>
		</div>
		<div id="buscarContent" style="border:1px solid #000099;display:block" >
					<form name="frminforme" id="frminforme" action="" method="post">
                    <table width="640" border="0" align="center" cellpadding="1" cellspacing="1">
                        <tr>
    						<td width="130">&nbsp;</td>
                            <td width="100">&nbsp;</td>
    						<td width="140">&nbsp;</td>
                            <td width="100">&nbsp;</td>
                            <td width="100">&nbsp;</td>
                            <td width="150">&nbsp;</td>
        				</tr>
<tr>
                            <td>Mes : </td>
					        <td><select id="SelMes" name="SelMes" tabindex="1">
                                    <OPTION VALUE="0">Todos</OPTION>
                                    <OPTION VALUE="1">Enero</OPTION>
                                    <OPTION VALUE="2">Febrero</OPTION>
                                    <OPTION VALUE="3">Marzo</OPTION>
                                    <OPTION VALUE="4">Abril</OPTION>
                                    <OPTION VALUE="5">Mayo</OPTION>
                                    <OPTION VALUE="6">Junio</OPTION>
                                    <OPTION VALUE="7">Julio</OPTION>
                                    <OPTION VALUE="8">Agosto</OPTION>
                                    <OPTION VALUE="9">Septiembre</OPTION>
                                    <OPTION VALUE="10">Octubre</OPTION>
                                    <OPTION VALUE="11">Noviembre</OPTION>
                                    <OPTION VALUE="12">Diciembre</OPTION>
                            </select></td>
                            <td>Año : </td>
                            <td><select id="selAno" name="selAno" tabindex="2">
                                            	<OPTION VALUE="0">Todos</OPTION>
											<%
                                            For ano = 2010 To cdbl(year(now)+1) Step 1
                                            %>
                                                <OPTION VALUE="<%=ano%>" <%if(ano=cdbl(year(now)))then%>selected="selected"<%end if%>><%=ano%></OPTION>
                                            <%
                                            Next
                                            %>
                            </select></td>
                             <td><input type="button" value="Ver Informe" onclick="VerInforme();" id="btnBuscar" name="btnBuscar"/></td>
                        </tr>           
						<!--<tr>
                            <td>Fecha de Inicio :</td>
					        <td><input id="txtFInicio" name="txtFInicio" type="text" maxlength="10" size="12"/> 
                            </td>
                            <td>Fecha de Término :</td>
					        <td><input id="txtFTermino" name="txtFTermino" type="text" maxlength="10" size="12"/>  
							</td>
                            <td><input type="button" value="Ver Informe" onclick="VerInforme();" id="btnBuscar" name="btnBuscar"/></td>
						</tr>
              			<tr>
                       		<td>Proyecto :</td>
                    		<td><select id="SelTipoProy" name="SelTipoProy" tabindex="5" onchange="tabla();">
                			<option value="0">Todos</option>
                    		<option value="1">Abierto</option>
                    		<option value="2">Cerrado</option>
                    		<option value="5">Cerrado - Collahuasi</option>                             
                    		<option value="3">Proyecto Almagro</option>
                    		<option value="4">Proyecto Spot Mutual</option>
                            	<option value="6">Proyecto Sigro</option>
                  			</select></td>
              			</tr>     -->  
				<tr>               
 				   <td>Ver : </td>
				   <td><select id="SelVer" name="SelVer" tabindex="1">
                                    			<OPTION VALUE="1" selected="selected">Cliente Actual</OPTION>
                                    			<OPTION VALUE="2">Consolidado</OPTION>
                            	       </select></td>     
				</tr>   
                            <tr>
                            	<td colspan="6">&nbsp;</td>
   			    </tr>
					</table>
                </form> 
			</div>	
		</div>
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2011 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>
