<!--#include file="cnn_string.asp"-->
<%
infoCont="0"
if(Session("countInfoB")<>"")then infoCont=Session("countInfoB") end if

empId="0"
if(Session("usuario")<>"")then empId=Session("usuario") end if
%>
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
<script src="js/jquery.js" language="javascript"></script>
<script src="js/i18n/grid.locale-sp.js" language="javascript"></script>
<script src="js/jquery-ui.js" language="javascript"></script>
<script src="js/jquery.jqGrid.js" language="javascript"></script>
<script src="js/jquery.tbltogrid.js" type="text/javascript" charset="utf-8"> </script> 
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script type="text/javascript">
$(document).ready(function(){
		//tableToGrid("#mytable");   
		//tableToGrid("#mytable2"); 
		
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
		
		$("#avisoInicial").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 460,
			width: 850,
			modal: true
		});
		
		Banner();

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
		
	  $("#Doc").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 500,
			width:1000,
			modal: true,
			overlay: {
				backgroundColor: '#000',
				opacity: 0.5
			},
			buttons: {
				'Cerrar': function() {
					$(this).dialog('close');
				}
			},
			title: 'Documento'
		});		
	 });
	
	 function validarFrm(){
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
		$.post("cambioContrasena/frmContrasenaEmpresa.asp",
			   function(f){
				   validarFrm();
				   $('#pantContrasena').html(f);
				   validarFrm();
		});
		$('#pantContrasena').dialog('open')
	}
	
	function documento(){
		$("#ifPagina").attr('src',"documentos/Instructivo.pdf");
		if(!$('#Doc').dialog('isOpen'))
		{
			$('#Doc').dialog('open');
		}
	}

	function Banner(){
		if(<%=infoCont%>>0){
 			$.post("banner/mensajeEmp.asp",{id:<%=empId%>},
			   function(f){
				    $('#avisoInicial').html(f);
			});
			$("#avisoInicial").dialog('open');
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
	'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
	oConn.Open(MM_cnn_STRING)
	Set DATOS = Server.CreateObject("ADODB.RecordSet")
	DATOS.CursorType=3
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO5,PERMISO6 from USUARIOS "
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
		<li><a href="consultas.asp" accesskey="5" title="">Consultas</a></li>
        <%end if			
	 DATOS.MoveNext
   WEND
   end if
   if(Session("tipoUsuario")="0")then%>
		<li><a href="empresas.asp" accesskey="5" title="" class="selItem">Empresas</a></li>
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
				Usuario : <strong><%=Session("nombre")%> (<%=Session("correo_user")%>)</strong>
      <br />
      <a href="#" onclick="passChange();"><b>Cambiar Contraseña</b></a>
      <br />
      <br />
      <button OnClick="document.location.href='index.asp';">Cerrar Sesión</button>
		</div>
		<h3>Opciones</h3>
		<div class="bg1">
				<ul>
              <%if(Session("tipo_user_empresa")="1")then%>
                   <li class="first"><a href="empresacalendario.asp">Inscripción de Cursos</a></li>
                   <li><a href="empresasinspendientes.asp">Inscripciones Pendientes</a></li>
                   <li><a href="empresainsactivas.asp">Inscripciones Autorizadas</a></li>
				   <li><a href="empresascertificados.asp">Certificados</a></li>
                   <li><a href="empresascuentas.asp">Cuenta Corriente</a></li>
					<li><a href="solicitudescreditoEmpresa.asp">Solicitud Crédito</a></li>
					<li><a href="empresasOC.asp">Ingreso de Orden de Compra</a></li>
               <%else
				   if(Session("cargo_user_empresa")="1")then%>
					   <li class="first"><a href="empresacalendario.asp">Inscripción de Cursos</a></li>
					   <li><a href="empresasinspendientes.asp">Inscripciones Pendientes</a></li>
					   <li><a href="empresainsactivas.asp">Inscripciones Autorizadas</a></li>
					   <li><a href="empresascertificados.asp">Certificados</a></li>
					   <li><a href="solicitudescreditoEmpresa.asp">Solicitud Crédito</a></li>
				   <%else%>
                   <li><a href="empresascuentas.asp">Cuenta Corriente</a></li>
                        <table width="100" height="100" border="0">
                          <tr>
                            <td>&nbsp;</td>
                          </tr>
                        </table>
				   <%end if%>
              <%end if%>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
        <h2><em style="text-transform: capitalize;">Estimados Clientes :</em></h2>
        <table width="700" border="0">
          <tr>
          	<td>
			<p><font size="2">Les informamos que :</font></p>
            	<table width="700" border="0">
            		<tr>
                		<td width="30">&nbsp;</td>  
                		<td width="670"><!--<p style="color:#F00"><font size="2">&bull; A partir del <strong>28 de Noviembre de 2016</strong> todas aquellas empresas que han presentado en nuestro portal bloqueos por concepto de atraso en el pago de facturas <strong>NO</strong> podrán hacer uso de <strong>Ordenes Compra</strong> como forma de pago en sus futuras inscripciones a cursos.</p>-->   
                        <p><font size="2">&bull; Se les informa que todas las solicitudes de inscripción a cursos que se realicen después de las 16:00hrs quedaran eliminadas.
Nuestro horario de atención para solicitud de inscripciones en la plataforma es el siguiente:
<br/>
<strong>Lunes a Jueves hasta las 16:00hrs.<br/>
Viernes hasta las 14:30hrs.</strong></font></p>
             <p><font size="2">&bull; El horario de atención a público, consultas y correos es de lunes a jueves 8:00 a 13:00 y 14:30 a 18:00 y los viernes de 8:00 a 13:00 y 14:30 a 16:00 Hrs.</font></p>
            <!--<p><font size="2">&bull; Para consultas relacionadas con la Programación y Coordinación de Cursos favor contactarse con la <b>Srta. Gladys Contreras</b> al E-mail <b><A HREF="mailto:gpcontreras@mutualcapacitacion.cl">gpcontreras@mutualcapacitacion.cl </A></b> o con la <b>Srta. Loreto Orrego</b> al E-mail <b><A HREF="mailto:lorrego@mutualcapacitacion.cl">lorrego@mutualcapacitacion.cl</A></b> al teléfono <b>(55) 2651283</b>.</font></p>
             <p><font size="2">&bull; Todas las modificaciones o anulaciones de inscripciones a cursos se deben solicitar al E-mail <b><A HREF="mailto:pcaceres@mutualcapacitacion.cl">pcaceres@mutualcapacitacion.cl</A></b> o al teléfono <b>(55) 2651286</b>.</font></p> 
            <p><font size="2">&bull; Para consultas relacionadas con órdenes de compras favor contactarse con la <b>Srta. Melissa Guadalupe</b>  al E-mail <b><A HREF="mailto:mguadalupe@mutualcapacitacion.cl">mguadalupe@mutualcapacitacion.cl</A></b> o al teléfono <b>(55) 2651285</b>.</font></p>
            <p><font size="2">&bull; El horario de entrega de credenciales de cursos realizados en Antofagasta es de 9:00 a 12:30 y 15:00 a 17: 30  de lunes a jueves y los días viernes de 9:00 a 12:30 y 15:00 a 16:00 Hrs.</font></p>
            <p><font size="2">&bull; El horario de entrega de credenciales de cursos realizados en Santiago es de 09:30 a 12:30  Hrs. de lunes a viernes en las oficinas ubicadas en <b>Paseo Bulnes 241 2do piso Oficina H</b>.</font></p>            
            <p><font size="2">&bull; Para consultas relacionadas a la entrega o despachos de credenciales de cursos realizados en Antofagsta  debe contactarse al E-mail <b><A HREF="mailto:mguadalupe@mutualcapacitacion.cl">mguadalupe@mutualcapacitacion.cl</A></b> o al teléfono <b>(55) 2651285</b>.</font></p> 
            <p><font size="2">&bull; Para consultas relacionadas a la entrega o despachos de credenciales de cursos realizados en Santiago debe contactarse al E-mail <b><A HREF="mailto:imontoya@mutualcapacitacion.cl">imontoya@mutualcapacitacion.cl</A></b> o al teléfono <b>(02) 26981523</b>.</font></p>           
            <p><font size="2">&bull; Para consultas relacionadas al portal favor contactarse con <b>Mario González</b> al E-mail <b><A HREF="mailto:magonzalezg@mutualasesorias.cl">magonzalezg@mutualasesorias.cl</A></b> o al teléfono <b>(2) 787 9946</b>.</font></p> --> 
            			</td> 
                   </tr>
				   </table> 
                  </td>  
          </tr>
		</table>
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="avisoInicial" title="Banner Informativo"></div>
</body>
</html>
