<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title></title>
<meta name="Keywords" content="" />
<meta name="Description" content="" />
<link href="css/default.css" rel="stylesheet" type="text/css" />
<link href="css/jquery-ui.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" type="text/css" href="css/ui.all.css"/>
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
<script src="js/jquery.jqGrid.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
$(document).ready(function(){					
	$("#codempresa").append("<option value=\"\">Seleccione</option>");
	$("#codTrabajador").append("<option value=\"\">Seleccione</option>");
	llena_curriculo(0);		   
    
	$("#Doc").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 600,
			width:1000,
			modal: true,
			overlay: {
				backgroundColor: '#000',
				opacity: 0.5
			},
			buttons: {
				'Aceptar': function() {
					$(this).dialog('close');
				}
			},
			title: 'Certificado Asistencia A Curso De Capacitación'
		});
	
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
	});	

	function llena_curriculo(id_curriculo){
			$("#codprog").html("");
			$("#codempresa").html("");
			$("#codTrabajador").html("");
			$("#codempresa").append("<option value=\"\">Seleccione</option>");
	        $("#codTrabajador").append("<option value=\"\">Seleccione</option>");
			
			$.get("libroclases/curriculo.asp",
						function(xml){
							$("#codprog").append("<option value=\"\">Seleccione</option>");
							$('row',xml).each(function(i) { 
								if(id_curriculo==$(this).find('ID_MUTUAL').text())
									$("#codprog").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" selected="selected">'+
																			$(this).find('NOMBRES').text()+ '</option>');
								else
									$("#codprog").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" >'+
																			$(this).find('NOMBRES').text()+ '</option>');
							});
						});
	}	   
	
	function llena_empresa(id_programa,empresa){
			$("#codempresa").html("");
			$("#codTrabajador").html("");

	        $("#codTrabajador").append("<option value=\"\">Seleccione</option>");
			$.get("libroclases/empresa.asp",
				   {id_programa:id_programa},
						function(xml){
							$("#codempresa").append("<option value=\"\">Seleccione</option>");
							
							$('row',xml).each(function(i) { 
								if(empresa==$(this).find('ID').text())
									$("#codempresa").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																			$(this).find('NOMBRES').text()+ '</option>');
								else
									$("#codempresa").append('<option value="'+$(this).find('ID').text()+'" >'+
																			$(this).find('NOMBRES').text()+ '</option>');
							});
						});
	}	   
	
	function llena_trabajador(id_empresa,trabajador){
			$("#codTrabajador").html("");
			$("#codTrabajador").append("<option value=\"\">Seleccione</option>");
			$.get("libroclases/trabajador.asp",
				   {id_empresa:id_empresa,id_programa:$("#codprog").val()},
						function(xml){
							$("#codTrabajador").append("<option value='0'>Todos</option>");
							$('row',xml).each(function(i) { 
								if(trabajador==$(this).find('ID').text())
									$("#codTrabajador").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																			$(this).find('NOMBRES').text()+ '</option>');
								else
									$("#codTrabajador").append('<option value="'+$(this).find('ID').text()+'" >'+
																			$(this).find('NOMBRES').text()+ '</option>');
							});
						});
	}	   
	
	function VerInforme()
	{
		validar();
		if($("#frminforme").valid())
		{
documento("libroclases/pdf.asp?prog="+$("#codprog").val()+"&empresa="+$("#codempresa").val()+"&trabajador="+$("#codTrabajador").val()+"&relator=0");
		}
	}
	
	function documento(arch){
		//alert("http://norte.otecmutual.cl/ordenes/"+arch)
		$("#ifPagina").attr('src',arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
	}
	
	function validar(){
		$("#frminforme").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			codprog:{
				required:true
			},
			codempresa:{
				required:true
			},
			codTrabajador:{
				required:true
			}
		},
		messages:{
			codprog:{
				required:"&bull; Seleccione Código de Programación."
			},
			codempresa:{
				required:"&bull; Seleccione Nombre de Empresa."
			},
			codTrabajador:{
				required:"&bull; Seleccione Nombre Trabajador."
			}
			
		}
	});
    }function validarFrm(){
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
	oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
	Set DATOS = Server.CreateObject("ADODB.RecordSet")
	DATOS.CursorType=3
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4 from USUARIOS "
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
        <li><a href="manejocursos.asp" accesskey="3" title="" class="selItem">Manejo de Cursos</a></li>
        <%end if
		if(DATOS("PERMISO4")<>"0")then
		%>
		<li><a href="finanzas.asp" accesskey="4" title="">Finanzas</a></li>
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
				<li class="first"><a href="manejocursoslibroclases.asp">Libro de Clases</a></li>
                <li><a href="manejoeva_cierre.asp">Evaluación y Cierre</a></li>
                <li><a href="manejocursoscertificados.asp">Certificados</a></li>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
        <h2><em style="text-transform: capitalize;">Certificados</em></h2>
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
					<table width="620" border="0" align="center" cellpadding="1" cellspacing="1">
                     	<tr>
    						<td width="150">&nbsp;</td>
                            <td width="150">&nbsp;</td>
                            <td width="100">&nbsp;</td>
                            <td width="100">&nbsp;</td>
                       
   						</tr>
						<tr>
                            <td>Código Programacion :</td>
                            <td><select name="codprog" id="codprog" onchange="llena_empresa(this.value,0)"></select></td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td rowspan="3"><input type="button" class="boton" value="Ver Informe" onclick="VerInforme();" id="btnBuscar" name="btnBuscar"/></td>
						</tr>
                        <tr>
						  <td>Empresa :</td>
						  <td colspan="3"><select name="codempresa" id="codempresa" onchange="llena_trabajador(this.value,0)">
					      </select></td>
						</tr>
                        <tr>
						  <td>Trabajador :</td>
						  <td><select name="codTrabajador" id="codTrabajador" class="formulario"></select></td>
					      <td>&nbsp;</td>
					      <td></td>
						</tr>
                        <tr>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
   					   </tr>
					</table>
				</form> 
                    
				</div>	
                <div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
                   	 <ul></ul> 
               		</div>		
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
</body>
</html>
