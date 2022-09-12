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
<link rel="stylesheet" type="text/css" href="css/ui.all.css"/>
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
<script src="js/jquery.jqGrid.js" type="text/javascript" charset="utf-8"></script>
<script src="js/ajaxfileupload.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
$(document).ready(function(){					
//window.open("inspendientes/listar.asp");

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
		$("#dialog").dialog({
			autoOpen: false,
			bgiframe: true,
			height:640,
			width: 1100,
			modal: true,
			overlay: {
				backgroundColor: '#f00',
				opacity: 0.5
			},
			buttons: {
				'Guardar': function() {
						if($("#frmempresaOC").valid())
						{
							
							subeDoc();
							$(this).dialog('close');
							$("#mensaje").dialog('open');

						}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	/*	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 400,
			width: 480,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmempresaOC").valid())
						{
							//window.open($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize());
							$.post($('#frmempresaOC').attr('action')+'?'+$('#frmempresaOC').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
																						   });
							$("#list1").trigger("reloadGrid");
							//document.getElementById("frmEmpresa").reset();
							$(this).dialog('close');
						}
				},
				Cancelar: function() {
					//document.getElementById("frmEmpresa").reset();
					
					$(this).dialog('close');
				}
			}
		});*/
		
	jQuery("#list1").jqGrid({ 
		url:'empresasOC/listar.asp', 
		datatype: "xml", 
		colNames:['NRO OC','DESCRIPCION', 'MONTO', 'MONTO UTILIZADO', 'MONTO RESTANTE', 'ARCHIVO'], 
		colModel:[
				   {name:'NRO_OC',index:'NRO_OC', width:55}, 
				   {name:'DESCRIPCION',index:'DESCRIPCION', width:55}, 
				   {name:'MONTO',index:'MONTO', width:59}, 
				   {name:'MONTO_UTILIZADO',index:'MONTO_UTILIZADO'}, 
				   {name:'MONTO_RESTANTE',index:'MONTO_RESTANTE'}, 
				   {align:"right",editable:true, width:20}
					], 
		rowNum:10, 
		autowidth: true, 
		rowList:[10,20,30], 
		pager: jQuery('#pager1'), 
		sortname: 'NRO_OC', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listados OC" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
															title:"Agregar nuevo registro", 
															buttonicon :'ui-icon-plus', 
															onClickButton:function(){
																$.post("empresasOC/frmempresaOC.asp",
																	   {id:0},
																	   function(f){
																		   $('#dialog').html(f);
																		
																});
																$('#dialog').dialog('open');
															} }); 
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list1").trigger("reloadGrid");
													  } }); 
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
			title: 'Documento'
		});
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 520,
			width: 1000,
			modal: true,
			buttons: {
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	
	
	
	
	function documento(arch){
		$("#ifPagina").attr('src',"http://norte.otecmutual.cl/ordenes/"+arch);
		if(!$('#Doc').dialog('isOpen'))
		{
			$('#Doc').dialog('open');
		}
	}
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
				  // validarFrm();
				   $('#pantContrasena').html(f);
				   //validarFrm();
		});
		$('#pantContrasena').dialog('open')
	}
	
	
	function subeDoc()
    {
                               $("#loading")
                               .ajaxStart(function(){
                                               $(this).show();
                               })
                               .ajaxComplete(function(){
                                               $(this).hide();
											   $(this).remove();
											   $("#dialog").dialog('close');
					                           $("#txtMensaje").html("Se ha ingresado de forma correcta la orden de compra.");
											   $("#mensaje").dialog('open');
                               });
							   
                               $.ajaxFileUpload
                               (
                                      {
                                           url:'empresasOC/subir.asp?',
                                           secureuri:false,
                                           fileElementId:'txtDoc',
                                           elements:
										   'txtNroOC;txtDescripcionOC;txtMon;txtArchi',										
										   dataType: 'json',
                                           success: function (data, status)
                                           {
											   if(typeof(data.error) != 'undefined')
											   {
												  if(data.error != '')
												  {
													 alert("Problemas "+data.error);
												  }
												  else      
												  {
													$("#list1").trigger("reloadGrid");
												  }
											   }
                                           },
                                           error: function (data, status, e)
                                           {
                                                 $("#list1").trigger("reloadGrid");
                                           }
                                         }
                               );
                               return false;
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
      <!--<%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%>-->
	  
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
			<h2><em style="text-transform: capitalize;">Listados Ordenes de Compra</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="INGRESO DE ORDEN DE COMPRA">
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="mensaje" title="Registro de Inscripción">
     <label id="txtMensaje" name="txtMensaje"></label>
</div>
</body>
</html>
