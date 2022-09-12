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
<script src="js/jquery.jqGrid.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
var emp_del;

$(document).ready(function(){					
	$('#txUsuario').defaultValue('Nombre de Usuario');
	$('#txPassword').defaultValue('Contraseña');
//window.open("mutuales/listar.asp");
	jQuery("#list1").jqGrid({ 
		url:'mutuales/listar.asp', 
		datatype: "xml", 
		colNames:['Rut', 'Razón Social', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'rut',index:'rut', width:22, align:'center'}, 
				   {name:'razon social',index:'R_SOCIAL'}, 
				   { align:"right",editable:true, width:10}, 
				   { align:"right",editable:true, width:9} ], 
		rowNum:30, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager1'), 
		sortname: 'id_otic', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Mutual" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("mutuales/frmMutual.asp",
															   {id:0},
															   function(f){
																   $('#dialog').html(f);
																   validar();
														});
														$('#dialog').dialog('open');
													} }); 

	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list1").trigger("reloadGrid");
													  } }); 
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
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 400,
			width: 930,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmMutual").valid())
						{
							$.post($('#frmMutual').attr('action')+'?'+$('#frmMutual').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
																						   });
							$("#list1").trigger("reloadGrid");
							//document.getElementById("frmOtic").reset();
							$(this).dialog('close');
						}
				},
				Cancelar: function() {
					//document.getElementById("frmOtic").reset();
					$(this).dialog('close');
				}
			}
		});
		
		$("#dialog_el").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
						$.post("mutuales/eliminar.asp",{id:emp_del});
						$(this).dialog('close');
						$("#list1").trigger("reloadGrid");
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	

	function eliminar(i){
	 emp_del=i;
	 $('#dialog_el').dialog('open');
	}

	function validar(){
		$("#frmMutual").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txRut:{
				required:true,
				rut:true
			},
			txtRsoc:{
				required:true
			},
			txtGiro:{
				required:true
			},
			txtDir:{
				required:true
			},
			txtCom:{
				required:true
			},
			txtCiu:{
				required:true
			},
			txtFon:{
				required:true
			},
			txtFax:{
				required:true
			},
			txtRut2:{
				required:true,
				rut:true
				//equalTo: "#txRut"
				//distinctTo: "#txRut",
			},
			txtNomb:{
				required:true
			},
			txtAp_pater:{
				required:true
			},
			txtAp_mater:{
				required:true
			},
			txtCel:{
				required:true
				//,number:true
			},
			txtMail:{
				required:true,
				email:true
			},
			txtCargo:{
				required:true
			}
		},
		messages:{
			txRut:{
				required:"&bull; Ingrese Rut Empresa.",
				rut:"&bull; Rut No valido"
			},
			txtRsoc:{
				required:"&bull; Ingrese Razón Social Empresa."
			},
			txtGiro:{
				required:"&bull; Ingrese Giro Empresa."
			},
			txtDir:{
				required:"&bull; Ingrese Dirección Empresa."
			},
			txtCom:{
				required:"&bull; Ingrese Comuna Empresa."
			},
			txtCiu:{
				required:"&bull; Ingrese Ciudad Empresa."
			},
			txtFon:{
				required:"&bull; Ingrese Teléfono Empresa.",
				number:"&bull; El Campo Teléfono solo Permite Números."
			},
			txtFax:{
				required:"&bull; Ingrese Fax Empresa.",
				number:"&bull; El Campo Fax solo Permite Números."
			},
			txtRut2:{
				required:"&bull; Ingrese Rut Contacto.",
				rut:"&bull; Rut No valido"
				//equalTo:"&bull; El Rut Contacto Debe Ser Distinto Al Rut OTIC",
				//distinctTo:"&bull; El Rut Contacto Debe Ser Distinto Al Rut OTIC",
			},
			txtNomb:{
				required:"&bull; Ingrese Nombre Contacto."
			},
			txtAp_pater:{
				required:"&bull; Ingrese Apellido Paterno Contacto."
			},
			txtAp_mater:{
				required:"&bull; Ingrese Apellido Materno Contacto."
			},
			txtCel:{
				required:"&bull; Ingrese Celular Contacto."
				//,number:"&bull; El Campo Celular solo Permite Números."
			},
			txtMail:{
				required:"&bull; Ingrese Correo Electrónico Contacto.",
				email:"&bull; Correo No Valido"
			},
			txtCargo:{
				required:"&bull; Ingrese Crago Contacto."
			}
		}
	});
    }
	
	function update(i){
		//document.getElementById("frmEmpresa").reset();
		$.post("mutuales/frmMutual.asp",
			   {id:i},
			   function(f){
				   validar();
				   $('#dialog').html(f);
				   validar();
		});
		 $('#dialog').dialog('open');
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
	'oConn.Open("Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas")
	oConn.Open(MM_cnn_STRING)
	Set DATOS = Server.CreateObject("ADODB.RecordSet")
	DATOS.CursorType=3
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO5,PERMISO6,PERMISO12 from USUARIOS "
	sql = sql&" where USUARIOS.ID_USUARIO='"&Session("usuarioMutual")&"'"

   DATOS.Open sql,oConn
   WHILE NOT DATOS.EOF
		if(DATOS("PERMISO1")<>"0")then
		%>
		<li><a href="administracion.asp" accesskey="1" title="" class="selItem">Administraci&oacute;n</a></li>
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
				<li class="first"><a href="administracionUsuario.asp">Usuarios</a></li>
				<li><a href="administracionEmpresa.asp">Empresas</a></li>
                <li><a href="administracionOtic.asp">OTIC</a></li>
				<li><a href="administracionmutual.asp">Mutuales</a></li>
                <li><a href="administracionInstructor.asp">Relatores</a></li>
                <li><a href="administracionCurriculo.asp">Portafolio</a></li>
                <li><a href="administracionsede.asp">Sedes</a></li>
			</ul>
		</div>		
	</div>
		<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Mutuales</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Mutual">
</div>
<div id="dialog_el" title="Eliminar Mutual">
	<p>Esta seguro de eliminar la Mutual.</p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>