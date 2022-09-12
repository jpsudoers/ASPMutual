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
//window.open("usuario/listar.asp");
	jQuery("#list1").jqGrid({ 
		url:'usuario/listar.asp', 
		datatype: "xml", 
		colNames:['Rut', 'Nombre', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'rut',index:'rut', width:22, align:'right'}, 
				   {name:'nom_usuario',index:'nombre_usuario'}, 
				   { align:"right",editable:true, width:10}, 
				   { align:"right",editable:true, width:10} ], 
		rowNum:100, 
		autowidth: true, 
		rowList:[100,300,500], 
		pager: jQuery('#pager1'), 
		sortname: 'nombre_usuario', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Usuarios" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	/**/jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														/**/$.post("usuario/frmUsuario.asp",
															   {id:0},
															   function(f){
																  $('#dialog').html(f);
																  //$('#txtFecha').datepicker();
																  $('#txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
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
			height: 595,
			width: 930,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmUsuario").valid())
						{
							//window.open($('#frmUsuario').attr('action')+'?'+$('#frmUsuario').serialize())
							$.post($('#frmUsuario').attr('action')+'?'+$('#frmUsuario').serialize(),function(d){
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
						$.post("usuario/eliminar.asp",{id:emp_del},function(d){
																			$("#list1").trigger("reloadGrid");
																			});
						$("#list1").trigger("reloadGrid");
						$(this).dialog('close');
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	

	function eliminar(i){
	 /*emp_del=i;
	 $('#dialog_el').dialog('open');*/
	}

	function Permisos(){
 	  	$('#p4').attr('checked', false);	
	  	if($('#p7').attr("checked") || $('#p8').attr("checked") || $('#p9').attr("checked") || $('#p10').attr("checked") || $('#p11').attr("checked")) 
		{
			$('#p4').attr('checked', true);
		}
	}

	function PermisosManuales(){
 	  	$('#p12').attr('checked', false);	
	  	if($('#p13').attr("checked") || $('#p14').attr("checked") || $('#p15').attr("checked")) 
		{
			$('#p12').attr('checked', true);
		}
	}

	function validar(){
		$("#frmUsuario").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txRut:{
				required:true,
				rut:true,
				remote: {
						url: "usuario/buscaUsuario.asp",
						type: "post",
						data: {
						  txtRut: function() {
								return $('#txRut').val();
						  }
						}
				}
			},
			txtNomb:{
				required:true
			},
			txtAp_pater:{
				required:true
			},
			txtEmail:{
				required:true,
				email:true
			},
			txtAp_mater:{
				required:true
			}
		},
		messages:{
			txRut:{
				required:"&bull; Ingrese Rut Usuario.",
				rut:"&bull; Rut No valido",
				remote:"&bull; Rut de Usuario ya Registrado."
			},
			txtNomb:{
				required:"&bull; Ingrese Nombre Usuario."
			},
			txtAp_pater:{
				required:"&bull; Ingrese Apellido Paterno Usuario."
			},
			txtEmail:{
				required:"&bull; Ingrese Correo Electrónico Usuario.",
				email:"&bull; Correo No Valido"
			},
			txtAp_mater:{
				required:"&bull; Ingrese Apellido Materno Usuario."
			}
		}
	});
    }
	
	/*function val(e) {
    tecla = (document.all) ? e.keyCode : e.which; // 2
    if (tecla==8) return true; // 3
    patron =/[A-Za-z\s]/; // 4
    te = String.fromCharCode(tecla); // 5
    return patron.test(te); // 6
	}*/
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}

	
	function update(i){
		//document.getElementById("frmEmpresa").reset();
		$.post("usuario/frmUsuario.asp",
			   {id:i},
			   function(f){
				   validar();
				   $('#dialog').html(f);
				   validar();
		});
		 $('#dialog').dialog('open');
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
			<h2><em style="text-transform: capitalize;">Usuarios</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Usuario">
</div>
<div id="dialog_el" title="Eliminar Usuario">
	<p>Esta seguro de eliminar el Usuario.</p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>