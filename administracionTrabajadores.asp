<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml"><head>
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
var emp_del;

$(document).ready(function(){					
	$('#txUsuario').defaultValue('Nombre de Usuario');
	$('#txPassword').defaultValue('Contraseña');

	//$("#frmTrabajador txtFecha").datepicker({changeMonth: true,changeYear: true,yearRange:"-100:+0"});
	//$("#txtFecha").val('2009-01-01');
	

	jQuery("#list1").jqGrid({ 
		url:'trabajadores/listar.asp', 
		datatype: "xml", 
		colNames:['Rut', 'Nombre', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'rut',index:'rut', width:70}, 
				   {name:'nombre_trabajador',index:'nombre_trabajador'}, 
				   { align:"right",editable:true, width:20}, 
				   { align:"right",editable:true, width:20} ], 
		rowNum:10, 
		autowidth: true, 
		rowList:[10,20,30], 
		pager: jQuery('#pager1'), 
		sortname: 'ID_TRABAJADOR', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Trabajadores" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("trabajadores/frmtrabajadores.asp",
															   {id:0},
															   function(f){
																   
																  $('#dialog').html(f);
																  $("#txtFecha").datepicker({changeMonth: true,changeYear: true,yearRange:"-100:+0"});
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
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 350,
			width: 1000,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmTrabajador").valid())
						{
							//window.open($('#frmTrabajador').attr('action')+'?'+$('#frmTrabajador').serialize());
			
							$.post($('#frmTrabajador').attr('action')+'?'+$('#frmTrabajador').serialize(),function(d){
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
						$.post("trabajadores/eliminar.asp",{id:emp_del});
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
		//$('#frmTrabajador txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		$("#frmTrabajador").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txRut:{
				required:true,
				rut:true
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
				required:true,
				number:true
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
				required:true,
				number:true
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
				required:"&bull; Ingrese Rut Trabajador.",
				rut:"&bull; Rut No valido"
			},
			txtDir:{
				required:"&bull; Ingrese Dirección Trabajador."
			},
			txtCom:{
				required:"&bull; Ingrese Comuna Trabajador."
			},
			txtCiu:{
				required:"&bull; Ingrese Ciudad Trabajador."
			},
			txtFon:{
				required:"&bull; Ingrese Teléfono Trabajador.",
				number:"&bull; El Campo Teléfono solo Permite Números."
			},
			txtNomb:{
				required:"&bull; Ingrese Nombre Trabajador."
			},
			txtAp_pater:{
				required:"&bull; Ingrese Apellido Paterno Trabajador."
			},
			txtAp_mater:{
				required:"&bull; Ingrese Apellido Materno Trabajador."
			},
			txtCel:{
				required:"&bull; Ingrese Celular Trabajador.",
				number:"&bull; El Campo Celular solo Permite Números."
			},
			txtMail:{
				required:"&bull; Ingrese Correo Electrónico Trabajador.",
				email:"&bull; Correo No Valido"
			},
			txtCargo:{
				required:"&bull; Ingrese Cargo Trabajador."
			}
		}
	});
    }
	
	function estudios(nro)
    {
         document.getElementById("txtEscol").selectedIndex = nro;
    }
	
	function val(e) {
    tecla = (document.all) ? e.keyCode : e.which; // 2
    if (tecla==8) {return true;}// 3
    patron =/[A-Za-z\s]/; // 4
    te = String.fromCharCode(tecla); // 5
    return patron.test(te); // 6
	}
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}

	function update(i){
		//document.getElementById("frmEmpresa").reset();
		$.post("trabajadores/frmtrabajadores.asp",
			   {id:i},
			   function(f){
				   validar();
				   $('#dialog').html(f);
				   $("#txtFecha").datepicker({changeMonth: true,changeYear: true,firstDay: 1,dateFormat: 'dd-mm-yy'});
				   $('#txRut').focus();
				   $('#txRut').blur();
				   validar();
		});
		 $('#dialog').dialog('open');
	}
	</script>
</head>
<body>
<div id="header">
	<h1><img src="images/logo.png"  /></h1>
	<ul>
		<li><a href="administracion.asp" accesskey="1" title="" class="selItem">Administraci&oacute;n</a></li>
		<li><a href="operacion.asp" accesskey="2" title="">Operaci&oacute;n</a></li>
        		<li><a href="manejocursos.asp" accesskey="3" title="">Manejo de Cursos</a></li>
		<li><a href="finanzas.asp" accesskey="4" title="">Finanzas</a></li>
		<li><a href="empresas.asp" accesskey="5" title="">Empresas</a></li>

	</ul>
</div>
<div id="content">
	<div id="colOne">
  	<h3>Login</h3>
		<div class="bg1">
			Usuario:<strong>Nombre de Usuario</strong>
      <i>Perfil de usuario</i><br />
      <%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%>
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
			</ul>
		</div>		
	</div>
		<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Trabajadores</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Trabajador">
</div>
<div id="dialog_el" title="Eliminar Trabajador">
	<p>Esta seguro de eliminar el Trabajador.</p>
</div>
</body>
</html>