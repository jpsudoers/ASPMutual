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
var emp_del;

$(document).ready(function(){					
	$('#txUsuario').defaultValue('Nombre de Usuario');
	$('#txPassword').defaultValue('Contraseña');

	jQuery("#list1").jqGrid({ 
		url:'solicitudes/listar.asp', 
		datatype: "xml", 
		colNames:['Rut', 'Nombre Empresa', 'Código Curso', 'Nombre Curso', 'Fecha Inicio', '&nbsp;'], 
		colModel:[
				   {name:'RUT',index:'RUT', width:58}, 
				   {name:'R_SOCIAL',index:'R_SOCIAL'}, 
				   {name:'CODIGO',index:'CODIGO', width:70}, 
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO'}, 
				   {name:'FECHA',index:'FECHA', width:71}, 
				   { align:"right",editable:true, width:20} ], 
		rowNum:10, 
		autowidth: true, 
		rowList:[10,20,30], 
		pager: jQuery('#pager1'), 
		sortname: 'SOLICITUD.id_solicitud', 
		viewrecords: true, 
		sortorder: "desc", 
		caption:"Listado de Solicitudes" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list1").trigger("reloadGrid");
													  } }); 
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height:400,
			width: 900,
			modal: true,
			buttons: {
				'Cerrar': function() {
					$(this).dialog('close');
				}
			}
		});
	});	

	function update(i){
		$.post("solicitudes/frmsolicitudes.asp",
			   {id:i},
			   function(f){
				    $('#dialog').html(f);
		});
		 $('#dialog').dialog('open');
	}
	</script>
</head>
<body>
<div id="header">
	<h1><img src="images/logo.png"  /></h1>
	<ul>
		<li><a href="administracion.asp" accesskey="1" title="">Administraci&oacute;n</a></li>
		<li><a href="#" accesskey="2" title="" class="selItem">Operaci&oacute;n</a></li>
        <li><a href="manejocursos.asp" accesskey="5" title="">Manejo de Cursos</a></li>
		<li><a href="finanzas.asp" accesskey="3" title="">Finanzas</a></li>
		<li><a href="empresas.asp" accesskey="4" title="">Empresas</a></li>
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
				<li class="first"><a href="operacionprogramacion.asp">Programación</a></li>
				<li><a href="operacionautorizacion.asp">Autorización</a></li>
				<li><a href="operacioninscripcion.asp">Inscripción</a></li>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Solicitudes</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro de Solicitud">
</div>
</body>
</html>
