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
<style type="text/css">
    .suggestionsBox {
        position: relative;
        left: 10px;
        margin: 10px 0px 0px 0px;
        width: 400px;
        background-color: #212427;
        -moz-border-radius: 7px;
        -webkit-border-radius: 7px;
        border: 2px solid #000;   
        color: #fff;
    }
   
    .suggestionList {
        margin: 0px;
        padding: 0px;
    }
   
    .suggestionList li {
       
        margin: 0px 0px 3px 0px;
        padding: 3px;
        cursor: pointer;
    }
   
    .suggestionList li:hover {
        background-color: #659CD8;
    }
</style>
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
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("solicitudesempresas/frmsolicitudes.asp",
															   {id:0},
															   function(f){
																  $('#dialog').html(f);
																  $('#txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
																  llena_curriculo(0);
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
			height:260,
			width: 900,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmSolicitudes").valid())
						{
							//window.open($('#frmSolicitudes').attr('action')+'?'+$('#frmSolicitudes').serialize());
							$.post($('#frmSolicitudes').attr('action')+'?'+$('#frmSolicitudes').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
																						   });
							$("#list1").trigger("reloadGrid");
							$(this).dialog('close');
						}
				},
				Cancelar: function() {
					//document.getElementById("frmOtic").reset();
					$(this).dialog('close');
				}
			}
		});
	});	

function validar(){
		//$('#frmAutorizacion txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		$("#frmSolicitudes").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txRut:{
				required:true
			},
			txtCurso:{
				required:true
			},
			txtFecha:{
				required:true
			},
			txtPart:{
				required:true
			}
		},
		messages:{
			txRut:{
				required:"&bull; Ingrese Rut de la Empresa."
			},
			txtCurso:{
				required:"&bull; Seleccione Curso."
			},
			txtFecha:{
				required:"&bull; Seleccione Fecha Requerida para el Curso."
			},
			txtPart:{
				required:"&bull; Seleccione Número de Participantes."
			}
		}
	});
    }

    function lookup(inputString) {
		if(inputString.length <=5) {
			$('#suggestions').hide();
		} else {
				$.post("autorizacion/sugEmpresa.asp", {txt: inputString}, function(data){
						if(data.length >0) {
								$('#suggestions').show();
								$('#autoSuggestionsList').html(data);
						}
				});
		}
	}

	function fill(id,rut) {
	   $('#Empresa').val(id);
	   cargaDatos(id);
	   $('#txRut').val(rut);
	   $('#suggestions').hide();
	}
	
	function llena_curriculo(id_curriculo){
		$("#txtCurso").html("");
		$.get("solicitud/curriculo.asp",
					function(xml){
						$("#txtCurso").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_curriculo==$(this).find('ID_MUTUAL').text())
								$("#txtCurso").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$("#txtCurso").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
					});
	}
	
	function cargaDatos(id)	
	{
		$("#txtRsocOtic").html("");
		$("#txtRutOtic").html("");
		$("#txtIdOtic").val("");
		$("#txRut").val("");
		$("#txtRsoc").html("");

		$.get("autorizacion/datosempresa.asp",
						 {id:id},
						 function(xml){
									$('row',xml).each(function(i) { 
									$("#txtRsoc").html($(this).find('RSOCIAL').text());
									$("#txRut").val($(this).find('RUT').text());
									$("#txtRsocOtic").html($(this).find('RSOCIALOTIC').text());
									$("#txtRutOtic").html($(this).find('RUTOTIC').text());
									$("#txtIdOtic").val($(this).find('IDOTIC').text());
						});
			});
	}

	function update(i){
		$.post("solicitudesempresas/frmsolicitudes.asp",
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
		<li><a href="operacion.asp" accesskey="2" title="">Operaci&oacute;n</a></li>
        		<li><a href="manejocursos.asp" accesskey="3" title="">Manejo de Cursos</a></li>
		<li><a href="finanzas.asp" accesskey="4" title="">Finanzas</a></li>
		<li><a href="#" accesskey="5" title="" class="selItem">Empresas</a></li>
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
              <%if(Session("tipo_user_empresa")="1")then%>
                   <li class="first"><a href="empresacalendario.asp">Inscripción de Cursos</a></li>
                   <li><a href="empresasinspendientes.asp">Inscripciones Pendientes</a></li>
                   <li><a href="empresainsactivas.asp">Inscripciones Autorizadas</a></li>
                   <li><a href="empresascertificados.asp">Certificados</a></li>
                   <li><a href="empresascuentas.asp">Cuenta Corriente</a></li>
               <%else
				   if(Session("cargo_user_empresa")="1")then%>
					   <li class="first"><a href="empresacalendario.asp">Inscripción de Cursos</a></li>
					   <li><a href="empresasinspendientes.asp">Inscripciones Pendientes</a></li>
					   <li><a href="empresainsactivas.asp">Inscripciones Autorizadas</a></li>
					   <li><a href="empresascertificados.asp">Certificados</a></li>
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
