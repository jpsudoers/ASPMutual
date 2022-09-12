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

//window.open("autorizacion/listar.asp");

	jQuery("#list1").jqGrid({ 
		url:'autorizacion/listar.asp', 
		datatype: "xml", 
		colNames:['Cód Auto.', 'Rut', 'Nombre Empresa', 'Nombre Curso', 'OC', 'Fecha', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'codigo',index:'codigo', width:55}, 
				   {name:'rut',index:'rut', width:59}, 
				   {name:'empresa',index:'empresa'}, 
				   {name:'curso',index:'curso', width:120}, 
				   {name:'orden',index:'orden', width:40}, 
				   {name:'FECHA__AUTORIZACION',index:'FECHA__AUTORIZACION', width:55}, 
				   { align:"right",editable:true, width:20}, 
				   { align:"right",editable:true, width:20} ], 
		rowNum:10, 
		autowidth: true, 
		rowList:[10,20,30], 
		pager: jQuery('#pager1'), 
		sortname: 'AUTORIZACION.ID_AUTORIZACION', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Autorizaciones" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("autorizacion/frmautorizacion.asp",
															   {id:0},
															   function(f){
																  $('#dialog').html(f);
																  $('#txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
																  llena_curriculo(0);
																  $("#solicitud").append("<option value=\"\">Seleccione</option>");
																  //llena_empresa(0);
																  llena_otic(0);
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
			height:350,
			width: 900,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmAutorizacion").valid())
						{
							//window.open($('#frmAutorizacion').attr('action')+'?'+$('#frmAutorizacion').serialize());
							$.post($('#frmAutorizacion').attr('action')+'?'+$('#frmAutorizacion').serialize(),function(d){
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
		
		$("#dialog_el").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
						$.post("autorizacion/eliminar.asp",{id:emp_del});
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
		//$('#frmAutorizacion txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		$("#frmAutorizacion").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			Curriculo:{
				required:true
			},
			Empresa:{
				required:true
			},
			txtFecha:{
				required:true
			},
			txtValor:{
				required:true
			},
			txtAutorizador:{
				required:true
			},
			txtValorAutorizador:{
				required:true
			},
			txtInscrito:{
				required:true
			}
		},
		messages:{
			Curriculo:{
				required:"&bull; Seleccione Curriculo."
			},
			Empresa:{
				required:"&bull; Seleccione Empresa."
			},
			txtFecha:{
				required:"&bull; Seleccione Fecha Autorización."
			},
			txtValor:{
				required:"&bull; Ingrese Valor."
			},
			txtAutorizador:{
				required:"&bull; Ingrese Autorizador."
			},
			txtValorAutorizador:{
				required:"&bull; Ingrese Valor Autorizador."
			},
			txtInscrito:{
				required:"&bull; Ingrese Inscritos."
			}
		}
	});
    }
	
	/*function llena_empresa(id_empresa){
		//window.open("autorizacion/empresa.asp");
		$("#Empresa").html("");
		$.get("autorizacion/empresa.asp",
					function(xml){
						$("#Empresa").append("<option value=\"\">Seleccione</option>");
						//$("#Empresa").append("<option value=170>Sin Empresa</option>");
						$('row',xml).each(function(i) { 
							if(id_empresa==$(this).find('ID').text())
								$("#Empresa").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																		$(this).find('razon').text()+ '</option>');
							else
								$("#Empresa").append('<option value="'+$(this).find('ID').text()+'" >'+
																		$(this).find('razon').text()+ '</option>');
						});
					});
	}*/
	
	function llena_curriculo(id_curriculo){
		$("#Curriculo").html("");
		$.get("autorizacion/curriculo.asp",
					function(xml){
						$("#Curriculo").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_curriculo==$(this).find('ID_MUTUAL').text())
								$("#Curriculo").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$("#Curriculo").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
					});
	}
	
	function llena_solicitud(id_empresa,id_prog,id_sol){
		$("#solicitud").html("");
		$.get("autorizacion/solicitud.asp",
			  		{id_empresa:id_empresa,id_prog:id_prog},
					function(xml){
						$("#solicitud").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_sol==$(this).find('ID').text())
								$("#solicitud").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																		$(this).find('CODIGO').text()+ '</option>');
							else
								$("#solicitud").append('<option value="'+$(this).find('ID').text()+'" >'+
																		$(this).find('CODIGO').text()+ '</option>');
						});
					});
	}
	
	function llena_otic(id_otic){
		//window.open("autorizacion/empresa.asp");
		$("#OTIC").html("");
		$.get("autorizacion/otic.asp",
					function(xml){
						$("#OTIC").append("<option value=\"\">Seleccione</option>");
						//$("#OTIC").append("<option value=5>Sin OTIC</option>");
						$('row',xml).each(function(i) { 
							if(id_otic==$(this).find('ID').text())
								$("#OTIC").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																		$(this).find('razon').text()+ '</option>');
							else
								$("#OTIC").append('<option value="'+$(this).find('ID').text()+'" >'+
																		$(this).find('razon').text()+ '</option>');
						});
					});
	}
	
	function val(e) {
    tecla = (document.all) ? e.keyCode : e.which; // 2
    if (tecla==8) {return true;} // 3
    patron =/[A-Za-z\s]/; // 4
    te = String.fromCharCode(tecla); // 5
    return patron.test(te); // 6
	}
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
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
	   llena_solicitud(id,$('#Curriculo').val(),0)
	   cargaDatos(id);
	   $('#txRut').val(rut);
	   $('#suggestions').hide();
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
	
	function cargaDatosInscritos(id_programa)	
	{
		$("#txtInscritos").html("");

		$.get("programacion/datosInscritos.asp",
						 {id_programa:id_programa},
						 function(xml){
								$('row',xml).each(function(i) { 
									$("#txtInscritos").html($(this).find('inscritos').text());
						});
			});
	}
	
	function update(i){
		$.post("autorizacion/frmautorizacion.asp",
			   {id:i},
			   function(f){
				  // validar();
				    $('#dialog').html(f);
				    $('#txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
					llena_curriculo($('#txtIdCurriculo').val());
					cargaDatosInscritos($('#txtIdCurriculo').val());
					$("#solicitud").append("<option value=\"\">Seleccione</option>");
					//llena_empresa($('#txtEmpresa').val());
					cargaDatos($('#Empresa').val());
					llena_solicitud($('#Empresa').val(),$('#txtIdCurriculo').val(),$('#txtSol').val());
					llena_otic($('#txtIdOtic').val());
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
		<li><a href="administracion.asp" accesskey="1" title="">Administraci&oacute;n</a></li>
		<li><a href="#" accesskey="2" title="" class="selItem">Operaci&oacute;n</a></li>
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
				<li class="first"><a href="operacionprogramacion.asp">Programación</a></li>
				<li><a href="operacionautorizacion.asp">Autorización</a></li>
				<li><a href="operacioninscripcion.asp">Inscripción</a></li>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Autorización</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Autorización">
</div>
<div id="dialog_el" title="Eliminar Autorización">
	<p>Esta seguro de eliminar la Autorización.</p>
</div>

</body>
</html>
