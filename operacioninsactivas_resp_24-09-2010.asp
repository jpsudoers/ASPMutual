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
<script src="js/jquery.tbltogrid.js" type="text/javascript" charset="utf-8"> </script> 
<script src="js/ajaxfileupload.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
var estadoFrmTrab=0;
var rutTrabAsig="";
var ins_del;
var insDelTrab=0;
var progSelCod=0;

$(document).ready(function(){					

	jQuery("#list1").jqGrid({ 
		url:'inscripcionactivas/listar.asp', 
		datatype: "xml", 
		colNames:['Cód. Prog.', 'Rut', 'Nombre Empresa', 'Cód. Curso', 'Nombre Curso', 'Relator', 'Sede/Sala','&nbsp;','&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'cod_prog',index:'cod_prog', width:45}, 
				   {name:'rut_empresa',index:'rut_empresa', width:50}, 
				   {name:'empresa',index:'empresa'}, 
				   {name:'cod_curso',index:'cod_curso', width:60}, 
				   {name:'nom_curso',index:'nom_curso', width:70}, 
				   {name:'relator',index:'relator', width:70}, 
				   {name:'sala',index:'sala', width:70}, 
				   {align:"right",editable:true, width:20},
				   {align:"right",editable:true, width:20},
				   {align:"right",editable:true, width:20}], 
		rowNum:30, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager1'), 
		sortname: 'AUTORIZACION.ID_PROGRAMA', 
		viewrecords: true, 
		sortorder: "DESC", 
		caption:"Listado de Inscripciones Autorizadas" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list1").trigger("reloadGrid");
													  } }); 

     $("#dialog_delIns").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 400,
			modal: true,
			buttons: {
				'Aceptar': function() {
						$.post("inscripcionactivas/eliminar.asp",{ins_del:ins_del},function(d){
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

    $("#InsBloque").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 320,
			width: 500,
			modal: true,
			buttons: {
				'Cerrar': function() {
					$(this).dialog('close');
				}
			}
		});

    $("#cursos").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 380,
			width:850,
			modal: true,
			buttons: {
				'Aceptar': function() {
						if($("#frmAsignar").valid())
						{
								$('#txtBloque').val($('#SelBloque').val());
								$('#txtSala').val($('#SelRelator').val());
								$('#txtRelator').val($('#SelSala').val());
								
								$('#txtRelProg').html($('#SelRelNom').val());
								$('#txtSalProg').html($('#SelSalaNom').val());
								$('#Programa').val($('#SelProg').val());
								
								$('#txtFechI').html($('#SelFechI').val());
								$('#txtFechT').html($('#SelFechT').val());
						
								$(this).dialog('close');
						}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			},
			title: 'Cambiar Relator / Sede'
		});

	$("#dialog_el").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 400,
			modal: true,
			buttons: {
				'Aceptar': function() {
							 $.post("inscripcionactivas/tab_eliminar.asp",
			       					 				{histId:insDelTrab,tabPart:$('#tabFechaPart').val()},
			                								function(d){
								 							    $("#list2").trigger("reloadGrid");							  
							  });
						      $(this).dialog('close');
				 			  numParticipantes(0);
							  $('#txtDocOculto').val("");
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
	});

	$("#dialogDelTrab").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 180,
			width: 450,
			modal: true,
			buttons: {
				'Aceptar': function() {
					$(this).dialog('close');
				}
			}
	});

	$("#dialog_trab").dialog({
			autoOpen: false,
			bgiframe: true,
			height:400,
			width: 500,
			modal: true,
			buttons: {
				'Agregar': function() {
						if($("#frmTrabajador").valid())
						{
                            $.post($('#frmTrabajador').attr('action')+'?'+$('#frmTrabajador').serialize(),function(d){
								 													       $("#list2").trigger("reloadGrid"); 
																						   });
							$("#list2").trigger("reloadGrid");
							$(this).dialog('close');
						}
				},
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
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
    
	$("#mensaje").dialog({
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
			height:580,
			width: 970,
			modal: true,
			buttons: {
				'Guardar': function() {
					if($("#frminscripcion").valid())
					{
						if($('#SubDocSelec').val()!="")
						{
							subeDoc();
						}
						else
						{
							//window.open($('#frminscripcion').attr('action')+'?'+$('#frminscripcion').serialize())
							
							$.post($('#frminscripcion').attr('action')+'?'+$('#frminscripcion').serialize(),function(d){
																							$("#list1").trigger("reloadGrid"); 
																								   });
							$("#txtMensaje").html("Inscripción Modificada Exitosamente");
							$("#mensaje").dialog('open');
							$(this).dialog('close');								
							$("#list1").trigger("reloadGrid"); 
						}
					}
				},
				'Cerrar': function() {
						$(this).dialog('close');
				}
			}
		});
	
	$("#Doc").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 500,
			width:900,
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
	});	
	
	function update(id){
		$.post("inscripcionactivas/frminscripcion.asp",
			   {id:id},
			   function(f){
				    $('#dialog').html(f);
					jQuery("#list2").jqGrid({ 
								scroll: 1, 					
url:'inscripcionactivas/listar2.asp?programa='+$('#Programa').val()+'&empresa='+$('#Empresa').val()+'&autorizacion='+$('#autorizacionId').val()+'&tabPart='+$('#tabFechaPart').val(),
								datatype: "xml", 
							    colNames:['Rut', 'Nombre, Apellido Paterno, Apellido Materno', 'Cargo En La Empresa', 'Escolaridad','&nbsp;','&nbsp;','&nbsp;'], 
								colModel:[
										{name:'rut',index:'rut', width:40}, 
										{name:'nombre',index:'nombre'}, 
										{name:'cargo',index:'cargo', width:80}, 
										{name:'escolaridad',index:'escolaridad', width:80},
										{align:"center",width:15},
				   						{align:"center",width:15}, 
				   						{align:"center",width:15}], 
								rowNum:100, 
								rownumbers: true, 
								rownumWidth: 40, 
								autowidth: true, 
								pager: jQuery('#pager2'), 
								sortname: 'PROGRAMA.ID_PROGRAMA', 
								viewrecords: true, 
								sortorder: "asc", 
								caption:"Nomina de Participantes"
								}); 
								llenaTipoCom($('#txtTipoDoc').val());
								valFrmIns();
		});
		 $('#dialog').dialog('open');
	}
	
	function llenaTipoCom(id_tipo){
		$("#Compromiso").html("");
		$("#Compromiso").append("<option value=\"\">Seleccione</option>");
		if(id_tipo==0){
			$("#Compromiso").append("<option value=\"4\">Carta Compromiso</option>");
			$("#Compromiso").append("<option value=\"0\" selected=\"selected\">Orden de Compra</option>");
			$("#Compromiso").append("<option value=\"1\">Vale Vista</option>");
			$("#Compromiso").append("<option value=\"2\">Dep&oacute;sito Cheque</option>");
			$("#Compromiso").append("<option value=\"3\">Transferencia</option>");
		}else if(id_tipo==1){
			$("#Compromiso").append("<option value=\"4\">Carta Compromiso</option>");
			$("#Compromiso").append("<option value=\"0\">Orden de Compra</option>");
			$("#Compromiso").append("<option value=\"1\" selected=\"selected\">Vale Vista</option>");
			$("#Compromiso").append("<option value=\"2\">Dep&oacute;sito Cheque</option>");
			$("#Compromiso").append("<option value=\"3\">Transferencia</option>");
		}else if(id_tipo==2){
			$("#Compromiso").append("<option value=\"4\" selected=\"selected\">Carta Compromiso</option>");
			$("#Compromiso").append("<option value=\"0\">Orden de Compra</option>");
			$("#Compromiso").append("<option value=\"1\">Vale Vista</option>");
			$("#Compromiso").append("<option value=\"2\" selected=\"selected\">Dep&oacute;sito Cheque</option>");
			$("#Compromiso").append("<option value=\"3\">Transferencia</option>");
		}else if(id_tipo==3){
			$("#Compromiso").append("<option value=\"4\" selected=\"selected\">Carta Compromiso</option>");
			$("#Compromiso").append("<option value=\"0\">Orden de Compra</option>");
			$("#Compromiso").append("<option value=\"1\">Vale Vista</option>");
			$("#Compromiso").append("<option value=\"2\">Dep&oacute;sito Cheque</option>");
			$("#Compromiso").append("<option value=\"3\" selected=\"selected\">Transferencia</option>");			
		}else if(id_tipo==4){
			$("#Compromiso").append("<option value=\"4\" selected=\"selected\">Carta Compromiso</option>");
			$("#Compromiso").append("<option value=\"0\">Orden de Compra</option>");
			$("#Compromiso").append("<option value=\"1\">Vale Vista</option>");
			$("#Compromiso").append("<option value=\"2\">Dep&oacute;sito Cheque</option>");
			$("#Compromiso").append("<option value=\"3\">Transferencia</option>");		
		}else{
			$("#Compromiso").append("<option value=\"4\">Carta Compromiso</option>");
			$("#Compromiso").append("<option value=\"0\">Orden de Compra</option>");
			$("#Compromiso").append("<option value=\"1\">Vale Vista</option>");
			$("#Compromiso").append("<option value=\"2\">Dep&oacute;sito Cheque</option>");
			$("#Compromiso").append("<option value=\"3\">Transferencia</option>");		
		}
	}
	
	
	function documento(arch){
		$("#ifPagina").attr('src',"http://norte.otecmutual.cl/ordenes/"+arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
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

	function llena_escolaridad(id_tipo){
		$("#escolaridadTrab").html("");
		$("#escolaridadTrab").append("<option value=\"\">Seleccione</option>");
		if(id_tipo==0){
			$("#escolaridadTrab").append("<option value=\"0\" selected=\"selected\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");
		}else if(id_tipo==1){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\" selected=\"selected\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");
		}else if(id_tipo==2){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\" selected=\"selected\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");
		}else if(id_tipo==3){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\" selected=\"selected\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");			
		}else if(id_tipo==4){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\" selected=\"selected\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");			
		}else if(id_tipo==5){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\" selected=\"selected\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");		
		}else if(id_tipo==6){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
		 $("#escolaridadTrab").append("<option value=\"6\" selected=\"selected\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");			
		}else if(id_tipo==7){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\" selected=\"selected\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");
		}else if(id_tipo==8){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\" selected=\"selected\">Universitaria Completa</option>");
		}else{
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">Básica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">Básica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior Técnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior Técnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");			
		}
	}

	function validarFrmTrab(){
		$("#frmTrabajador").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txtRutTrab:{
				required:true,
				rut:true
			},
			txtPasTrab:{
				required:true
			},
			txtNomTrab:{
				required:true
			},
			txtAPaterTrab:{
				required:true
			},
			txtAMaterTrab:{
				required:true
			},
			txtCargoTrab:{
                required:true                        
            },
			insGrilla:{
                required:true                        
            }
		},
		messages:{
			txtRutTrab:{
				required:"&bull; Ingrese Rut del Trabajador.",
				rut:"&bull; Rut No Valido."
			},
			txtPasTrab:{
				required:"&bull; Ingrese el Número de Pasaporte del Trabajador."
			},
			txtNomTrab:{
				required:"&bull; Ingrese Nombre del Trabajador."
			},
			txtAPaterTrab:{
				required:"&bull; Ingrese Número Apellido Paterno del Trabajador."
			},
			txtAMaterTrab:{
				required:"&bull; Ingrese Número Apellido Materno del Trabajador."
			},
			txtCargoTrab: {
                required:"&bull; Ingrese Cargo del Trabajador."                        
            },
			insGrilla:{
                required:"&bull; El Trabajador ya se Encuentra Inscrito."              
            }
		}
	});
	}
	
	function datosTrabajador(rut)	
	{
		$("#txtNomTrab").val("");
		$("#txtAPaterTrab").val("");
		$("#txtAMaterTrab").val("");
		$("#txtCargoTrab").val("");
		$("#txtTrabID").val("0");
		$("#txtIdEscolaridad").val("");
		llena_escolaridad($("#txtIdEscolaridad").val());
		
		$.get("inscripcionactivas/datosTrabajador.asp",
				{rut:rut},
						 function(xml){
									$('row',xml).each(function(i) { 
									$("#txtNomTrab").val($(this).find('NOM_TRAB').text());
									$("#txtAPaterTrab").val($(this).find('APATERTRAB').text());
									$("#txtAMaterTrab").val($(this).find('AMATERTRAB').text());
									$("#txtCargoTrab").val($(this).find('CARGO_EMPRESA').text());
									$("#txtTrabID").val($(this).find('ID_TRABAJADOR').text());
									$("#txtIdEscolaridad").val($(this).find('ESCOLARIDAD').text());
									llena_escolaridad($("#txtIdEscolaridad").val());
      inscritosHistTrab($("#txtTrabProg").val(),$(this).find('ID_TRABAJADOR').text(),$("#autoIdTxt").val(),$("#TrabFecha").val());
							});
			});
	}
	
	function inscritosHistTrab(id,trab,idauto,trabfech)	
	{
		$.get("inscripcionactivas/datosTrabajadorHist.asp",
				{id:id,trab:trab,idauto:idauto,trabfech:trabfech},
						 function(xml){
									$('row',xml).each(function(i) { 
															   if($(this).find('THisTrab').text()=="0")
															   {
																	$('#insGrilla').val("1");
															   }
															   else
															   {
																	$('#insGrilla').val("");  
															   }
							});
			});
	}

	function formatCurrency(num) {
		num = num.toString().replace(/$|,/g,'');
		if(isNaN(num))
		{
		num = "0";
		}
		sign = (num == (num = Math.abs(num)));
		num = Math.floor(num*100+0.50000000001);
		cents = num%100;
		num = Math.floor(num/100).toString();
		if(cents<10)
		{
		cents = "0" + cents;
		}
		for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
		{
		num = num.substring(0,num.length-(4*i+3))+'.'+num.substring(num.length-(4*i+3));
		}
	  return (((sign)?'':'-') + '$ ' + num);
	}

	function numParticipantes(estado)	
	{
							if(estado==1)
							{
		   						  $('#NParticipantes').val(parseInt($('#NParticipantes').val())+1);
							}
							else
							{
									$('#NParticipantes').val(parseInt($('#NParticipantes').val())-1);	
							}
			                      
								  
								  $('#txtNParticipantes').html($('#NParticipantes').val());
									
								  $("#ValorTotal").val($('#NParticipantes').val()*$('#Valor').val());
								  $("#txtValorTotal").html(formatCurrency($("#ValorTotal").val()));
	}	

	function cambiarCampo(){
		if($("#extranjero").attr("checked")) 
		{
			OcultarFilaRut('rutTrab');
			MostrarFilaPas('pasTrab');
			$('#checkExtran').val("1");
		}
		else
		{
			MostrarFilaRut('rutTrab');
			OcultarFilaPas('pasTrab');
			$('#checkExtran').val("0");
		}
	}
	
	function OcultarFilaRut(Fila) {
   			var elementos = document.getElementsByName(Fila);
   			for (k = 0; k < elementos.length; k++)
      		elementos[k].style.display = "none";
			$('#txtRutTrab').val("");
			if(rutTrabAsig=="")
			{
				generarRut();
			}
			else
			{
				$('#txtRutTrab').val(rutTrabAsig);
			}
			datosTrabajador("1");
	}
	
	function MostrarFilaRut(Fila) {
			$('#txtRutTrab').val("");
  			var elementos = document.getElementsByName(Fila);
   			for (i = 0; i < elementos.length; i++)
      		elementos[i].style.display ="";
			$('#txtRutTrab').val("");
	}
	
	function generarRut()
	{
		$.get("inscripcionactivas/numRut.asp",
						          function(xml){
								   $('row',xml).each(function(i) { 
									rutTrabAsig=$(this).find('num').text()+'-'+getDigito($(this).find('num').text()+"");
									$('#txtRutTrab').val(rutTrabAsig);
							      });
	    });
	}

	function getDigito(rut)
    {
          var dvr = '0';
          suma = 0;
          mul  = 2;
          for(i=rut.length -1;i >= 0;i--) 
          { 
            suma = suma + rut.charAt(i) * mul;    
            if (mul == 7)
            {
              mul = 2;
            }   
            else
            {         
              mul++;
            } 
          }
          res = suma % 11;  
          if (res==1)
          {
            return 'k';
          } 
          else if(res==0)
          {   
            return '0';
          } 
          else  
          {   
            return 11-res;
          }
    }
	
	function OcultarFilaPas(Fila) {
   			var elementos = document.getElementsByName(Fila);
   			for (k = 0; k < elementos.length; k++)
      		elementos[k].style.display = "none";
			$('#txtPasTrab').val("1");
			
	}
	
	function MostrarFilaPas(Fila) {
  			var elementos = document.getElementsByName(Fila);
   			for (i = 0; i < elementos.length; i++)
      		elementos[i].style.display ="";
			$('#txtPasTrab').val("");
	}
	
	function reempTrab(histId)
	{
		$.post("inscripcionactivas/frmTrabajador.asp",
	 {histId:histId,trabId:0,histProgId:$('#Programa').val(),tabPart:$('#tabFechaPart').val(),autoId:$('#autorizacionId').val()},
										   function(f){
												$('#dialog_trab').html(f);
												estadoFrmTrab=0;
												rutTrabAsig="";
												llena_escolaridad($("#txtIdEscolaridad").val());
												validarFrmTrab();
												});
											$('#dialog_trab').dialog('open');
	}
		
	function upTrabajador(idTrab)
	{
		 $.post("inscripcionactivas/frmTrabajador.asp",
					{trabId:idTrab},
						function(f){
							$('#dialog_trab').html(f);
								llena_escolaridad($("#txtIdEscolaridad").val());
								estadoFrmTrab=1;
								rutTrabAsig="";
								validarFrmTrab();
							});
		$('#dialog_trab').dialog('open');
	}
	
	function delTrab(histId)
	{
			if(parseInt($("#NParticipantes").val())>1)
			{
				insDelTrab=histId;
				$('#dialog_el').dialog('open');
			}
			else
			{
				$("#dialogDelTrab").dialog('open');
			}
	}
	
	function delInscripcion(IdInscripcion)
	{
		    ins_del=IdInscripcion;
			$('#dialog_delIns').dialog('open');
	}

	function DocCambiado(arch){
		$('#txtDocOculto').val(arch);
	}
	
	function valFrmIns(){
		$("#frminscripcion").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			Compromiso:{
				required:true
			},
			txtNOrden:{
				required:true
			},
			txtDocOculto: {
                required: true,
                accept: "xls|doc|pdf|jpg|tif|png|gif"                          
            }
		},
		messages:{
			Compromiso:{
				required:"&bull; Seleccione el Tipo Compromiso."
			},
			txtNOrden:{
				required:"&bull; Ingrese Número de Documento."
			},
			txtDocOculto: {
                required:"&bull; Seleccione Documento de Compromiso",
                accept:"&bull; El Formato del Documento de Compromiso no es Válido."                          
            }
		}
	});
    }
	
	function subeDoc()
    {
                               $("#loading")
                               .ajaxStart(function(){
                                               $(this).show();
                               })
                               .ajaxComplete(function(){
                                               $(this).hide();
											   $("#dialog").dialog('close');
 											   $("#txtMensaje").html("Inscripción Modificada Exitosamente");
											   $("#mensaje").dialog('open');
	                           });
							   
                               $.ajaxFileUpload
                               (
                                      {
                                           url:'inscripcionactivas/subir.asp?',
                                           secureuri:false,
                                           fileElementId:'SubDocSelec',
elements:'txtNOrden;Compromiso;ValorTotal;NParticipantes;txtBloque;txtSala;txtRelator;Empresa;Programa;autorizacionId;ConOtic;tabFechaPart',
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
	
	function cambiaRelSede(mostrar)
    {
		if(parseInt(mostrar)==0)
		{
			progSelCod=$('#Programa').val();
		}
		else
		{
			progSelCod="";
		}
		
	 $.post("inscripcionactivas/frmAsignar.asp",
							   {IdProg:progSelCod,IdAuto:$('#autorizacionId').val()},
							   function(f){
									$('#cursos').html(f);
									tableToGrid("#mytable"); 
									valSelBloq();
								});
							
							$('#cursos').dialog('open');
							$(this).dialog('close');
	}
	
	function valSelBloq(){
		$("#frmAsignar").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			SelBloque:{
				required:true
			}
		},
		messages:{
			SelBloque:{
				required:"&bull; Seleccione Bloque"
			}
		}
	});
    }
	
	function checkear(seleccionado,bloque,relator,sala,nomrel,nomsal,programa,fechI,fechT){
			for (i=1; i<=$('#countFilas').val(); i++)
			{
				$('#check'+i).attr('checked', false);	
			}
			
			$('#check'+seleccionado).attr('checked', true);
			$('#SelBloque').val(bloque);
			$('#SelRelator').val(relator);
			$('#SelSala').val(sala);
			$('#SelRelNom').val(nomrel);
			$('#SelSalaNom').val(nomsal);
			$('#SelProg').val(programa);
			$('#SelFechI').val(fechI);
			$('#SelFechT').val(fechT);
   	}
	
	function MostrarInscritos(bloque){
		$.post("inscripcionactivas/frmMostrar.asp",
			   function(f){
				    $('#InsBloque').html(f);
					jQuery("#listMostrar").jqGrid({ 
								scroll: 1, 					
								url:'inspendientes/listarBloque.asp?bloId='+bloque, 
								datatype: "xml", 
							    colNames:['Rut', 'Nombre'], 
								colModel:[
										{name:'rut',index:'rut', width:40}, 
										{name:'nombre',index:'nombre'}], 
								rowNum:100, 
								rownumbers: true, 
								rownumWidth: 40, 
								autowidth: true, 
								pager: jQuery('#pagerMostrar'), 
								sortname: 'bloque_programacion.ID_BLOQUE', 
								viewrecords: true, 
								sortorder: "asc", 
								caption:"Nomina de Participantes"
								}); 
		});
		 $('#InsBloque').dialog('open');
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
		<li><a href="operacion.asp" accesskey="2" title="" class="selItem">Operaci&oacute;n</a></li>
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
				<li class="first"><a href="operacionsolempresa.asp">Solicitud de Nueva Empresa</a></li>
                <li><a href="operacionprogramacion.asp">Programación de Cursos</a></li>
                <li><a href="operacioncalendario.asp">Inscripción de Cursos</a></li>
				<li><a href="operacioninspendientes.asp">Autorizar Inscripciones</a></li>
				<li><a href="operacioninsactivas.asp">Inscripciones Autorizadas</a></li>
                <li><a href="operacioncierre.asp">Revisión y Cierre</a></li>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Inscripciones Autorizadas</em></h2>
          <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Inscripciones Autorizadas">
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="dialog_trab" title="Registro de Trabajador">
</div>
<div id="dialog_el" title="Eliminar Trabajador">
	<p>Esta seguro de eliminar al trabajador.</p>
</div>
<div id="dialog_delIns" title="Eliminar Inscripción Autorizada">
	<p>Esta seguro de eliminar la inscripción.</p>
</div>
<div id="dialogDelTrab" title="Eliminar Participante">
	<p>No se puede eliminar al participante, ya que la inscripción al menos debe tener a un participante.</p>
</div>
<div id="mensaje" title="Registro de Inscripción">
     <label id="txtMensaje" name="txtMensaje"></label>
</div>
<div id="cursos" title="Cambiar Relator / Sede">
</div>
<div id="InsBloque" title="Inscritos en Bloque">
</div>
</body>
</html>