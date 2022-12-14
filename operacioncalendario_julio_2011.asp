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
<style type="text/css">
    .suggestionsBox {
        position: relative;
        left: 10px;
        margin: 10px 0px 0px 0px;
        width: 500px;
        background-color: #212427;
        -moz-border-radius: 7px;
        -webkit-border-radius: 7px;
        border: 2px solid #000;   
        color: #fff;
		height:150px;
		overflow:auto;
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
	
	.suggestionsBox2 {
        position: relative;
        left: 10px;
        margin: 10px 0px 0px 0px;
        width: 500px;
        background-color: #212427;
        -moz-border-radius: 7px;
        -webkit-border-radius: 7px;
        border: 2px solid #000;   
        color: #fff;
		height:200px;
		overflow:auto;
    }
   
    .suggestionList2 {
        margin: 0px;
        padding: 0px;
    }
   
    .suggestionList2 li {
       
        margin: 0px 0px 3px 0px;
        padding: 3px;
        cursor: pointer;
    }
   
    .suggestionList2 li:hover {
        background-color: #659CD8;
    }
</style>
<script type="text/javascript">
var amtval="";
var fran="";
var i=1;
var costo=0;
var escolaridades=0;
var rutRepetido=0;
var estadoFrmTrab=0;
var rutTrabAsig="";
$(document).ready(function(){					

	jQuery("#list1").jqGrid({ 
	    url:'emp_programacion/listarOperacion.asp', 
		datatype: "xml", 
		colNames:['C??d. Prog.','C??digo Curso', 'Sence', 'Nombre Curso', 'Vac.', 'Inicio', '&nbsp;'], 
		colModel:[
				   {name:'ID_PROGRAMA',index:'ID_PROGRAMA', width:25, align:'center'}, 
				   {name:'CODIGO',index:'CODIGO', width:42}, 
				   {name:'SENCE',index:'SENCE', width:19, align:'center'}, 
				   {name:'CURSO',index:'CURSO'}, 
				   {name:'CUPOS',index:'CUPOS', width:15, align:'center'}, 
				   {name:'FECHA',index:'FECHA', width:33, align:'center'}, 
				   {align:"right",editable:true, width:8}], 
		rowNum:30, 
		autowidth: true, 
		rowList:[30,50,100],  
		pager: jQuery('#pager1'), 
		sortname: 'PROGRAMA.FECHA_INICIO_', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Cursos Programados" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list1").trigger("reloadGrid");
													  } }); 
	
	function subeDoc()
    {
                               $("#loading")
                               .ajaxStart(function(){
                                               $(this).show();
                               })
                               .ajaxComplete(function(){
                                               $(this).hide();
											   $("#dialog").dialog('close');
					                           $("#txtMensaje").html("La Inscripci??n a Curso ha sido Enviada Exitosamente.");
											   $("#mensaje").dialog('open');
                               });
							   
                               $.ajaxFileUpload
                               (
                                      {
                                           url:'emp_programacion/subir.asp?',
                                           secureuri:false,
                                           fileElementId:'txtDoc',
                                           elements:
'Empresa;programa;EmpMan;Compromiso;txtNum;tabProgId;NParticipantes;Valor;ValorTotal;TipoEmpresa;Costo',										
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
	
	$("#dialog").dialog({
			autoOpen: false,
			bgiframe: true,
			height:600,
			width: 1000,
			modal: true,
			overlay: {
				backgroundColor: '#f00',
				opacity: 0.5
			},
			buttons: {
				'Guardar': function() {
					if($("#terminos").attr("checked")) 
					{
						if($("#frmProgEmp").valid())
						{
							//$(this).dialog('close');
							//$("#mensaje").dialog('open');
							subeDoc();
						}
					}
				},
				Cancelar: function() {
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
							//window.open($('#frmTrabajador').attr('action')+'?'+$('#frmTrabajador').serialize())
							$.post($('#frmTrabajador').attr('action')+'?'+$('#frmTrabajador').serialize(),function(d){
																					if(estadoFrmTrab==0)
																					{
																						   numParticipantes(1);	
																					}
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
							$("#txtmContrasena").html("Contrase??a Modificada Exitosamente.");
							$("#mContrasena").dialog('open');
						}
						else
						{
							$("#txtmContrasena").html("La Contrase??a Antigua Ingresada es Incorrecta.");
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
	
	$("#datosTerminos").dialog({
			autoOpen: false,
			bgiframe: true,
			height:450,
			width: 1000,
			modal: true,
			buttons: {
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	


	function validar(){
		$("#frmProgEmp").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txRut:{
				required:true
			},
			Compromiso:{
				required:true
			},
			txtNum:{
				required:true
			},
			txtDoc: {
                required: true,
                accept: "xls|doc|pdf|jpg|tif|png|gif"                          
            },
			NParticipantes:{
				max: 30,
                min: 1
			}
		},
		messages:{
			txRut:{
				required:"&bull; Ingrese Rut Empresa."
			},
			Compromiso:{
				required:"&bull; Seleccione el Tipo Compromiso."
			},
			txtNum:{
				required:"&bull; Ingrese N??mero de Documento."
			},
			txtDoc: {
                required:"&bull; Seleccione Documento de Compromiso",
                accept:"&bull; El Formato del Documento de Compromiso no es V??lido."                          
            },
			NParticipantes:{
				max: "&bull; El N??mero de Participantes debe ser menor o Igual a 30",
                min: "&bull; No hay Participantes Ingresados"
			}
		}
	});
    }


	 function lookup(inputString) {
		if(inputString.length <5) {
			$('#suggestions').hide();
			$("#txtRsoc").html("");
			$("#Empresa").html("");
			$("#txtRutOtic").html("");
			$("#txtRsocOtic").html("");
			$("#EmpMan").html("");
			OcultarFilas("EmpOtic");
		} else {
				$.post("emp_programacion/sugEmpresaSinOtic.asp", {txt: inputString}, function(data){
						if(data.length >0) {
								$('#suggestions').show();
								$('#autoSuggestionsList').html(data);
						}
				});
		}
	}

    function cargaDatosEmpresa(id)	
	{
		$("#txRut").val("");
		$("#txtRsoc").html("");
		$("#Empresa").html("");
		$("#txtRutOtic").html("");
		$("#txtRsocOtic").html("");
		$("#EmpMan").html("");
		
 
		$.get("emp_programacion/datosempresaoperacion.asp",
						 {id:id},
						 function(xml){
									$('row',xml).each(function(i) { 
									$("#txtRsoc").html($(this).find('RSOCIAL').text());
									//$("#txRut").val($(this).find('RUT').text());
									//$("#Empresa").val($(this).find('IDEMPRESA').text());
									
									if($(this).find('RUTOTIC').text()!=0)
									{
										MostrarFilas("EmpOtic");
									}
									else
									{
										OcultarFilas("EmpOtic");
									}

									$("#txtRutOtic").html($(this).find('RUTOTIC').text());
									$("#txtRsocOtic").html($(this).find('RSOCIALOTIC').text());
									$("#EmpMan").val($(this).find('IDOTIC').text());
									$("#TipoEmpresa").val($(this).find('TIPO').text());
						});
			});
	}

	function fill(id,rut) {
	   $('#Empresa').val(id);
	   cargaDatosEmpresa(id);
	   $('#txRut').val(rut);
	   $('#suggestions').hide();
	}

    var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
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
	
		function grilla(id)
	{
		                  jQuery("#list2").jqGrid({ 
						  url:'emp_programacion/listar2.asp?id='+id, 
						  datatype: "xml", 
		colNames:['Rut', 'Nombre, Apellido Paterno, Apellido Materno', 'Cargo En La Empresa', 'Escolaridad', '&nbsp;','&nbsp;'], 
						  colModel:[
									{name:'rut',index:'rut', width:40}, 
									{name:'nombre',index:'nombre'}, 
									{name:'cargo',index:'cargo', width:80}, 
									{name:'escolaridad',index:'escolaridad', width:100},
									{align:"right",editable:true, width:20}, 
				   					{align:"right",editable:true, width:20}], 
						  rowNum:30, 
						  rownumbers: true, 
						  rownumWidth: 40, 
						  autowidth: true, 
						  rowList:[10,20,30],  
						  pager: jQuery('#pager2'), 
						  viewrecords: true, 
						  sortorder: "asc",
						  caption:"Nomina de Inscritos" 
		}); 
						  
		jQuery("#list2").jqGrid('navGrid','#pager2',{edit:false,add:false,del:false,search:false,refresh:false});
		jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
															title:"Agregar nuevo trabajador", 
															buttonicon :'ui-icon-plus', 
															onClickButton:function(){
																	$.post("emp_programacion/frmTrabajador.asp",
																		   {preinsId:id,trabId:0},
																		   function(f){
																			   $('#dialog_trab').html(f);
																			   estadoFrmTrab=0;
																			   rutTrabAsig="";
																			   llena_escolaridad($("#txtIdEscolaridad").val());
																			   validarFrmTrab();
																	});
																$('#dialog_trab').dialog('open');
															}}); 
		
	    jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list2").trigger("reloadGrid");
													  } }); 			  
	}
	
	function llena_escolaridad(id_tipo){
		$("#escolaridadTrab").html("");
		$("#escolaridadTrab").append("<option value=\"\">Seleccione</option>");
		if(id_tipo==0){
			$("#escolaridadTrab").append("<option value=\"0\" selected=\"selected\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");
		}else if(id_tipo==1){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\" selected=\"selected\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");
		}else if(id_tipo==2){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\" selected=\"selected\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");
		}else if(id_tipo==3){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\" selected=\"selected\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");			
		}else if(id_tipo==4){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\" selected=\"selected\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");			
		}else if(id_tipo==5){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\" selected=\"selected\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");		
		}else if(id_tipo==6){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
		 $("#escolaridadTrab").append("<option value=\"6\" selected=\"selected\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");			
		}else if(id_tipo==7){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\" selected=\"selected\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");
		}else if(id_tipo==8){
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\" selected=\"selected\">Universitaria Completa</option>");
		}else{
			$("#escolaridadTrab").append("<option value=\"0\">Sin Escolaridad</option>");
			$("#escolaridadTrab").append("<option value=\"1\">B??sica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"2\">B??sica Completa</option>");
			$("#escolaridadTrab").append("<option value=\"3\">Media Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"4\">Media Completa</option>");
			$("#escolaridadTrab").append("<option value=\"5\">Superior T??cnica Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"6\">Superior T??cnica Profesional Completa</option>");
			$("#escolaridadTrab").append("<option value=\"7\">Universitaria Incompleta</option>");
			$("#escolaridadTrab").append("<option value=\"8\">Universitaria Completa</option>");			
		}
	}
	
	function validarFrmTrab(){
		$("#frmTrabajador").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txtRutTrab:{
				required:true,
				rut:true
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
			txtNomTrab:{
				required:"&bull; Ingrese Nombre del Trabajador."
			},
			txtAPaterTrab:{
				required:"&bull; Ingrese N??mero Apellido Paterno del Trabajador."
			},
			txtAMaterTrab:{
				required:"&bull; Ingrese N??mero Apellido Materno del Trabajador."
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
	
	function eliminar_trab(idTrab,idPreins)
	{
		$.post("emp_programacion/eliminar_tab.asp",
			        {idTrab:idTrab,idPreins:idPreins},
			                function(d){
								  $("#list2").trigger("reloadGrid");							  
							});
		numParticipantes(0);
	}
	
	function update_trab(idTrab,idPreins)
	{
		 $.post("emp_programacion/frmTrabajador.asp",
					{preinsId:idPreins,trabId:idTrab},
						function(f){
							$('#dialog_trab').html(f);
								llena_escolaridad($("#txtIdEscolaridad").val());
								estadoFrmTrab=1;
								rutTrabAsig="";
								validarFrmTrab();
							});
		$('#dialog_trab').dialog('open');
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
		
		$.get("emp_programacion/datosTrabajador.asp",
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
									inscritosPreInsTrab($("#txtTrabPreins").val(),$(this).find('ID_TRABAJADOR').text());
							});
			});
	}
	
	function inscritosPreInsTrab(id,trab)	
	{
		$.get("emp_programacion/datosTrabajadorPreins.asp",
				{id:id,trab:trab},
						 function(xml){
									$('row',xml).each(function(i) { 
															   if($(this).find('TPreinsTrab').text()=="0")
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
									
								  $('#Costo').val($("#ValorTotal").val());
	}	
	
	/*function grilla(id)
	{
		i=1;
		amtval="";
		fran="";
		costo=0;
		escolaridades=0;
		rutRepetido=0;
		
														jQuery("#list2").jqGrid({ 
														scroll:1, 					
														url:'emp_programacion/listar2.asp?id='+id, 
														datatype: "xml", 
														colNames:['Rut', 'Nombre, Apellido Paterno, Apellido Materno', 'Cargo En La Empresa', 'Escolaridad', 'Franquicia'], 
														colModel:[
																   {name:'rut',index:'rut', width:40, editable:true}, 
																   {name:'nombre',index:'nombre', editable:true}, 
																   {name:'cargo',index:'cargo', width:80, editable:true}, 
															{name:'escolaridad',index:'escolaridad', width:100, editable:true,edittype:"select",
																   editoptions:{value:"0:Sin Escolaridad;1:B??sica Incompleta;2:B??sica Completa;3:Media Incompleta;4:Media Completa;5:Superior T??cnica Incompleta;6:Superior T??cnica Profesional Completa;7:Universitaria Incompleta;8:Universitaria Completa"}},
																   {name:'franquicia',index:'franquicia', width:35, editable:true,editoptions:{size:"20",maxlength:"3",onKeyPress:"return acceptNum(event)"}}], 
														rowNum:100, 
														rownumbers: true, 
														rownumWidth: 40, 
														autowidth: true, 
														pager: jQuery('#pager2'), 
														viewrecords: true, 
														caption:"Nomina de Inscritos" ,
														forceFit : true, 
														cellEdit: true, 
														cellsubmit: 'clientArray',
														afterSaveCell : function(rowid,name,val,iRow,iCol) { 
														
														if(name=='rut') 
														{ 
														    rutRepetido=0;
														    for (fila=1; fila<i; fila++)
															{
												               if(jQuery("#list2").jqGrid('getCell',fila,1)==val)
															   {
																  rutRepetido=1;
					    $("#txtMensaje").html("El Rut "+jQuery("#list2").jqGrid('getCell',rowid,1)+" ya se encuentra en la Nomina de Asistentes.");
																$("#mensaje").dialog('open');
															   }
															}
															
															if(rutRepetido===0)
															{
																if(Valida_Rut(jQuery("#list2").jqGrid('getCell',rowid,1))!==true)
																{
													 $("#txtMensaje").html("El Rut "+jQuery("#list2").jqGrid('getCell',rowid,1)+" No es V??lido.");
																$("#mensaje").dialog('open');
																}
															}
														}
														
														if(name=='escolaridad') 
														{ 
															escolaridades=val;
														}
														
														if(name=='franquicia') 
														{ 

														if(mycheck(jQuery("#list2").jqGrid('getCell',rowid,5))===false)
														{
															$("#txtMensaje").html("La Franquicia debe ser Mayor de 0 y Menor o Igual 100.");
															$("#mensaje").dialog('open');
														}
														else
														{
														
														 if(Valida_Rut(jQuery("#list2").jqGrid('getCell',rowid,1))===true && rutRepetido===0)
														 {
															    if(jQuery("#list2").jqGrid('getCell',rowid,1)!="" && 
																jQuery("#list2").jqGrid('getCell',rowid,2)!="" && 
																jQuery("#list2").jqGrid('getCell',rowid,3)!="" && 
																jQuery("#list2").jqGrid('getCell',rowid,4)!="" && 
																jQuery("#list2").jqGrid('getCell',rowid,5)!="")
																{
																	amtval += "'" + jQuery("#list2").jqGrid('getCell',rowid,1) + 
																			"','" + jQuery("#list2").jqGrid('getCell',rowid,2) + 
																			"','" + jQuery("#list2").jqGrid('getCell',rowid,3) + 
																			"'," + escolaridades + '/'; 

																	fran += "'" + jQuery("#list2").jqGrid('getCell',rowid,5) + "'/";
																	costo=0;
																	
																	for (fila=1; fila<=i; fila++)
																	{
												                  costo=costo+($('#Valor').val()*(jQuery("#list2").jqGrid('getCell',fila,5)/100));
																			$('#txtCosto').html(formatCurrency(costo));
																			$('#Costo').val(costo);
																			jQuery("#list2").jqGrid("setSelection",i,true);
																	}
																	
																	    if(jQuery("#list2").jqGrid('getCell',i,1)!="" && 
																		jQuery("#list2").jqGrid('getCell',i,2)!="" && 
																		jQuery("#list2").jqGrid('getCell',i,3)!="" && 
																		jQuery("#list2").jqGrid('getCell',i,4)!="" && 
																		jQuery("#list2").jqGrid('getCell',i,5)!="")
																		{
																		$("#datostabla").val(amtval);
																	    $("#datosfran").val(fran);
																	
																	    $('#txtNParticipantes').html(i);
																	    $('#NParticipantes').val(i);
																	  
																	    $('#txtValorTotal').html(formatCurrency(i*$('#Valor').val()));
																	    $('#ValorTotal').val(i*$('#Valor').val());	
																		
																		if(i<$('#Vacantes').val())
																		{
																			i+=1;
																			jQuery("#list2").jqGrid('addRowData',i, mydata1[0]); 
																			jQuery("#list2").jqGrid("setSelection",i,true);
																		}
																		}
																	escolaridades=0;
																}
																else
																{
																	$("#txtMensaje").html("Por Favor, Ingrese todos los Campos Requeridos.");
																	$("#mensaje").dialog('open');
																}
																
																}
																else
																{
																	if(rutRepetido===0)
															        {
											         $("#txtMensaje").html("El Rut "+jQuery("#list2").jqGrid('getCell',rowid,1)+" No es V??lido.");
																    $("#mensaje").dialog('open');
															        }
																	else
																	{
					 $("#txtMensaje").html("El Rut "+jQuery("#list2").jqGrid('getCell',rowid,1)+" ya se encuentra en la Nomina de Asistentes.");
																$("#mensaje").dialog('open');
																	}
																}
														}
															}}
														}); 
																  
	}

	function update(i,v){
		$.post("emp_programacion/frmprogrope.asp",
			   {id:i,vacantes:v},
			   function(f){
				    $('#dialog').html(f);
					grilla(0);
				    validar();
		});
		 $('#dialog').dialog('open');
	}*/
	
	function update(i,v){
		//alert(i+' '+v)
		$.post("emp_programacion/frmprogrope.asp",
			   {id:i,vacantes:v},
			   function(f){
				    $('#dialog').html(f);
					grilla($('#tabProgId').val());
				    validar();
		});
		 $('#dialog').dialog('open');
	}
	
	function abreTerminos(){
		$.post("emp_programacion/terminoscurso.asp",
			   function(f){
				    $('#datosTerminos').html(f);
		});
		 $('#datosTerminos').dialog('open');
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
				required:"&bull; Ingrese la Contrase??a Antigua."
			},
			NuevaCont:{
				required:"&bull; Ingrese la Nueva Contrase??a.",
				rangelength:"&bull; La Nueva Contrase??a debe Contener al Menos 5 Caracteres."
			},
			RepConNueva:{
				required:"&bull; Vuelva a Ingresar la Contrase??a Nueva.",
				equalTo:"&bull; Las Contrase??as Nuevas no Coinciden."
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
	
	function MostrarFilas(Fila) {
    var elementos = document.getElementsByName(Fila);
  		 for (i = 0; i < elementos.length; i++)
      		elementos[i].style.display ="";
	}

	function OcultarFilas(Fila) {
   	var elementos = document.getElementsByName(Fila);
  	 	for (k = 0; k < elementos.length; k++)
      	elementos[k].style.display = "none";
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
		$.get("emp_programacion/numRut.asp",
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
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO5,PERMISO6,PERMISO12 from USUARIOS "
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
      <a href="#" onclick="passChange();"><b>Cambiar Contrase??a</b></a>
      <br />
      <br />
      <button OnClick="document.location.href='index.asp';">Cerrar Sesi??n</button>
		</div>
		<h3>Opciones</h3>
		<div class="bg1">
				<ul>
				<li class="first"><a href="operacionsolempresa.asp">Solicitud de Nueva Empresa</a></li>
                <li><a href="operacionprogramacion.asp">Programaci??n de Cursos</a></li>
                <li><a href="operacioncalendario.asp">Inscripci??n de Cursos</a></li>
				<li><a href="operacioninspendientes.asp">Autorizar Inscripciones</a></li>
				<li><a href="operacioninsactivas.asp">Inscripciones Autorizadas</a></li>
                <li><a href="operacioncierre.asp">Revisi??n y Cierre</a></li>
                <li><a href="operacionhistins.asp">Historico de Inscripciones</a></li>                
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Calendario de Cursos</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
            <br />
            <h2><label style="font-size:14px; text-transform: none;color: #31576F; font-style:italic"><B>Nota: Los Cursos en Negrita son Exclusivos para una Empresa.</B></label></h2>            
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright ?? 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro de Inscripci??n">
</div>
<div id="dialog_trab" title="Agregar Trabajador">
</div>
<div id="mensaje" title="Registro de Inscripci??n">
     <label id="txtMensaje" name="txtMensaje"></label>
</div>
<div id="datosTerminos" title="Terminos y  Condiciones">
</div>
<div id="pantContrasena" title="Cambiar Contrase??a">
</div>
<div id="mContrasena" title="Cambiar Contrase??a">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>
