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
var emp_del;
var emp_del_tab;
var bus_rel;
var del_idBloque;
var del_idprog;
$(document).ready(function(){					
//window.open("programacion/listar.asp");

	jQuery("#list1").jqGrid({ 
		url:'programacion/listar.asp', 
		datatype: "xml", 
		colNames:['F. Inicio','Cód.','Código Curso', 'Nombre Curso', 'Cupos', 'Ins.', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'FECHA_INICIO_',index:'FECHA_INICIO_', width:42, align:'center'}, 				  
				   {name:'ID_PROGRAMA',index:'ID_PROGRAMA', width:22, align:'center'}, 
				   {name:'CODIGO',index:'CODIGO', width:48, align:'center'}, 
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO'}, 
				   {name:'CUPOS',index:'CUPOS', width:22, align:'center'}, 
				   {name:'INSCRITOS',index:'INSCRITOS', width:22, align:'center'}, 
				   { align:"right",editable:true, width:13}, 
				   { align:"right",editable:true, width:10} ], 
		rowNum:300, 
		height:350,
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'PROGRAMA.FECHA_INICIO_', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Programaciones" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("programacion/frmprogramacion.asp",
															   {id:0},
															   function(f){
																  $('#dialog').html(f);
																  $('#txtFechInicio').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
																  $('#txtFechTermino').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
																  $('#txtFechApertura').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
																  $('#txtFechCierre').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
																  llena_curriculo(0);
																  llena_Tipo(0);
																  //$("#tdLabValEsp").hide();
																  //$("#tdValEsp").hide();																  
																  mostrarFilaEmpresa(0);
																  $('#numBloques').val("");
																  $('#txtInscritos').html("0");
																  $('#txtVacantes').html("0");
																  grilla();
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
			height:570,
			width: 900,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmProgramacion").valid())
						{
							//window.open($('#frmProgramacion').attr('action')+'?'+$('#frmProgramacion').serialize());
							$.post($('#frmProgramacion').attr('action')+'?'+$('#frmProgramacion').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
																						   });
							$("#list1").trigger("reloadGrid");
							$(this).dialog('close');
						}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
		
			$("#dialogBloque").dialog({
			bgiframe: true,
			autoOpen: false,
			height:350,
			width: 550,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmBloque").valid())
						{
						//window.open($('#frmBloque').attr('action')+'?'+$('#frmBloque').serialize());
						$.post($('#frmBloque').attr('action')+'?'+$('#frmBloque').serialize(),function(d){
			   																				$("#list2").trigger("reloadGrid"); 
																						   });
							$("#list2").trigger("reloadGrid");
							$('#numBloques').val("1");
							$(this).dialog('close');
						}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	
		
		$("#dialog_el_bloque").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 400,
			modal: true,
			buttons: {
				'Aceptar': function() {
					//alert(del_idprog+' '+del_idBloque)
						$.post("programacion/cerrarEliminar.asp",{programa:del_idprog,bloque:del_idBloque},function(d){
			   									$("#list2").trigger("reloadGrid"); 
											});
						
						$(this).dialog('close');
				},
				Cancelar: function() {
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
						$.post("programacion/eliminar.asp",{id:emp_del});
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
		//$('#frmProgramacion txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		$("#frmProgramacion").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			Curriculo:{
				required:true
			},
			Tipo:{
				required:true
			},
			id_empresa:{
				required: {depends:function(element){if($("#Tipo").val()==1) return true ;else return false;}}
			},			
			txtFechApertura:{
				required:true
			},
			txtFechCierre:{
				required:true
			},
			txtFechInicio:{
				required:true
			},
			txtFechTermino:{
				required:true
			},
			txtValor:{
				required:true
			},
			txtCupo:{
				required:true
			},
			numBloques:{
				required:true
			},
			txtValEsp:{
				required:true,//required: {depends:function(element){if($("#Tipo").val()==1) return true ;else return false;}},
				number:true
			}		
		},
		messages:{
			Curriculo:{
				required:"&bull; Seleccione Curriculo."
			},
			Tipo:{
				required:"&bull; Seleccione Tipo de Programa."
			},
			id_empresa:{
				required:"&bull; Ingrese Rut o Razón de la Empresa."
			},				
			txtFechApertura:{
				required:"&bull; Seleccione Fecha De Apertura."
			},
			txtFechCierre:{
				required:"&bull; Seleccione Fecha De Cierre."
			},
			txtFechInicio:{
				required:"&bull; Seleccione Fecha De Inicio."
			},
			txtFechTermino:{
				required:"&bull; Seleccione Fecha De Termino."
			},
			txtValor:{
				required:"&bull; Ingrese Valor Programa."
			},
			txtCupo:{
				required:"&bull; Ingrese Cupo Programa."
			},
			numBloques:{
				required:"&bull; Ingrese al Menos un Bloque."
			},
			txtValEsp:{
				required:"&bull; Ingrese Valor Especial del Curso.",
				number:"&bull; Ingrese Solo Números en el Valor Especial del Curso."
			}
		}
	});
    }

	function llena_curriculo(id_curriculo){
		$("#Curriculo").html("");
		$.get("programacion/curriculo.asp",
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
	
	function llena_Tipo(id_tipo){
		$("#Tipo").html("");
		sSel = " selected=\"selected\" "
		$("#Tipo").append("<option value=\"\">Seleccione</option>");
		if(id_tipo==1){
			$("#Tipo").append("<option value=\"1\" selected=\"selected\">Para Una Empresa</option>");
			$("#Tipo").append("<option value=\"2\" >Para Varias Empresas</option>");
		}else if(id_tipo==2){
			$("#Tipo").append("<option value=\"1\">Para Una Empresa</option>");
			$("#Tipo").append("<option value=\"2\" selected=\"selected\">Para Varias Empresas</option>");
		}else{
			$("#Tipo").append("<option value=\"1\">Para Una Empresa</option>");
			$("#Tipo").append("<option value=\"2\">Para Varias Empresas</option>");
		}
	}
	
	function mostrarFilaEmpresa(tipo){
		if(tipo==1)
		{
			$("#filaEmpresa").show();
			//$("#tdLabValEsp").show();
			//$("#tdValEsp").show();
			$("#labValor").html("Valor Referencia :");
			$("#txtValEsp").val("");
		}
		else
		{
			$("#filaEmpresa").hide();
			//$("#tdLabValEsp").hide();
			//$("#tdValEsp").hide();	
			$("#labValor").html("Valor :");			
			//$("#txtValEsp").val("");				
			$("#txRut").val("");
			$("#txtRsoc").html("");
			$("#id_empresa").val("");
			//if($("#txtValor").html()!="&nbsp;"){
			//$("#txtValEsp").val($("#txtValor").html().replace('$ ','').replace('.','').replace('&nbsp;',''));//}
		}
	}
	
	function lookup(inputString) {
		if(inputString.length <3) {
			$('#suggestions').hide();
			$("#txtRsoc").html("");
			$("#id_empresa").val("");
		} else {
				$.post("programacion/sugEmpresa.asp", {txt: inputString}, function(data){
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
 
		$.get("emp_programacion/datosempresaoperacion.asp",{id:id},
						 function(xml){
									$('row',xml).each(function(i) { 
									$("#txtRsoc").html($(this).find('RSOCIAL').text());
						});
			});
	}

	function fill(id,rut) {
	   $('#id_empresa').val(id);
	   cargaDatosEmpresa(id);
	   $('#txRut').val(rut);
	   $('#suggestions').hide();
	}

	
	function cargaDatos(id_curso)	
	{
		$("#txtCurso").html("");
		$("#txtValor").html("");
		
		$.get("programacion/datoscurso.asp",
						 {id_curso:id_curso},
						 function(xml){
								$('row',xml).each(function(i) { 
									$("#txtCurso").html($(this).find('curso').text());
									if($(this).find('sence').text()=="1"){
										$("#Sence_NO").attr("checked","checked");
										$("#Sence_SI").attr("checked","");
									}
									
									if($(this).find('sence').text()=="0"){
										$("#Sence_NO").attr("checked","");
										$("#Sence_SI").attr("checked","checked");
									}
									
									$("#txtValor").html($(this).find('valor').text());
									/*$("#txtValEsp").val("");
									if($("#Tipo").val()!=1)
									{
										$("#txtValEsp").val($(this).find('valor').text().replace('$ ','').replace('.',''));
									}*/
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
									cupo=$('#txtCupo').val();
									ins=$(this).find('inscritos').text();
									$('#txtVacantes').html(cupo-ins);
						});
			});
	}
	
	function calcula(){
		 $('#txtVacantes').html($('#txtCupo').val());
	}
	
	function calVacantes(){
		 $('#txtVacantes').html(parseInt($('#txtCupo').val())-parseInt($('#txtInscritos').html()));
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

	function grilla()
	{
		//window.open("programacion/listaTabTemp.asp?id_prog=" + $('#tabFecha').val())
		jQuery("#list2").jqGrid({ 
		url:'programacion/listaTabTemp.asp?id_prog=' + $('#tabFecha').val(), 
		datatype: "xml", 
		colNames:['Relator (es)', 'Sala/Sede', 'Cupos', 'Inscri.','&nbsp;','&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'CODIGO',index:'CODIGO'}, 
				   {name:'SENCE',index:'SENCE'}, 
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO', width:19, align:'center'}, 
				   {name:'CUPOS',index:'CUPOS', width:19, align:'center'},
				   {align:"right",editable:true, width:12},
				   {align:"right",editable:true, width:11},
				   {align:"right",editable:true, width:11}], 
		rowNum:30, 
        rownumbers: true, 
        rownumWidth: 25, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager2'), 
		sortname: 'INSTRUCTOR_RELATOR.RUT', 
		viewrecords: true, 
		sortorder: "desc", 
		caption:"Listado de Bloques" 
		}); 
	
	   jQuery("#list2").jqGrid('navGrid','#pager2',{edit:false,add:false,del:false,search:false,refresh:false});
	   jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("programacion/frmBloque.asp",
															   {id:$('#tabFecha').val(),bloque:"0",estado:"0",totalVacantes:$('#txtCupo').val()},
															   function(f){
																 $('#dialogBloque').html(f);
																 bus_rel=0;
																 $('#filaRelSeg').hide();
																 $('#filaDir').hide();
																 llena_sede(0);
																 llena_relator($('#tabRelator').val(),$('#tabRelSeg').val());
																 $('#AgregaRel').hide();
																 llena_relator_seg($('#tabRelator').val(),$('#tabRelSeg').val());
																 validarFrmTab();
														});
														$('#dialogBloque').dialog('open');
													} }); 

	jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list2").trigger("reloadGrid");
													  } }); 
	}

	function llena_sede(id_sede){
		$("#salaSede_frm").html("");
		$.get("programacion/sedes_frm.asp?idSede="+id_sede+"&progId="+$("#tabFecha").val(),
					function(xml){
						$("#salaSede_frm").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_sede==$(this).find('ID').text())
								$("#salaSede_frm").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																		$(this).find('NOMBRE').text()+ '</option>');
							else
								$("#salaSede_frm").append('<option value="'+$(this).find('ID').text()+'" >'+
																		$(this).find('NOMBRE').text()+ '</option>');
						});
					});
	}
	
	function mostrarDirSede(id_sede){
		if(id_sede=="27")
		{
			$("#filaDir").show();
			//$("#txtDir_frm").focus();
		}
		else
		{
			$("#filaDir").hide();
			$("#txtDir_frm").val("");
		}
	}

	function llena_relator(id_relator,Relatorseg){
		$("#relator_frm").html("");
		$.get("programacion/instructor_frm.asp?idRelator="+id_relator+"&progId="+$("#tabFecha").val()+"&Relatorseg="+Relatorseg,
					function(xml){
						$("#relator_frm").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_relator==$(this).find('ID_INSTRUCTOR').text())
								$("#relator_frm").append('<option value="'+$(this).find('ID_INSTRUCTOR').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$("#relator_frm").append('<option value="'+$(this).find('ID_INSTRUCTOR').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
					});
	}
	
	function elimina_relator(){
		$('#filaRelSeg').hide();
		$('#AgregaRel').show();
		bus_rel=1;
		llena_relator_seg($('#tabRelator').val(),0);
		$('#txtNRelator').val("1");
	}
	
	function agrega_relator(){
		$('#filaRelSeg').show();
		$('#AgregaRel').hide();
		llena_relator_seg($('#tabRelator').val(),$('#tabRelSeg').val());
		$('#txtNRelator').val("2");
	}	
	
	function llena_seg_rel(relator){
		if(relator!="" && relator!=null){
			if($('#txtNRelator').val()==1){
				$('#AgregaRel').show();
			}
		}
		else
		{
			$('#AgregaRel').hide();
		}
	}	
	
	function llena_relator_seg(id_relator,Relatorseg){
		$("#relator_frm_seg").html("");
		$.get("programacion/instructor_frm.asp?idRelator="+id_relator+"&progId="+$("#tabFecha").val()+"&Relatorseg="+Relatorseg,
					function(xml){
						$("#relator_frm_seg").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) {
							if(bus_rel==1)
							{
								Relatorseg=0;
							}
												   
							if(Relatorseg==$(this).find('ID_INSTRUCTOR').text())
								$("#relator_frm_seg").append('<option value="'+$(this).find('ID_INSTRUCTOR').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$("#relator_frm_seg").append('<option value="'+$(this).find('ID_INSTRUCTOR').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
					});
	}
	
	function update(i){
		$.post("programacion/frmprogramacion.asp",
			   {id:i},
			   function(f){
				    $('#dialog').html(f);
				    $('#txtFechInicio').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
					$('#txtFechTermino').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
					$('#txtFechApertura').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
					$('#txtFechCierre').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
					llena_curriculo($('#txtIdCurriculo').val());
					cargaDatos($('#txtIdCurriculo').val());
					llena_Tipo($('#txtTipo').val());
					//$("#tdLabValEsp").hide();
					//$("#tdValEsp").hide();						
					mostrarFilaEmpresa($('#txtTipo').val());
					cargaDatosInscritos($('#txtId').val())
					grilla();
				    validar();
		});
		 $('#dialog').dialog('open');
	}
	
	function update_tab(i,bloque){
		$.post("programacion/frmBloque.asp",
			   {id:i,bloque:bloque,estado:"1",totalVacantes:$('#txtCupo').val()},
			   function(f){
				    $('#dialogBloque').html(f);
					$('#filaRelSeg').hide();
					
                    llena_sede($('#tabSala').val());
					//llena_relator($('#tabRelator').val());
					bus_rel=0;
					llena_relator($('#tabRelator').val(),$('#tabRelSeg').val());
					
					if($('#tabRelSeg').val()!=null && $('#tabRelSeg').val()!=""){
								$('#filaRelSeg').show();
								$('#AgregaRel').hide();
								//llena_relator_seg($('#tabRelSeg').val(),$('#tabRelator').val());
								llena_relator_seg($('#tabRelator').val(),$('#tabRelSeg').val());
								$('#txtNRelator').val("2");
					}
					
					mostrarDirSede($('#tabSala').val());
					
					validarFrmTab();
		});
		 $('#dialogBloque').dialog('open');
	}
	
	function cerrar_bloque(i,bloque,estado){
		//alert(i+' '+bloque+' '+estado);
		$.post("programacion/cerrarBloque.asp",{prog:i,bloque:bloque,estado:estado},function(d){
			   																		$("#list2").trigger("reloadGrid"); 
																				 });
	}
	
	function eliminarBloque(programa,bloque){
		 //alert(programa + ' ' +bloque);
		 del_idBloque=bloque;
		 del_idprog=programa;
		 $('#dialog_el_bloque').dialog('open');
		 
		/*$.post("programacion/cerrarEliminar.asp",{programa:programa,bloque:bloque},function(d){
			   									$("#list2").trigger("reloadGrid"); 
											});*/
		
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

	function validarFrmTab(){
		$("#frmBloque").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			relator_frm:{
				required:true
			},
			relator_frm_seg:{
				required: {depends:function(element){if($("#txtNRelator").val()==2) return true ;else return false;}}
			},
			selRelSeg:{
				required: {depends:function(element){if($("#relator_frm").val()!=null && $("#relator_frm_seg").val()!=null && $("#relator_frm").val()!="" && $("#relator_frm_seg").val()!="" && $("#relator_frm").val()==$("#relator_frm_seg").val()) return true ;else return false;}}
			},			
			salaSede_frm:{
				 required:true
			},
			txtDir_frm:{
				required: {depends:function(element){if($("#salaSede_frm").val()==27) return true ;else return false;}}
			},
			txtCupo_frm:{
				required:true,
				max: $('#totProgDisp').val(),
				min: 1
			}
		},
		messages:{
			relator_frm:{
				required:"&bull; Seleccione el Primer Relator."
			},
			relator_frm_seg:{
				required:"&bull; Seleccione el Segundo Relator."
			},	
			selRelSeg:{
				required:"&bull; El Segundo Relator debe ser Distinto al Primer Relator."
			},
			salaSede_frm:{
				required:"&bull; Seleccione Sede."
			},
			txtDir_frm:{
				required:"&bull; Ingrese Dirección del Curso."
			},			
			txtCupo_frm:{
				required:"&bull; Ingrese Capacidad Bloque.",
				max: "&bull; El cupo disponible para este bloque es igual a "+$('#totProgDisp').val(),
				min: "&bull; El cupo para este bloque debe ser mayor o igual a 1"
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
                <li><a href="operacionhistins.asp">Historico de Inscripciones</a></li>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Programación </em><em>de</em><em style="text-transform: capitalize;"> Cursos</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Programación">
</div>
<div id="dialog_el" title="Eliminar Programación">
	<p>Esta seguro de eliminar la Programación.</p>
</div>
<div id="dialog_el_bloque" title="Eliminar Bloque">
	<p>Esta seguro de eliminar el bloque.</p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="dialogBloque" title="Registro de Bloque">
</div>
</body>
</html>
