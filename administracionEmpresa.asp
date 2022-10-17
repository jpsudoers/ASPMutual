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
var cargada=false;
$(document).ready(function(){					
tabla();
//actualizarGrilla2();
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 635,
			width: 1020,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmEmpresa").valid())
						{
							//window.open($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize());
							$.post($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize(),function(d){
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
		
		$("#msnValida").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 180,
			width: 400,
			modal: true,
			buttons: {
				Aceptar: function() {
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
		
		$("#dialog_espera").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 180,
			width: 400,
			modal: true
		});
		
		$("#dialog_el").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
						$.post("empresa/eliminar.asp",{id:emp_del},function(f){
																			tabla();
																			});
						$(this).dialog('close');
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
		$("#dialogUsuario").dialog({
			bgiframe: true,
			autoOpen: false,
			height:350,
			width: 550,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmAgregarUsuario").valid())
						{
						//window.open($('#frmBloque').attr('action')+'?'+$('#frmBloque').serialize());
						$.post($('#frmAgregarUsuario').attr('action')+'?'+$('#frmAgregarUsuario').serialize(),function(d){
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
		$("#dialogUsuario2").dialog({
			bgiframe: true,
			autoOpen: false,
			height:350,
			width: 550,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmAgregarUsuario").valid())
						{
						//window.open($('#frmBloque').attr('action')+'?'+$('#frmBloque').serialize());
						$.post($('#frmAgregarUsuario').attr('action')+'?'+$('#frmAgregarUsuario').serialize(),function(d){
																							$("#list3").trigger("reloadGrid"); 
																						   });
						    //grilla2();
							$("#list3").trigger("reloadGrid");
							$('#numBloques').val("1");
							$(this).dialog('close');
						}
				},
				Cancelar: function() {
					//grilla2();
					$(this).dialog('close');
				}
			}
		});
		$("#dialog_elUsuario").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
						$.post("empresa/eliminarUsuario.asp",{id:usuarioEmpresa});
						$(this).dialog('close');
						$("#list2").trigger("reloadGrid");
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	

	function tabla()
	{
		//alert($('[name="rbtn_buscar"]:checked').val()+' '+$('[name="rbtn_tipo"]:checked').val()+' '+$("#txt_buscar").val());
		sUrl='empresa/listar.asp?campo='+$('[name="rbtn_buscar"]:checked').val()+'&tipo='+$('[name="rbtn_tipo"]:checked').val()+'&texto='+$("#txt_buscar").val();
		//window.open(sUrl)
		if(cargada){
			jQuery("#list1").jqGrid('setGridParam',{url:sUrl});
			jQuery("#list1").trigger("reloadGrid"); 
		}else{
			 cargada=true;
   			 jQuery("#list1").jqGrid({ 
				url:sUrl, 
				datatype: "xml", 
				colNames:['Rut', 'Razón Social', '&nbsp;','&nbsp;','&nbsp;'], 
				colModel:[
						   {name:'rut',index:'rut', width:30, align:'center'}, 
						   {name:'r_social',index:'r_social'}, 
						   { align:"right",editable:true, width:10}, 
						   { align:"right",editable:true, width:10}, 
						   { align:"right",editable:true, width:9} ], 
				rowNum:300, 
				autowidth: true, 
				rowList:[300,500,1000], 
				pager: jQuery('#pager1'), 
				sortname: 'rut', 
				viewrecords: true, 
				sortorder: "asc", 
				caption:"Listado de Empresas" 
				}); 
			
			jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
			jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
															title:"Agregar nuevo registro", 
															buttonicon :'ui-icon-plus', 
															onClickButton:function(){
																$.post("empresa/frmEmpresa.asp",
																	   {id:0},
																	   function(f){
																		   $('#dialog').html(f);
																		   //tabla2();
																		   //grilla2();
																		   llena_mutual(0);
																		   llena_otic(0);
																		   validar();
																});
																$('#dialog').dialog('open');
															} }); 
		
			jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
															  title:"Refrescar", 
															  buttonicon :'ui-icon-refresh', 
															  onClickButton:function(){ 
															     $("#txt_buscar").val("");
																 //$("#list1").trigger("reloadGrid");
																 tabla();
															  } }); 
															  
			jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
															  title:"Exportar a Excel", 
															  buttonicon :'ui-icon-script',
															  onClickButton:function(){ 
															     window.open("empresa/xlsEmp.asp?c=1",'Informe')
															  } }); 																  
		  }
	}
	
	function eliminar(i){
	 emp_del=i;
	 $('#dialog_el').dialog('open');
	}

	function validar(){
		$("#frmEmpresa").validate({
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
				required:"&bull; Ingrese Teléfono Empresa."
			}
		}
	});
    }
	
	function val(e) {
		tecla = (document.all) ? e.keyCode : e.which; // 2
		if (tecla==8){return true;} // 3
		patron =/[A-Za-z\s]/; // 4
		te = String.fromCharCode(tecla); // 5
		return patron.test(te); // 6
	}
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}
	
	function llena_mutual(id_mutual){
		$("#txtMut").html("");
		$.get("empresa/mutual.asp",
					function(xml){
						$("#txtMut").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_mutual==$(this).find('ID_MUTUAL').text())
								$("#txtMut").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$("#txtMut").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
						 if(id_mutual=='0')
							 	$("#txtMut").append("<option value='0' selected='selected'>Sin Mutual</option>");
							 else
								$("#txtMut").append("<option value='0'>Sin Mutual</option>");
					});
	}
	
	function llena_otic(id_otic){
		//window.open("autorizacion/empresa.asp");
		$("#OTIC").html("");
		$.get("empresa/otic.asp",
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
						 if(id_otic=='0')
							 	$("#OTIC").append("<option value='0' selected='selected'>Sin OTIC</option>");
							 else
								$("#OTIC").append("<option value='0'>Sin OTIC</option>");
					});
	}
	
	function update(i){
		//document.getElementById("frmEmpresa").reset();
		$.post("empresa/frmEmpresa.asp",
			   {id:i},
			   function(f){
				   validar();
				   $('#dialog').html(f);
				    llena_mutual($('#txtIdMut').val());
					llena_otic($('#txtIdOtic').val());
					grilla(i);
					//verificaContactos();
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
	
	function actDes_Empresa(i,estado){
		/*$.post("empresa/actDesEmpresa.asp",{empresa:i,estado:estado},function(d){
			   																		$("#list1").trigger("reloadGrid"); 
																				 });*/
	}
	
	/*function verificaContactos()
	{
		if($("#mismoContacto").attr("checked")) 
		{
			$("#contactoIgual").val("1");
			$("#txtPassConta").val("");	
			$("#filaPassConta").hide();
			
			$('#txtNombConta').attr("readOnly", true);
			$("#txtMailConta").attr("readOnly", true);
			$("#txtContaFono").attr("readOnly", true);
			$("#txtCargoConta").attr("readOnly", true);
		} 
		else 
		{
			$("#contactoIgual").val("0");
			$("#filaPassConta").show();			
			
			$('#txtNombConta').attr("readOnly", false);
			$("#txtMailConta").attr("readOnly", false);
			$("#txtContaFono").attr("readOnly", false);
			$("#txtCargoConta").attr("readOnly", false);
		}
	}	*/
	
	/*function AgregarDatos()
	{
		if($("#mismoContacto").attr("checked")) 
		{
			$("#contactoIgual").val("1");
			//alert($("#mismoContacto").val());
			$("#txtNombConta").val($("#txtNomb").val());
			$("#txtMailConta").val($("#txtMail").val());			
			$("#txtContaFono").val($("#txtFonoCont").val());
			$("#txtCargoConta").val($("#txtCargo").val());
			$("#txtPassConta").val("");	
			$("#filaPassConta").hide();
			
			$('#txtNombConta').attr("readOnly", true);
			$("#txtMailConta").attr("readOnly", true);
			$("#txtContaFono").attr("readOnly", true);
			$("#txtCargoConta").attr("readOnly", true);
		} 
		else 
		{
			$("#contactoIgual").val("0");
			$("#txtNombConta").val("");
			$("#txtMailConta").val("");			
			$("#txtContaFono").val("");
			$("#txtCargoConta").val("");
			$("#txtPassConta").val("");	
			$("#filaPassConta").show();			
			
			$('#txtNombConta').attr("readOnly", false);
			$("#txtMailConta").attr("readOnly", false);
			$("#txtContaFono").attr("readOnly", false);
			$("#txtCargoConta").attr("readOnly", false);
		}
	}
	*/
	function envioPass()
	{
		if($("#frmAgregarUsuario").valid())
		{
			$.post($('#frmAgregarUsuario').attr('action')+'?'+$('#frmAgregarUsuario').serialize(),function(d){
			   																				 
																						   });
				rutEmpresa = $('#txRut').val();
				nomUser = $('#txtNomb').val(); 
				correoUser = $('#txtMail').val(); 
				passUser = $('#txtPassCord').val();

			
			
			$('#dialog_espera').dialog('open');
$.post("empresa/EnvioContrasena.asp",{rutEmpresa:rutEmpresa,nomUser:nomUser,correoUser:correoUser,passUser:passUser},function(d){
									             $('row',d).each(function(i) { 
														$('#dialog_espera').dialog('close');
																								
														if($(this).find('Valido').text()=="1")
														{
								  				            $('#msnValidaTxt').html("Información Enviada Exitosamente.");
															$('#msnValida').dialog('open');
														}
												 });
									});
		}
	}
	
	/*function copiaDatos()
	{	
	    if($("#mismoContacto").attr("checked")) 
		{
			$("#txtNombConta").val($("#txtNomb").val());
			$("#txtMailConta").val($("#txtMail").val());			
			$("#txtContaFono").val($("#txtFonoCont").val());
			$("#txtCargoConta").val($("#txtCargo").val());
		} 
	}*/
	
	function MuestraRef(){
		$('#lbl1').hide();
		$('#lbl2').hide();
		$('#lblHES').hide();
		$('#lblMIGO').hide();
		$('#lbl5').hide();

		if($('#cond_oc').attr("checked"))
		{
			$('#lbl1').show();
			$('#lbl2').show();
			$('#lblHES').show();
			$('#lblMIGO').show();
			$('#lbl5').show();
		}
	}
	
	
	function grilla(id)
	{
		
		jQuery("#list2").jqGrid({ 
		url:'empresa/listadoUsuariosEmpresa.asp?id='+id, 
		datatype: "xml", 
		colNames:['Nombre', 'CARGO', 'TELEFONO', 'EMAIL','&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'Nombre',index:'Nombre'}, 
				   {name:'Cargo',index:'Cargo'}, 
				   {name:'Telefono',index:'Telefono', align:'center'}, 
				   {name:'Email',index:'Email', align:'center'},
				   {align:"right",editable:true, width:12},
				   {align:"right",editable:true, width:11}
	
				  ],
		rowNum:30, 
        rownumbers: true, 
        rownumWidth: 25, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager2'), 
		sortname: 'Nombre', 
		viewrecords: true, 
		sortorder: "desc", 
		caption:"Listado de Usuarios" 
		}); 
	
	   jQuery("#list2").jqGrid('navGrid','#pager2',{edit:false,add:false,del:false,search:false,refresh:false});
	   jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
													title:"Agregar nuevo usuario", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("empresa/frmAgregarUsuario.asp",
														{idEmpresa:id},
															   function(f){
																 $('#dialogUsuario').html(f);
																
														});
														$('#dialogUsuario').dialog('open');
													} }); 

	jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list2").trigger("reloadGrid");
													  } }); 
	}

	function eliminar2(i){
	 usuarioEmpresa=i;
	 $('#dialog_elUsuario').dialog('open');
	}
	function update2(i){
		$.post("empresa/frmAgregarUsuario.asp",
			   {id:i},
			   function(f){
				    $('#dialogUsuario').html(f);
					//grilla(i);
				    validar();
		});
		 $('#dialogUsuario').dialog('open');
	
}
function grilla2()
{		sUrl='empresa/listadoUsuariosNuevosEmpresa.asp?rutEmpresa=' + $('#txRut').val();
		//window.open(sUrl)
		console.log($("#txRut").val());
		jQuery("#list3").jqGrid({ 
		url:sUrl,//'empresa/listadoUsuariosNuevosEmpresa.asp?rutEmpresa=' + $('#txRut').val(), 
		//url:sUrl,//'empresa/listadoUsuariosNuevosEmpresa.asp?rutEmpresa=' + $('#txRut').val(), 
		//data: $("#txRut").val(),
		datatype: "xml", 
		colNames:['rut_Empresa','Nombre', 'CARGO', 'TELEFONO', 'EMAIL','CONTRASEÑA','&nbsp;'], 
		colModel:[
				   {name:'Rut_Empresa',index:'Rut_Empresa'}, 
				   {name:'NOMBRE',index:'NOMBRE'}, 
				   {name:'CARGO',index:'CARGO'}, 
				   {name:'TELEFONO',index:'TELEFONO', align:'center'}, 
				   {name:'EMAIL',index:'EMAIL', align:'center'},
				   {name:'CONTRASEÑA',index:'CONTRASEÑA', align:'center'},
				   {align:"right",editable:true, width:12}
				  ],
		rowNum:30, 
        rownumbers: true, 
        rownumWidth: 25, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager3'), 
		sortname: 'Nombre', 
		viewrecords: true, 
		sortorder: "desc", 
		caption:"Listado de nuevos usuarios" 
		}); 
	
	   jQuery("#list3").jqGrid('navGrid','#pager3',{edit:false,add:false,del:false,search:false,refresh:false});
	   jQuery("#list3").jqGrid('navButtonAdd',"#pager3",{caption:"",
													title:"Agregar nuevo usuario", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("empresa/frmAgregarUsuario.asp",
														{rutEmpresa:$("#txRut").val()},
															   function(f){
																 validar();
																 //grilla2();
																 // actualizarGrilla2();
																 $('#dialogUsuario2').html(f);
																 
														});
														$('#dialogUsuario2').dialog('open');
													} }); 

	jQuery("#list3").jqGrid('navButtonAdd',"#pager3",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
													  //grilla2();
													  actualizarGrilla2();
 														// $("#list3").trigger("reloadGrid");
													  } }); 
	}
	
	
	function validarCamposNuevaEmpresa(){
	
	   var validacion = Errores();
	   //console.log(validacion);
	   if  ($("#txRut").val() != "" 
	   &&  $("#txtRsoc").val() != "" 
	   &&  $("#txtGiro").val() != "" 
	   &&  $("#txtDir").val() != "" 
	   &&  $("#txtCom").val() != "" 
	   &&  $("#txtCiu").val() != "" 
	   &&  $("#txtFon").val() != "" 
	   &&  $("#txtFax").val() != ""
	   && !validacion
	   ){
			grilla2();
	   }
	}
	function validarMail(){
		if($("#txtMail").val().indexOf('@', 0) == -1 || $("#txtMail").val().indexOf('.', 0) == -1) {
            $("#messageBox1 ul").append("<li>Email tiene formato incorrecto</li>");
            return false;
        }else{
			$("#messageBox1 ul").find("li").remove();
		}
	
	}
		function Errores(){
        count=0;
			$('#messageBox1 ul').each(function(i) {
				var len = $(this).find('li').size();
				//console.log(len);
					if(len>0) count++;
				});
			if (count > 0){
				var prueba = $('#messageBox1 ul li').css('display');
				if(prueba != 'none'){
					return true;
				}else {
					return false;
				}
			}
			else{
				return false;
			}
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
			<h2><em style="text-transform: capitalize;">Empresas</em></h2>
           <center><table width="709" border="0">
                    <tr>
                    <td width="650"><div id="buscarContent" style="border:1px solid #000099;display:block" >
					<div>
						<table cellpadding="0" cellspacing="0" style="border:0px;width:100%;">
							<tr>
								<td align="left"><h3><em style="text-transform: capitalize;">Busqueda</em></h3></td>
							</tr>
						</table>
					</div>
					<table border="0" align="center" cellpadding="1" cellspacing="1">
						<tr>
							  <td width="400" align="center">
								<!--<label id="campoBuscar">-->
								<input name="rbtn_buscar" id="rbtn_buscarR" type="radio" value="RUT" checked="checked" />Por Rut
           				        <input name="rbtn_buscar" id="rbtn_buscarN" type="radio" value="R_SOCIAL"/>Por Razón Social
								<!--</label>--></td>
						</tr>
						<tr>
							<td><input name="txt_buscar" type="text" id="txt_buscar" size="80"/></td>
							<td><input type="button" class="boton" value="Buscar" id="btnBuscar" onclick="tabla();"/></td>
						</tr>
						<tr>
                          <td align="center"><!--<label id="tipoBuscar">-->
								<input name="rbtn_tipo" type="radio" value=" = '"/>Es exacta
								<input name="rbtn_tipo" type="radio" value=" LIKE '%" checked="checked"/>Que contienen
							    <!--</label>-->	  
						  </td>
						</tr>
					</table>
				</div></td>
              </tr>
            </table></center>
            <table width="200" border="0">
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog_espera" title="Envio de Contraseñas">
<center>
<table width="100" border="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><center><font size="2"><b>Espere ...</b></font></center></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><center><img id="loading" src="images/loader.gif"/></center></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</center>
</div>
<div id="msnValida" title="Envio de Contraseñas">
     <label id="msnValidaTxt" name="msnValidaTxt"></label>
</div>
<div id="dialog" title="Registro Empresa">
</div>
<div id="dialog_el" title="Eliminar Empresa">
	<p>Esta seguro de eliminar la empresa.</p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="dialog_elUsuario" title="Eliminar Usuario Empresa">
	<p>Esta seguro de eliminar el usuario de la empresa.</p>
</div>
<div id="dialogUsuario" title="Usuario Empresa">
</div>
<div id="dialogUsuario2" title="Usuario Empresa">
</div>
</body>
</html>