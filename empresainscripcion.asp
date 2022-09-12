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
var amtval="";
var sol=0;
var i=1;
var dir='inscripcionempresa/listar2.asp?id='+sol
$(document).ready(function(){					
	$('#txUsuario').defaultValue('Nombre de Usuario');
	$('#txPassword').defaultValue('Contraseña');

	var amtval="";
//window.open("inscripcionempresa/listar.asp");

	jQuery("#list1").jqGrid({ 
		url:'inscripcionempresa/listar.asp', 
		datatype: "xml", 
		colNames:['Cód. Prog.', 'Rut', 'Nombre Empresa', 'Cód. Curso', 'Nombre Curso'], 
		colModel:[
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO', width:50}, 
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO', width:50}, 
				   {name:'empresa',index:'empresa'}, 
				   {name:'otic',index:'otic', width:50}, 
				   {name:'FECHA__AUTORIZACION',index:'FECHA__AUTORIZACION', width:70}], 
		rowNum:10, 
		autowidth: true, 
		rowList:[10,20,30], 
		pager: jQuery('#pager1'), 
		sortname: 'AUTORIZACION.ID_AUTORIZACION', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Inscripciones" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("inscripcionempresa/frminscripcion.asp",
															   {id:0},
															   function(f){
																  $('#dialog').html(f);
																  $("#autorizacion").append("<option value=\"\">Seleccione</option>");
																  grilla(0);
																 
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
			height:540,
			width: 900,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frminscripcion").valid())
						{
							//window.open($('#frminscripcion').attr('action')+'?'+$('#frminscripcion').serialize());
							$.post($('#frminscripcion').attr('action')+'?'+$('#frminscripcion').serialize(),function(d){
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
						$.post("inscripcionempresa/eliminar.asp",{id:emp_del});
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
		//$('#frminscripcion txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		$("#frminscripcion").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			autorizacion:{
				required:true
			}
		},
		messages:{
			autorizacion:{
				required:"&bull; Seleccione Código Autorización."
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
	   cargaDatos2(id);
	  // alert(id);
	   llena_autorizacion(0,id);
	   $('#txRut').val(rut);llena_autorizacion
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
	
	function cargaDatos2(id)	
	{
		$("#txRut").val("");
		$("#txtRsoc").html("");

		$.get("autorizacion/datosempresa.asp",
						 {id:id},
						 function(xml){
									$('row',xml).each(function(i) { 
									$("#txtRsoc").html($(this).find('RSOCIAL').text());
									$("#txRut").val($(this).find('RUT').text());
									// llena_autorizacion(0);
						});
			});
	}


	function llena_autorizacion(id_autorizacion,empresa){
		//alert(id_autorizacion,empresa);
		$("#autorizacion").html("");
		$.get("inscripcionempresa/autorizacion.asp",
			  		{empresa:empresa},
					function(xml){
						$("#autorizacion").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_autorizacion==$(this).find('ID').text())
								$("#autorizacion").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																		$(this).find('CODIGO').text()+ '</option>');
							else
								$("#autorizacion").append('<option value="'+$(this).find('ID').text()+'" >'+
																		$(this).find('CODIGO').text()+ '</option>');
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

	function cargaDatos(id_autorizacion)	
	{
		$("#txtRutEmpresa").html("");
		$("#txtRSocial").html("");
		$("#txtIdEmpresa").val("");
		$("#nomCurso").html("");
		$("#desCurso").html("");
		
		$("#txtRutInstructor").html("");
		$("#nomInstructor").html("");
		$("#txtIdPrograma").val("");
		$("#txtSede").html("");
		$("#txtDir").html("");
		
		$.get("inscripcionempresa/datosautorizacion.asp",
						 {id_autorizacion:id_autorizacion},
						 function(xml){
								$('row',xml).each(function(i) { 
									$("#txtRutEmpresa").html($(this).find('RUT').text());
									$("#txtRSocial").html($(this).find('RSOCIAL').text());
									$("#txtIdEmpresa").val($(this).find('IDEMPRESA').text());
									$("#nomCurso").html($(this).find('NOMBRE').text());
									$("#desCurso").html($(this).find('DESCRIPCION').text());
									cargaDatosPrograma($(this).find('IDMUTUAL').text())
									//alert(sol);
									//sol=50;
									//alert(sol);
									//dir='inscripcionempresa/listar2.asp?id='+sol
									
						});
						//$("#nombre").focus();	
			});
	}
	
	function grilla(id)
	{
	
	//alert(id);
	//window.open('inscripcionempresa/listar2.asp?id='+id);
	
														jQuery("#list2").jqGrid({ 
														scroll: 1, 					
														url:'inscripcionempresa/listar2.asp?id='+id, 
														datatype: "xml", 
														colNames:['Rut', 'Nombre, Apellido Paterno, Apellido Materno', 'Cargo En La Empresa', 'Escolaridad'], 
														colModel:[
																   {name:'rut',index:'rut', width:40, editable:true}, 
																   {name:'nombre',index:'nombre', editable:true}, 
																   {name:'cargo',index:'cargo', width:80, editable:true}, 
															{name:'escolaridad',index:'escolaridad', width:80, editable:true,edittype:"select",
																   editoptions:{value:"0:Sin Escolaridad;1:Básica Incompleta;2:Básica Completa;3:Media Incompleta;4:Media Completa;5:Superior Técnica Incompleta;6:Superior Técnica Profesional Completa;7:Universitaria Incompleta;8:Universitaria Completa"}}], 
														rowNum:100, 
														rownumbers: true, 
														rownumWidth: 40, 
														autowidth: true, 
														rowList:[10,20,30], 
														pager: jQuery('#pager2'), 
														sortname: 'PROGRAMA.ID_PROGRAMA', 
														viewrecords: true, 
														sortorder: "asc", 
														caption:"Nomina de Inscritos" ,
														forceFit : true, 
														cellEdit: true, 
														cellsubmit: 'clientArray',
														afterSaveCell : function(rowid,name,val,iRow,iCol) { 
														if(name=='escolaridad') 
														{ 
															
															amtval += "'" + jQuery("#list2").jqGrid('getCell',rowid,1) + 
																	"','" + jQuery("#list2").jqGrid('getCell',rowid,2) + 
																	"','" + jQuery("#list2").jqGrid('getCell',rowid,3) + 
																	"'," + val + '/'; 
															//alert(amtval);
															$("#datostabla").val(amtval);
															
															i+=1;
															jQuery("#list2").jqGrid('addRowData',i, mydata1[0]); 
															
															}}
														}); 
																  
	}
	
	
	var mydata1 = [ 
					   {rut:"",nombre:"",cargo:'',escolaridad:''}
					   ]; 

	
	function cargaDatosPrograma(id_programa)	
	{
		$("#txtRutInstructor").html("");
		$("#nomInstructor").html("");
		$("#txtIdPrograma").val("");
		$("#txtSede").html("");
		$("#txtDir").html("");
		
		
		$.get("inscripcionempresa/datosprograma.asp",
						 {id_programa:id_programa},
						 function(xml){
								$('row',xml).each(function(i) { 
									$("#txtRutInstructor").html($(this).find('RUT').text());
									$("#nomInstructor").html($(this).find('instructor').text());
									$("#txtIdPrograma").val($(this).find('IDPROGRAMA').text());
									$("#txtSede").html($(this).find('SEDE').text());
									$("#txtDir").html($(this).find('DIRECCION').text());
									// llena_programa($(this).find('IDMUTUAL').text(),0);
						});
						//$("#nombre").focus();	
			});
	}
	
	function update(i){
		$.post("inscripcionempresa/frminscripcion.asp",
			   {id:i},
			   function(f){
				    $('#dialog').html(f);
					jQuery("#list2").jqGrid({ 
														scroll: 1, 					
														url:'solicitud/listar.asp', 
														datatype: "xml", 
														colNames:['Rut', 'Nombre, Apellido Paterno, Apellido Materno', 'Cargo En La Empresa', 'Escolaridad'], 
														colModel:[
																   {name:'rut',index:'rut', width:40, editable:true}, 
																   {name:'nombre',index:'nombre', editable:true}, 
																   {name:'cargo',index:'cargo', width:80, editable:true}, 
																   {name:'escolaridad',index:'escolaridad', width:80, editable:true,edittype:"select",
																   editoptions:{value:"0:Sin Escolaridad;1:Básica Incompleta;2:Básica Completa;3:Media Incompleta;4:Media Completa;5:Superior Técnica Incompleta;6:Superior Técnica Profesional Completa;7:Universitaria Incompleta;8:Universitaria Completa"}}], 
														rowNum:100, 
														rownumbers: true, 
														rownumWidth: 40, 
														autowidth: true, 
														rowList:[10,20,30], 
														pager: jQuery('#pager2'), 
														sortname: 'PROGRAMA.ID_PROGRAMA', 
														viewrecords: true, 
														sortorder: "asc", 
														caption:"Nomina de Inscritos" ,
														forceFit : true, 
														cellEdit: true, 
														cellsubmit: 'clientArray',
														afterSaveCell : function(rowid,name,val,iRow,iCol) { 
														if(name=='escolaridad') 
														{ 
															
															amtval += "'" + jQuery("#list2").jqGrid('getCell',rowid,1) + 
																	"','" + jQuery("#list2").jqGrid('getCell',rowid,2) + 
																	"','" + jQuery("#list2").jqGrid('getCell',rowid,3) + 
																	"'," + val + '/'; 
															$("#datostabla").val(amtval);
															}}
														}); 
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
			<h2><em style="text-transform: capitalize;">Inscripción</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Inscripción">
</div>
<div id="dialog_el" title="Eliminar Inscripción">
	<p>Esta seguro de eliminar la Inscripción.</p>
</div>

</body>
</html>
