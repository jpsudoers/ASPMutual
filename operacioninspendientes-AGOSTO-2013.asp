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
<script src="js/jquery.js" language="javascript"></script>
<script src="js/i18n/grid.locale-sp.js" language="javascript"></script>
<script src="js/jquery-ui.js" language="javascript"></script>
<script src="js/jquery.jqGrid.js" language="javascript"></script>
<script src="js/jquery.tbltogrid.js" type="text/javascript" charset="utf-8"> </script> 
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script type="text/javascript">
var id_prog=0;
var id_preinscripcion=0;
var cargada=false;
var sUrl=""
var ins_tot=1;
var amtval="";
var comienzo=0;
var ingresados=0;
var elimina_instructor="";
var preins_del=0
$(document).ready(function(){					

	jQuery("#list1").jqGrid({ 
		url:'inspendientes/listar.asp', 
		datatype: "xml", 
		colNames:['Fecha', 'Cód.', 'Rut', 'Nombre Empresa', 'Nombre Curso', 'Part.', '&nbsp;', '&nbsp;'], 
		colModel:[
				   {name:'fecha',index:'fecha', width:52,align:'center'}, 
				   {name:'programa',index:'programa', width:30,align:'center'}, 
				   {name:'rut',index:'rut', width:55,align:'right'}, 
				   {name:'empresa',index:'empresa'}, 
				   {name:'curso',index:'curso', width:120}, 
				   {name:'nparticipantes',index:'nparticipantes', width:24, align:'center'},
				   {align:"right",editable:true, width:15},
				   {align:"right",editable:true, width:13}], 
		rowNum:300, 
		height:350,
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'preinscripciones.id_preinscripcion', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Inscripciones Pendientes" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
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
		
		$("#dialog_el").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 300,
			width: 490,
			modal: true,
			buttons: {
				'Aceptar': function() {
						if($("#frmElimAuto").valid())
						{
								//window.open($('#frmElimAuto').attr('action')+'?'+$('#frmElimAuto').serialize())
								$('#sendMsn').html('Espere Mientras se Envia Notificación.');
								$("#sendConf").dialog('open');
								
								$.post($('#frmElimAuto').attr('action')+'?'+$('#frmElimAuto').serialize(),function(d){
																							$("#sendConf").dialog('close');
																							$('#sendMsn').html('');
																							$("#list1").trigger("reloadGrid"); 
																								   });
								$(this).dialog('close');								
								$("#list1").trigger("reloadGrid"); 
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
			height:100,
			width: 600,
			modal: true,
			buttons: {
				Aceptar: function() {
					$(this).dialog('close');
				}
			}
		});
	
	  $("#sendConf").dialog({
			autoOpen: false,
			bgiframe: true,
			height:150,
			width: 350,
			modal: true,
			overlay: {
				backgroundColor: '#f00',
				opacity: 0.5
			},			
		    closeOnEscape: false,
		    open: function(event, ui) { $(".ui-dialog-titlebar-close").hide(); }
	  });	
	
	  $("#Doc").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 550,
			width:1000,
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
	
	 $("#InsBloque").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 400,
			width:500,
			modal: true,
			overlay: {
				backgroundColor: '#000',
				opacity: 0.5
			},
			buttons: {
				'Cerrar': function() {
					$(this).dialog('close');
				}
			}
		});
	
	
	 $("#cursos").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 430,
			width:850,
			modal: true,
			buttons: {
				'Aceptar': function() {
						if($("#frmAsignar").valid())
						{
								//window.open($('#frmAsignar').attr('action')+'?'+$('#frmAsignar').serialize())
								$('#sendMsn').html('Espere Mientras se Envia Confirmación.');
								$("#sendConf").dialog('open');
								
								$.post($('#frmAsignar').attr('action')+'?'+$('#frmAsignar').serialize(),function(d){
																							$('#txtpreinscripcion').val("");
																							$("#list1").trigger("reloadGrid"); 
																							$("#sendConf").dialog('close');
																							$('#sendMsn').html('');
																								   });
								$(this).dialog('close');								
								$("#list1").trigger("reloadGrid"); 
						}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			},
			title: 'Asignar Participantes'
		});
	
	$("#dialog_Rechazo").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 300,
			width: 490,
			modal: true,
			buttons: {
				'Aceptar': function() {
						if($("#frmRechAuto").valid())
						{
								//window.open($('#frmElimAuto').attr('action')+'?'+$('#frmElimAuto').serialize())
								$('#sendMsn').html('Espere Mientras se Envia Notificación.');
								$("#sendConf").dialog('open');
								
								$.post($('#frmRechAuto').attr('action')+'?'+$('#frmRechAuto').serialize(),function(d){
																							$("#sendConf").dialog('close');
																							$('#sendMsn').html('');									
																							$("#list1").trigger("reloadGrid"); 
																								   });
								$(this).dialog('close');								
								$("#list1").trigger("reloadGrid"); 
						}
						
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 550,
			width: 1000,
			modal: true,
			buttons: {
				'Guardar': function() {
					if($("#terminos").attr("checked")) 
					{
						/*$.post("inspendientes/correoRechazo.asp",{prog:id_prog,id:id_preinscripcion},function(d){
			   																		$("#list1").trigger("reloadGrid"); 
																				 });
						
						$("#list1").trigger("reloadGrid");*/
						 $.post("inspendientes/frmRechAuto.asp",
							   {id:$('#txtpreinscripcion').val()},
							   function(f){
									$('#dialog_Rechazo').html(f);
									$('#razonRech').focus();
									valFrmRech();
						});
						 
						$(this).dialog('close');
						$('#dialog_Rechazo').dialog('open');
					}
					else
					{
						if($("#frmAutorizacion").valid())
						{
							   $.post("inspendientes/frmAsignar.asp",
	  {prog:id_prog,id:id_preinscripcion,con_otic:$('#ConOtic').val(),con_fran:$('#Confran').val(),reg_sence:$('#Reg_Fran').val(),id_proyecto:$('#id_proyecto').val()},
							   function(f){
									$('#cursos').html(f);
									tableToGrid("#mytable"); 
									//$("#Sede").append("<option value=\"\">Seleccione</option>");
									//llena_relator(0);
									//total_trabajadores();
									validar();
									//tabla(0);
								});
							
							$('#cursos').dialog('open');
							$(this).dialog('close');
						}
					}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	

    function tabla(sedeId)
	{
		$('#totalInsAsig').val("0")
		total_inscritos();
		ins_tot=parseInt($('#totalTrabAsig').val())-parseInt(ingresados);
		sUrl='inspendientes/listarIncritos.asp?prog='+$("#progId").val()+"&sede="+sedeId+"&relator="+$("#Relator").val();
		if(cargada){
			jQuery("#list3").jqGrid('setGridParam',{url:sUrl}).trigger("reloadGrid");
			jQuery("#list3").trigger("reloadGrid"); 
		}else{
			 cargada=true;
   			 jQuery("#list3").jqGrid({ 
				scroll: 1, 
				url:sUrl, 
				datatype: "xml", 
				colNames:['Rut', 'Nombre, Apellido Paterno, Apellido Materno', 'Cargo En La Empresa', 'Escolaridad'], 
				colModel:[
						{name:'rut',index:'T.rut', width:40}, 
						{name:'nombre',index:'T.nombre'}, 
						{name:'cargo',index:'T.cargo', width:80}, 
						{name:'escolaridad',index:'T.escolaridad', width:80}], 
					rowNum:30, 
					autowidth: true, 
					rownumbers: true, 
					pager: jQuery('#pager3'),
					sortname: 'rut',
					viewrecords: true,
					sortorder: "asc",
				    caption:"Nomina de Participantes Inscritos"
				}).navGrid('#pager3',{edit:false,add:false,del:false,search:false,refresh:false});
		  }
		  validar();
	}

	function validar(){
		$("#frmAsignar").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			bloqueSel:{
				required:true
			}
		},
		messages:{
			bloqueSel:{
				required:"&bull; Seleccione Bloque"
			}
		}
	});
    }

	function verifica_sede(relator){
		$.get("inspendientes/verificaSedes.asp?prog="+id_prog+"&relator="+relator,
					function(xml){
						$('row',xml).each(function(i) { 
							if($(this).find('totalSede').text()>0)
								llena_sede(relator);
							else
								llena_sede_disponibles();
						});
					});
	}

	function llena_relator(id_relator){
		$("#Relator").html("");
		$("#Relator").append("<option value=\"\">Seleccione</option>");
		$.get("inspendientes/instructor.asp?instructor="+id_prog+"&id_relator="+elimina_instructor,
					function(xml){
						$('row',xml).each(function(i) { 
							if(id_relator==$(this).find('ID_INSTRUCTOR').text())
								$("#Relator").append('<option value="'+$(this).find('ID_INSTRUCTOR').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$("#Relator").append('<option value="'+$(this).find('ID_INSTRUCTOR').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
					});
	}
	
	function llena_sede(instructor){
		$("#Sede").html("");
		$.get("inspendientes/sedes.asp?instructor="+instructor,
					function(xml){
						$("#Sede").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							$("#Sede").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																		$(this).find('NOMBRE').text()+ '</option>');
							 tabla($(this).find('ID').text());
						});
					});
		$('#Sede').attr("disabled", true);
	}
	
	function llena_sede_disponibles(){
		$("#Sede").html("");
		$("#Sede").append("<option value=\"\">Seleccione</option>");
		$.get("inspendientes/sedesDisponibles.asp?prog="+id_prog,
					function(xml){
						$('row',xml).each(function(i) { 
							$("#Sede").append('<option value="'+$(this).find('ID').text()+'" >'+
																		$(this).find('NOMBRE').text()+ '</option>');
						});
					});
		$('#Sede').removeAttr("disabled"); 
		tabla(0);
	}
	
	function total_trabajadores(){
		$.get("inspendientes/totalTrabajadores.asp?preinscripcion="+$('#txtpreinscripcionAsig').val(),
					function(xml){
						$('row',xml).each(function(i) { 
								$('#totalTrabAsig').val($(this).find('total').text());
						});
					});
	}
	
	function total_inscritos(){
		$.get("inspendientes/totalInscritos.asp?programa="+$('#progId').val()+"&sede="+$('#Sede').val()+"&relator="+$('#Relator').val(),
					function(xml){
						$('row',xml).each(function(i) { 
								$('#totalInsAsig').val($(this).find('total').text());
						});
					});
	}
	
	
	function MostrarInscritos(bloque){
		$.post("inspendientes/frmMostrar.asp",
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
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}
	
	function update(i){
		id_preinscripcion=i;
		$.post("inspendientes/frmautorizacion.asp",
			   {id:i},
			   function(f){
				    $('#dialog').html(f);
					id_prog=$('#Programa').val();
					if($('#Confran').val()=="1")
					{
						$('#Oticpreg').show();
						$('#pregOtic').show();
						$('#Regpreg').show();
						$('#RegFran').show();
					}
					else
					{
						$('#Oticpreg').hide();
						$('#pregOtic').hide();
						$('#Regpreg').hide();
						$('#RegFran').hide();
					}
					jQuery("#list2").jqGrid({ 
								scroll: 1, 					
								url:'inspendientes/listar2.asp?preins='+i, 
								datatype: "xml", 
							  colNames:['Rut', 'Nombre, Apellido Paterno, Apellido Materno', 'Cargo En La Empresa', 'Escolaridad'],
								colModel:[
										{name:'rut',index:'rut', width:40}, 
										{name:'nombre',index:'nombre'}, 
										{name:'cargo',index:'cargo', width:80}, 
										{name:'escolaridad',index:'escolaridad', width:80}], 
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
					cargada=false;
					sUrl="";
					ins_tot=1;
					ingresados=0;
					amtval="";
					comienzo=0;
					elimina_instructor="";
					//id_preinscripcion=0;
		});
		 $('#dialog').dialog('open');
	}
	
	function documento(arch){
		$("#ifPagina").attr('src',"http://norte.otecmutual.cl/ordenes/"+arch);
		if(!$('#Doc').dialog('isOpen'))
		{
			$('#Doc').dialog('open');
		}
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
	
	function checkear(seleccionado){
			
			for (i=1; i<=$('#countFilas').val(); i++)
			{
				$('#check'+i).attr('checked', false);	
			}
			
			$('#check'+seleccionado).attr('checked', true);
			$('#bloqueSel').val($('#Bloque'+seleccionado).val());
			$('#Relator').val($('#relId'+seleccionado).val());
			$('#Sede').val($('#Sala'+seleccionado).val());
   	   }
	
	function eliminar(i){
	  $.post("inspendientes/frmElimAuto.asp",
			   {id:i},
			   function(f){
				    $('#dialog_el').html(f);
					$('#razonElim').focus();
					valFrmElim();
		});
	    $('#dialog_el').dialog('open');
	}
	
	function valFrmElim(){
		$("#frmElimAuto").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			razonElim:{
				required:true
			}
		},
		messages:{
			razonElim:{
				required:"&bull; Ingrese las Razones por las que se Elimina la Solicitud."
			}
		}
	});
    }

	function valFrmRech(){
		$("#frmRechAuto").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			razonRech:{
				required:true
			}
		},
		messages:{
			razonRech:{
				required:"&bull; Ingrese las Razones por las que se Rechaza la Solicitud."
			}
		}
	});
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
			<h2><em style="text-transform: capitalize;">Autorizar Inscripciones</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Autorización de Inscripción">
</div>
<div id="cursos" title="Asignar Participantes">
</div>
<div id="InsBloque" title="Inscritos en Bloque">
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="mensaje" title="Asignar Participantes">
     <label id="txtMensaje" name="txtMensaje"></label>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="sendConf" title="Enviando">
     <p align="center"><label id="sendMsn" name="sendMsn"></label></br></br><img src="images/loadfbk.gif"/></p>
</div>
<div id="dialog_el" title="Eliminar Solicitud de Inscripción a Curso">
</div>
<div id="dialog_Rechazo" title="Rechazar Solicitud de Inscripción a Curso">
</div>
</body>
</html>
