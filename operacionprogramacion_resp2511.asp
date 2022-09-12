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
var emp_del_tab;
$(document).ready(function(){					
//window.open("programacion/listar.asp");

	jQuery("#list1").jqGrid({ 
		url:'programacion/listar.asp', 
		datatype: "xml", 
		colNames:['Cód. Prog.','Código Curso', 'Nombre Curso', 'Cupos', 'Inscritos', 'Fecha Inicio', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'ID_PROGRAMA',index:'ID_PROGRAMA', width:50}, 
				   {name:'CODIGO',index:'CODIGO', width:62}, 
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO'}, 
				   {name:'CUPOS',index:'CUPOS', width:29, align:'center'}, 
				   {name:'INSCRITOS',index:'INSCRITOS', width:29, align:'center'}, 
				   {name:'FECHA_INICIO_',index:'FECHA_INICIO_', width:60}, 
				   { align:"right",editable:true, width:20}, 
				   { align:"right",editable:true, width:20} ], 
		rowNum:30, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager1'), 
		sortname: 'PROGRAMA.FECHA_INICIO_', 
		viewrecords: true, 
		sortorder: "desc", 
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
			height:560,
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
			height:300,
			width: 400,
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
			}
		},
		messages:{
			Curriculo:{
				required:"&bull; Seleccione Curriculo."
			},
			Tipo:{
				required:"&bull; Seleccione Tipo de Programa."
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
		colNames:['Rut', 'Relator', 'Sala/Sede', 'Cupos', 'Inscritos','&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'ID_PROGRAMA',index:'ID_PROGRAMA', width:50}, 
				   {name:'CODIGO',index:'CODIGO'}, 
				   {name:'SENCE',index:'SENCE'}, 
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO', width:30}, 
				   {name:'CUPOS',index:'CUPOS', width:30},
				   {align:"right",editable:true, width:20},
				   {align:"right",editable:true, width:20}], 
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
																 llena_sede(0);
																 llena_relator(0);
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
	
		function llena_relator(id_relator){
		$("#relator_frm").html("");
		$.get("programacion/instructor_frm.asp?idRelator="+id_relator+"&progId="+$("#tabFecha").val(),
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
                    llena_sede($('#tabSala').val());
					llena_relator($('#tabRelator').val());
					validarFrmTab();
		});
		 $('#dialogBloque').dialog('open');
	}
	
	function cerrar_bloque(i,bloque,estado){
		//alert(i+' '+bloque+' '+estado);
		$.post("programacion/cerrarBloque.asp",{prog:i,bloque:bloque,estado:estado},function(d){
			   																		$("#list2").trigger("reloadGrid"); 
																				 });
	}function validarFrm(){
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
			salaSede_frm:{
				 required:true
			},
			txtCupo_frm:{
				required:true,
				max: $('#totProgDisp').val(),
				min: 1
			}
		},
		messages:{
			relator_frm:{
				required:"&bull; Seleccione Relator."
			},
			salaSede_frm:{
				required:"&bull; Seleccione Sede."
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
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="dialogBloque" title="Registro de Bloque">
</div>
</body>
</html>
