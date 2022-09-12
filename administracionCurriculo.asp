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

$(document).ready(function(){					
//window.open("curriculo/listar.asp")
	jQuery("#list1").jqGrid({ 
		url:'curriculo/listar.asp', 
		datatype: "xml", 
		colNames:['Código', 'Sence', 'Nombre','&nbsp;', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'CODIGO',index:'CODIGO', width:30, align:'center'}, 
				   {name:'SENCE',index:'SENCE', width:13, align:'center'}, 
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO'}, 
				   { align:"right",width:8}, 				   
				   { align:"right",width:8}, 
				   { align:"right",width:7} ], 
		rowNum:30, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager1'), 
		sortname: 'NOMBRE_CURSO', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Portafolios" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("curriculo/frmcurriculo.asp",
															   {id:0},
															   function(f){
																  $('#dialog').html(f);
																  //$('#txtFecha').datepicker();
																  $('#txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
																  llena_vigencia(0);
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
	
	
	$("#dialogDetAct").dialog({
			bgiframe: true,
			autoOpen: false,
			height:490,
			width: 860,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmActividad").valid())
						{
							$.post($('#frmActividad').attr('action')+'?'+$('#frmActividad').serialize(),function(d){
			   																		$("#list2").trigger("reloadGrid"); 
																					});
							$("#list2").trigger("reloadGrid");
							$(this).dialog('close');
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
			height:490,
			width: 860,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmCurriculo").valid())
						{
							//window.open($('#frmCurriculo').attr('action')+'?'+$('#frmCurriculo').serialize());
			
							$.post($('#frmCurriculo').attr('action')+'?'+$('#frmCurriculo').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
																						   });
							$("#list1").trigger("reloadGrid");
							//document.getElementById("frmOtic").reset();
							$(this).dialog('close');
						}
				},
				Cancelar: function() {
					//document.getElementById("frmOtic").reset();
					$(this).dialog('close');
				}
			}
		});
		
		$("#dialogAct").dialog({
			bgiframe: true,
			autoOpen: false,
			height:490,
			width: 860,
			modal: true,
			buttons: {
				Cerrar: function() {
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
						$.post("curriculo/eliminar.asp",{id:emp_del},function(d){
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
	});	

	function eliminar(i){
	 emp_del=i;
	 $('#dialog_el').dialog('open');
	}

	function eliminarAct(i,b){
		$.post("curriculo/eliminarAct.asp",{id:i,b:b},function(d){
											$("#list2").trigger("reloadGrid");
																});
	}

	function validarAct(){
		$("#frmActividad").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			actDia:{
				required:true
			},
			actBloque:{
				required:true
			},
			actHoras:{
				required:true
			}
		},
		messages:{
			actDia:{
				required:"&bull; Seleccione Dia."
			},
			actBloque:{
				required:"&bull; Seleccione Bloque."
			},
			actHoras:{
				required:"&bull; Seleccione Hora."
			}
		}
	});
    }

	function validar(){
		//$('#frmCurriculo txtFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		$("#frmCurriculo").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txtNom:{
				required:true
			},
			txtDesc:{
				required:true
			},
			txtObj:{
				required:true
			},
			txtAud:{
				required:true
			},
			txtHor:{
				required:true,
				number:true
			},
			txtVig:{
				required:true
			},
			txtCod:{
				required:true
			},
			txtValor:{
				required:true,
				number:true
			},
			txtValAfiliados:{
				required:true,
				number:true
			}
		},
		messages:{
			txtNom:{
				required:"&bull; Ingrese Nombre Curso."
			},
			txtDesc:{
				required:"&bull; Ingrese Descripción Curso."
			},
			txtObj:{
				required:"&bull; Ingrese Objetivo Curso."
			},
			txtAud:{
				required:"&bull; Ingrese Audiencia Curso."
			},
			txtHor:{
				required:"&bull; Ingrese Horario Curso.",
				number:"&bull; Ingrese Solo Números en el Número de Horas."
			},
			txtVig:{
				required:"&bull; Seleccione Vigencia Curso."
			},
			txtCod:{
				required:"&bull; Ingrese Código Curso."
			},
			txtValor:{
				required:"&bull; Ingrese Valor del Curso.",
				number:"&bull; Ingrese Solo Números en Valor del Curso."
			},
			txtValAfiliados:{
				required:"&bull; Ingrese Valor Especial del Curso para Afiliados a Mutual.",
				number:"&bull; Ingrese Solo Números en Valor Especial del Curso para Afiliados a Mutual."
			}
		}
	});
    }
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){ 
	// NOTE: Backspace = 8, Enter = 13, '0' = 48, '9' = 57 
	var key = nav4 ? evt.which : evt.keyCode; 
	return (key <= 13 || (key >= 48 && key <= 57));
	}
	
	function llena_vigencia(vigencia){
		$("#txtVig").html("");
		$("#txtVig").append("<option value=\"\">Seleccione</option>");
		if(vigencia==1){
			$("#txtVig").append("<option value=\"1\" selected=\"selected\">6 Meses</option>");
			$("#txtVig").append("<option value=\"2\">12 Meses</option>");
			$("#txtVig").append("<option value=\"3\">18 Meses</option>");
			$("#txtVig").append("<option value=\"4\">24 Meses</option>");
			$("#txtVig").append("<option value=\"5\">48 Meses</option>");
		}else if(vigencia==2){
			$("#txtVig").append("<option value=\"1\">6 Meses</option>");
			$("#txtVig").append("<option value=\"2\" selected=\"selected\">12 Meses</option>");
			$("#txtVig").append("<option value=\"3\">18 Meses</option>");
			$("#txtVig").append("<option value=\"4\">24 Meses</option>");
			$("#txtVig").append("<option value=\"5\">48 Meses</option>");
		}else if(vigencia==3){
			$("#txtVig").append("<option value=\"1\">6 Meses</option>");
			$("#txtVig").append("<option value=\"2\">12 Meses</option>");
			$("#txtVig").append("<option value=\"3\" selected=\"selected\">18 Meses</option>");
			$("#txtVig").append("<option value=\"4\">24 Meses</option>");
			$("#txtVig").append("<option value=\"5\">48 Meses</option>");
		}else if(vigencia==4){
			$("#txtVig").append("<option value=\"1\">6 Meses</option>");
			$("#txtVig").append("<option value=\"2\">12 Meses</option>");
			$("#txtVig").append("<option value=\"3\">18 Meses</option>");
			$("#txtVig").append("<option value=\"4\" selected=\"selected\">24 Meses</option>");
			$("#txtVig").append("<option value=\"5\">48 Meses</option>");
		}else if(vigencia==5){
			$("#txtVig").append("<option value=\"1\">6 Meses</option>");
			$("#txtVig").append("<option value=\"2\">12 Meses</option>");
			$("#txtVig").append("<option value=\"3\">18 Meses</option>");
			$("#txtVig").append("<option value=\"4\">24 Meses</option>");
			$("#txtVig").append("<option value=\"5\" selected=\"selected\">48 Meses</option>");
		}else{
			$("#txtVig").append("<option value=\"1\">6 Meses</option>");
			$("#txtVig").append("<option value=\"2\">12 Meses</option>");
			$("#txtVig").append("<option value=\"3\">18 Meses</option>");
			$("#txtVig").append("<option value=\"4\">24 Meses</option>");
			$("#txtVig").append("<option value=\"5\">48 Meses</option>");
		}
	}
	
	function update(i){
		//document.getElementById("frmEmpresa").reset();
		$.post("curriculo/frmcurriculo.asp",
			   {id:i},
			   function(f){
				   validar();
				   $('#dialog').html(f);
				   llena_vigencia($('#txtVigId').val());
				   validar();
		});
		 $('#dialog').dialog('open');
	}
	
	function updateAct(i, id_mutual){
		$.post("curriculo/frmActividad.asp",{id:i,id_mutual:id_mutual},
			   function(f){
				   $('#dialogDetAct').html(f);
				   listaCombos('actDia', 1, $('#txtactDiaId').val());
				   listaCombos('actBloque', 2, $('#txtactBloqueId').val());
				   listaCombos('actHoras', 3, $('#txtactHorasId').val());
				   validarAct();
		});
		 $('#dialogDetAct').dialog('open');
	}
	
	function listaCombos(pre, tipo, id){
		$('#' + pre).html("");
		$.get("curriculo/listaCombos.asp?t="+tipo,
					function(xml){
						$('#' + pre).append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id==$(this).find('ID').text())
								$('#' + pre).append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$('#' + pre).append('<option value="'+$(this).find('ID').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
					});
	}	
	
	function detActividades(i){
		$.post("curriculo/frmActividades.asp",{id:i},
			   function(f){
				   $('#dialogAct').html(f);
							jQuery("#list2").jqGrid({ 
							url:'curriculo/listar2.asp?id='+i,
								datatype: "xml", 
							    colNames:['Dia', 'Bloque', 'Tema', 'Actividad', 'Horas', '&nbsp;', '&nbsp;'], 
								colModel:[
										{name:'DIA',index:'DIA', width:25}, 
										{name:'NOMBRE_BLOQUE',index:'NOMBRE_BLOQUE', width:35}, 										
										{name:'TEMA',index:'TEMA'}, 
										{name:'ACTIVIDAD',index:'ACTIVIDAD'}, 
										{name:'horas',index:'horas', width:25},
				   						{align:"center",width:15}, 
				   						{align:"center",width:15}], 
								rowNum:100, 
								rownumbers: true, 
								rownumWidth: 20, 
								height:220,
								width:800,
								pager: jQuery('#pager2'), 
								sortname: 'cb.NOMBRE_BLOQUE', 
								viewrecords: true, 
								sortorder: "asc", 
								caption:"Detalle Actividades"
								}); 				
								
								jQuery("#list2").jqGrid('navGrid','#pager2',{edit:false,add:false,del:false,search:false,refresh:false});
								jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
													title:"Agregar Nueva Actividad", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("curriculo/frmActividad.asp",{id:0},
															   function(f){
																  $('#dialogDetAct').html(f);
																  listaCombos('actDia', 1, 0);
				   												  listaCombos('actBloque', 2, 0);
				   												  listaCombos('actHoras', 3, 0);
																  $('#txtIdMutual').val(i);
																  validarAct();
														});
														$('#dialogDetAct').dialog('open');
													} }); 

								jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list2").trigger("reloadGrid");
													  } }); 								   
		});
		
		$('#dialogAct').dialog('open');
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
			<h2><em style="text-transform: capitalize;">Portafolio</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro de Portafolio">
</div>
<div id="dialogAct" title="Registro de Actividades">
</div>
<div id="dialogDetAct" title="Registro de Actividad">
</div>
<div id="dialog_el" title="Eliminar Portafolio">
	<p>Esta seguro de eliminar el Portafolio.</p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>