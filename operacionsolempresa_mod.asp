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

	jQuery("#list1").jqGrid({ 
		url:'solempresa/listar.asp', 
		datatype: "xml", 
		colNames:['F. Solicitud', 'Rut', 'Razón Social', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'FECHA_SOLICITUD',index:'FECHA_SOLICITUD', width:25, align:'center'}, 				  
				   {name:'rut',index:'rut', width:25, align:'center'}, 
				   {name:'R_SOCIAL',index:'R_SOCIAL'}, 
				   { align:"right",editable:true, width:8}, 
				   { align:"right",editable:true, width:7} ], 
		rowNum:30, 
		height:350,
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager1'), 
		sortname: 'id_empresa', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Solicitudes" 
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
	
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 580,
			width: 1000,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmEmpresa").valid())
						{
							//window.open($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize());
							$('#sendMsn').html('Espere Mientras se Envian las Cuentas de Usuario.');
							$("#sendConf").dialog('open');
							
							$.post($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize(),function(d){
																							$("#sendConf").dialog('close');
																							$('#sendMsn').html('');									
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
		
		$("#dialog_el").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
						$('#sendMsn').html('Espere Mientras se Envia Notificación.');
						$("#sendConf").dialog('open');
						
						$.post("solempresa/eliminar.asp",{id:emp_del},function(d){
																				  $("#sendConf").dialog('close');
																				  $('#sendMsn').html('');	
																				  $("#list1").trigger("reloadGrid");
																				 });
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

	function estado_solicitud(){
	     if($("#rechazar").attr("checked")) 
		 {
			$("#rechazo").val("1");
		 }
		 else
		 {
			 $("#rechazo").val("0");
		 }
		 //alert($("#rechazo").val());
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
			},
			txtNomb:{
				required:true
			},
			txtMail:{
				required:true,
				email:true
			},
			txtCargo:{
				required:true
			},
			txtNombConta:{
				required:true
			},
			txtMailConta:{
				required:true,
				email:true
			},
			txtCargoConta:{
				required:true
			},
			txtFonoCont:{
				required:true
			},
			txtContaFono:{
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
			},
			txtNomb:{
				required:"&bull; Ingrese Nombre Contacto."
			},
			txtMail:{
				required:"&bull; Ingrese Correo Electrónico Contacto.",
				email:"&bull; Correo No Valido"
			},
			txtCargo:{
				required:"&bull; Ingrese Cargo Contacto."
			},
			txtNombConta:{
				required:"&bull; Ingrese Nombre Contacto Contabilidad."
			},
			txtMailConta:{
				required:"&bull; Ingrese Correo Electrónico Contacto Contabilidad.",
				email:"&bull; Correo No Valido"
			},
			txtCargoConta:{
				required:"&bull; Ingrese Cargo Contacto Contabilidad."
			},
			txtFonoCont:{
				required:"&bull; Ingrese Télefono Contacto Empresa."
			},
			txtContaFono:{
				required:"&bull; Ingrese Télefono Contacto Contabilidad."
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
		$.get("solempresa/mutual.asp",
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
		$.get("solempresa/otic.asp",
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
		$.post("solempresa/frmEmpresa.asp",
			   {id:i},
			   function(f){
				   validar();
				   $('#dialog').html(f);
				    llena_mutual($('#txtIdMut').val());
					llena_otic($('#txtIdOtic').val());
				   validar();
		});
		 $('#dialog').dialog('open');
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
			<h2><em style="text-transform: capitalize;">Solicitud de Nueva Empresa</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Empresa">
</div>
<div id="dialog_el" title="Eliminar Solicitud de Inscripción de Empresa">
	<p>Esta Seguro de Eliminar la Solicitud de Inscripción de Empresa.</p>
</div>
<div id="sendConf" title="Enviando">
     <p align="center"><label id="sendMsn" name="sendMsn"></label></br></br><img src="images/loadfbk.gif"/></p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>