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
var art_del;

$(document).ready(function(){					
	jQuery("#list1").jqGrid({ 
		url:'bodegas/listar.asp', 
		datatype: "xml", 
		colNames:['Código','Ubicación','Dirección','Responsable', '&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'ID_BODEGA',index:'B.ID_BODEGA', width:40, align:'center'},
				   {name:'UBICACION',index:'B.UBICACION'}, 
				   {name:'DIRECCION',index:'B.DIRECCION'}, 
				   {name:'NOMBRES',index:'U.NOMBRES'},
				   { align:"right",editable:true, width:15}, 
				   { align:"right",editable:true, width:13} ], 
		rowNum:30, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager1'), 
		sortname: 'B.ID_BODEGA', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Bodegas" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("bodegas/frmbodegas.asp",
															   {id:0},
															   function(f){
																  $('#dialog').html(f);
																  llena_responsable(0);
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
			height:310,
			width: 470,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmbodegas").valid())
						{
							$.post($('#frmbodegas').attr('action')+'?'+$('#frmbodegas').serialize(),function(d){
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
			width: 400,
			modal: true,
			buttons: {
				'Aceptar': function() {
							$.post("bodegas/eliminar.asp",{id:art_del},function(f){
																			$("#list1").trigger("reloadGrid");
																			});
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

	function llena_responsable(id_responsable){
		$("#Responsable").html("");
		$.get("bodegas/responsable.asp",
					function(xml){
						$("#Responsable").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_responsable==$(this).find('ID_USUARIO').text())
							 $("#Responsable").append('<option value="'+$(this).find('ID_USUARIO').text()+'" selected="selected">'+
																		$(this).find('USUARIO').text()+ '</option>');
							else
								$("#Responsable").append('<option value="'+$(this).find('ID_USUARIO').text()+'" >'+
																		$(this).find('USUARIO').text()+ '</option>');
						});
					});
	}

	function eliminar(i){
	 art_del=i;
	 $('#dialog_el').dialog('open');
	}

	function validar(){
		$("#frmbodegas").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			UbiBdg:{
				required:true
			},
			DirBdg:{
				required:true
			},
			ResBdg:{
				required:true
			}
		},
		messages:{
			UbiBdg:{
				required:"&bull; Ingrese la Ubicación de la Bodega."
			},
			DirBdg:{
				required:"&bull; Ingrese la Dirección de la Bodega."
			},
			ResBdg:{
				required:"&bull; Ingrese el Responsable de la Bodega."
			}
		}
	});
    }

	function update(i){
		$.post("bodegas/frmbodegas.asp",
			   {id:i},
			   function(f){
				   $('#dialog').html(f);
				   llena_responsable($('#IdResp').val());
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
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO6,PERMISO12,PERMISO13,PERMISO14,PERMISO15 from USUARIOS "
	sql = sql&" where USUARIOS.ID_USUARIO='"&Session("usuarioMutual")&"'"

   DATOS.Open sql,oConn
		if(DATOS("PERMISO1")<>"0")then
		%>
		<li><a href="administracion.asp" accesskey="1" title="">Administraci&oacute;n</a></li>
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
		<li><a href="manuales.asp" accesskey="6" title="" class="selItem">Manuales</a></li>
        <%end if	
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
            	<%
				dim totOpciones				
				Dim OPCIONES
				OPCIONES=0
				
				if(DATOS("PERMISO13")<>"0")then
				%>
				<li class="first"><a href="manualesbodega.asp">Bodegas</a></li>
                <li><a href="manualesarticulos.asp">Articulos</a></li>
                <%else
						OPCIONES=OPCIONES+2
				end if
                if(DATOS("PERMISO14")<>"0")then
		        %>
                <li><a href="manualesmov.asp">Ingresos</a></li>
                <%else
						OPCIONES=OPCIONES+1
				end if				
                if(DATOS("PERMISO15")<>"0")then
		        %>
                <li><a href="manualesegresos.asp">Salidas</a></li>
                <%else
						OPCIONES=OPCIONES+1
				end if				
                if(DATOS("PERMISO13")<>"0")then
		        %>    
                <li><a href="manualesajustes.asp">Ajustes</a></li>  
                <li><a href="manualinformes.asp">Informes de Movimientos</a></li> 
                <li><a href="manualexistencias.asp">Informes de Existencias</a></li>   
                <%else
						OPCIONES=OPCIONES+2
				end if	
				
									
				if(OPCIONES>0)then
					For totOpciones = 1 To OPCIONES Step 1
				    %>
						<p>&nbsp;</p>
				    <%
					Next
				end if			
		        %> 
			</ul>
		</div>		
	</div>
		<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Bodegas</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro de Bodega">
</div>
<div id="dialog_el" title="Eliminar Bodega">
	<p>Esta seguro de eliminar la Bodega.</p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>