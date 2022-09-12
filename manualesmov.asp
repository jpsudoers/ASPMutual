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
<link rel="stylesheet" type="text/css" href="css/ui.all.css"/>
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
<script src="js/jquery.jqGrid.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
$(document).ready(function(){					
	jQuery("#list1").jqGrid({ 
		url:'manualesmov/listar.asp', 
		datatype: "xml", 
		colNames:['Fecha','Tipo','Articulo','Bodega','Cant.'], 
		colModel:[
				   {name:'FECHA',index:'M.FECHA', width:45, align:'center'}, 
				   {name:'TIPO_MOVIMIENTO',index:'M.TIPO_MOVIMIENTO', width:45},				   
				   {name:'DESC_ARTICULO',index:'A.DESC_ARTICULO'}, 
				   {name:'UBICACION',index:'B.UBICACION'},
				   {name:'CANTIDAD',index:'M.CANTIDAD', width:20, align:'center'}], 
		rowNum:30, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager1'), 
		sortname: 'M.FECHA', 
		viewrecords: true, 
		sortorder: "desc", 
		caption:"Listado de Ingresos" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													title:"Agregar nuevo registro", 
													buttonicon :'ui-icon-plus', 
													onClickButton:function(){
														$.post("manualesmov/frmmanualesmov.asp",
															   {id:0},
															   function(f){
																  $('#dialog').html(f);
																  llena_articulos(0,0);
																  llena_bodegas(0);
																  llena_tipo_mov(0);
											                      $('#MovFecha').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy'});
																  $('#MovFechaDoc').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy'});
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
			height:445,
			width: 700,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmmanualesmov").valid())
						{
							//window.open($('#frmmanualesmov').attr('action')+'?'+$('#frmmanualesmov').serialize())
							$.post($('#frmmanualesmov').attr('action')+'?'+$('#frmmanualesmov').serialize(),function(d){
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

	function llena_articulos(id_articulo,id_bodega){
		$("#MovArt").html("");
		$.get("manualesmov/articulos.asp",{id_bodega:id_bodega},
					function(xml){
						$("#MovArt").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_articulo==$(this).find('ID_ARTICULO').text())
								$("#MovArt").append('<option value="'+$(this).find('ID_ARTICULO').text()+'" selected="selected">'+
																		$(this).find('DESC_ARTICULO').text()+ '</option>');
							else
								$("#MovArt").append('<option value="'+$(this).find('ID_ARTICULO').text()+'" >'+
																		$(this).find('DESC_ARTICULO').text()+ '</option>');
						});
					});
	}
	
	function llena_tipo_mov(id_movimiento){
		$("#MovTipo").html("");
		$.get("manualesmov/tipo_mov.asp",
					function(xml){
						$("#MovTipo").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_movimiento==$(this).find('ID_TIPO_MOV').text())
								$("#MovTipo").append('<option value="'+$(this).find('ID_TIPO_MOV').text()+'" selected="selected">'+
																		$(this).find('NOMBRE_MOV').text()+ '</option>');
							else
								$("#MovTipo").append('<option value="'+$(this).find('ID_TIPO_MOV').text()+'" >'+
																		$(this).find('NOMBRE_MOV').text()+ '</option>');
						});
					});
	}
	
	function llena_bodegas(id_bodega){
		$("#MovBdg").html("");
		$.get("manualesmov/bodegas.asp",
					function(xml){
						$("#MovBdg").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_bodega==$(this).find('ID_BODEGA').text())
								$("#MovBdg").append('<option value="'+$(this).find('ID_BODEGA').text()+'" selected="selected">'+
																		$(this).find('BODEGA').text()+ '</option>');
							else
								$("#MovBdg").append('<option value="'+$(this).find('ID_BODEGA').text()+'" >'+
																		$(this).find('BODEGA').text()+ '</option>');
						});
					});
	}

	function validar(){
		$("#frmmanualesmov").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			MovBdg:{
				required:true
			},
			MovArt:{
				required:true
			},
			MovTipo:{
				required:true
			},
			MovTipoDoc:{
				required:true
			},			
			MovCantidad:{
				required:true,
			    digits: true
			},
			MovPrecio:{
				required:true,
			    number: true
			},
			MovFecha:{
				required:true
			},
			MovFechaDoc:{
				required:true
			}
		},
		messages:{
			MovBdg:{
				required:"&bull; Seleccione Bodega."
			},
			MovArt:{
				required:"&bull; Seleccione Articulo."
			},
			MovTipo:{
				required:"&bull; Seleccione el Tipo de Movimiento."
			},
			MovTipoDoc:{
				required:"&bull; Seleccione el Tipo de Documento."
			},				
			MovCantidad:{
				required:"&bull; Ingrese la Cantidad de Articulos.",
			    digits:"&bull; Ingrese Solo Números en el Campo Cantidad."
			},
			MovPrecio:{
				required:"&bull; Ingrese el Precio Unitario Sin Iva.",
			    number:"&bull; Ingrese Solo Números en el Campo Precio Unitario."
			},
			MovFecha:{
				required:"&bull; Seleccione la Fecha de Ingreso de los Articulos."
			},
			MovFechaDoc:{
				required:"&bull; Seleccione la Fecha del Documento."
			}
		}
	});
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
			<h2><em style="text-transform: capitalize;">Ingresos</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro de Ingreso">
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>