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
		height:200px;
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
</style>
<script type="text/javascript">
var i=1;
var cargada=false;
$(document).ready(function(){					
	$('#txRut').defaultValue('Todos');
	tabla();
	
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
	
	$("#Doc").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 500,
			width:900,
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
			title: 'Certificado de Asistencia a Curso de Capacitación'
		});
	});	
	
	function certificados2(programa,empresa,trabajador,relator){
		documento("libroclases/Certificado.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);
	}
	
	function certificados(programa,empresa,trabajador,relator,curso){
		/*if(curso=='82'){
			documento("libroclases/CertificadoCodelco.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);	
		}
		else*/ if(curso=='52'){
			documento("libroclases/CertificadoCMCC.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);	
		}		
		else{
			documento("libroclases/Certificado.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);
		}
	}
	
	
	function documento(arch){
		$("#ifPagina").attr('src',arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
	}
	
	function lookup(inputString) {
		if(inputString.length <=2) {
			$('#suggestions').hide();
			$("#txtNombre").html("");
		    $("#txtTrabajador").html("");
			$('#Trabajador').val("");
			if(inputString.length <=0)
			{
				tabla();
			}
		} else {
				$.post("empcertificados/sugTrabajador.asp", {txt: inputString}, function(data){
						if(data.length >0) {
								$('#suggestions').show();
								$('#autoSuggestionsList').html(data);
						}
				});
		}
	}

	function fill(id,rut) {
	   $('#Trabajador').val(id);
	   cargaDatosTrabajador(id);
	   $('#suggestions').hide();
	}

	function cargaDatosTrabajador(id)	
	{
		$("#txRut").val("");
		$("#txtNombre").html("");
		$("#txtTrabajador").html("");
 
		$.get("bhp/datosTrabajador.asp",
						 {id:id},
						 function(xml){
									$('row',xml).each(function(i) { 
									$("#txtNombre").html("Nombre :");
									$("#txtTrabajador").html($(this).find('NOMBRE').text());
									$("#txRut").val($(this).find('RUT').text());
									tabla();
						});
			});
	}

	function tabla()
	{
		//{name:'vigencia',index:'vigencia', width:20, align:"center"}, 	
		if(!cargada){
			jQuery("#list1").jqGrid({ 
			url:'empcertificados/listar.asp?trabajador=0', 
			datatype: "xml", 
			colNames:['Rut/ID','Nombre Trabajador', 'Nombre del Curso', 'Fecha', '&nbsp;'], 
			colModel:[
					   {name:'RUT',index:'T.RUT', width:22},
					   {name:'NOMBRES',index:'T.NOMBRES', width:60}, 
					   {name:'NOMBRE_CURSO',index:'C.NOMBRE_CURSO', width:60}, 
					   {name:'FECHA_TERMINO',index:'P.FECHA_TERMINO', width:20, align:"center"}, 
					   { align:"right",editable:true, width:4}], 
			rowNum:100, 
			autowidth: true, 
			rowList:[100,150,200], 
			pager: jQuery('#pager1'), 
			sortname: 'T.NOMBRES', 
			viewrecords: true, 
			sortorder: "asc", 
			caption:"Listado de Certificados" 
			}); 
	
			jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
			jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
															  title:"Refrescar", 
															  buttonicon :'ui-icon-refresh', 
															  onClickButton:function(){ 
															  $('#Trabajador').val("");
															  $('#txRut').val("");
															  $('#txtTrabajador').html("");
															  $("#txtNombre").html("");
															  $('#txRut').defaultValue('Todos');	
jQuery("#list1").jqGrid('setGridParam',{url:"empcertificados/listar.asp?trabajador="+$('#Trabajador').val()}).trigger("reloadGrid")
															  } }); 
															  
			jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"Excel",
															  title:"Exportar a Excel", 
															  buttonicon :'ui-icon-script',
															  onClickButton:function(){ 
															  	if($('#Empresa').val()!=''){
												 			   		window.open("empcertificados/xls2.asp?e="+$('#Empresa').val()+"&t="+$('#Trabajador').val(),'Informe');
															  	}
															  } }); 
											
			if(<%=Session("usuario")%> == '17765'){											  
					jQuery("#list1").jqGrid('navButtonAdd','#pager1',{
						   caption:"Exportar Histórico", 
						   buttonicon :'ui-icon-script',
						   onClickButton : function () { 
									window.open("empcertificados/xls.asp","Reporte");
						   } 
					});			
			}
			cargada=true;
		}else{
jQuery("#list1").jqGrid('setGridParam',{url:"empcertificados/listar.asp?trabajador="+$('#Trabajador').val()}).trigger("reloadGrid")
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
		$.post("cambioContrasena/frmContrasenaEmpresa.asp",
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
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO5,PERMISO6 from USUARIOS "
	sql = sql&" where USUARIOS.ID_USUARIO='"&Session("usuarioMutual")&"'"

   DATOS.Open sql,oConn
   WHILE NOT DATOS.EOF
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
	 DATOS.MoveNext
   WEND
   end if
   if(Session("tipoUsuario")="0")then%>
		<li><a href="empresas.asp" accesskey="5" title="" class="selItem">Empresas</a></li>
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
				Usuario : <strong><%=Session("nombre")%> (<%=Session("correo_user")%>)</strong>
      <br />
      <a href="#" onclick="passChange();"><b>Cambiar Contraseña</b></a>
      <br />
      <br />
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
<li><a href="solicitudescreditoEmpresa.asp">Solicitud Crédito</a></li>
<li><a href="empresasOC.asp">Ingreso de Orden de Compra</a></li>
               <%else
				   if(Session("cargo_user_empresa")="1")then%>
					   <li class="first"><a href="empresacalendario.asp">Inscripción de Cursos</a></li>
					   <li><a href="empresasinspendientes.asp">Inscripciones Pendientes</a></li>
					   <li><a href="empresainsactivas.asp">Inscripciones Autorizadas</a></li>
					   <li><a href="empresascertificados.asp">Certificados</a></li>
<li><a href="solicitudescreditoEmpresa.asp">Solicitud Crédito</a></li>
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
			<h2><em style="text-transform: capitalize;">Certificados</em></h2>
            <br />
            <table border="0">
              <tr>
                <td width="100">Rut Trabajador :</td>
                <td width="100"><input id="txRut" name="txRut" type="text" tabindex="1" maxlength="20" size="20" onkeyup="lookup(this.value);"/><div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:1;left:530px">
                        <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow"/>
                        <div class="suggestionList" id="autoSuggestionsList">&nbsp;</div>
                   </div><input type="HIDDEN" id="Trabajador" name="Trabajador"/>
				   <input type="hidden" id="Empresa" name="Empresa" value="<%=Session("usuario")%>"></input>
				   </td>
                   <td width="80" align="right"><label id="txtNombre" name="txtNombre"></label></td>
                   <td width="300">&nbsp;<label id="txtTrabajador" name="txtTrabajador"></label></td>
              </tr>
            </table>
            <br />
	        <table id="list1"></table> 
            <div id="pager1"></div> 
            <br />
            <h2><label style="font-size:14px; text-transform: none;color: #31576F; font-style:italic"><B>Nota: Los Certificados desplegados en Rojo han Expirado.</B></label></h2>
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro de Evaluación y Cierre de Curso">
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>
