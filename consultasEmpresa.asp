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
var cargada=false;
$(document).ready(function(){					
tabla();
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 470,
			width: 990,
			modal: true,
			buttons: {
				Cerrar: function() {
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

	function tabla()
	{
		sUrl='finanzasempresa/listar.asp?campo='+$('[name="rbtn_buscar"]:checked').val()+'&tipo='+$('[name="rbtn_tipo"]:checked').val()+'&texto='+$("#txt_buscar").val();
		if(cargada){
			jQuery("#list1").jqGrid('setGridParam',{url:sUrl});
			jQuery("#list1").trigger("reloadGrid"); 
		}else{
			 cargada=true;
   			 jQuery("#list1").jqGrid({ 
				url:sUrl, 
				datatype: "xml", 
				colNames:['Rut', 'Razón Social', '&nbsp;'], 
				colModel:[
						   {name:'rut',index:'rut', width:25, align:"right"}, 
						   {name:'r_social',index:'r_social'}, 
						   {align:"right",editable:true, width:10} ], 
				rowNum:30, 
				autowidth: true, 
				rowList:[30,50,100], 
				pager: jQuery('#pager1'), 
				sortname: 'rut', 
				viewrecords: true, 
				sortorder: "asc", 
				caption:"Listado de Empresas" 
				}); 
			
			jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
			jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
															  title:"Refrescar", 
															  buttonicon :'ui-icon-refresh', 
															  onClickButton:function(){ 
															     $("#txt_buscar").val("");
																 tabla();
															  } }); 
		  }
	}

	function update(i){
		//window.open("finanzasempresa/frmEmpresa.asp?id=i");
		$.post("finanzasempresa/frmEmpresa.asp",{id:i},
			   function(f){
				   $('#dialog').html(f);
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
		<li><a href="consultas.asp" accesskey="5" title="" class="selItem">Consultas</a></li>
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
				<li class="first"><a href="consultasEmpresa.asp">Empresas</a></li>
			</ul>
		</div>		
	</div>
		<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Empresas</em></h2>
           <center><table width="650" border="0">
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
								<label id="campoBuscar">
								<input name="rbtn_buscar" id="rbtn_buscar" type="radio" value="RUT" checked="checked" />Por Rut
           				        <input name="rbtn_buscar" id="rbtn_buscar" type="radio" value="R_SOCIAL"/>Por Razón Social
								</label></td>
						</tr>
						<tr>
							<td><input name="txt_buscar" type="text" id="txt_buscar" size="80"/></td>
							<td><input type="button" class="boton" value="Buscar" id="btnBuscar" onclick="tabla();"/></td>
						</tr>
						<tr>
                          <td align="center"><label id="tipoBuscar">
								<input name="rbtn_tipo" type="radio" value=" = '"/>Es exacta
								<input name="rbtn_tipo" type="radio" value=" LIKE '%" checked="checked"/>Que contienen
							    </label>	  
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
<div id="dialog" title="Registro Empresa">
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>