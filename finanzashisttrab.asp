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
$(document).ready(function(){
		$('#txRut').defaultValue('Todas');
		llena_curriculo(0);
		
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
			title: 'Informe de Movimientos'
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

	function VerInforme()
	{
		//alert($("#SelMes").val()+ ' ' + $("#selAno").val()+ ' ' + $("#selIns").val());
			//if ($("input[@name='optVer']:checked").val()=="1")	
				//documento("manualesinforme/pdf.asp?tipo="+$('#Tipo_Mov').val());
			//else
		window.open('consultashistoricos/xls.asp?mes='+$("#SelMes").val()+'&ano='+$("#selAno").val()+'&insc='+$("#selIns").val()+'&id_empresa='+$("#Empresa").val()+'&id_mutual='+$('#Curriculo').val(),'Informe')
	}
	
	function documento(arch){
		$("#ifPagina").attr('src',arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
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
	
	function lookup(inputString) {
		if(inputString.length <=3) {
			$("#txtRsoc").html("");
			$("#txtRsocCam").html("&nbsp;");
			$('#Empresa').val("");
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
	   cargaDatos(id);
	   $('#txRut').val(rut);
	   $('#suggestions').hide();
	}
	
	function cargaDatos(id)	
	{
		$("#txRut").val("");
		$("#txtRsoc").html("");
		$("#txtRsocCam").html("&nbsp;");
	
		$.get("cuentacorriente/datosempresa.asp",
						 {id:id},
						 function(xml){
									$('row',xml).each(function(i) { 
									$("#txtRsocCam").html("Razón Social : ");
									$("#txtRsoc").html($(this).find('RSOCIAL').text());
									$("#txRut").val($(this).find('RUT').text());
						});
			});
	}	
	
	function llena_curriculo(id_curriculo){
		$("#Curriculo").html("");
		$.get("consultascertificados/curriculo.asp",
					function(xml){
						$("#Curriculo").append("<option value='0'>Seleccione</option>");
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
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO6,PERMISO7,PERMISO8,PERMISO9,PERMISO10,PERMISO11,PERMISO12 from USUARIOS "
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
		<li><a href="finanzas.asp" accesskey="4" title="" class="selItem">Finanzas</a></li>
        <%end if
		if(DATOS("PERMISO6")<>"0")then
		%>
		<li><a href="consultas.asp" accesskey="5" title="" >Consultas</a></li>
        <%end if
		if(DATOS("PERMISO12")<>"0")then
		%>
		<li><a href="manuales.asp" accesskey="6" title="">Manuales</a></li>
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
			 <!--#include file="menu_izquierdo/menuFinanzas.asp"-->		
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
    	<h2><em style="text-transform: capitalize;">Historico</em><b> de </b><em style="text-transform: capitalize;">Cursos</em></h2>
        <br />
        <div>
						<table cellpadding="0" cellspacing="0" style="border:0px;width:100%;" class="cabezaMsj">
							<tr class="buscarTitulo">
								<td><h3><em style="text-transform: capitalize;">Informes</em></h3></td>
							</tr>
						</table>
		</div>
		<div id="buscarContent" style="border:1px solid #000099;display:block" >
					<form name="frminforme" id="frminforme" action="" method="post">
                    <br />
                    <table width="620" border="0" align="center" cellpadding="1" cellspacing="2">
                        <tr>
                            <td colspan="4">Rut Empresa :
          				   <input id="txRut" name="txRut" type="text" tabindex="1" maxlength="20" size="20" onkeyup="lookup(this.value);"/>
                            <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:100px;left:550px">
                                <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
                                <div class="suggestionList" id="autoSuggestionsList">
                                  &nbsp;
                                </div>
                             </div>
                             	<input type="hidden" id="Empresa" name="Empresa"/>
                            </td>
                            <td colspan="4"><label id="txtRsocCam" name="txtRsocCam">&nbsp;</label><label id="txtRsoc" name="txtRsoc"></label></td>
                        </tr>
                        <tr>
                             <td colspan="8">Curso :  <select id="Curriculo" name="Curriculo" tabindex="1" style="width:34em;"></select></td>
                        </tr>
						<tr>
                            <td width="30">Mes : </td>
					        <td width="100"><select id="SelMes" name="SelMes" tabindex="2">
                                    <OPTION VALUE="0">Todos</OPTION>
                                    <OPTION VALUE="1">Enero</OPTION>
                                    <OPTION VALUE="2">Febrero</OPTION>
                                    <OPTION VALUE="3">Marzo</OPTION>
                                    <OPTION VALUE="4">Abril</OPTION>
                                    <OPTION VALUE="5">Mayo</OPTION>
                                    <OPTION VALUE="6">Junio</OPTION>
                                    <OPTION VALUE="7">Julio</OPTION>
                                    <OPTION VALUE="8">Agosto</OPTION>
                                    <OPTION VALUE="9">Septiembre</OPTION>
                                    <OPTION VALUE="10">Octubre</OPTION>
                                    <OPTION VALUE="11">Noviembre</OPTION>
                                    <OPTION VALUE="12">Diciembre</OPTION>
                            </select></td>
                            <td width="30">Año : </td>
                            <td width="80"><select id="selAno" name="selAno" tabindex="3">
                                            	<OPTION VALUE="0">Todos</OPTION>
											<%
                                            For ano = 2010 To cdbl(year(now)) Step 1
                                            %>
                                                <OPTION VALUE="<%=ano%>"><%=ano%></OPTION>
                                            <%
                                            Next
                                            %>
                            </select></td>
                            <td width="80">Inscripciones : </td>
                            <td width="200"><select id="selIns" name="selIns" tabindex="4">
										<OPTION VALUE="Todas">Todas</OPTION>
                                        <OPTION VALUE="0">Liberadas</OPTION>
                                        <OPTION VALUE="2">Pendientes de Liberar</OPTION>                                        
                            </select></td>                            
                        <td><input type="button" value="Ver Informe" onclick="VerInforme();" id="btnBuscar" name="btnBuscar"/></td>
						</tr>
					</table>
                    <br />                    
                </form> 
			</div>	
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
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
