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
			/*if ($("#selIns").val()=="1"	
					window.open('consultasfinanzas/xls.asp?mes='+$("#SelMes").val()+'&ano='+$("#selAno").val()+'&insc='+$("#selIns").val(),'Informe Ventas Mes')
			else*/
			//alert('consultasRptLF/xls.asp?mes='+$("#SelMes").val()+'&ano='+$("#selAno").val()+'&insc='+$("#selIns").val());
					window.open('consultasRptLF/xls.asp?mes='+$("#SelMes").val()+'&ano='+$("#selAno").val()+'&insc='+$("#selIns").val(),'Informe')
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
		<li><a href="consultas.asp" accesskey="5" title="">Consultas</a></li>
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
    	<h2><em style="text-transform: capitalize;">Liberación &amp; Facturación</em></h2>
        <br />
        <div>
						<table cellpadding="0" cellspacing="0" style="border:0px;width:100%;" class="cabezaMsj">
							<tr class="buscarTitulo">
								<td><h3><em style="text-transform: capitalize;">Reportes</em></h3></td>
							</tr>
						</table>
		</div>
		<div id="buscarContent" style="border:1px solid #000099;display:block" >
					<form name="frminforme" id="frminforme" action="" method="post">
				    <br />                     
                    <table width="620" border="0" align="center" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="30">Mes : </td>
					        <td width="100"><select id="SelMes" name="SelMes" tabindex="1">
                                    <!--<OPTION VALUE="0">Todos</OPTION>-->
                                <option value="1" <%if(cdbl(month(now))=1)then%>selected="selected"<%end if%>>Enero</option>
                                <option value="2" <%if(cdbl(month(now))=2)then%>selected="selected"<%end if%>>Febrero</option>
                                <option value="3" <%if(cdbl(month(now))=3)then%>selected="selected"<%end if%>>Marzo</option>
                                <option value="4" <%if(cdbl(month(now))=4)then%>selected="selected"<%end if%>>Abril</option>
                                <option value="5" <%if(cdbl(month(now))=5)then%>selected="selected"<%end if%>>Mayo</option>
                                <option value="6" <%if(cdbl(month(now))=6)then%>selected="selected"<%end if%>>Junio</option>
                                <option value="7" <%if(cdbl(month(now))=7)then%>selected="selected"<%end if%>>Julio</option>
                                <option value="8" <%if(cdbl(month(now))=8)then%>selected="selected"<%end if%>>Agosto</option>
                                <option value="9" <%if(cdbl(month(now))=9)then%>selected="selected"<%end if%>>Septiembre</option>
                                <option value="10" <%if(cdbl(month(now))=10)then%>selected="selected"<%end if%>>Octubre</option>
                                <option value="11" <%if(cdbl(month(now))=11)then%>selected="selected"<%end if%>>Noviembre</option>
                                <option value="12" <%if(cdbl(month(now))=12)then%>selected="selected"<%end if%>>Diciembre</option>
                            </select></td>
                            <td width="30">Año : </td>
                            <td width="80"><select id="selAno" name="selAno" tabindex="2">
                                            	<!--<OPTION VALUE="0">Todos</OPTION>-->
											<%
                                            For ano = 2016 To cdbl(year(now)) Step 1
                                            %>
                                                <OPTION VALUE="<%=ano%>" <%if(ano=cdbl(year(now)))then%>selected="selected"<%end if%>><%=ano%></OPTION>
                                            <%
                                            Next
                                            %>
                            </select></td>
                            <td width="80">Inscripciones : </td>
                            <td width="200"><select id="selIns" name="selIns" tabindex="3">
										<OPTION VALUE="1">Ventas Mes</OPTION>
                                        <OPTION VALUE="2">Pendiente de Liberar - Facturar</OPTION>
                                        <%if(Session("usuarioMutual")="8" or Session("usuarioMutual")="142" or Session("usuarioMutual")="145" or Session("usuarioMutual")="36" or Session("usuarioMutual")="1238" or Session("usuarioMutual")="1244" or Session("usuarioMutual")="1251" or Session("usuarioMutual")="209" or Session("usuarioMutual")="1269")then%>
                                        <OPTION VALUE="3">Informe de Facturac&oacute;n Mensual</OPTION> 
                                        <%end if%>       
										<!--<OPTION VALUE="2">Liberadas No Facturadas</OPTION>  -->                                
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
