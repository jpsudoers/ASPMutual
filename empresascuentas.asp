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
		 $('#fechaInicio').val("");
		 $('#fechaTermino').val("");			   
		 $('#fechaInicio').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		 $('#fechaTermino').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });

         validar();
		
		$("#Doc").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 600,
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
			title: 'Cuenta Corriente'
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

    function documento(arch){
		//alert("http://norte.otecmutual.cl/ordenes/"+arch)
		$("#ifPagina").attr('src',arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
	}

	function VerInforme()
	{
		if(validaFecha($('#fechaInicio').val(),$('#fechaTermino').val())==true && $('#fechaInicio').val()!="" && $('#fechaTermino').val()!="")
		{
		    $('#validaFechas').val("1");
			//alert($('#validaFechas').val());
		}
		else
		{
		    $('#validaFechas').val("");
			//alert($('#validaFechas').val());
		}
		
		validar();
		
		if($("#frminforme").valid())
		{
		   //alert(validaFecha($('#txtFechI').val(),$('#txtFechF').val()));
		   documento("cuentacorriente/pdf.asp?empresa="+$('#Empresa').val());
		}
	}

	function validar(){
		$("#frminforme").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			fechaInicio:{
				required:true
			},
			fechaTermino:{
				required:true
			},
			Empresa:{
				required:true
			},
			validaFechas:{
				required:true
			}
		},
		messages:{
			fechaInicio:{
				required:"&bull; Seleccione Fecha de Inicio."
			},
			fechaTermino:{
				required:"&bull; Seleccione Fecha de Término."
			},
			Empresa:{
				required:"&bull; empresa."
			},
			validaFechas:{
				required:"&bull; La Fecha de Inicio debe ser Menor o Igual a la Fecha de Término."
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
	
	 var formato = 1;  
 
	 function invFecha(nTipFormat,dFecIni){  
		  var dFecIni = dFecIni.replace(/-/g,"/"); 
		  var nPosUno  = ponCero(dFecIni.substr(0,dFecIni.indexOf("/")));  
		  var nPosDos  = ponCero(dFecIni.substr(parseInt(dFecIni.indexOf("/")) + 1,parseInt(dFecIni.lastIndexOf("/")) - parseInt(dFecIni.indexOf("/")) - 1));
		  var nPosTres = ponCero(dFecIni.substr(parseInt(dFecIni.lastIndexOf("/")) + 1));  
	 
		  switch(nTipFormat){  
			 case 1 :   
				 dReturnFecha = nPosTres + "" + nPosDos + "" + nPosUno;  
				 break;  
			 case 2 :     
				 dReturnFecha = nPosTres + "" + nPosUno + "" +nPosDos;  
				 break;  
			 case 3 :   
				 dReturnFecha = nPosUno + "" + nPosDos + "" +nPosTres;  
				 break;  
			 case 4 :   
				 dReturnFecha = nPosUno + "" + nPosTres + "" +nPosDos;  
				 break;  
		 }  
		 return dReturnFecha;       
	 }  
	   
	 function ponCero(strPon){  
		 if(parseInt(strPon.length) < 2)  
			 strPon = "0" + strPon;  
		 return strPon;  
	 }  
	
	 function comparaFecha(dFormat,dFecMenor, dFecMayor){  
		 dFecMenor = invFecha(dFormat,dFecMenor);  
		 dFecMayor = invFecha(dFormat,dFecMayor);  
	   
		 if(dFecMenor > dFecMayor)  
			 return false;  
		 else  
			 return true;  
	 }  
	   
	 function validaFecha(fechaInicio,FechaTermino){  
		 if(comparaFecha(formato,fechaInicio,FechaTermino) == true)  
			 return true;
		 else  
			 return false; 
	 }  
     </script>
</head>
<body>
<div id="header">
	<h1><img src="images/logo.png"  /></h1>
	<ul>
    <% if(Session("tipoUsuario")="1")then
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
		   <li><a href="solicitudescreditoEmpresa.asp">Solicitud de Credito</a></li>
		   <li><a href="empresasOC.asp">Ingreso de Orden de Compra</a></li>
               <%else
				   if(Session("cargo_user_empresa")="1")then%>
					   <li class="first"><a href="empresacalendario.asp">Inscripción de Cursos</a></li>
					   <li><a href="empresasinspendientes.asp">Inscripciones Pendientes</a></li>
					   <li><a href="empresainsactivas.asp">Inscripciones Autorizadas</a></li>

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
        <h2><em style="text-transform: capitalize;">Cuenta Corriente</em></h2>
        <br />
       <div>
						<table cellpadding="0" cellspacing="0" style="border:0px;width:100%;" class="cabezaMsj">
							<tr class="buscarTitulo">
								<td><h3><em style="text-transform: capitalize;">Informes</em></h3></td>
							</tr>
						</table>
					</div>
			<div id="buscarContent" style="border:1px solid #000099;display:block" >
                    <form name="frminforme" id="frminforme" action="cambioContrasena/modificarEmpresa.asp" method="post">
                    <table width="620" border="0" align="center" cellpadding="1" cellspacing="1">
                        <tr>
    						<td width="80">&nbsp;</td>
                            <td width="150">&nbsp;</td>
                            <td width="80">&nbsp;</td>
                            <td width="150">&nbsp;</td>
   				        </tr>
						<tr>
                            <td>Desde :</td>
                            <td><input type="text" id="fechaInicio" name="fechaInicio" value=""/></td>
                            <td>Hasta :</td>
                            <td><input type="text" id="fechaTermino" name="fechaTermino" value=""/></td>
                            <td><input type="button" value="Ver Informe" onclick="VerInforme();" id="btnBuscar" name="btnBuscar"/></td>
						</tr>
                       <tr>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td><input type="hidden" id="Empresa" name="Empresa" value="<%=Session("usuario")%>"/></td>
							<td><input type="hidden" id="validaFechas" name="validaFechas" value=""/></td>
                            <td>&nbsp;</td>
   					   </tr>
					</table>
                </form> 
				</div>	
                <div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
                   	 <ul></ul> 
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
