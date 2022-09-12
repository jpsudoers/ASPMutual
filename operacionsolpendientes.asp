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
	$('#fechaI').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
	$('#fechaT').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
	});	

	function VerInforme()
	{
		//validar();
		//if($("#frminforme").valid())
		//{
		window.open("solpendientes/pdf.asp?fechaI="+$('#fechaI').val()+"&fechaT="+$('#fechaT').val(),'Informe')
		//}
	}
	
		function validar(){
		$("#frminforme").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			fechaI:{
				required:true
			},
			fechaT:{
				required:true
			}
		},
		messages:{
			fechaI:{
				required:"&bull; Seleccione Fecha Desde."
			},
			fechaT:{
				required:"&bull; Seleccione Fecha Hasta."
			}
		}
	});
    }
	
	</script>
</head>
<body>
<div id="header">
	<h1><img src="images/logo.png"  /></h1>
	<ul>
	<li><a href="administracion.asp" accesskey="1" title="">Administraci&oacute;n</a></li>
		<li><a href="#" accesskey="2" title="" class="selItem">Operaci&oacute;n</a></li>
        		<li><a href="manejocursos.asp" accesskey="3" title="">Manejo de Cursos</a></li>
		<li><a href="finanzas.asp" accesskey="4" title="">Finanzas</a></li>
		<li><a href="empresas.asp" accesskey="5" title="">Empresas</a></li>
		
	</ul>
</div>
<div id="content">
	<div id="colOne">
  	<h3>Login</h3>
		<div class="bg1">
			Usuario:<strong>Nombre de Usuario</strong>
      <i>Perfil de usuario</i><br />
      <%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%>
      <button OnClick="document.location.href='index.asp';">Cerrar Sesión</button>
		</div>
		<h3>Opciones</h3>
		<div class="bg1">
			<ul>
				<li class="first"><a href="operacionsolicitudes.asp">Solicitudes</a></li>
                <li><a href="operacionsolpendientes.asp">Solicitudes Pendientes</a></li>
                <li><a href="operacionprogramacion.asp">Programación</a></li>
				<li><a href="operacionautorizacion.asp">Autorización</a></li>
				<li><a href="operacioninscripcion.asp">Inscripción</a></li>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
        <h2><em style="text-transform: capitalize;">Solicitudes Pendientes</em></h2>
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
					<table width="620" border="0" align="center" cellpadding="1" cellspacing="1">
                   <tr>
    						<td width="50">&nbsp;</td>
                            <td width="170">&nbsp;</td>
                            <td width="50">&nbsp;</td>
                            <td width="170">&nbsp;</td>
                       
   						</tr>
						<tr>
                            <td>Desde :</td>
                            <td><input id="fechaI" name="fechaI" type="text" tabindex="1" maxlength="50" size="15"/></td>
                            <td>Hasta :</td>
                            <td><input id="fechaT" name="fechaT" type="text" tabindex="2" maxlength="50" size="15"/></td>
                            <td width="164" rowspan="1"><input type="button" value="Ver Informe" onclick="VerInforme();" id="btnBuscar" name="btnBuscar"/></td>
						</tr>
                        <tr>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
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
</body>
</html>
