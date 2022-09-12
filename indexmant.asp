<% Session.Abandon %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Mutual Capacitación</title>
<meta name="Keywords" content="" />
<meta name="Description" content="" />
<link href="css/default.css" rel="stylesheet" type="text/css" />
<link href="css/jquery-ui.css" rel="stylesheet" type="text/css"/>
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
var error="";
$(document).ready(function(){	
	$('#txEmpresa').defaultValue('Rut de Empresa');
	$('#txUsuario').defaultValue('Rut de Usuario');
	$('#txPassword').defaultValue('Contraseña');
	MostrarFilas('Empr');
	});	

	function aMays(e, elemento) {
	tecla=(document.all) ? e.keyCode : e.which;
	if(tecla == 13) 
	  Incio();
	}

   function OcultarFilas(Fila) {
   			var elementos = document.getElementsByName(Fila);
   			for (k = 0; k < elementos.length; k++)
      		elementos[k].style.display = "none";
			$('#txEmpresa').val("");
			$('#txUsuario').val("");
			$('#txPassword').val("");
			$('#txEmpresa').defaultValue('Rut de Empresa');
			$('#txUsuario').defaultValue('Rut de Usuario');
			//$('#txPassword').defaultValue('Contraseña');
	}
	
	function MostrarFilas(Fila) {
  			var elementos = document.getElementsByName(Fila);
   			for (i = 0; i < elementos.length; i++)
      		elementos[i].style.display ="";
			$('#txEmpresa').val("");
			$('#txUsuario').val("");
			$('#txPassword').val("");
			$('#txEmpresa').defaultValue('Rut de Empresa');
			$('#txUsuario').defaultValue('ID de Usuario');
			//$('#txPassword').defaultValue('Contraseña');
	}

    function Incio()
	{
		$("#mensaje").html("");
//window.open("funciones/login.asp?rut_empresa="+$("#txEmpresa").val()+"&rut_usuario="+$("#txUsuario").val()+"&pass="+$("#txPassword").val()+"&tipo="+$("input[@name='Tipo']:checked").val())
		$.get("funciones/login.asp",
						 {rut_empresa:$("#txEmpresa").val(),rut_usuario:$("#txUsuario").val(),
						 pass:$("#txPassword").val(),tipo:$("input[@name='Tipo']:checked").val()},
						 function(xml){
								$('row',xml).each(function(i) { 
									if($(this).find('REGISTRO').text()=='1')
									{
										    document.location.href='inicio.asp';
									}
									else
									{
										$("#mensaje").html("La Información Ingresada no es Valida.");
									}
						});
		});
 }
</script>
</head>
<body onload="Javascript:history.go(1);" onunload="Javascript:history.go(1);">
<div id="header">
	<h1><img src="images/logo.png"/></h1>
</div>
<div id="content">


	  <div class="bg2">
			<h2><center><em style="text-transform: capitalize"><font size="+6">Sitio en Mantención</font></em>
			</center></h2>
            <br />
            <p>
              <font size="+1">Estimados Clientes, en nuestros constantes esfuerzos por mejorar nuestros servicios, hemos iniciado un plan de cambio de infraestructura, que permitirá hacer más eficiente el proceso de sus requerimientos.</font>
            </p>
			<br />
            <p><font size="+1">Por lo anterior, el servicio no estará disponible   hasta el Lunes 21 de marzo a las 08:00 hrs.</font></p>
            <br />
            <br />
            <br />            
<h2><center><em style="text-transform: capitalize"><font size="+1">Mesa</font></em><b><font size="+1"> de </font></b><em style="text-transform: capitalize"><font size="+1">Ayuda</font></em>
            </center></h2>
            <h2><center><em style="text-transform: capitalize"><font size="+1">Email : </font></em><b><font size="+1"><a href="mailto:mas.ayuda@subcontrataley.cl">mas.ayuda@subcontrataley.cl</a></font></b>
			</center></h2>
            <br />
            <br /> 
	  </div>
	</div>

<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
</body>
</html>