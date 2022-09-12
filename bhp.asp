<% Session.Abandon %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Mutual Capacitación</title>
<meta name="Keywords" content="" />
<meta name="Description" content="" />
<link rel="shortcut icon" href="images/IcoMutual.ico" />
<link href="css/default.css" rel="stylesheet" type="text/css" />
<link href="css/jquery-ui.css" rel="stylesheet" type="text/css"/>
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
var error="";
$(document).ready(function(){	
	$('#txUsuario').defaultValue('Rut de Usuario');
	$('#txPassword').defaultValue('Contraseña');
	});	

	function aMays(e, elemento) {
	tecla=(document.all) ? e.keyCode : e.which;
	if(tecla == 13) 
	  Incio();
	}

    function Incio()
	{
		$("#mensaje").html("");
//window.open("funciones/login.asp?rut_empresa="+$("#txEmpresa").val()+"&rut_usuario="+$("#txUsuario").val()+"&pass="+$("#txPassword").val()+"&tipo="+$("input[@name='Tipo']:checked").val())
		$.get("funciones/login_cursos.asp",
						 {rut_usuario:$("#txUsuario").val(),pass:$("#txPassword").val()},
						 function(xml){
								$('row',xml).each(function(i) { 
									if($(this).find('REGISTRO').text()=='1')
									{
										    document.location.href='inicio_cursos.asp';
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
	<div id="colOne">
		<h3><em style="text-transform: capitalize">Acceso</em></h3>
		<div class="bg1">
        <center>
          <table cellspacing="0" cellpadding="1" border=0>
           <tr>
            <tr>
                <td height="30"><center><input type="text" id="txUsuario" name="txUsuario"/></center></td>
            </tr>
            <tr>
                <td height="30"><center><input type="password" id="txPassword" name="txPassword" onkeypress="aMays(event, this)"/></center></td>
            </tr>
            <tr>
                <td height="30"><center><button id="btnIngreso" OnClick="Incio();">Ingresar</button></center></td>
            </tr>
             <tr>
                <td height="30"><label id="mensaje" style="color:#F00"></label></td>
            </tr>
          </table>
          </center>
		</div>
	</div>
	<div id="colTwo">
	  <div class="bg2">
			<h2><center><em style="text-transform: capitalize"><font size="+6">BIENVENIDOS</font></em>
			</center></h2>
            <h2><center><em style="text-transform: capitalize"><font size="+3">&nbsp;</font></em>
			</center></h2>
            <h2><font size="+3"><center><em style="text-transform: capitalize;">Sistema</em><b> de </b><em style="text-transform: capitalize;">administración</em></center></font></h2>
            <h2><center>
              <font size="+3"><b>y</b><em style="text-transform: capitalize;"> Control</em><b> de </b><em style="text-transform: capitalize;">cursos</em></font>
        </center></h2>
            <h2><center><em style="text-transform: capitalize"><font size="+3">&nbsp;</font></em>
			</center></h2>
            <h2><center><em style="text-transform: capitalize"><font size="+3">&nbsp;</font></em>
			</center></h2>
            <h2><center><em style="text-transform: capitalize"><font size="+6">&nbsp;</font></em>
			</center></h2>
            <h2><center><em style="text-transform: capitalize"></em>
			</center></h2>
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
</body>
</html>