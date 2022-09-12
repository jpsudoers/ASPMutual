<% 
Session.Contents.RemoveAll()
Session.Abandon %>
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
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
<script type="text/javascript">
var error="";
$(document).ready(function(){	
	$('#txEmpresa').defaultValue('Rut de Empresa');
	$('#txUsuario').defaultValue('ID de Usuario');
	$('#txPassword').defaultValue('Contraseña');
	MostrarFilas('Empr');
	    
	  $("#Doc").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 500,
			width:1000,
			modal: true,
			overlay: {
				backgroundColor: '#000',
				opacity: 0.5
			},
			buttons: {
				'Cerrar': function() {
					$(this).dialog('close');
				}
			},
			title: 'Documento'
		});	
		
		$("#dialog_espera").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 180,
			width: 400,
			modal: true
		});
		
		$("#avisoInicial").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 600,
			width: 1200,
			modal: true
		});
		
		$("#avisoInicial").dialog('open');
		
		$("#msnValida").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 180,
			width: 400,
			modal: true,
			buttons: {
				Aceptar: function() {
						$(this).dialog('close');
				}
			}
		});
		
		 $("#ResContrasena").dialog({
			autoOpen: false,
			bgiframe: true,
			height:400,
			width: 500,
			modal: true,
			buttons: {
				Restablecer: function() {
					if($("#frmResContrasena").valid())
					{
						$('#dialog_espera').dialog('open');
						$(this).dialog('close');
						$.post($('#frmResContrasena').attr('action')+'?'+$('#frmResContrasena').serialize(),function(d){
									              $('row',d).each(function(i) { 
														$('#dialog_espera').dialog('close');
																								
														if($(this).find('Valido').text()=="0")
														{
				 $('#msnValidaTxt').html("El Rut de Empresa o el Correo Electrónico de Usuario ingresado no es Válido.");
															$('#msnValida').dialog('open');
														}
														else
														{
		         $('#msnValidaTxt').html("Su Contraseña ha sido Restablecida Exitosamente. Por Favor, Revise su Correo Electrónico.");
															$('#msnValida').dialog('open');
														}
												  });
																						
																													   });
					}
				},
				Cerrar: function() {
						$(this).dialog('close');
				}
			}
		});
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
			OcultarResContra('resFila');
	}
	
	 function OcultarResContra(Fila) {
   			var elementos = document.getElementsByName(Fila);
   			for (k = 0; k < elementos.length; k++)
      		elementos[k].style.display = "none";
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
			MostrarResContra('resFila');
	}

	function MostrarResContra(Fila) {
  			var elementos = document.getElementsByName(Fila);
   			for (i = 0; i < elementos.length; i++)
      		elementos[i].style.display ="";
	}

    function Incio()
	{
		$("#mensaje").html("");
		$.get("funciones/login.asp",
						 {rut_empresa:$("#txEmpresa").val(),rut_usuario:$("#txUsuario").val(),
						 pass:$("#txPassword").val(),tipo:$("input[@name='Tipo']:checked").val()},
						 function(xml){
								$('row',xml).each(function(i) { 
									if($(this).find('REGISTRO').text()!='0')
									{
										    document.location.href='inicio.asp';
									}
									else
									{
										$("#mensaje").html("La Información Ingresada no es Válida. Inténtelo de nuevo");
									}
						});
		});
 }
 
 function RestablecerPass(){
		$.post("resContrasena/frmResContrasena.asp",
			  {empresa:$("#txEmpresa").val(),usuario:$("#txUsuario").val()},
			   function(f){
				   $('#ResContrasena').html(f);
				   $('#resRutEmpresa').focus();
				   validarResFrm();
		});
		$('#ResContrasena').dialog('open')
 }
 
 function validarResFrm(){
		$("#frmResContrasena").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			resRutEmpresa:{
				required:true,
				rut:true
			},
			resEmailEmpresa:{
				 required:true,
				 email:true
			}
		},
		messages:{
			resRutEmpresa:{
				required:"&bull; Ingrese el Rut de la Empresa.",
				rut:"&bull; Rut No válido"
			},
			resEmailEmpresa:{
				required:"&bull; Ingrese el Correo Electrónico del Usuario.",
				email:"&bull; Correo No Válido"
			}
		}
	});
    }
	
	function documento(doc){
		$("#ifPagina").attr('src','documentos/'+doc);
		if(!$('#Doc').dialog('isOpen'))
		{
			$('#Doc').dialog('open');
		}

	}	
</script>
</head>
<body onload="Javascript:history.go(1);" onunload="Javascript:history.go(1);">
<div id="header">
	<h1><img src="images/logo.png"/></h1>
    <ul>
          <li><a href="cert_online.asp" accesskey="3" class="selItem" onclick="">Certificados On line</a></li>
    </ul>    
</div>
<div id="content">
	<div id="colOne">
		<h3><em style="text-transform: capitalize">Acceso</em></h3>
		<div class="bg1">
        <center>
          <table border=0>
           <tr>
             <td><center><input type="radio" name="Tipo" id="Tipo" value="0" onclick="MostrarFilas('Empr');" checked>Empresa
		  		 <input type="radio" name="Tipo" id="Tipo" value="1" onclick="OcultarFilas('Empr');">Mutual</center><input type="hidden" id="u" /></td>
			</tr>
			<tr name="Empr" id="Empr">
                <td height="30"><center><input type="text" id="txEmpresa" name="txEmpresa"/></center></td>
            </tr>
            <tr>
                <td height="30"><center><input type="text" id="txUsuario" name="txUsuario"/></center></td>
            </tr>
            <tr>
                <td><center><input type="password" id="txPassword" name="txPassword" onkeypress="aMays(event, this)"/></center></td>
            </tr>
            <tr name="resFila" id="resFila">
                <td height="20"><center><a href="#" onclick="RestablecerPass();">¿Ha olvidado la contraseña?</a></center></td>
            </tr>
            <tr>
                <td height="30"><center><button id="btnIngreso" OnClick="Incio();">Ingresar</button></center></td>
            </tr>
             <tr>
                <td height="30"><label id="mensaje" style="color:#F00"></label></td>
            </tr>
	    <tr align="center">
            <td><font size="-1" color="#FF0000"><strong>Ver Video Tutorial</strong></font></br></br> 
		<video controls width="150">
	 		<source src="video/video_x264.mp4" type="video/mp4">
	 		<source src="video/video_x264.ogg" type="video/ogg">	
				<p>tu navegador no soporta este formato</p>
		</video>
		</td>
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
        
 <h2><center>
            <em style="text-transform: capitalize"><font size="+1">Horarios</font></em><b><font size="+1"> de </font></b><em style="text-transform: capitalize"><font size="+1">Atención :</font></em><em style="text-transform: capitalize"><font size="+1"> Lunes</font></em><b><font size="+1"> a </font></b><em style="text-transform: capitalize"><font size="+1">Jueves</font></em><b><font size="+1"> de 08:00 a 18:00 </font></b><em style="text-transform: capitalize"><font size="+1">Horas.</font></em>
            </center></h2>    
            <h2><center>
            <em style="text-transform: capitalize"><font size="+1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Viernes</font></em><b><font size="+1"> de 08:00 a </font></b><em style="text-transform: capitalize"><font size="+1">16:00 Horas.</font></em>
            </center></h2>                     
    <h2><center>
            <em style="text-transform: capitalize"><font size="+1">Horario límite para inscribir : </font></em><b><em style="text-transform: capitalize"><font size="+1"> Lunes</font></em><b><font size="+1"> a </font></b><em style="text-transform: capitalize"><font size="+1">Jueves</font></em><b><font size="+1"> de 17:00 </font></b><em style="text-transform: capitalize"><font size="+1">Horas.</font></em>
            </center></h2>    
            <h2><center>
            <em style="text-transform: capitalize"><font size="+1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Viernes</font></em><font size="+1"> de 15:00 </font></b><em style="text-transform: capitalize"><font size="+1">Horas.</font></em>
            </center></h2>     
         
            <!--<h2><center><em style="text-transform: capitalize"><font size="+1">Email : </font></em><b><font size="+1"><a href="mailto:mas.ayuda@subcontrataley.cl">mas.ayuda@subcontrataley.cl</a></font></b>
			</center></h2>-->
            <h2><center><em style="text-transform: capitalize"><font size="+3">&nbsp;</font></em>
			</center></h2>
            <h2><center><em style="text-transform: capitalize"><font size="+1"><a href="operacionsolicitud.asp">Solicitud de Inscripción de Empresa</a></font></em>
			</center></h2>
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="ResContrasena" title="Restablecer la contraseña">
</div>
<div id="dialog_espera" title="Restablecer la contraseña">
<center>
<table width="200" border="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><center><font size="2"><b><i>Espere, mientras se restablece su contraseña.</i></b></font></center></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><center><img id="loading" src="images/loader.gif"/></center></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</center>
</div>
<div id="avisoInicial" title="Aviso">
<b><font size="+1">Información General:</font></b><br/><br/>
<center>
<table border=1 cellpadding=1 cellspacing=1 width=988>
<td colspan="4" bgcolor="#000066" style="color:white" ><b>Contactos Coordinadores de Cursos</b></td>
  <!--<tr>
  <td width="250" align="center"><b>Inducción Camino a Cero Daño<br> 
  		Minera Escondida Ltda</b></td>
  <td width="250" align="center"><b>Inducción Mantos Blancos y Pampa Norte</b></td>
  <td width="250" align="center"><b>Inducción SSO VP Codelco y MantoVerde</b></td>
 </tr>-->
 <tr>
  <td align="center"><b>Karen Cofré Ojeda</b><br/>
  <a href="mailto:kcofre@mutualcapacitacion.cl" 
    target="_parent">kcofre@mutualcapacitacion.cl</a><br/><b>Cursos de VP Codelco</b><br/>
(55) 2 651285<br/>(9) 31790442
  <td align="center"><b>Gonzalo Retamal A.</b><br/><a href="mailto:gretamal@mutualcapacitacion.cl" 
    target="_parent">gretamal@mutualcapacitacion.cl</a><br/><a href="mailto:cursosgoldfields@mutualcapacitacion.cl" 
    target="_parent">cursosgoldfields@mutualcapacitacion.cl</a><br/><b>Cursos de Minera Gold Fields</b><br/>
(55) 2 651277<br/>
(9) 42939922
</td>
  <td align="center"><b>Maricel Veas V.</b><br/><a href="mailto:mveas@mutualcapacitacion.cl" 
    target="_parent">mveas@mutualcapacitacion.cl</a><br/><a href="mailto:coordinacioncursos@mutualcapacitacion.cl" 
    target="_parent">coordinacioncursos@mutualcapacitacion.cl</a><br/><b>Cursos de Minera Escondida, Spence</b><br/>
(55) 2 651283<br/>
(9) 79506230
</td>
  <td align="center"><b>Angelo Medero M.</b><br/><a href="mailto:amedero@mutualcapacitacion.cl" 
    target="_parent">amedero@mutualcapacitacion.cl</a><br/><b>Cursos de VP Codelco</b><br/>
(9) 31230741
</td>
</tr>
<tr>
  <td colspan="1" align="center"><b>Marco Paredes P.</b><br/><a href="mailto:mparedes@mutualcapacitacion.cl" 
    target="_parent">mparedes@mutualcapacitacion.cl</a><br/><a href="mailto:cursosmantoscopper@mutualcapacitacion.cl" 
    target="_parent">cursosmantoscopper@mutualcapacitacion.cl</a><br/><b>Cursos de Mantos Blancos y MantoVerde</b><br/>
(9) 32202412
</td>
  <td colspan="2"  align="center"><b>Dylan Varas G.</b><br/><a href="mailto:dvaras@mutualcapacitacion.cl" 
    target="_parent">dvaras@mutualcapacitacion.cl</a><br/><a href="mailto:cursosangloamerican@mutualcapacitacion.cl" 
    target="_parent">cursosangloamerican@mutualcapacitacion.cl</a><br/><b>Cursos de Anglo American</b><br/>
(9) 50040576
</td>
  <td align="center"><b>Estefanía Sagua V.</b><br/><a href="mailto:esagua@mutualcapacitacion.cl" 
    target="_parent">esagua@mutualcapacitacion.cl</a><br/><b>Asistente Comercial</b><br/>
(9) 40730678
</td>
 </tr>
 </table>
</center>
 <br/>
 <br/>
<center>
 <table border="1" cellpadding="1" cellspacing="1" width="930">
 <tr>
 	<td colspan="3" bgcolor="#000066" style="color:white"><b>Facturación <!---->/ Despacho</b></td>
    <tr>
  <td align="center" colspan="1"><b>Facturación</b></td>
  <td align="center"><b>Cobranza</b></td>
  <!--<td align="center"><b>Mesa Central</b></td>-->
   </tr>
    <tr>
  <td align="center"><b>César González H.<br>
  </b>
  <b>Ejecutivo de Liberación y Facturación</b><br>
 (<a href="mailto:cmgonzalez@mutualcapacitacion.cl" target="_parent">cmgonzalez@mutualcapacitacion.cl</a>)<br>+569 790 63 583  
 </td>
  <td align="center"><b>Pedro Godoy G.<br>
  </b>
  <b>Ejecutivo</b><br>
 (<a href="mailto:pgodoyg@mutualcapacitacion.cl" target="_parent">pgodoyg@mutualcapacitacion.cl</a>)<br> 
 (55) 265 1276
 </td>  
   </tr>
  </tr>
 </table>
  <br/>
 <table border="1" cellpadding="1" cellspacing="1" width="930">
 <tr>
 	<td colspan="3" bgcolor="#000066" style="color:white"><b>Contacto Comercial y Proyectos Especiales</b></td>
    <tr>
          <td align="center"><b>Mariela Lobos Bernal<br>
          </b>
          <b>Directora Nacional de Capacitación Minería</b><br>
         (<a href="mailto:mlobos@mutualcapacitacion.cl" target="_parent">mlobos@mutualcapacitacion.cl</a>)<br>(9) 93094294 
         </td>
          <td align="center"><b>Chaylin Huanchicay Barraza<br>
          </b>
          <b>Jefe de Coordinación Minería</b><br>
         (<a href="mailto:chuanchicay@mutualcapacitacion.cl" target="_parent">chuanchicay@mutualcapacitacion.cl</a>)<br> 
         (55) 2651281 ó 993217512
         </td>  
          <td align="center"><b>Nelson Maturana Hoyos<br>
          </b>
<b>Jefe de Proyectos VP Codelco</b><br>
         (<a href="mailto:nmaturana@mutualcapacitacion.cl" target="_parent">nmaturana@mutualcapacitacion.cl</a>)<br> 
         (9) 98715152
         </td>  
   </tr>
  </tr>
 </table>  </center>
  <hr />
<font size="+2" color="#FF0000"><i>Información Importante</i></span></font><br/><br/>
 <hr />
<p><b><font size="+1">Estimadas Empresas:</font></b><br/><br/>
<b><u>Informamos que desde el 1 de junio comenzaremos a realizar cursos de Inducción HSE Minera Escondida (Cero Daño) en modalidad asincrónica e-learning.</u></b><br/><br/>
La Programación se hará el lunes 30 de mayo y esta consiste en crear cursos de lunes a viernes con capacidad de 100 cupos por días, para ello las empresas deben ingresar a programación de curso e inscribir. En un plazo de 24 horas se habilitará el curso durante 7 días para que el participante pueda realizarlo. A cada empresa les llegará un correo de confirmación con más detalle de plataforma de ingreso usuario y clave para cada participante, instrucciones de ingreso a esta.<br/><br/>

Es por lo que invitamos cordialmente a nuestros clientes para informarse con los correspondientes coordinadores de cursos.

</p><hr />
<p>
<p><b><font size="+1">Estimadas Empresas:</font></b><br/><br/>
Les informamos que para pago de cursos con transferencia lo deben hacer: <br/><br/>
<b>A nombre de : </b>Mutual de Seguridad Capacitación S.A. <br/>
<b>Rut : </b>76.410.180-4<br/>
<b>Número de Cuenta : </b>10675612<br/>
<b>Tipo de Cuenta : </b>Cuenta Corriente<br/>
<b>Banco : </b>BCI<br/>
</p>
<!--<hr />
<p>
<b><font size="+1">Estimada Empresa:</font></b><br/><br/>
Desde <b>Abril</b> el curso <b>Inducción de Mantos Copper - Div Mantos Blancos</b> se realizará los días <b>Lunes y Martes</b>.<br/>
</p>
<hr />
<h1 style="font-size:16px">
<strong>Estimadas Empresas: </strong></h1>
<p>Les recordamos notificar a los coordinadores respectivos todas aquellas actualizaciones asociadas a la información de su empresa, tales como cambio de razón social, giro, dirección, teléfonos, procedimientos de facturación (inclusión de HES, lugar de recepción, horarios, etc.), contacto Coordinador de cursos y finanzas, con la finalidad de actualizar nuestros sistemas y mejorar la calidad de servicio que les brindamos a ustedes. </p>
<br/><hr />
<p>
<b><font size="+1">Estimada Empresa:</font></b><br/><br/>
Informamos que a contar del <b>Viernes 5 de Octubre 2018</b> se modifica el nombre de los cursos de:<br/><br/>
- Inducción en seguridad, salud, medio ambiente y comunidad para alcanzar el camino cero daño Por <b>Inducción HSE Minera Escondida.</b><br/>
- Técnicas De Seguridad Y Salud Ocupacional Para Actividades Mineras Por <b>Inducción HSE Pampa Norte.</b><br/>
</p>
<hr />
<p>
<p>
<b><font size="+1">Estimada Empresa:</font></b><br/><br/>
Por motivos de cambios en nuestros procesos de obtención de órdenes de compra OTIC, informamos que para solo aquellas empresas que inscriben con dichas órdenes de compra, los certificados estarán disponibles como máximo 5 días después de haber terminado el curso. Es por este motivo, solicitamos que las inscripciones realizadas adjunten su respectiva orden de compra según la OTIC y no respaldo preliminar de inscripción.</b>
<br/>
</p>-->
<hr />
<p>
<p>
<b><font size="+1">Estimada Empresa:</font></b><br/><br/>
Para modificar o anular inscripciones tienen un plazo de 24 horas previas al comienzo del curso, de lo contrario se procederá al cobro de dicha inscripción. Favor comunicarse con coordinador respectivo a través de correo electrónico.
<br/>
</p>
<hr />
<!--<p>
<b><font size="+1">Estimadas Empresas:</font></b><br/><br/>
Les recordamos notificar al correo <b><a href="mailto:magonzalezg@mutualasesorias.cl" target="_parent">magonzalezg@mutualasesorias.cl</a></b> todas aquellas actualizaciones asociadas a la información de su empresa.</p>
<br/>
<b>Atte.<br/>
Mutual Capacitaci&oacute;n.</b><br/><br/>-->
</div>
<div id="msnValida" title="Restablecer la contraseña">
     <label id="msnValidaTxt" name="msnValidaTxt"></label>
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
</body>
</html>