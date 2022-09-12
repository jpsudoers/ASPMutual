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
<script type="text/javascript">
estado=1;
$(document).ready(function(){					
	llena_otic(0);
	llena_mutual(0);
	$("#txRut").focus();
	validar();

			$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 200,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
					$('#dialog').dialog('close');
					document.location.href='index.asp';
				}
			}
		});
		
			$("#dialog_espera").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 200,
			width: 600,
			modal: true
		});
		
		$("#rut1").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
					$('#rut1').dialog('close');
					$("#txRut").val("");
					$("#txRut").focus();
				}
			}
		});
		
			$("#datosTerminos").dialog({
			autoOpen: false,
			bgiframe: true,
			height:450,
			width: 1000,
			modal: true,
			buttons: {
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	


	function RutEmpresa()	
	{
		$.get("solicitud/rutempresa.asp",
						 {rut:$("#txRut").val()},
						 function(xml){
									$('row',xml).each(function(i) { 
									if($(this).find('total').text()!="0")
									{
										$("#txRut").val("");
										$('#rut1').dialog('open');
										$("#txRut").focus();
										//$("#txtRsoc").focus();
									}
						});
			});
	}

	function validar(){
		$("#frmEmpresa").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txRut:{
				required:true,
				rut:true
			},
			txtRsoc:{
				required:true
			},
			txtGiro:{
				required:true
			},
			txtDir:{
				required:true
			},
			txtCom:{
				required:true
			},
			txtCiu:{
				required:true
			},
			txtFon:{
				required:true
			},
			txtOtic:{
				required:true
			},
			txtMutual:{
				required:true
			},
			txtNomb:{
				required:true
			},
			txtCargo:{
				required:true
			},
			txtEmail:{
				required:true,
				email:true,
				distinctTo:"#txtEmailCont"
			},
			txtEmailRepetir:{
				required:true,
				email: true,
				equalTo: "#txtEmail"
			},
			txtFono:{
				required:true
			},
			txtNombCont:{
				required:true
			},
			txtCargoCont:{
				required:true
			},
			txtEmailCont:{
				required:true,
				email:true
			},
			txtEmailContRepetir:{
				required:true,
				email:true
			},
			txtFonoCont:{
				required:true
			}
		},
		messages:{
			txRut:{
				required:"&bull; Ingrese Rut Empresa.",
				rut:"&bull; Rut No valido"
			},
			txtRsoc:{
				required:"&bull; Ingrese Razón Social Empresa."
			},
			txtGiro:{
				required:"&bull; Ingrese el Giro Empresa."
			},
			txtDir:{
				required:"&bull; Ingrese Dirección Empresa."
			},
			txtCom:{
				required:"&bull; Ingrese Comuna Empresa."
			},
			txtCiu:{
				required:"&bull; Ingrese Ciudad Empresa."
			},
			txtFon:{
				required:"&bull; Ingrese Teléfono Empresa."
			},
			txtOtic:{
				required:"&bull; Seleccione Otic a la que Pertenece Empresa."
			},
			txtMutual:{
				required:"&bull; Seleccione Mutual a la que Pertenece Empresa."
			},
			txtNomb:{
				required:"&bull; Ingrese Nombre Contacto."
			},
			txtCargo:{
				required:"&bull; Ingrese Cargo Contacto."
			},
			txtEmail:{
				required:"&bull; Ingrese Correo Electrónico Contacto.",
				email:"&bull; Correo No Valido",
				distinctTo:"&bull; El Correo del Coordinador de Cursos debe ser distinto al correo del Contacto de Contabilidad."
			},
			txtEmailRepetir:{
				required:"&bull; Ingrese Nuevamente el Correo Electrónico del Contacto.",
				email:"&bull; Correo No Valido",
				equalTo: "&bull; Los Correos del Coordinador de Cursos deben ser Iguales."
			},
			txtFono:{
				required:"&bull; Ingrese Teléfono del Contacto."
			},
			txtNombCont:{
				required:"&bull; Ingrese Nombre del Contacto de Contabilidad."
			},
			txtCargoCont:{
				required:"&bull; Ingrese Cargo del Contacto de Contabilidad."
			},
			txtEmailCont:{
				required:"&bull; Ingrese Correo Electrónico del Contacto de Contabilidad.",
				email:"&bull; Correo No Valido"
			},
			txtEmailContRepetir:{
				required:"&bull; Ingrese Nuevamente Correo Electrónico del Contacto de Contabilidad.",
				email:"&bull; Correo No Valido"
			},
			txtFonoCont:{
				required:"&bull; Ingrese Teléfono del Contacto de Contabilidad."
			}
		}
	});
    }

	function llena_mutual(id_mutual){
		$("#txtMutual").html("");
		$.get("empresa/mutual.asp",
					function(xml){
						$("#txtMutual").append("<option value=\"\">Seleccione</option>");
						$('row',xml).each(function(i) { 
							if(id_mutual==$(this).find('ID_MUTUAL').text())
							{
$("#txtMutual").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" selected="selected">'+$(this).find('NOMBRES').text()+ '</option>');
							}
							else
							{
								$("#txtMutual").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
							}
						});
						$("#txtMutual").append("<option value='0'>Sin Mutual</option>");
					});
	}
	
	function llena_otic(id_otic){
		$("#txtOtic").html("");
		$.get("empresa/otic.asp",
					function(xml){
						$("#txtOtic").append("<option value=\"\">Seleccione</option>");
						//$("#OTIC").append("<option value=5>Sin OTIC</option>");
						$('row',xml).each(function(i) { 
							if(id_otic==$(this).find('ID').text())
							{
								$("#txtOtic").append('<option value="'+$(this).find('ID').text()+'" selected="selected">'+
																		$(this).find('razon').text()+ '</option>');
							}
							else
							{
								$("#txtOtic").append('<option value="'+$(this).find('ID').text()+'" >'+
																		$(this).find('razon').text()+ '</option>');
							}
						});
						$("#txtOtic").append("<option value='0'>Sin OTIC</option>");
					});
	    }
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}
	
	function estadoBoton()
	{
		if(estado==1)
		{
			document.getElementById("enviar").disabled=false;
			estado=0;
		}
		else
		{
			document.getElementById("enviar").disabled=true;
			estado=1;
		}
	}
	
	function AgregarDatos()
	{
		if($("#mismoContacto").attr("checked")) 
		{
			$("#contactoIgual").val("1");
			$("#txtNombCont").val($("#txtNomb").val());
			$("#txtCargoCont").val($("#txtCargo").val());
			$("#txtEmailCont").val($("#txtEmail").val());
			$("#txtEmailContRepetir").val($("#txtEmailRepetir").val());
			$("#txtFonoCont").val($("#txtFono").val());
			
			$('#txtNombCont').attr("readOnly", true);
			$("#txtCargoCont").attr("readOnly", true);
			$("#txtEmailCont").attr("readOnly", true);
			$("#txtEmailContRepetir").attr("readOnly", true);
			$("#txtFonoCont").attr("readOnly", true);
		} 
		else 
		{
			$("#contactoIgual").val("0");
			$("#txtNombCont").val("");
			$("#txtCargoCont").val("");
			$("#txtEmailCont").val("");
			$("#txtEmailContRepetir").val("");
			$("#txtFonoCont").val("");
			
			$('#txtNombCont').attr("readOnly", false);
			$("#txtCargoCont").attr("readOnly", false);
			$("#txtEmailCont").attr("readOnly", false);
			$("#txtEmailContRepetir").attr("readOnly", false);
			$("#txtFonoCont").attr("readOnly", false);
		}
	}
	
	function guardar()
	{
		if($("#frmEmpresa").valid())
		{
							//window.open($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize());
							$('#dialog_espera').dialog('open');
							$.post($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize(),function(d){
																							 $('#dialog_espera').dialog('close');		 																									$('#dialog').dialog('open');			
																						   });
		}
	}
	
	function abreTerminos(){
		//setInterval("window.scrollTo(0,0)",100)

		$.post("solicitud/terminosempresa.asp",
			   function(f){
				    $('#datosTerminos').html(f);
		});
		 $('#datosTerminos').dialog('open');
	}
	</script>
</head>
<body>
<div id="header">
	<h1><img src="images/logo.png"  /></h1>
	<ul>
	</ul>
</div>
<div id="content" class="bg1">
  <h2><em style="text-transform: capitalize;">Solicitud de Inscripción de Empresa</em></h2>
<form name="frmEmpresa" id="frmEmpresa" action="solicitud/insertar.asp" method="post">
	<table cellspacing="0" cellpadding="1" border=0>
    <tr>
    	<td width="105">&nbsp;</td>
      	<td width="200">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="80">&nbsp;</td>
        <td width="200">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="60">&nbsp;</td>
        <td width="200">&nbsp;</td> 
   	</tr>
     <tr>
    	<td colspan="8"><center><h3><em style="text-transform: capitalize;">Antecedentes de la Empresa</em></h3></center></td>
   	</tr>
    <tr>
    	<td>&nbsp;</td>
      	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
    <tr>
    	<td>Rut :</td>
      	<td><input id="txRut" name="txRut" type="text" tabindex="1" size="12" maxlength="11" onblur="RutEmpresa();"/></td>
        <td>&nbsp;</td>
        <td>Raz&oacute;n Social :</td>
        <td><input id="txtRsoc" name="txtRsoc" type="text" tabindex="2" maxlength="99" size="40"/></td>
        <td>&nbsp;</td>
        <td>Giro :</td>
        <td><input id="txtGiro" name="txtGiro" type="text" tabindex="3"  size="40"  maxlength="99"/></td>
   	</tr>
	<tr>
        <td>Direcci&oacute;n :</td>
        <td colspan="1"><input id="txtDir" name="txtDir" type="text" tabindex="4" maxlength="99" size="40"/></td>
        <td>&nbsp;</td>
        <td>Comuna :</td>
        <td><input id="txtCom" name="txtCom" type="text" tabindex="5" size="40" maxlength="40"/></td>
        <td>&nbsp;</td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="6" size="40" maxlength="40"/></td>
    </tr>
	<tr>
       <td>Fono :</td>
       <td><input id="txtFon" name="txtFon" type="text" tabindex="7" size="20" maxlength="20"/></td>
       <td></td>
       <td>Fax :</td>
       <td><input id="txtFax" name="txtFax" type="text" tabindex="8" maxlength="20" size="20"/></td>
       <td colspan="3"></td>
    </tr>
    <tr>
    	<td>Mutual :</td>
        <td><select id="txtMutual" name="txtMutual" tabindex="10"></select></td>
        <td>&nbsp;</td>
        <td>OTIC :</td>
        <td colspan="4"><select id="txtOtic" name="txtOtic" tabindex="11" style="width:37em;"></select></td>
    </tr>
     <tr>
    	<td>&nbsp;</td>
      	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
    <tr>
    	<td colspan="8"><center>
    	  <h3><em style="text-transform: capitalize;">Contacto de Coordinación de Cursos</em></h3></center></td>
   	</tr>
     <tr>
    	<td>&nbsp;</td>
      	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
    <tr>
        <td>Nombre :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="13" maxlength="40" size="40"/></td>
        <td>&nbsp;</td>
        <td></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
     <tr>
       <td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="14" maxlength="50" size="40"/></td>
        <td>&nbsp;</td>
        <td>Teléfono :</td>
        <td><input id="txtFono" name="txtFono" type="text" tabindex="15" size="40" maxlength="50"/></td>
           <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
     <tr>
        <td>Email : </td>
        <td><input id="txtEmail" name="txtEmail" type="text" tabindex="16" maxlength="50" size="40" onpaste="return false" oncopy="return false;"/></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
     <tr>
        <td>Repetir Email :</td>
        <td><input id="txtEmailRepetir" name="txtEmailRepetir" type="text" tabindex="17" maxlength="50" size="40" onpaste="return false" oncopy="return false;"/></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
    <tr>
    	<td>&nbsp;</td>
      	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
    <tr>
    	<td colspan="8"><center>
    	  <h3><em style="text-transform: capitalize;">Contacto de Contabilidad</em></h3></center></td>
   	</tr>
     <tr>
    	<td>&nbsp;</td>
      	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
   <!-- 	<tr>
    	<td colspan="3"><label>
    	<input type="checkbox" name="mismoContacto" id="mismoContacto" onclick="AgregarDatos();" tabindex="18"/>Repetir los Datos del Contacto de Coordinación de Cursos</label></td>
   	</tr>-->
    <tr>
    	<td><input type="hidden" id="contactoIgual" name="contactoIgual" value="0"/></td>
      	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
     <tr>
        <td>Nombre :</td>
        <td><input id="txtNombCont" name="txtNombCont" type="text" tabindex="20" maxlength="33" size="40"/></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
     <tr>
        <td>Cargo :</td>
        <td><input id="txtCargoCont" name="txtCargoCont" type="text" tabindex="21" maxlength="50" size="40"/></td>
        <td>&nbsp;</td>
        <td>Télefono :</td>
        <td><input id="txtFonoCont" name="txtFonoCont" type="text" tabindex="22" size="40" maxlength="50"/></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
    <tr>
    	<td>Email :</td>
      	<td><input id="txtEmailCont" name="txtEmailCont" type="text" tabindex="23" maxlength="50" size="40" onpaste="return false" oncopy="return false;"/></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
     <tr>
    	<td>Repetir Email :</td>
      	<td><input id="txtEmailContRepetir" name="txtEmailContRepetir" type="text" tabindex="24" maxlength="50" size="40" onpaste="return false" oncopy="return false;"/></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
     <tr>
    	<td>&nbsp;</td>
      	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
   	</tr>
	<tr>
    	<td colspan="3"><label>
    	<input type="checkbox" name="terminos" id="terminos" onclick="estadoBoton();" tabindex="25"/>Acepto los <a href="#" onclick="window.scrollBy(0,50);abreTerminos();">T&eacute;rminos y Condiciones</a></label></td>
   	</tr>
    <tr>
    	<td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td align="right"><INPUT TYPE=button OnClick="document.location.href='index.asp';" name="btn_volver" id="btn_volver" VALUE="Volver"></td>
        <td>&nbsp;</td>
        <td><INPUT TYPE=submit OnClick="guardar();" id="enviar" name="enviar" VALUE="Enviar" disabled="disabled"></td>
        <td>&nbsp;</td> 
   	</tr>
</table>
</form>  
<div id="messageBox1" style="height:100px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div>   
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Solicitud de Inscripción Empresa">
	<p>La Solicitud de Inscripción de Empresa fue Enviada Exitosamente. La Aprobación o Rechazo de la Solicitud Sera Notificada a Través de un Correo Electrónico al Coordinador de Cursos de su Organización.</p>
</div>
<div id="dialog_espera" title="Solicitud de Inscripción Empresa">
<center>
<table width="200" border="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><center><font size="2"><b>Espere ...</b></font></center></td>
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
<div id="rut1" title="Solicitud de Inscripción Empresa">
	<p>El Rut de Empresa Ingresado, ya se Encuentra Registrado.</p>
</div>
<div id="datosTerminos" title="Terminos y  Condiciones">
</div>
</body>
</html>
