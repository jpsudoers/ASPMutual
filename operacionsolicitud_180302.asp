<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title></title>
<meta name="Keywords" content="" />
<meta name="Description" content="" />

<link href="css/default.css" rel="stylesheet" type="text/css" />
<link href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" type="text/css" href="css/ui.all.css"/>
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<!--script src="js/jquery.js" type="text/javascript"></script-->
<script src="js/jquery3.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui3.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript" src="js/jquery.mask.min.js"></script>



<script type="text/javascript">
estado=1;

$(document).ready(function(){	


	$('.telefonos').mask('(00) 0000-0000');		
	$('.telefonos2').mask('(00) 0000-0000');				

	llena_mutual(0);
	$("#txRut").focus();
	//validar();
	
	
	$("#rut1").dialog(
		{
		  modal: true,
		  bgiframe: false,
		  autoOpen: false,
		  height:340,
		  width: 500,
		  buttons: {
			Ok: function() {
			  $( this ).dialog( "close" );
			}
		  }
		}
	);
	
	
	
		$("#dialog").dialog({
			bgiframe: false,
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
			bgiframe: false,
			autoOpen: false,
			height: 200,
			width: 600,
			modal: true
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

		
});	// (document).ready(function()

llena_comuna();

function llena_comuna(){
	$("#txtCom").html("");
	//alert(id_mutual)
	$.get("solicitud/comuna.asp",
	function(xml){
		$("#txtCom").append("<option value=\"\">Seleccione..</option>");
		$('row',xml).each(function(i) 
		{ 
			
			$("#txtCom").append('<option value="'+$(this).find('id').text()+'" >'+ $(this).find('desc').text()+ '</option>');

		});

	});
}



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
							
						}
			});
		});
	}
	
	
	function validarF(){
		var men = '';
		if ( $("#txRut").val().length < 9 ){ 				men = men + '- Ingrese Rut.<br>'}
		if ( $("#txtRsoc").val().length < 5 ){ 				men = men + '- Ingrese Razón Social Empresa.<br>'}
		if ( $("#txtGiro").val().length < 5 ){ 				men = men + '- Ingrese el Giro Empresa.<br>'}
		if ( $("#txtDir").val().length < 5 ){ 				men = men + '- Ingrese Dirección Empresa.<br>'}
		//if ( $("#txtCom").length() < 5 ){ 				men = men + '- Ingrese Comuna Empresa.'}
		if ( $("#txtCiu").val().length < 5 ){ 				men = men + '- Ingrese Ciudad Empresa.<br>'}
		if ( $("#txtFon").val().length < 5 ){ 				men = men + '- Ingrese Teléfono Empresa.<br>'}
		//if ( $("#txtMutual").length() < 5 ){ 			men = men + '- Seleccione Mutual a la que Pertenece Empresa.'}
		if ( $("#txtCargo").val().length < 5 ){				men = men + '- Ingrese Cargo Contacto.<br>'}
		if ( $("#txtEmail").val().length < 5 ){				men = men + '- Ingrese Correo Electrónico Contacto.<br>'}
		if ( $("#txtEmailRepetir").val().length < 5 ){ 		men = men + '- Ingrese Nuevamente el Correo Electrónico del Contacto.<br>'}
		if ( $("#txtFono").val().length < 5 ){ 				men = men + '- Ingrese Teléfono del Contacto.<br>'}
		if ( $("#txtNombCont").val().length < 5 ){ 			men = men + '- Ingrese Nombre del Contacto de Contabilidad.<br>'}
		if ( $("#txtEmailCont").val().length < 5 ){ 		men = men + '- Ingrese Correo Electrónico del Contacto de Contabilidad.<br>'}
		if ( $("#txtEmailContRepetir").val().length < 5 ){ 	men = men + '- Ingrese Nuevamente Correo Electrónico del Contacto de Contabilidad.<br>'}
		if ( $("#txtFonoCont").val().length < 5 ){ 			men = men + '- Ingrese Teléfono del Contacto de Contabilidad.<br>'}
		
		return men
		
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
	


	function guardar()
	{
		var menr = RutEmpresa();	
		var mens = validarF();


		if( $('#terminos').prop('checked') == false ) {
			alert('Acepte los terminos y condiciones');
			return;
		}
		
		if (mens!=''){
			$( "#validaF" ).html(mens);
			$( "#validaF" ).dialog(
				{
				  modal: true,
				  bgiframe: false,
				  height:340,
				  width: 500,
				  buttons: {
					Ok: function() {
					  $( this ).dialog( "close" );
					}
				  }
				}
			);
			return;	
		}
		
		
	

		//window.open($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize());
		$('#dialog_espera').dialog('open');
		$.post($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize(),function(d){
		$('#dialog_espera').dialog('close');
		$('#dialog').dialog('open');			
		});

		
		
		
/*		if($("#frmEmpresa").valid())
		{
						window.open($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize());
						$('#dialog_espera').dialog('open');
						$.post($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize(),function(d){
						$('#dialog_espera').dialog('close');
						$('#dialog').dialog('open');			
						});
		}*/
	}
	
	function abreTerminos(){
		//setInterval("window.scrollTo(0,0)",100)

		$.post("solicitud/terminosempresa.asp",
			   function(f){
				    $('#datosTerminos').html(f);
		});
		 $('#datosTerminos').dialog('open');
	}
	$('#existeRut').dialog('open');
	



function llena_mutual(id_mutual){
	$("#txtMutual").html("");
	//alert(id_mutual)
	$.get("empresa/mutual.asp",
	function(xml){
		$("#txtMutual").append("<option value=\"\">Seleccione</option>");
		$('row',xml).each(function(i) { 
			if(id_mutual==$(this).find('ID_MUTUAL').text())
			{
				$("#txtMutual").val($(this).find('ID_MUTUAL_').text());
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
	
	</script>
</head>
<body>
<div id="validaF" title="Validaciones">
</div>

<div id="header">
	<h1><img src="images/logo.png"  /></h1>
	<ul>
	</ul>
</div>
<div id="content" class="bg1">
  <h2><em style="text-transform: capitalize;">Solicitud de Inscripción de Empresa</em></h2>
<form name="frmEmpresa" id="frmEmpresa" action="solicitud/insertar.asp" method="post" onSubmit="return false">
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
    	<td colspan="8">&nbsp;</td>
   	</tr>
    <tr>
    	<td>Rut :</td>
      	<td><input id="txRut" name="txRut" type="text" tabindex="1" size="12" maxlength="11" onblur="RutEmpresa();" placeholder="Ej: 11111111-1"/></td>
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
        <td><select name="txtCom" id="txtCom"></select><!--input id="txtCom" name="txtCom" type="text" tabindex="5" size="40" maxlength="40"/--></td>
        <td>&nbsp;</td>
        <td>Ciudad :</td>
        <td><input id="txtCiu" name="txtCiu" type="text" tabindex="6" size="40" maxlength="40"/></td>
    </tr>
	<tr>
       <td>Fono :</td>
       <td><input id="txtFon" name="txtFon" type="text" tabindex="7" size="20" maxlength="20" class="telefonos"/>

           <label>
             <input name="FM1" type="radio" id="FM1_0" value="1" checked="checked" />
             Fijo</label>

           <label>
             <input type="radio" name="FM1" value="2" id="FM1_1" />
             Móvil</label>
</td>
       <td></td>
       <td>Fax :</td>
       <td><input id="txtFax" name="txtFax" type="text" tabindex="8" maxlength="20" size="20"  class="telefonos"/></td>
       <input type="hidden" id="u" />
       <td>&nbsp;</td>
       <td>Mutual :</td>
       <td><select id="txtMutual" name="txtMutual" tabindex="10"></select></td>
    </tr>
<!--<tr>
    	<td>Mutual :</td>
        <td><select id="txtMutual" name="txtMutual" tabindex="10"></select></td>
        <td>&nbsp;</td>
        <td>OTIC :</td>
        <td colspan="4"><select id="txtOtic" name="txtOtic" tabindex="11" style="width:37em;"></select></td>
    </tr>-->
     <tr>
    	<td colspan="8">&nbsp;</td>
   	</tr>
    <tr>
    	<td colspan="8"><center>
    	  <h3><em style="text-transform: capitalize;">Contacto de Coordinación de Cursos</em></h3></center></td>
   	</tr>
    <tr>
    	<td colspan="8">&nbsp;</td>
   	</tr>
    <tr>
        <td>Nombre :</td>
        <td><input id="txtNomb" name="txtNomb" type="text" tabindex="13" maxlength="40" size="40"/></td>
        <td>&nbsp;</td>
        <td>Cargo :</td>
        <td><input id="txtCargo" name="txtCargo" type="text" tabindex="14" maxlength="50" size="40"/></td>
        <td>&nbsp;</td>
        <td>Teléfono :</td>
        <td><input id="txtFono" name="txtFono" type="text" tabindex="15" size="20" maxlength="20"  class="telefonos2"/>

           <label>
             <input name="FM2" type="radio" id="FM2_0" value="1" />
             Fijo</label>

           <label>
             <input name="FM2" type="radio" id="FM2_1" value="2" checked="checked" />
        Móvil</label></td>
    </tr>
     <tr>
        <td>Email : </td>
        <td><input id="txtEmail" name="txtEmail" type="text" tabindex="16" maxlength="50" size="40" onpaste="return false" oncopy="return false;"/></td>
        <td colspan="6">&nbsp;</td>
    </tr>
     <tr>
        <td>Repetir Email :</td>
        <td><input id="txtEmailRepetir" name="txtEmailRepetir" type="text" tabindex="17" maxlength="50" size="40" onpaste="return false" oncopy="return false;" onblur="igualdad(1);"/></td>
        <td colspan="6"><div id="msj1" style="color:#F00"></div></td>
    </tr>
    <tr>
    	<td colspan="8">&nbsp;</td>
   	</tr>
    <tr>
    	<td colspan="8"><center>
    	  <h3><em style="text-transform: capitalize;">Contacto de Contabilidad</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="8">&nbsp;</td>
   	</tr>

    <tr>
    	<td colspan="8"><input type="hidden" id="contactoIgual" name="contactoIgual" value="0"/></td>
   	</tr>
     <tr>
        <td>Nombre :</td>
        <td><input id="txtNombCont" name="txtNombCont" type="text" tabindex="20" maxlength="33" size="40"/></td>
        <td>&nbsp;</td>
        <td>Cargo :</td>
        <td><input id="txtCargoCont" name="txtCargoCont" type="text" tabindex="21" maxlength="50" size="40"/></td>
        <td>&nbsp;</td>
        <td>Télefono :</td>
        <td><input id="txtFonoCont" name="txtFonoCont" type="text" tabindex="22" size="20" maxlength="20"  class="telefonos2"/>

           <label>
             <input name="FM3" type="radio" id="FM3_0" value="1" />
             Fijo</label>

           <label>
             <input name="FM3" type="radio" id="FM3_1" value="2" checked="checked" />
        Móvil</label></td>
    </tr>
    <tr>
    	<td>Email :</td>
      	<td><input id="txtEmailCont" name="txtEmailCont" type="text" tabindex="23" maxlength="50" size="40" onpaste="return false" oncopy="return false;"/></td>
        <td colspan="6">&nbsp;</td>
   	</tr>
     <tr>
    	<td>Repetir Email :</td>
      	<td><input id="txtEmailContRepetir" name="txtEmailContRepetir" type="text" tabindex="24" maxlength="50" size="40" onpaste="return false" oncopy="return false;" onblur="igualdad(2);"/></td>
        <td colspan="6"><div id="msj2" style="color:#F00"></div></td>
   	</tr>
     <tr>
    	<td colspan="8">&nbsp;</td>
   	</tr>
	<tr>
    	<td colspan="8"><label>
    	<input type="checkbox" name="terminos" id="terminos"  tabindex="25"/>Acepto los <a href="#" onclick="window.scrollBy(0,50);abreTerminos();">T&eacute;rminos y Condiciones</a></label></td>
   	</tr>
    <tr>
    	<td colspan="4">&nbsp;</td>
        <td align="right"><INPUT TYPE=button OnClick="document.location.href='index.asp';" name="btn_volver" id="btn_volver" VALUE="Volver"></td>
        <td>&nbsp;</td>
        <td><INPUT TYPE=submit OnClick="guardar();" id="enviar" name="enviar" VALUE="Enviar"></td>
        <td>&nbsp;</td> 
   	</tr>
</table>
</form>  
</div>
<div id="messageBox1" style="height:100px;overflow:auto;width:400px;"><ul></ul></div>   

<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>

<div id="dialog" title="Solicitud de Inscripción Empresa">
	<p>La Solicitud de Inscripción de Empresa fue Enviada Exitosamente. La Aprobación o Rechazo de la Solicitud Sera Notificada a Través de un Correo Electrónico al Coordinador de Cursos de su Organización.</p>
</div>

<div id="dialog_espera" title="Solicitud de Inscripción Empresa">
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
</div>


<div id="dialog_termino" title="Aceptar los Términos y Condiciones"></div>


<div id="rut1" title="Solicitud de Inscripción Empresa">
	<p>El Rut de Empresa Ingresado, ya se Encuentra Registrado.</p>
</div>
<div id="datosTerminos" title="Terminos y  Condiciones"></div>


<script>

function igualdad(val){
	var mail1='';
	var mail2='';
	if (parseInt(val) == 1){
		mail1 = $('#txtEmail').val();
		mail2 =$('#txtEmailRepetir').val();
		
		if (mail1 != mail2){
			$('#msj1').html('Los Email no coinciden..');
			$('#txtEmailRepetir').val('') ;
			$('#txtEmail').focus();
		}else{$('#msj1').html('');}	
	}
	
	if (parseInt(val) == 2){
		mail1 = $('#txtEmailCont').val();
		mail2 = $('#txtEmailContRepetir').val();
		
		if (mail1 != mail2){
			$('#msj2').html('Los Email no coinciden..');
			$('#txtEmailContRepetir').val('') ;
			$('#txtEmailCont').focus();
		}else{$('#msj2').html('');}
		
	}
	
	
}
$('#dialog').dialog('close');
$('#dialog_espera').dialog('close');


</script>
</body>
</html>
