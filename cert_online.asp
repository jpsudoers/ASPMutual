<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Mutual Capacitación</title>
<meta name="Keywords" content="" />
<meta name="Description" content="" />
<link rel="shortcut icon" href="images/IcoMutual.ico" />
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
var cargada=false;
$(document).ready(function(){	
			$('#txRut').val("");
			$('#txRutT').val("");
			$('#txCodigo').val("");
			$('#ERS1').attr('checked', false);
			$('#ERS0').attr('checked', false);
				
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
		
	$("#dialog_carga").dialog({
			autoOpen: false,
			bgiframe: true,
			height:200,
			width: 500,
			modal: true,
			buttons: {
				Salir: function() {
					$(this).dialog('close');
				}
			}
		});		
		
		$("#Doc2").dialog({
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
			title: 'Certificado de Asistencia a Curso de Capacitación'
		});		
	});	


	function documento2(cod_auto)	
	{
		$.get("online/verificar.asp",{c:cod_auto},
						 function(xml){
									$('DATOS',xml).each(function(i) { 
										if($(this).find('total').text()=='1')
										{
											$("#doc").show();
											$("#espacio_doc").hide();
											$("#ifPagina").attr('src','libroclases/Certificado.asp?prog='+$(this).find('p').text()+'&empresa='+$(this).find('e').text()+'&trabajador='+$(this).find('t').text()+'&relator='+$(this).find('r').text());
										}
										else
										{
											$("#doc").hide();
											$("#espacio_doc").show();
											$("#img_espacio_doc").show();
											$("#lb_espacio_doc").show();
										}
						});
			});
	}

	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}

	function tabla()
	{
		$("#tabladoc").show();
		$('#espacio_doc').hide();
		if(!cargada){
			jQuery("#list1").jqGrid({ 
			url:'consultascertificados/listar_cur_abiertos.asp?trabajador='+$('#txRutT').val()+'&emp='+$('#txRut').val(), 
			datatype: "xml", 
			colNames:['Rut/ID','Nombre Trabajador', 'Curso', 'Fecha Curso', 'Vigencia', 'Est. Cert.', '&nbsp;'], 
			colModel:[
					   {name:'RUT',index:'T.RUT', width:22},
					   {name:'NOMBRES',index:'T.NOMBRES', width:60}, 
					   {name:'R_SOCIAL',index:'E.R_SOCIAL', width:60}, 
					   {name:'FECHA_TERMINO',index:'P.FECHA_TERMINO', width:20, align:"center"}, 
					   {name:'vigencia',index:'vigencia', width:20, align:"center"}, 
					   {name:'estado',index:'estado', width:20, align:"center"}, 					   
					   { align:"right",editable:true, width:4}], 
			rowNum:100, 
			autowidth: true, 
			rowList:[100,150,200], 
			pager: jQuery('#pager1'), 
			sortname: 'T.NOMBRES', 
			viewrecords: true, 
			sortorder: "asc", 
			caption:"Listado de Certificados" 
			}); 
	
			jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});

			cargada=true;
		}else{
jQuery("#list1").jqGrid('setGridParam',{url:"consultascertificados/listar_cur_abiertos.asp?trabajador="+$('#txRutT').val()+"&emp="+$('#txRut').val()}).trigger("reloadGrid")
		}
	}
	
	function certificados(programa,empresa,trabajador,relator){
		//alert(programa+' '+empresa+' '+trabajador+' '+relator);
		documento("libroclases/Certificado.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);
	}
	
	function documento(arch){
		$("#ifPagina2").attr('src',arch);
		if(!$('#Doc2').dialog('isOpen'))
			$('#Doc2').dialog('open');
	}
	/*function guardar()
	{
		if($("#frmEmpresa").valid())
		{
							//window.open($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize());
							$('#dialog_espera').dialog('open');
							$.post($('#frmEmpresa').attr('action')+'?'+$('#frmEmpresa').serialize(),function(d){
																							 $('#dialog_espera').dialog('close');		 																									$('#dialog').dialog('open');			
																						   });
		}
	}*/
	</script>
</head>
<body>
<div id="header">
	<h1><img src="images/logo.png"  /></h1>
    <ul>
        <li><a href="index.asp" accesskey="3" class="selItem">Volver</a></li>
    </ul>
</div>
<div id="content" class="bg1">
  <h2><em style="text-transform: capitalize;">Certificados On-Line</em></h2>
<!--<form name="frmEmpresa" id="frmEmpresa" action="solicitud/insertar.asp" method="post">-->
	<table cellspacing="2" cellpadding="4" border=0>
    <tr>
    	<td width="210">&nbsp;</td>
      	<td width="270">&nbsp;</td>
        <td width="540">&nbsp;</td>
   	</tr>
    <!----><tr>
    	<td colspan="3"><center><h3><em style="text-transform: capitalize;">Certificados de Asistencia a Cursos</em></h3></center></td>
   	</tr>
    <!--<tr>
    	<td colspan="3">&nbsp;</td>
   	</tr>-->
    <tr> <!--id="pregRnv" style="display:none"-->
       <td><strong>Seleccione Tipo de Busqueda :</strong></td>
       <td colspan="2"><input type="radio" name="ERS" id="ERS1" value="1" onclick="$('#ER').val('1');$('#bCodAuten').show();$('#bRutET').hide();$('#tabladoc').hide();;$('#espacio_doc').show();$('#txRutT').val('');$('#txRut').val('');"/>Por C&oacute;digo de Antentificaci&oacute;n	<input type="radio" name="ERS" id="ERS0" value="0" onclick="$('#ER').val('0');$('#bCodAuten').hide();$('#bRutET').show();$('#doc').hide();$('#espacio_doc').show();$('#txCodigo').val('');//$('#dialog_carga').dialog('open');"/>Por Rut de Empresa y Trabajador</td><!--$('#ER').val('0');$('#bCodAuten').hide();$('#bRutET').show();$('#doc').hide();$('#espacio_doc').show();$('#txCodigo').val('');-->
     </tr>      
    <tr id="bCodAuten" style="display:none">
    	<td><strong>Ingrese C&oacute;digo de Autentificaci&oacute;n :</strong></td>
      	<td><input id="txCodigo" name="txCodigo" type="text" tabindex="1" size="30" maxlength="23"/></td>
        <td><input id="enviar" name="enviar" VALUE="Validar" type="button" OnClick="documento2($('#txCodigo').val());" /></td>
   	</tr>
    <tr id="bRutET" style="display:none">
    	<td><strong>Ingrese Rut de Empresa :</strong></td>
      	<td><input id="txRut" name="txRut" type="text" tabindex="1" maxlength="14" size="13"/>&nbsp;&nbsp;Ej: 16313889-4</td>
        <td><strong>Ingrese Rut del Trabajador :&nbsp;&nbsp;&nbsp;</strong><input id="txRutT" name="txRutT" type="text" tabindex="2" maxlength="14" size="15"/>&nbsp;&nbsp;Ej: 16313889-4&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input id="enviar" name="enviar" VALUE="Validar" type="button" OnClick="tabla();" /></td>
   	</tr>    
    <!--<tr>
    	<td colspan="3">&nbsp;</td>
   	</tr>-->
    <tr>
    	<td colspan="3" style="display:none" height="350" id="doc"><iframe style="width:100%;height:100%" id="ifPagina"></iframe></td>
        <td colspan="3" style="display:none" height="350" id="tabladoc">
	    	<table id="list1"></table> 
        	<div id="pager1"></div>
        </td>
   	</tr>
    <tr>
    	<td colspan="3" height="350" id="espacio_doc"><center><img src="images/error_busqueda.png" width="150" height="160" id="img_espacio_doc" name="img_espacio_doc" style="display:none"/><br /><label id="lb_espacio_doc" style="display:none; color:#F00" >No sean han encontrado resultados.</label></center></td>
   	</tr>    
</table>
<!--</form>  -->
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="Doc2" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina2"></iframe>
</div>
<div id="dialog_carga" title="Información">
<table width="400" border="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><font size="2"><b>Por favor acceda al sistema con su usuario y clave para obtener información.</b></font></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</div>
<!--<div id="dialog_espera" title="Solicitud de Inscripción Empresa">
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
</div>-->
</body>
</html>