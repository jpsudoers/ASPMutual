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
<link rel="stylesheet" type="text/css" href="css/ui.all.css"/>
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
<script src="js/jquery.jqGrid.js" type="text/javascript" charset="utf-8"></script>
<script src="js/ajaxfileupload.js" type="text/javascript" charset="utf-8"></script>
<style type="text/css">
    .suggestionsBox {
        position: relative;
        left: 10px;
        margin: 10px 0px 0px 0px;
        width: 500px;
        background-color: #212427;
        -moz-border-radius: 7px;
        -webkit-border-radius: 7px;
        border: 2px solid #000;   
        color: #fff;
		height:150px;
		overflow:auto;
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
	
	.suggestionsBox2 {
        position: relative;
        left: 10px;
        margin: 10px 0px 0px 0px;
        width: 500px;
        background-color: #212427;
        -moz-border-radius: 7px;
        -webkit-border-radius: 7px;
        border: 2px solid #000;   
        color: #fff;
		height:200px;
		overflow:auto;
    }
   
    .suggestionList2 {
        margin: 0px;
        padding: 0px;
    }
   
    .suggestionList2 li {
       
        margin: 0px 0px 3px 0px;
        padding: 3px;
        cursor: pointer;
    }
   
    .suggestionList2 li:hover {
        background-color: #659CD8;
    }
	
	.fto_adjunta_table
	{
		float: left; 
		border:0px; 
		width:100px; 
		font-size: 11px; 
		font-family: arial;
		text-align: center;
	}

	.fto_adjunta {
	float: left;
	width: 100px;
	height: 100px;
	background-image: url(images/img_ejemplo.jpg);
	margin: 13px;
	}

	.eliminar_img_doc {
		background-color: #1f2e5e;
		width: 20px;
		height: 20px;
		float: right;
		padding: 7px;
		padding-right: 7px;
		padding-bottom: 7px;
		padding-right: 7px;
		padding-bottom: 7px;
		padding-right: 10px;
		padding-bottom: 10px;
		position: absolute;
	}

</style>
<script type="text/javascript">
var amtval="";
var fran="";
var i=1;
var costo=0;
var escolaridades=0;
var rutRepetido=0;
var estadoFrmTrab=0;
var rutTrabAsig="";
var cargada = false;

$(document).ready(function(){					
 
	llena_estado();
	
	$("#dialog_carga").dialog({
			autoOpen: false,
			bgiframe: true,
			height:500,
			width: 1200,
			modal: true,
			buttons: {
				Salir: function() {
					$(this).dialog('close');
				}
			}
		});

	$("#dialog").dialog({
			autoOpen: false,
			bgiframe: true,
			height:600,
			width: 800,
			modal: true,
			overlay: {
				backgroundColor: '#f00',
				opacity: 0.5
			},
			buttons: {
				'Aprobar': function() {
					 cambiarEstado(true)
				},
				'Rechazar': function() {
					 cambiarEstado(false)
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	
 	$("#dialog-eliminar").dialog({
		  autoOpen: false,
		  bgiframe: true,
		  height: "auto",
		  width: 400,
		  modal: true,
		  buttons: {
			"Borrar Solicitud": function() {
			
				$('#operacion').val('delSolicitud');
				$.post("solicitud/solicitudcreditoAjax.asp",
							{idSolicitud:$('#solId').val(), operacion: $('#operacion').val()},
									function(d){
										  $("#list1").trigger("reloadGrid",[{page:1}]);							  
									});
			
			  $(this).dialog('close');
			},
			Cancel: function() {
			  $(this).dialog('close');
			}
		  }
		});
	
	$("#mensaje").dialog({
			autoOpen: false,
			bgiframe: true,
			height:220,
			width: 650,
			modal: true,
			buttons: {
				Aceptar: function() {
					$(this).dialog('close');
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

	$("#dialog_mensaje").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 200,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
					$('#dialog_mensaje').dialog('close');
					$("#list1").trigger("reloadGrid",[{page:1}]);
					 
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

	function validarExcel(){
		$("#frmCargaTrab").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txtXLS: {
                required: true,
                accept: "xls"                          
            }
		},
		messages:{
			txtXLS: {
                required:"&bull; Seleccione Documento",
                accept:"&bull; El Formato del Documento seleccionado no es Válido."                          
            }
		}
	});
    }

	function validar(){
		$("#frmProgEmp").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txRut:{
				required:true
			},
			Compromiso:{
				required:true
			},
			txtNum:{
				required:true
			},
			txtDoc: {
                required: true,
                accept: "xls|doc|pdf|jpg|tif|png|gif"                          
            },
			NParticipantes:{
				max: parseInt($('#Vacantes').val()),
                min: 1
			},
			COfran:{
				required:true
			},
			ConOtic:{
				required:true
			},
			NReg_Fran:{
				required: {depends:function(element){return $("#COtic0").is(":checked")}},
				number: true
			},
			docExiste:{
				required:true
			},
			proPert:{
				required: {depends:function(element){if($("#id_curso").val()=="39" || $("#id_curso").val()=="40") return true ;else return false;}}
			},
			txtOtic:{
				required: {depends:function(element){return $("#COtic1").is(":checked")}}
			}
		},
		messages:{
			txRut:{
				required:"&bull; Ingrese Rut Empresa."
			},
			Compromiso:{
				required:"&bull; Seleccione el Tipo Compromiso."
			},
			txtNum:{
				required:"&bull; Ingrese Número de Documento."
			},
			txtDoc: {
                required:"&bull; Seleccione Documento de Compromiso",
                accept:"&bull; El Formato del Documento de Compromiso no es Válido."                          
            },
			NParticipantes:{
				max: "&bull; El Número de Participantes Supero las Vacantes Disponibles para este Curso (Vacantes Disponibles : "+$('#Vacantes').val()+")",
                min: "&bull; No hay Participantes Ingresados"
			},
			COfran:{
				 required:"&bull; Seleccione SI o No dependiendo si Utilizara Franquicia Sence"
			},
			ConOtic:{
				 required:"&bull; Seleccione SI o No dependiendo si el curso está inscrito por OTIC"
			},
			NReg_Fran:{
				required:"&bull; Ingrese Número de Registro Sence",
				number: "&bull; En el N° de Registro Sence solo deben Ingresar Números."
			},
			docExiste:{
				required:"&bull; Número de Documento ya Utilizado."
			},
			proPert:{
				required:"&bull; Ingrese proyecto al cual pertenece."
			},
			txtOtic:{
				required:"&bull; Seleccione Otic."
			}
		}
	});
    }

  
	function fill(id,rut) {
	   $('#Empresa').val(id);
	   $('#txRut').val(rut);
	   $('#suggestions').hide();
	}

	function formatCurrency(num) {
		num = num.toString().replace(/$|,/g,'');
		if(isNaN(num))
		{
		num = "0";
		}
		sign = (num == (num = Math.abs(num)));
		num = Math.floor(num*100+0.50000000001);
		cents = num%100;
		num = Math.floor(num/100).toString();
		if(cents<10)
		{
		cents = "0" + cents;
		}
		for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
		{
		num = num.substring(0,num.length-(4*i+3))+'.'+num.substring(num.length-(4*i+3));
		}
	  return (((sign)?'':'-') + '$ ' + num);
	}
	
 
	function validarFrmTrab(){
		$("#frmTrabajador").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txtRutTrab:{
				required:true,
				rut:true
			},
			txtPasTrab:{
				required:true
			},
			txtNomTrab:{
				required:true
			},
			txtAPaterTrab:{
				required:true
			},
			txtAMaterTrab:{
				required:true
			},
			txtCargoTrab:{
                		required:true                        
            		},
	    		txtMail:{
				required:true/*,
				email:true*/
			},
	    		txtEmailTrab:{
				required:true,
				email:true
			},	
			insGrilla:{
                required:true                        
            },
			insTbjCurso:{
                required:true                        
            }
		},
		messages:{
			txtRutTrab:{
				required:"&bull; Ingrese Rut del Trabajador.",
				rut:"&bull; Rut No Valido."
			},
			txtPasTrab:{
				required:"&bull; Ingrese el Número de Pasaporte del Trabajador."
			},
			txtNomTrab:{
				required:"&bull; Ingrese Nombre del Trabajador."
			},
			txtAPaterTrab:{
				required:"&bull; Ingrese Número Apellido Paterno del Trabajador."
			},
			txtAMaterTrab:{
				required:"&bull; Ingrese Número Apellido Materno del Trabajador."
			},
			txtCargoTrab: {
                		required:"&bull; Ingrese Cargo del Trabajador."                        
            		},
	    		txtMail:{
				required:"&bull; Ingrese Teléfono."/*,
				email:"&bull; Correo No Valido"*/
			},
			txtEmailTrab:{
				required:"&bull; Ingrese Correo Electrónico del Trabajador.",
				email:"&bull; Correo No Valido"
			},
			insGrilla:{
                required:"&bull; Trabajador ingresado en Inscripción Actual."              
            },
			insTbjCurso:{
                required:"&bull; Trabajador con Inscripción Autorizada o pendiente de Autorizar."              
            }
		}
	});
	}
	
	function eliminar(idSolicitud)
	{
		$('#solId').val(id);
		
		$("#dialog-eliminar").dialog("open");
	}
	
 
  
	function update(id,rutemp,nomemp,estado,linea,obs1,obs2){
		 $('#operacion').val('upEstado');
		 $('#solId').val(id);
		 $('#lblRutEmpresa').text(rutemp); 
		 $('#lblNombreEmpresa').text(nomemp); 
		 $('#estadoOriginal').val(estado);

		var Date =  linea;
		var elem = Date.split('/');
		$('#lineaCred1').text(elem[0]);
		$('#lineaCred2').text(elem[1]);
		$('#lineaCred3').text(elem[2]);

		$('#lineaCred4').text(elem[3]);
		$('#lineaCred5').text(elem[4]);
		$('#lineaCred6').text(elem[5]);

		$('#lineaCred7').text(elem[6]);
		$('#lineaCred8').text(elem[7]);
		$('#lineaCred9').text(elem[8]);

		$('#lineaCred10').text(elem[9]);
		$('#lineaCred11').text(elem[10]);
		$('#lineaCred12').text(elem[11]);

		$('#lineaCred13').text(elem[12]);

		$('#lineaCred14').text(obs1);
		$('#lineaCred15').text(obs2);

		 //cargadocumentos(id);
		 $('#dialog').dialog('open');

	}
	
	function cambiarEstado(aceptaRechaza)
	{
		 
		estadoOriginal = $('#estadoOriginal').val();

		valoractualizo = "4";

		if(estadoOriginal == '1') aceptaRechaza ? valoractualizo = "2" : valoractualizo = "4";
		if(estadoOriginal == '2') aceptaRechaza ? valoractualizo = "3" : valoractualizo = "4";
		if(estadoOriginal == '4') aceptaRechaza ? valoractualizo = "1" : valoractualizo = "4";


		$('#selEstadoSolicitud').val(valoractualizo);


		$('#operacion').val('upEstado');
		$('#dialog_espera').dialog('open');
		
		$.post($('#frmSolicitud').attr('action')+'?'+$('#frmSolicitud').serialize(),function(d){
			$('#dialog_espera').dialog('close');
			$('#dialog_mensaje').dialog('open');
			$('#dialog').dialog('close');			
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
		$.post("cambioContrasena/frmContrasenaEmpresa.asp",
			   function(f){
				   validarFrm();
				   $('#pantContrasena').html(f);
				   validarFrm();
		});
		$('#pantContrasena').dialog('open')
	}
	
	function bloqueo(){
      $("#txtMensaje").html("Estimado Cliente, su empresa está bloqueada para solicitar acceder. Para mayor información en: <br/><br/>&bull; Contactarse al Fono 55 2 651276 o al correo <a href='mailto:cobranza.masc@mutualasesorias.cl' target='_parent'>cobranza.masc@mutualasesorias.cl.</a>");
$("#mensaje").dialog('open');
	}
	
	function cierraSession(){
		document.location.href='index.asp';
	}

	function cargadocumentos(id){

		$('#documentosContent').html("");
		$.post("solicitud/solicitudcreditoAjax.asp", {solId:id, operacion: 'listaDocumentos', borrar: '0'}, function(data){
						if(data.length >0) {
								$('#documentosContent').html(data);
						}
				});
	}
 
	function llena_estado(){
		$("#solEstado").html("");
		$.get("solicitud/solicitudcreditoAjax.asp?operacion=listaEstados",
					function(xml){
						$("#solEstado").append("<option value='0'>Todas</option>");
							
						$('row',xml).each(function(i) { 
							
							$("#solEstado").append('<option value="'+$(this).find('ID_ESTADO').text()+'" >'+
																		$(this).find('NOMBRE_ESTADO').text()+ '</option>');
						});
					});
		tabla(0);
		cargada = true;
	}
	
		function tabla(id)
		{
			
			if(cargada){

				sUrl = 'solicitud/listarsolcredito.asp?estado=' + id
				jQuery("#list1").jqGrid('setGridParam',{url:sUrl});
				$("#list1").trigger("reloadGrid",[{page:1}]);
			} else
			{
				jQuery("#list1").jqGrid({ 
				url:'solicitud/listarsolcredito.asp', 
				datatype: "xml", 
				colNames:['Cod. Sol.','Rut Empresa', 'Nombre Empresa', 'Fecha Solicitud', 'Estado','&nbsp;','&nbsp;'], 
				colModel:[
						{name:'Cod. Sol.',index:'ID_SOLICITUD', width:10}, 
						{name:'Rut Empresa',index:'RUT_EMPRESA', width:20}, 
						{name:'Nombre Empresa',index:'NOMBRE_EMPRESA', width:20}, 
						{name:'Fecha Solicitud',index:'FECHA_SOLICITUD', width:20}, 
						{name:'Estado',index:'ESTADO_DESC', width:20}, 
						{align:"right",editable:true, width:5}, 
						{align:"right",editable:true, width:5}], 
				rowNum:50, 
				height:350,
				autowidth: true, 
				rowList:[50,100,200], 
				pager: jQuery('#pager1'), 
				sortname: 'SOLICITUDES_CREDITO.FECHA_SOLICITUD', 
				viewrecords: true, 
				sortorder: "asc", 
				caption:"Listado de Solicitudes de Crédito" 
				}); 

				jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
				jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
																title:"Refrescar", 
																buttonicon :'ui-icon-refresh', 
																onClickButton:function(){ 
																	tabla($("#solEstado").val());
																} }); 


				}
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
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO5,PERMISO6,PERMISO12 from USUARIOS "
	sql = sql&" where USUARIOS.ID_USUARIO='"&Session("usuarioMutual")&"'"

   DATOS.Open sql,oConn
   WHILE NOT DATOS.EOF
		if(DATOS("PERMISO1")<>"0")then
		%>
		<li><a href="administracion.asp" accesskey="1" title="">Administraci&oacute;n</a></li>
        <%end if
		if(DATOS("PERMISO2")<>"0")then
		%>
		<li><a href="operacion.asp" accesskey="2" title="" class="selItem">Operaci&oacute;n</a></li>
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
		if(DATOS("PERMISO12")<>"0")then
		%>
		<li><a href="manuales.asp" accesskey="6" title="">Manuales</a></li>
        <%end if		
	 DATOS.MoveNext
   WEND
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
				<li class="first"><a href="operacionsolempresa.asp">Solicitud de Nueva Empresa</a></li>
<li><a href="solicitudescredito.asp">Solicitud Crédito</a></li>
                <li><a href="operacionprogramacion.asp">Programación de Cursos</a></li>
                <li><a href="operacioncalendario.asp">Inscripción de Cursos</a></li>
				<li><a href="operacioninspendientes.asp">Autorizar Inscripciones</a></li>
				<li><a href="operacioninsactivas.asp">Inscripciones Autorizadas</a></li>
                <li><a href="operacioncierre.asp">Revisión y Cierre</a></li>
                <li><a href="operacionhistins.asp">Historico de Inscripciones</a></li>                
              
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Solicitudes de Creditos</em></h2>
			</br>
			<table width="300" border="0">
              <tbody>
              <tr>
 				<td>Estado :</td>
                <td colspan="4"><select id="solEstado" name="solEstado" tabindex="3" onchange="tabla(this.value);"></select></td>              
              </tr>
            </tbody>
			</table>
			</br>
            <table id="list1"></table> 
            <div id="pager1"></div> 
            <br />
            <h2><label style="font-size:14px; text-transform: none;color: #31576F; font-style:italic"><B></B></label></h2>
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Cambiar Estado Solicitud">
	<form name="frmSolicitud" id="frmSolicitud" action="solicitud/solicitudcreditoAjax.asp" method="post">
		<table cellspacing="3" cellpadding="3" border=0 width="100%">
		<tr>
			<td width="12%">&nbsp;</td>
			<td width="12%">&nbsp;</td>
			<td width="12%">&nbsp;</td>
			<td width="12%">&nbsp;</td>
			<td width="12%">&nbsp;</td>
			<td width="12%">&nbsp;</td>
			<td width="12%">&nbsp;</td>
			<td width="12%">&nbsp;</td>
		</tr>
		 <tr>
			<td colspan="8">
				<input type="hidden" id="solId" name="solId" /> 
				<input type="hidden" id="operacion" name="operacion" value="upEstado"/>
				<input type="hidden" id="selEstadoSolicitud" name="selEstadoSolicitud" />
				<input type="hidden" id="estadoOriginal" name="estadoOriginal" />
			</td>
		</tr>
		<tr>
			<td colspan="3"><b>Rut Empresa:</b> <label id="lblRutEmpresa" name="lblRutEmpresa"> </label></td>
			<td colspan="5"><b>Nombre Empresa:</b> <label id="lblNombreEmpresa" name="lblNombreEmpresa"> </label></td>
		</tr>
		<tr>
			<td colspan="3"><b>Linea de Credito Solicitado:</b> <label id="lineaCred1" name="lineaCred1"></label></td>
			<td colspan="3"><b>Plazo De Pago Solicitado:</b> <label id="lineaCred2" name="lineaCred2"></label></td>
			<td colspan="2"><b>Día De Pago :</b> <label id="lineaCred3" name="lineaCred3"></label></td>
		</tr>
		<tr>
			<td colspan="8">&nbsp;</td>
		</tr>
	        <tr>
			<td colspan="8"><center>
			  <h3><em style="text-transform: capitalize;">Información Solicitada</em></h3></center></td>
		</tr>
		<tr>
			<td colspan="3"><b>Encargado Compras :</b> <label id="lineaCred4" name="lineaCred4"></label></td>
			<td colspan="3"><b>Correo Electrónico :</b> <label id="lineaCred5" name="lineaCred5"></label></td>
			<td colspan="2"><b>Teléfono :</b> <label id="lineaCred6" name="lineaCred6"></label></td>
		</tr>
		<tr>
			<td colspan="3"><b>Encargado Compras :</b> <label id="lineaCred7" name="lineaCred7"></label></td>
			<td colspan="3"><b>Correo Electrónico :</b> <label id="lineaCred8" name="lineaCred8"></label></td>
			<td colspan="2"><b>Teléfono :</b> <label id="lineaCred9" name="lineaCred9"></label></td>
		</tr>
	        <tr>
			<td colspan="8"><b>Observaciones (Horario De Atención A Proveedores) :</b> <label id="lineaCred14" name="lineaCred14"></label></td>
		</tr>
		<tr>
			<td colspan="3"><b>Representante Legal :</b> <label id="lineaCred10" name="lineaCred10"></label></td>
			<td colspan="3"><b>Correo Electrónico :</b> <label id="lineaCred11" name="lineaCred11"></label></td>
			<td colspan="2"><b>Teléfono :</b> <label id="lineaCred12" name="lineaCred12"></label></td>
		</tr>
                <tr>
			<td colspan="8"><b>Condiciones Particulares Para Facturar, Señale Requisitos Especiales:</b> <label id="lineaCred13" name="lineaCred13"></label></td>
		</tr>
		<tr>
			<td colspan="8"><b>Información Adicional (Señale Observaciones Para Considerar) :</b> <label id="lineaCred15" name="lineaCred15"></label></td>
		</tr>
		<tr>
			<td colspan="8"><center>
			  <h3><em style="text-transform: capitalize;">Documentos</em></h3></center></td>
		</tr>
		<tr>
			<td colspan="8">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="8">Ver Documento : <a href="http://norteqa.otecmutual.cl/solicitud/documentos/SolicitudCredito_20211026122712.pdf" target="_blank">Ver</a>
			<div id="documentosContent" id="documentosContent" ></div>
			</td>
		</tr>
		<tr>
			<td colspan="8">&nbsp;</td>
		</tr>

		<!--
		<tr>
			<td colspan="8"><center>
			  <h3><em style="text-transform: capitalize;">Actualizar Estado</em></h3></center></td>
		</tr>
		<tr>
			<td colspan="8">&nbsp;</td>
		</tr>
		 <tr>
			<td colspan="2">&nbsp;</td>
			<td colspan="3" style="text-align: right;">Cambiar Estado:</td>
			<td colspan="3">
				<select id="selEstadoSolicitud" name="selEstadoSolicitud" style="width:200px"> 
					<option value="1">INGRESADO</option>
					<option value="2">APROBADO</option>
				</select>
			</td>
		</tr>
		--> 
	</table>
	
	</form>
</div>
<div id="dialog-eliminar" title="Eliminar">
  <p><span class="ui-icon ui-icon-alert" style="float:left; margin:12px 12px 20px 0;"></span>Esta seguro de eliminar la solicitud de crédito?</p>
</div>
<div id="mensaje" title="Registro de Inscripción">
     <label id="txtMensaje" name="txtMensaje"></label>
</div>
<div id="datosTerminos" title="Terminos y  Condiciones">
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>

<div id="dialog_mensaje" title="Solicitud de Crédito Empresa">
	<p>La Solicitud de Crédito de Empresa fue Actualizada Exitosamente. </p>
</div>
<div id="dialog_espera" title="Solicitud de Crédito Empresa">
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

</body>
</html>
