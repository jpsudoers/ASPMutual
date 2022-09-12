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
<script src="js/jquery.jPrintArea.js" type="text/javascript"></script>
<script src="js/NumALetras.js" type="text/javascript"></script>
<script type="text/javascript">
var emp_del;
var texto="";
var vacios=0;
var trabajador="";
var totSel=0;
var estadoLimpiar=0;
var estadoFacEmp=0;
var estadoFacOtic=0;
var cargada=false;
$(document).ready(function(){
	tabla();

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
	
	 $("#detCompCancel").dialog({
			autoOpen: false,
			bgiframe: true,
			height:80,
			width: 450,
			modal: true,
			buttons: {
				Si: function() {
						$.post($('#frmDetalle').attr('action')+'?'+$('#frmDetalle').serialize(),function(d){
																					     $("#list1").trigger("reloadGrid");
																						   });
						$('#frmDetalle')[0].reset();
						$("#detCancel").dialog('close');
						$(this).dialog('close');
						//$("#listEmpresas").trigger("reloadGrid");
						$("#list1").trigger("reloadGrid");
						estadoFacEmp=0;
				},
				No: function() {
					estadoFacEmp=0;
					$(this).dialog('close');
				}
			}
		});
	
	$("#detCompCancelOtic").dialog({
			autoOpen: false,
			bgiframe: true,
			height:80,
			width: 450,
			modal: true,
			buttons: {
				Si: function() {
						$.post($('#frmDetalleOtic').attr('action')+'?'+$('#frmDetalleOtic').serialize(),function(d){
																					      $("#list1").trigger("reloadGrid");
																						   });
					    $('#frmDetalleOtic')[0].reset();	
						$("#detCancelOtic").dialog('close');
						$(this).dialog('close');
						//$("#listEmpresas").trigger("reloadGrid");
						$("#list1").trigger("reloadGrid");
						estadoFacOtic=0;
				},
				No: function() {
					estadoFacOtic=0;
					$(this).dialog('close');
				}
			}
		});
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height:520,
			width: 925,
			modal: true,
			buttons: {
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
		});
	
		$("#detCancel").dialog({
			bgiframe: true,
			autoOpen: false,
			height:470,
			width: 650,
			modal: true,
			buttons: {
				Imprimir: function() {
					if($("#frmDetalle").valid())
					{
						if(estadoFacEmp==0)
						{
							estadoFacEmp=1;
							$.jPrintArea('#imprimeFactura');
							setTimeout("$('#detCompCancel').dialog('open');",6000)
						}
					}
				},
				Cerrar: function() {
					if(estadoFacEmp==0)
					{
						$('#frmDetalle')[0].reset();
						$(this).dialog('close');
					}
				}
			}
		});
		
		$("#detCancelOtic").dialog({
			bgiframe: true,
			autoOpen: false,
			height:550,
			width: 870,
			modal: true,
			buttons: {
				Imprimir: function() {
					if($("#frmDetalleOtic").valid())
					{
						if(estadoFacOtic==0)
						{
							estadoFacOtic=1;
							$.jPrintArea('#imprimeFacturaOtic');
							setTimeout("$('#detCompCancelOtic').dialog('open');",6000)
						}
					}
				},
				Cerrar: function() {
					if(estadoFacOtic==0)
					{
						$('#frmDetalleOtic')[0].reset();
						$(this).dialog('close');
					}
				}
			}
		});
	
	$("#Orden").dialog({
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
			title: 'Documento'
		});
	
	$("#Cert").dialog({
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
			title: 'Documento'
		});
	
	$("#Doc").dialog({
			autoOpen: false,
			bgiframe: true,
			height: 500,
			width:1000,
			modal: true,
			buttons: {
				'Imprimir': function() {
					$.jPrintArea('#Doc');
					$(this).dialog('close');
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			},
			title: 'Imprimir Factura'
		});
	});	

	function tabla()
	{
	 //if(!cargada){
	   jQuery("#list1").jqGrid({ 
		url:'finazasFacturacion/listatotales.asp', 
		datatype: "xml", 
		colNames:['Código Curso','Nombre Curso', 'Pendientes','&nbsp;'], 
		colModel:[
				   {name:'CODIGO',index:'C.CODIGO', width:33, align:'center'}, 				  
				   {name:'nombre',index:'nombre_curso'}, 
				   {name:'TOT',index:'TOTAL', width:27, align:'right'},
				   {align:"right",editable:true, width:8, align:'center'}], 
		rowNum:300, 
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'nombre_curso', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Cursos" 
		}); 
	
	   jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	   jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
															$("#list1").trigger("reloadGrid");
													   } });
	   /* cargada=true;
		}
		else
		{
jQuery("#list1").jqGrid('setGridParam',{url:"finazasFacturacion/listaEmpresas.asp?IDMutual="}).trigger("reloadGrid")
		}*/
	}

    function grilla()
	{
		jQuery("#listPart").jqGrid({ 
		url:'finazasFacturacion/listaPart.asp?IDAuto='+$('#id_autorizacion').val(), 
		datatype: "xml", 
		colNames:['Rut / Ident.', 'Alumno', 'Asistencia (%)', 'Calificación (%)', 'Evaluación'], 
		colModel:[
				   {name:'RUT',index:'RUT', width:30, align:'center'}, 
				   {name:'NOMBRES',index:'NOMBRES'},
				   {name:'ASISTENCIA',index:'ASISTENCIA', width:40, align:'Center'}, 
				   {name:'CALIFICACION',index:'CALIFICACION', width:40, align:'Center'}, 
				   {name:'EVALUACION',index:'EVALUACION', width:40, align:'center'}], 
		rowNum:60, 
        rownumbers: true, 
        rownumWidth: 30, 
		autowidth: true, 
		pager: jQuery('#pagePart'), 
		sortname: 'TRABAJADOR.NOMBRES', 
		sortorder: "desc", 
		caption:"Listado de Participantes" 
		}); 
	
	   jQuery("#listPart").jqGrid('navGrid','#pagePart',{edit:false,add:false,del:false,search:false,refresh:false});
	}
	
	function update(i){
		//alert(i)
		$.post("finazasFacturacion/frmInscripcion.asp",
			   {id:i},
			   function(f){
				    $('#dialog').html(f);
					grilla();
		});
		 $('#dialog').dialog('open');
	}
	
	function detFactura(id){
		$.post("finazasFacturacion/frmDetalle.asp",{id:id},
			   function(f){
				    $('#detCancel').html(f);
					$('#txtFechIngreso').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
					if($('#Pagada').val()=="0")
					{
						OcultarFechPago('filaFechPago');
						$('#txtFechIngreso').val('01-01-1990');
					}
					$('#NfacturaEmp').focus();
					estadoFacEmp=0;
					validaDet();
		});
		 $('#detCancel').dialog('open');
	}
	
	function OcultarFechPago(Fila) {
   			var elementos = document.getElementsByName(Fila);
   			for (k = 0; k < elementos.length; k++)
      		elementos[k].style.display = "none";
	}
	
	function MostrarFechPago(Fila) {
  			var elementos = document.getElementsByName(Fila);
   			for (i = 0; i < elementos.length; i++)
      		elementos[i].style.display ="";
	}
	
	function detFacturaOtic(id){
		$.post("finazasFacturacion/frmDetalleOtic.asp",{id:id},
			   function(f){
				    $('#detCancelOtic').html(f);
					$('#ValorEmpresa').focus();
					estadoLimpiar=0;
					MostrarFila('filaNFactConN');
					OcultarFila('filaNFactSinN');
					//MostrarFila('filaNFactSinNOtic');
					//OcultarFila('filaNFactConNOtic');
					estadoFacOtic=0;
					validaDetOtic();
		});
		 $('#detCancelOtic').dialog('open');
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

	function calcularOtic(){
		if($('#ValorEmpresa').val()!='' && parseInt($('#ValorEmpresa').val())>=0)
		{
			if(parseInt($('#V_Inscripcion').val())-parseInt($('#ValorEmpresa').val())>=0)
			{
				$('#ValorOtic').html(formatCurrency(parseInt($('#V_Inscripcion').val())-parseInt($('#ValorEmpresa').val())));
				$('#ValOtic').val(parseInt($('#V_Inscripcion').val())-parseInt($('#ValorEmpresa').val()));

				$('#txtValFacEmp').html(formatCurrency($('#ValorEmpresa').val()).replace("$ ",""));
				$('#txtTotFacEmp').html(formatCurrency($('#ValorEmpresa').val()).replace("$ ",""));
				$('#txtValFacOtic').html(formatCurrency($('#ValOtic').val()).replace("$ ",""));
				$('#txtTotFacOtic').html(formatCurrency($('#ValOtic').val()).replace("$ ",""));

				$('#txtFacEmp').html(covertirNumLetras(String($('#ValorEmpresa').val()+'.01')));
				$('#txtFacOtic').html(covertirNumLetras(String($('#ValOtic').val()+'.01')));
				
				if(parseInt($('#ValorEmpresa').val())>0)
				{
					if(estadoLimpiar==1)
					{
						$('#Nfactura').val("");
						estadoLimpiar=0;
					}
					
					$('#facEmpresa').val("");
					MostrarFila('filaNFactConN');
					OcultarFila('filaNFactSinN');
					
					//MostrarFila('filaNFactSinNOtic');
					//OcultarFila('filaNFactConNOtic');
					
					$('#NfacturaOticCon').val("");
					$('#NfacturaOtic').val("");
					$('#NfactOtic').html("");
					$('#facOtic').html("");
					$('#facOtic2').html("");
					calFacOtic();
					
					MostrarFila('imprimeFila1');
					OcultarFila('imprimeFila2');
				}
				else
				{
					estadoLimpiar=1;
					$('#Nfactura').val("1");
					$('#facEmpresa').val("");
					MostrarFila('filaNFactSinN');
					OcultarFila('filaNFactConN');
					
					//MostrarFila('filaNFactConNOtic');
					//OcultarFila('filaNFactSinNOtic');
					
					$('#NfacturaOticCon').val("");
					$('#NfacturaOtic').val("");
					$('#NfactOtic').html("");
					$('#facOtic').html("");
					$('#facOtic2').html("");
					
					MostrarFila('imprimeFila2');
					OcultarFila('imprimeFila1');					
				}
				
				if(parseInt($('#ValorEmpresa').val())==0)
				{
					ocultar('factEmpTab');
					mostrar('factOticTab');
				}
				else
				{
					mostrar('factEmpTab');
				}
				
				if(parseInt($('#ValOtic').val())==0)
				{
					mostrar('factEmpTab');
					ocultar('factOticTab');
				}
				else
				{
					mostrar('factOticTab');
				}
			}
			else
			{
				$('#ValorOtic').html("$ 0");
			    $('#ValOtic').val("");
				
				$('#txtValFacEmp').html("");
				$('#txtTotFacEmp').html("");
				$('#txtValFacOtic').html("");
				$('#txtTotFacOtic').html("");

				$('#txtFacEmp').html("");
				$('#txtFacOtic').html("");
			}
		}
		else
		{
			MostrarFila('filaNFactConN');
			OcultarFila('filaNFactSinN');
					
			//MostrarFila('filaNFactSinNOtic');
			//OcultarFila('filaNFactConNOtic');
			
			$('#ValorOtic').html("$ 0");
			$('#ValOtic').val("");
			
			$('#txtValFacEmp').html("");
			$('#txtTotFacEmp').html("");
			$('#txtValFacOtic').html("");
			$('#txtTotFacOtic').html("");

			$('#txtFacEmp').html("");
			$('#txtFacOtic').html("");
			
			$('#Nfactura').val("");
			$('#NfactOtic').html("");
		}
	}
	
	function OcultarFila(Fila) {
   			var elementos = document.getElementsByName(Fila);
   			for (k = 0; k < elementos.length; k++)
      		elementos[k].style.display = "none";
	}
	
	function MostrarFila(Fila) {
  			var elementos = document.getElementsByName(Fila);
   			for (i = 0; i < elementos.length; i++)
      		elementos[i].style.display ="";
	}
	
	function mostrar(idCapa)
    {
			document.getElementById(idCapa).style.display='block';
    }
	
	function ocultar(idCapa)
    {
			document.getElementById(idCapa).style.display='none';
    }
	
	function calFacOtic(){
		if($('#Nfactura').val()!='' && parseInt($('#Nfactura').val())>0)
		{
			$('#facEmpresa').html($('#Nfactura').val());
			
			if(parseInt($('#V_Inscripcion').val())-parseInt($('#ValorEmpresa').val())>0)
			{
				/*$('#NfacturaOtic').val(parseInt($('#Nfactura').val())+1);
				$('#NfactOtic').html($('#NfacturaOtic').val());
				$('#facOtic').html($('#NfacturaOtic').val());
				$('#facOtic2').html($('#NfacturaOtic').val());*/
			}
			else
			{
				if(parseInt($('#V_Inscripcion').val())-parseInt($('#ValorEmpresa').val())==0)
				{
					//$('#NfactOtic').html("No Aplica");
				}
				else
				{
					$('#NfactOtic').html("");
				}
				
				$('#NfacturaOtic').val("1");
				$('#facOtic').html("");
				$('#facOtic2').html("");
			}
		}
		else
		{
			$('#NfactOtic').html("");
			//$('#NfacturaOtic').val("1");
			$('#facOtic').html("");
			$('#facOtic2').html("");
		}
	}
	
	function calFacOticCon(){
		if($('#NfacturaOticCon').val()!='' && parseInt($('#NfacturaOticCon').val())>=0)
		{
			$('#NfacturaOtic').val($('#NfacturaOticCon').val());
			$('#NfactOtic').html($('#NfacturaOticCon').val());
			$('#facOtic').html($('#NfacturaOticCon').val());
			$('#facOtic2').html($('#NfacturaOticCon').val());
		}
		else
		{
			$('#NfactOtic').html("");
			$('#NfacturaOtic').val("");
			$('#facOtic').html("");
			$('#facOtic2').html("");
		}
	}
	
	function documento(arch){
		$.post("finazasFacturacion/docFactura.asp",
			   function(f){
				    $('#Doc').html(f);
		});
		 $('#Doc').dialog('open');
	}
	
	function docOC(arch){
		$("#ifPagina").attr('src',"http://norte.otecmutual.cl/ordenes/"+arch);
		if(!$('#Orden').dialog('isOpen'))
			$('#Orden').dialog('open');
	}
	
	function verManual(){
		$("#ifPagina").attr('src',"http://norte.otecmutual.cl/documentos/Ajuste_de_Papel.pdf");
		if(!$('#Orden').dialog('isOpen'))
			$('#Orden').dialog('open');
	}
	
	function docCert(arch){
		$("#ifPagCert").attr('src',arch);
		if(!$('#Cert').dialog('isOpen'))
			$('#Cert').dialog('open');
	}
	
	function validaDetOtic(){
		$("#frmDetalleOtic").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			Nfactura:{
				required:true,
				min: 1
			},
			NfacturaOtic:{
				required:true,
				min: 1
			},
			ValorEmpresa:{
				required:true,
				max: $('#V_Inscripcion').val(),
				min: 0
			}
		},
		messages:{
			Nfactura:{
				required:"&bull; Ingrese Número de Factura.",
				min: "&bull; El Número de Factura debe ser mayor o igual a 1"
			},
			NfacturaOtic:{
				required:"&bull; Ingrese Número de Factura para la OTIC.",
				min: "&bull; El Número de Factura para la OTIC debe ser mayor o igual a 1"
			},
			ValorEmpresa:{
				required:"&bull; Ingrese el Monto a Cancelar por la Empresa.",
				max: "&bull; El Monto a Cancelar por la Empresa debe ser menor o igual a "+formatCurrency($('#V_Inscripcion').val()),
				min: "&bull; El Monto a Cancelar por la Empresa debe ser mayor o igual a $ 0"
			}
		}
	});
    }
	
	function validaDet(){
		$("#frmDetalle").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			NfacturaEmp:{
				required:true,
				min: 1
			},
			txtFechIngreso:{
				required:true
			}
		},
		messages:{
			NfacturaEmp:{
				required:"&bull; Ingrese Número de Factura.",
				min: "&bull; El Número de Factura debe ser mayor o igual a 1"
			},
			txtFechIngreso:{
				required:"&bull; Seleccione la Fecha de Pago."
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
	
sql="select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO6,PERMISO7,PERMISO8,PERMISO9,PERMISO10,PERMISO11,PERMISO12 from USUARIOS "
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
			<h2><em style="text-transform: capitalize;">Facturación</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Inscripción">
</div>
<div id="detCancel" title="Facturación">
</div>
<div id="detCompCancel" title="Facturación">
<label id="msnFactEmp" name="msnFactEmp">La Impresión de la Factura se Realizó Exitosamente?.</label>
</div>
<div id="detCompCancelOtic" title="Facturación">
<label id="msnFactOtic" name="msnFactOtic">La Impresión de las Facturas se Realizó Exitosamente?.</label>
</div>
<div id="detCancelOtic" title="Facturación">
</div>
<div id="Doc" style="display:none;border:none">
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="Orden" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="Cert" title="PagCert">
	<iframe style="width:100%;height:100%" id="ifPagCert"></iframe>
</div>
</body>
</html>
