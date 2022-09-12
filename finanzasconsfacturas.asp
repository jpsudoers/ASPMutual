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
<script src="js/jquery.tbltogrid.js" type="text/javascript" charset="utf-8"> </script> 
<script src="js/ajaxfileupload.js" type="text/javascript" charset="utf-8"></script>
<style type="text/css">
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
		height:150px;
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
.suggestionsBox21 {        position: relative;
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
</style>
<script type="text/javascript">
var cargada=false;

$(document).ready(function(){					
		$('#txRutEmpresa').defaultValue('Ingrese Empresa');
		$('#txtFCursoHasta').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		$('#txtFCursoDesde').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });	
		$('#txtFEmiDesde').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
		$('#txtFEmiHasta').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });	
		llena_curriculo(0);

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

	$("#nTickets").dialog({
			autoOpen: false,
			bgiframe: true,
			height:200,
			width: 450,
			modal: true,
			buttons: {
				'Generar Tickets': function() {
					if($("#frmntickets").valid()){
	documento('libroventas/Tickets.asp?id='+$('#idAutoTickets').val()+'&ntickets='+$('#txtNTickets').val()+'&nfactura='+$('#nFac').val());
						//$(this).dialog('close');
//jQuery("#list1").jqGrid('setGridParam',{url:"libroventas/listar.asp?empresa="+$('#EmprBuscar').val()+"&selMes="+$('#SelMes').val()+"&selAno="+$('#selAno').val()}).trigger("reloadGrid")
					}
					jQuery("#list1").jqGrid('setGridParam',{url:"libroventas/listar.asp?empresa="+$('#EmprBuscar').val()+"&selMes="+$('#SelMes').val()+"&selAno="+$('#selAno').val()}).trigger("reloadGrid")					
					$(this).dialog('close');
				},
				'Cerrar': function() {
					$(this).dialog('close');
				}
			}
	});

	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height:580,
			width: 980,
			modal: true,
			buttons: {
				'Cerrar': function() {
						$(this).dialog('close');
				}
			}
		});

		$("#dialog_softland").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 250,
			width: 400,
			modal: true,
			buttons: {
				'Softland': function() {
		window.open("libroventas/xls_Softland.asp?i="+$('#f_inicio_soft').val()+"&t="+$('#f_hasta_soft').val(),'Softland')
								$(this).dialog('close');								
				},
				'Control Diario': function() {
		window.open("libroventas/xls_control_diario.asp?i="+$('#f_inicio_soft').val()+"&t="+$('#f_hasta_soft').val(),'Control Diario')
								$(this).dialog('close');								
				},	
				'Genera TXT': function() {
		window.open("TXT/txt.asp?i="+$('#f_inicio_soft').val()+"&t="+$('#f_hasta_soft').val(),'Genera TXT')
								$(this).dialog('close');								
				},									
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
	});
	
		$("#dialog_Anular").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 450,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
						if($("#frmFactAnular").valid())
						{
								//window.open($('#frmFactAnular').attr('action')+'?'+$('#frmFactAnular').serialize())
								$.post($('#frmFactAnular').attr('action')+'?'+$('#frmFactAnular').serialize(),function(d){
																							$("#list1").trigger("reloadGrid"); 
																								   });
								$(this).dialog('close');								
								$("#list1").trigger("reloadGrid");
						}
						
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
	});
	
	$("#Doc").dialog({
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
	});	
	
	function update(i){
		$.post("libroventas/frmInscripcion.asp",
			   {id:i},
			   function(f){
				    $('#dialog').html(f);
					grilla();
		});
		 $('#dialog').dialog('open');
	}
	
	function grilla()
	{
		jQuery("#listPart").jqGrid({ 
		url:'libroventas/listaPart.asp?IDAuto='+$('#id_autorizacion').val(), 
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
	
	function documento(arch){
		$("#ifPagina").attr('src',arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
	}
	
	function tickets(id,n_tickets,n_factura){
		if(n_tickets!="")
		{
			documento('libroventas/Tickets.asp?id='+id+'&ntickets='+n_tickets+'&nfactura='+n_factura);
		}
		else
		{
			$.post("libroventas/frmNTickets.asp",{id:id,fac:n_factura},
				   function(f){
						$('#nTickets').html(f);
						validarFrmTickets();
			});
			$('#nTickets').dialog('open');
		}
	}	
	
	function validarFrmTickets(){
		$("#frmntickets").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txtNTickets:{
				required:true,
				number: true
			}
		},
		messages:{
			txtNTickets:{
				required:"&bull; Ingrese Número de Tickets.",
				number: "&bull; Debe Ingresar Solo Números."
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

	function lookup2(inputString) {
		if(inputString.length <=2) {
			$('#suggestions2').hide();
		    $("#txtRazon").html("");
			$("#txtRsocEmpresa").html("");
	        $('#EmprBuscar').val("");
			if(inputString.length <=0)
			{
			tabla();
			}
		} else {
				$.post("libroventas/sugEmpresa.asp", {txt: inputString}, function(data){
						if(data.length >0) {
								$('#suggestions2').show();
								$('#autoSuggestionsList2').html(data);
						}
				});
		}
	}

	function fill2(id,rut) {
	   $('#EmprBuscar').val(id);
	   cargaDatosEmpresa2(id);
	   $('#suggestions2').hide();
	}

	function cargaDatosEmpresa2(id)	
	{
		$("#txRutEmpresa").val("");
		$("#txtRazon").html("");
		$("#txtRsocEmpresa").html("");
 
		$.get("bhp/datosempresa.asp",
						 {id:id},
						 function(xml){
									$('row',xml).each(function(i) {
									$("#txtRazon").html("Razón Social :");
									$("#txtRsocEmpresa").html($(this).find('RSOCIAL').text());
									$("#txRutEmpresa").val($(this).find('RUT').text());
									tabla();
						});
			});
	}

	function tabla()
	{
		//window.open("libroventas/listar.asp?empresa="+$('#EmprBuscar').val()+"&selMes="+$('#SelMes').val()+"&selAno="+$('#selAno').val())
		if(!cargada){
		jQuery("#list1").jqGrid({ 
		url:'libroventas/listarBuscar.asp?txDoc=&e=0&fch=&fcd=&fed=&feh=&c=0', 
		datatype: "xml", 
		colNames:['F.Emisión', 'Rut', 'Nombre de Empresa', 'N° Factura', 'Estado','&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'fecha_inicio',index:'F.FECHA_EMISION', width:37, align:'center'}, 
				   {name:'rut_empresa',index:'E.RUT', width:40, align:'right'}, 
				   {name:'empresa',index:'E.R_SOCIAL'}, 
				   {name:'FACTURA',index:'F.FACTURA', width:22, align:"right"},	
				   {name:'ESTADO',index:'F.ESTADO', width:29, align:'center'},	
				   {align:"right",editable:true, width:10},
				   {align:"right",editable:true, width:8}], 
		rowNum:300, 
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'F.FECHA_EMISION', 
		viewrecords: true, 
		sortorder: "DESC", 
		caption:"Listado de Ventas" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
														$("#txtRazon").html("");
														$("#txtRsocEmpresa").html("");
														$('#EmprBuscar').val("");
														$('#txRutEmpresa').val("");
														$('#txRutEmpresa').defaultValue('Ingrese Empresa');
														document.getElementById("Curriculo").selectedIndex = 0;
														/*document.getElementById("SelMes").selectedIndex = 0;
														document.getElementById("selAno").selectedIndex = 0;*/
														
jQuery("#list1").jqGrid('setGridParam',{url:"libroventas/listarBuscar.asp?txDoc=&e=0&fch=&fcd=&fed=&feh=&c=0"}).trigger("reloadGrid")
													  } });
													  
jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"Excel",
													  title:"Expotar a Excel", 
													  buttonicon :'ui-icon-script', 
													  onClickButton:function(){ 
 														 window.open("libroventas/xlsFacturacion.asp?txDoc="+$('#txDoc').val()+"&e="+$('#EmprBuscar').val()+"&fch="+$('#txtFCursoHasta').val()+"&fcd="+$('#txtFCursoDesde').val()+"&fed="+$('#txtFEmiDesde').val()+"&feh="+$('#txtFEmiHasta').val()+"&c="+$('#Curriculo').val(),'Informe')
													  } });										  
	    cargada=true;
		}
		else
		{
jQuery("#list1").jqGrid('setGridParam',{url:"libroventas/listarBuscar.asp?txDoc="+$('#txDoc').val()+"&e="+$('#EmprBuscar').val()+"&fch="+$('#txtFCursoHasta').val()+"&fcd="+$('#txtFCursoDesde').val()+"&fed="+$('#txtFEmiDesde').val()+"&feh="+$('#txtFEmiHasta').val()+"&c="+$('#Curriculo').val()}).trigger("reloadGrid")
		}
	}		
	
	/*function Anular(nfactura,id_autorizacion,id_factura){
	  $.post("libroventas/frmAnular.asp",{id_factura:id_factura,id_autorizacion:id_autorizacion,nfactura:nfactura},
			   function(f){
				    $('#dialog_Anular').html(f);
					$('#razonAnular').focus();
					valFrmAnular();
		});
	    $('#dialog_Anular').dialog('open');
	}*/
	
	function valFrmAnular(){
		$("#frmFactAnular").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			razonAnular:{
				required:true
			},
			tipo_baja:{
				required:true
			}		
		},
		messages:{
			razonAnular:{
				required:"&bull; Ingrese las Razones."
			},
			tipo_baja:{
				required:"&bull; Seleccione la acción a realizar con la Factura."
			}		
		}
	});
    }
	
	function llena_curriculo(id_curriculo){
		$("#Curriculo").html("");
		$.get("consultascertificados/curriculo.asp",
					function(xml){
						$("#Curriculo").append("<option value='0'>Todos</option>");
						$('row',xml).each(function(i) { 
							if(id_curriculo==$(this).find('ID_MUTUAL').text())
								$("#Curriculo").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$("#Curriculo").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
					});
			tabla();
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
			<h2><em style="text-transform: capitalize;">Buscar factura</em></h2>
            <br>
           <table width="950" border="0">
              <tr>
                <td width="90">Factura :</td>
                <td><input id="txDoc" name="txDoc" type="text" maxlength="20" size="20" onkeyup="tabla();"/></td>
              </tr>
              <tr>
                <td>Fecha Curso : </td>
                <td>Desde 
                <input id="txtFCursoDesde" name="txtFCursoDesde" type="text" maxlength="50" size="12" onchange="tabla();"/>
                &nbsp;&nbsp;Hasta&nbsp;&nbsp;
                <input id="txtFCursoHasta" name="txtFCursoHasta" type="text" maxlength="50" size="12" onchange="tabla();"/></td>
              </tr>  
              <tr>
                <td>Fecha Emisión : </td>
                <td>Desde 
                <input id="txtFEmiDesde" name="txtFEmiDesde" type="text" maxlength="50" size="12" onchange="tabla();"/>
                &nbsp;&nbsp;Hasta&nbsp;&nbsp;
                <input id="txtFEmiHasta" name="txtFEmiHasta" type="text" maxlength="50" size="12" onchange="tabla();"/></td>
              </tr>  
              <tr>
                <td>Curso :</td>
                <td><select id="Curriculo" name="Curriculo" onchange="tabla();" style="width:33em;"></select></td>
              </tr>    
              <tr>
                <td colspan="2">Rut de Empresa :&nbsp;&nbsp;&nbsp;<input id="txRutEmpresa" name="txRutEmpresa" type="text" maxlength="20" size="20" onkeyup="lookup2(this.value);"/>
                  <div class="suggestionsBox21" id="suggestions2" style="display: none;position:absolute;z-index:2;left:522px"> <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
                    <div class="suggestionList2" id="autoSuggestionsList2"> &nbsp; </div>
                  </div><input type="hidden" id="EmprBuscar" name="EmprBuscar"/>&nbsp;&nbsp;<label id="txtRazon" name="txtRazon"></label>&nbsp;&nbsp;<label id="txtRsocEmpresa" name="txtRsocEmpresa"></label></td>
              </tr>   
            </table>
            <br>            
          <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Inscripciones Facturadas">
</div>
<div id="dialog_softland" title="Descargar a Softland">
</div>
<div id="dialog_Anular" title="Anular o Refacturar">
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="nTickets" title="Ingresar Número de Tickets">
</div>
</body>
</html>