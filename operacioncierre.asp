<!--#include file="cnn_string.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

userID="0"
if(Session("usuarioMutual")<>"")then userID=Session("usuarioMutual") end if
%>
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
	
	.suggestionsBox3 {
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
   
    .suggestionList3 {
        margin: 0px;
        padding: 0px;
    }
   
    .suggestionList3 li {
       
        margin: 0px 0px 3px 0px;
        padding: 3px;
        cursor: pointer;
    }
   
    .suggestionList3 li:hover {
        background-color: #659CD8;
    }	
</style>
<script type="text/javascript">
var emp_del;
var texto="";
var vacios=0;
var cargada=false;
$(document).ready(function(){					
	/*jQuery("#list1").jqGrid({ 
		url:'revision_cierre/listar.asp', 
		datatype: "xml", 
		colNames:['F. Inicio','Rut', 'Nombre de Empresa','Nombre del Curso', '&nbsp;', '&nbsp;'], 
		colModel:[
				   {name:'FECHA_INICIO_',index:'CONVERT(date,P.FECHA_INICIO_)', width:48, align:'center'},
				   {name:'RUT',index:'E.RUT', width:45}, 
				   {name:'R_SOCIAL',index:'dbo.MayMinTexto (E.R_SOCIAL)'}, 
				   {name:'NOMBRE_CURSO',index:'nombre'},
				   { align:"right",editable:true, width:13},
				   { align:"right",editable:true, width:13}], 
		rowNum:300, 
		height:350,
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'CONVERT(date,P.FECHA_INICIO_)', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Cursos Evaluados" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list1").trigger("reloadGrid");
													  } }); 
	*/
	$('#txRutEmpresa').val('');
	$('#txRutEmpresa').defaultValue('Todas');
	$('#EmprBuscar').val('');
	llena_condicionComercial();				
	llenaGrilla();

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
	
	$("#datos").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
						$(this).dialog('close');
				}
			}
		});
		
	$("#msn").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 130,
			width: 600,
			modal: true,
			buttons: {
				'Cerrar': function() {
						$(this).dialog('close');
				}
			}
	});		
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height:560,
			width: 925,
			modal: true,
			buttons: {
				'Guardar': function() {
					vacios=0;
					for (i=1; i<=$('#countFilas').val(); i++)
					{
						if($('#AsTra'+i).val()=='' || $('#caTra'+i).val()=='')
						{
						vacios=1;
						}
					}
					
					if(vacios==0)
					{
						if($("#frmcierreeval").valid())
						{
							
							//window.open($('#frmcierreeval').attr('action')+'?'+$('#frmcierreeval').serialize());
							//$.post($('#frmcierreeval').attr('action')+'?'+$('#frmcierreeval').serialize(),function(d){
							$.post($('#frmcierreeval').attr('action'),$('#frmcierreeval').serialize(),function(d){								
			   																				$("#list1").trigger("reloadGrid"); 
																						   });
							$("#list1").trigger("reloadGrid");
							$(this).dialog('close');
						}
					}
					else
					{
						$("#datos").dialog('open');
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
				'Cerrar': function() {
					$(this).dialog('close');
				}
			},
			title: 'Documento'
		});	
	});	


	/*function cambiaEstado()
	{
		for (i=1; i<=$('#countFilas').val(); i++)
		{
		  if($('#AsTra'+i).val()!='' && $('#AsTra'+i).val()==100 && $('#caTra'+i).val()!='' && $('#caTra'+i).val()>=70)
			{
				$('#evaluacion'+i).html("Aprobado");
				$('#eval'+i).val("Aprobado");
			}
			else
			{
				$('#evaluacion'+i).html("Reprobado");
				$('#eval'+i).val("Reprobado");
			}
		}
	}*/
	
	function cambiaEstado()
	{
		for (i=1; i<=$('#countFilas').val(); i++)
		{
		    if($('#txtIdMutual').val()=="33"){
				if($('#AsTra'+i).val()!='' && $('#AsTra'+i).val()==100 && $('#caTra'+i).val()!='' && $('#caTra'+i).val()>=65)
				{
					if($('#caTra'+i).val()>=65 && $('#caTra'+i).val()<=69){
						$('#evaluacion'+i).html("Con Obs.");
						$('#eval'+i).val("Con Obs.");
					}
					else
					{
						$('#evaluacion'+i).html("Aprobado");
						$('#eval'+i).val("Aprobado");
					}
				}
				else
				{
					$('#evaluacion'+i).html("Reprobado");
					$('#eval'+i).val("Reprobado");
				}
		   }
		   else if ($('#txtIdMutual').val()=="2317" || $('#txtIdMutual').val()=="2319" || $('#txtIdMutual').val()=="2324" || $('#txtIdMutual').val()=="2325" || $('#txtIdMutual').val()=="2332" || $('#txtIdMutual').val()=="2334" || $('#txtIdMutual').val()=="2335" || $('#txtIdMutual').val()=="2318" || $('#txtIdMutual').val()=="2341" || $('#txtIdMutual').val()=="2342" || $('#txtIdMutual').val()=="2343" || $('#txtIdMutual').val()=="2326" || $('#txtIdMutual').val()=="2320" || $('#txtIdMutual').val()=="2321" || $('#txtIdMutual').val()=="2322" || $('#txtIdMutual').val()=="2333" || $('#txtIdMutual').val()=="2330" || $('#txtIdMutual').val()=="2331" || $('#txtIdMutual').val()=="2327" || $('#txtIdMutual').val()=="2328" || $('#txtIdMutual').val()=="82" || $('#txtIdMutual').val()=="2316" || $('#txtIdMutual').val()=="2337" || $('#txtIdMutual').val()=="2338" || $('#txtIdMutual').val()=="2339" || $('#txtIdMutual').val()=="2340"){
			   	if($('#AsTra'+i).val()!='' && $('#AsTra'+i).val()==100 && $('#caTra'+i).val()!='' && $('#caTra'+i).val()>=80)
				{
					$('#evaluacion'+i).html("Aprobado");
					$('#eval'+i).val("Aprobado");
				}
				else
				{
					$('#evaluacion'+i).html("Reprobado");
					$('#eval'+i).val("Reprobado");
				}
		   }
		   else
		   {
			   	if($('#AsTra'+i).val()!='' && $('#AsTra'+i).val()==100 && $('#caTra'+i).val()!='' && $('#caTra'+i).val()>=70)
				{
					$('#evaluacion'+i).html("Aprobado");
					$('#eval'+i).val("Aprobado");
				}
				else
				{
					$('#evaluacion'+i).html("Reprobado");
					$('#eval'+i).val("Reprobado");
				}
		   }			

		}
	}
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}

	function update(id){
		/*$.post("revision_cierre/frmcierreeval.asp",
			   {id:i,relator:relator},
			   function(f){
				    $('#dialog').html(f);
					tableToGrid("#mytable"); 
		});
		 $('#dialog').dialog('open');*/
		 //if(<%=userID%>=='8' || <%=userID%>=='169'){
			 $.post("revision_cierre/frmcierreeval.asp",{id:id},
				   function(f){
						$('#dialog').html(f);
						tableToGrid("#mytable"); 
			});
		 	
			$('#dialog').dialog('open');
		/* }
		 else
		 {
			 $("#msn").dialog('open');
		 }*/
	}
	
	function documento(arch){
		$("#ifPagina").attr('src',"http://norte.otecmutual.cl/ordenes/"+arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
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

	function eliberarsin(){
		$('#txtSoloCert').val("0");
		if($('#liberarsin').attr('checked')) {
    			//alert('Seleccionado');
			$('#txtSoloCert').val("1");
		}
	}

      function llena_condicionComercial(){
		$("#condicionesComerciales").html("");
		$.get("revision_cierre/condiciones_comerciales.asp",
					function(xml){
						$("#condicionesComerciales").append("<option value=''>Seleccione</option>");
						$('row',xml).each(function(i) { 
							$("#condicionesComerciales").append('<option value="'+$(this).find('ID_CONDICION_COMERCIAL').text()+'" >'+
																		$(this).find('NOMBRE').text()+ '</option>');
						});
					});
	}

	function llenaGrilla(){

		sUrl='revision_cierre/listar.asp?condicion='+$('#condicionesComerciales').val()+'&tipoventa='+$('#tipoVenta').val()+'&estado='+$('#estado').val()+'&empB='+$('#EmprBuscar').val();
		if(cargada){
			jQuery("#list1").jqGrid('setGridParam',{url:sUrl});
			jQuery("#list1").trigger("reloadGrid"); 
		}else{
			cargada=true;
		jQuery("#list1").jqGrid({ 
		url:sUrl,
		datatype: "xml", 
		colNames:['F. Inicio','Rut', 'Nombre de Empresa','Nombre del Curso', '&nbsp;', '&nbsp;'], 
		colModel:[
				   {name:'FECHA_INICIO_',index:'CONVERT(date,P.FECHA_INICIO_)', width:48, align:'center'},
				   {name:'RUT',index:'E.RUT', width:45}, 
				   {name:'R_SOCIAL',index:'dbo.MayMinTexto (E.R_SOCIAL)'}, 
				   {name:'NOMBRE_CURSO',index:'nombre'},
				   { align:"right",editable:true, width:13},
				   { align:"right",editable:true, width:13}], 
		rowNum:300, 
		height:350,
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'CONVERT(date,P.FECHA_INICIO_)', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Cursos Evaluados" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
														llenaGrilla();	
													  // $("#list1").trigger("reloadGrid");
													  } }); 
	}
}

 function lookup2(inputString) {
		if(inputString.length <=2) {
			$('#suggestions2').hide();
		    $("#txtRazon").html("");
			$("#txtRsocEmpresa").html("");
	        $('#EmprBuscar').val("");
			if(inputString.length <=0)
			{
			//tabla();
			llenaGrilla();
			}
		} else {
				$.post("bhp/sugEmpresa.asp", {txt: inputString}, function(data){
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
									//tabla();
									llenaGrilla();
						});
			});
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
				Usuario : <strong><%=Session("nombre")%></strong>
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
			<h2><em style="text-transform: capitalize;">Revisión y Cierre</em></h2>
              <table width="650" border="0">
              <tr>
                <td width="120"><b style="font-size:12px; ">Rut de Empresa :</b> </td>
                <td width="100"><input id="txRutEmpresa" name="txRutEmpresa" type="text" tabindex="1" maxlength="20" size="20" onkeyup="lookup2(this.value);" class="empty">
                   <div class="suggestionsBox2" id="suggestions2" style="display: none;position:absolute;z-index:2;left:522px">
                        <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow">
                        <div class="suggestionList2" id="autoSuggestionsList2">
                          &nbsp;
                        </div>
                   </div></td>
                    <td width="15"><input type="hidden" id="EmprBuscar" name="EmprBuscar"></td>
                	<td width="80"><label id="txtRazon" name="txtRazon"></label></td>
                    <td width="315"><label id="txtRsocEmpresa" name="txtRsocEmpresa"></label></td>
              </tr>
              </table>
<b style="font-size:12px; ">Condiciones Comerciales</b> 
			<select id="condicionesComerciales" name="condicionesComerciales" tabindex="10" style="width:10em;" onchange="llenaGrilla()">
				<option value="" selected>Seleccione</option>
				</select>
&nbsp; 
<b style="font-size:12px; ">Tipo Venta</b> 
			<select id="tipoVenta" name="tipoVenta" tabindex="10" style="width:10em;" onchange="llenaGrilla()">
				<option value="" selected>Seleccione</option>
				<option value="1" >Venta Directa</option>
				<option value="2">Venta Broker</option>
				</select>
				&nbsp; 
<b style="font-size:12px;">Estado:</b> 
<select id="estado" name="estado" tabindex="10" style="width:10em;" onchange="llenaGrilla()">
	<option value="" selected>Seleccione</option>
	<option value="1" >En Proceso de Liberación</option>
	<option value="2">Liberadas - Regresadas</option>
	</select>
	<br style="margin-bottom: 5px!important;" />
	<br style="margin-bottom: 5px!important;" />
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Revisión y Cierre de Curso">
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="datos" title="Revisión y Cierre de Curso">
	<p>Por Favor, Ingresar Toda la Información Requerida.</p>
</div>
<div id="msn" title="Revisión y Cierre de Curso">
	<p>Acción no permitida, Por Favor Contactarse con departamento Facturación.</p>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>
