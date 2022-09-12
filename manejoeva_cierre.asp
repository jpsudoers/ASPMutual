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
<script type="text/javascript">
var emp_del;
var texto="";
var vacios=0;

$(document).ready(function(){					

	jQuery("#list1").jqGrid({ 
		url:'evaluacioncierre/listar.asp', 
		datatype: "xml", 
		colNames:['Cód. Prog.','Código Curso', 'Sence', 'Nombre Curso','Relator','Fecha Inicio', '&nbsp;'], 
		colModel:[
				   {name:'PROGRAMA',index:'PROGRAMA', width:35},
				   {name:'CODIGO',index:'CODIGO', width:53}, 
				   {name:'SENCE',index:'SENCE', width:25}, 
				   {name:'NOMBRE_CURSO',index:'NOMBRE_CURSO'}, 
				   {name:'instructor',index:'instructor', width:50}, 
				   {name:'FECHA_INICIO_',index:'FECHA_INICIO_', width:45},
				   { align:"right",editable:true, width:20}], 
		rowNum:30, 
		autowidth: true, 
		rowList:[30,50,100], 
		pager: jQuery('#pager1'), 
		sortname: 'PROGRAMA.ID_PROGRAMA', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Evaluaciones y Cierres" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list1").trigger("reloadGrid");
													  } }); 
	
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
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height:550,
			width: 1000,
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
							$.post($('#frmcierreeval').attr('action')+'?'+$('#frmcierreeval').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
          //documento("evaluacioncierre/pdf.asp?prog="+$("#txtId").val()+"&empresa=0&trabajador=0&relator="+$("#ins_relator").val());		

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
			height: 600,
			width:1000,
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
			title: 'Certificado Asistencia A Curso De Capacitación'
		});
	});	


	function cambiaEstado()
	{
		for (i=1; i<=$('#countFilas').val(); i++)
		{
		  if($('#AsTra'+i).val()!='' && $('#AsTra'+i).val()>=60 && $('#caTra'+i).val()!='' && $('#caTra'+i).val()>=60)
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
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}

	function update(i,relator){
		$.post("evaluacioncierre/frmcierreeval.asp",
			   {id:i,relator:relator},
			   function(f){
				    $('#dialog').html(f);
					//cargaDatos($('#txtIdCurriculo').val());
				    //validar();
		});
		 $('#dialog').dialog('open');
	}
	
	function documento(arch){
		//alert(arch)
		//window.open(arch)
		$("#ifPagina").attr('src',arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
	}function validarFrm(){
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
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4 from USUARIOS "
	sql = sql&" where USUARIOS.ID_USUARIO='"&Session("usuarioMutual")&"'"

   DATOS.Open sql,oConn
   WHILE NOT DATOS.EOF
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
        <li><a href="manejocursos.asp" accesskey="3" title="" class="selItem">Manejo de Cursos</a></li>
        <%end if
		if(DATOS("PERMISO4")<>"0")then
		%>
		<li><a href="finanzas.asp" accesskey="4" title="">Finanzas</a></li>
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
				<li class="first"><a href="manejocursoslibroclases.asp">Libro de Clases</a></li>
                <li><a href="manejoeva_cierre.asp">Evaluación y Cierre</a></li>
        
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Evaluación y Cierre</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro de Evaluación y Cierre de Curso">
</div>
<div id="datos" title="Registro de Evaluación y Cierre de Curso">
	<p>Por Favor, Ingresar Toda la Información Requerida.</p>
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>
