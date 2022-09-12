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
	
	$("#dialogFac").dialog({
			bgiframe: true,
			autoOpen: false,
			height:380,
			width: 940,
			modal: true,
			buttons: {
				'Cerrar': function() {
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
		$.post("inscripcioneshistorico/frmInscripcion.asp",
			   {id:i},
			   function(f){
				    $('#dialog').html(f);
					grilla();
		});
		 $('#dialog').dialog('open');
	}
	
	function detFactura(i){
		$.post("inscripcioneshistorico/frmInscripcionFac.asp",{id:i},
			   function(f){
				    $('#dialogFac').html(f);
					tableToGrid("#mytable"); //grilla();
		});
		 $('#dialogFac').dialog('open');
	}		
	
	function grilla()
	{
		jQuery("#listPart").jqGrid({ 
		url:'inscripcioneshistorico/listaPart.asp?IDAuto='+$('#id_autorizacion').val(), 
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

	function lookup2(inputString) {
		if(inputString.length <=2) {
			$('#suggestions2').hide();
		    $("#txtRazon").html("");
			$("#txtRsocEmpresa").html("");
	        $('#EmprBuscar').val("0");
			if(inputString.length <=0)
			{
			tabla();
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
									tabla();
						});
			});
	}

	function tabla()
	{
		//window.open("inscripcioneshistorico/listar.asp?empresa=&programa=0")
		if(!cargada){
		jQuery("#list1").jqGrid({ 
		url:'inscripcioneshistorico/listar.asp?empresa=0&selMes=0&selAno=0', 
		datatype: "xml", 
		colNames:['Fecha', 'Rut', 'Nombre de Empresa', 'Cód. Curso', 'Nombre del Curso','Part', '&nbsp;','&nbsp;','&nbsp;'], 
		colModel:[
				   {name:'fecha_inicio',index:'PROGRAMA.FECHA_INICIO_', width:51, align:'center'}, 
				   {name:'rut_empresa',index:'EMPRESAS.RUT', width:52, align:'center'}, 
				   {name:'empresa',index:'EMPRESAS.R_SOCIAL'}, 
				   {name:'codigo',index:'CURRICULO.CODIGO', width:60},				   
				   {name:'nom_curso',index:'CURRICULO.NOMBRE_CURSO', width:100}, 
				   {name:'N_PARTICIPANTES',index:'AUTORIZACION.N_PARTICIPANTES', width:20, align:"right"},				   
				   {align:"right",editable:true, width:15},
				   {align:"right",editable:true, width:15},
				   {align:"right",editable:true, width:12}], 
		rowNum:300, 
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'PROGRAMA.FECHA_INICIO_', 
		viewrecords: true, 
		sortorder: "ASC", 
		caption:"Listado de Inscripciones" 
		}); 
	
	jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
														$("#txtRazon").html("");
														$("#txtRsocEmpresa").html("");
														$('#EmprBuscar').val("0");
														$('#txRutEmpresa').val("");
														$('#txRutEmpresa').defaultValue('Ingrese Empresa');
														document.getElementById("SelMes").selectedIndex = 0;
														document.getElementById("selAno").selectedIndex = 0;
														
jQuery("#list1").jqGrid('setGridParam',{url:"inscripcioneshistorico/listar.asp?empresa=0&selMes=0&selAno=0"}).trigger("reloadGrid")
													  } });
	    cargada=true;
		}
		else
		{
jQuery("#list1").jqGrid('setGridParam',{url:"inscripcioneshistorico/listar.asp?empresa="+$('#EmprBuscar').val()+"&selMes="+$('#SelMes').val()+"&selAno="+$('#selAno').val()}).trigger("reloadGrid")
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
			<h2><em style="text-transform: capitalize;">Historico de Inscripciones</em></h2>
            <br>
            <table width="650" border="0">
              <tr>
                <td width="90">Rut de Empresa :</td>
                <td width="100"><input id="txRutEmpresa" name="txRutEmpresa" type="text" tabindex="1" maxlength="20" size="20" onkeyup="lookup2(this.value);"/>
                  <div class="suggestionsBox21" id="suggestions2" style="display: none;position:absolute;z-index:2;left:522px"> <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
                    <div class="suggestionList2" id="autoSuggestionsList2"> &nbsp; </div>
                  </div></td>
                <td width="15"><input type="hidden" id="EmprBuscar" name="EmprBuscar"/></td>
                <td width="80"><label id="txtRazon" name="txtRazon"></label></td>
                <td width="315"><label id="txtRsocEmpresa" name="txtRsocEmpresa"></label></td>
              </tr>
              <tr>
                <td>Mes del Curso :</td>
                <td><select id="SelMes" name="SelMes" tabindex="2" onchange="tabla();">
                <option value="0">Todos</option>
                    <option value="1">Enero</option>
                    <option value="2">Febrero</option>
                    <option value="3">Marzo</option>
                    <option value="4">Abril</option>
                    <option value="5">Mayo</option>
                    <option value="6">Junio</option>
                    <option value="7">Julio</option>
                    <option value="8">Agosto</option>
                    <option value="9">Septiembre</option>
                    <option value="10">Octubre</option>
                    <option value="11">Noviembre</option>
                    <option value="12">Diciembre</option>
                  </select></td>
                <td colspan="3">&nbsp;</td>
              </tr>
              <tr>
                <td>Año del Curso : </td>
                <td colspan="4"><select id="selAno" name="selAno" tabindex="3" onchange="tabla();">
                  <option value="0">Todos</option>
                  <%
                                            For ano = 2010 To cdbl(year(now)) Step 1
                                            %>
                  <option value="<%=ano%>"><%=ano%></option>
                  <%
                                            Next
                                            %>
                </select></td>
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
<div id="dialog" title="Inscripciones Autorizadas">
</div>
<div id="dialogFac" title="Datos Facturación">
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