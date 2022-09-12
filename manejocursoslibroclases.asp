<!--#include file="cnn_string.asp"-->
<%
RelUser="0"
If Not IsNull(Session("relUsuario"))then RelUser=Session("relUsuario") end if

TipoUser="0"
If(Session("userTipo")<>"")then TipoUser=Session("userTipo") end if

IDUser="0"
If(Session("usuarioMutual")<>"")then IDUser=Session("usuarioMutual") end if
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
<script type="text/javascript">
var vacios=0;

$(document).ready(function(){					

	jQuery("#list1").jqGrid({ 
		url:'evaluacioncierre/lista_cursos.asp?r=<%=RelUser%>&tp=<%=TipoUser%>&usr=<%=IDUser%>', 
		datatype: "xml", 
		colNames:['Inicio','Cód. Curso', 'Nombre Curso','Relator','Ejec.','LC','PE', 'EC', 'ER', 'XLS'], 
		colModel:[
				   {name:'FECHA_INICIO_',index:'CONVERT(date,PROGRAMA.FECHA_INICIO_, 105)', width:50, align:'center'},				
				   {name:'CODIGO',index:'CURRICULO.CODIGO', width:64, align:'center'}, 
				   {name:'NOMBRE_CURSO',index:'dbo.MayMinTexto(CURRICULO.NOMBRE_CURSO)'}, 
				   {name:'instructor',index:'instructor'}, 
				   {name:'DIR_EJEC',index:'PROGRAMA.DIR_EJEC', width:26}, 				   
				   {align:"right",editable:true, width:15}, 
				   {align:"right",editable:true, width:15}, 				   
				   {align:"right",editable:true, width:15},
				   {align:"right",editable:true, width:15}, 				   
				   {align:"right",editable:true, width:15}], 
		rowNum:300, 
		height:350,		
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'CONVERT(date,PROGRAMA.FECHA_INICIO_, 105)', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Libros de Clases" 
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
		
	  $("#sendConf").dialog({
			autoOpen: false,
			bgiframe: true,
			height:150,
			width: 350,
			modal: true,
			overlay: {
				backgroundColor: '#f00',
				opacity: 0.5
			},			
		    closeOnEscape: false,
		    open: function(event, ui) { $(".ui-dialog-titlebar-close").hide(); }
	  });			
	
		$("#datos").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 200,
			width: 600,
			modal: true,
			buttons: {
				'Aceptar': function() {
						$(this).dialog('close');
				}
			}
		});

		$("#dialogCDN").dialog({
			bgiframe: true,
			autoOpen: false,
			height:380,
			width: 1050,
			modal: true,
			buttons: {
				Cerrar: function() {
					$(this).dialog('close');
				},
				'Guardar': function() {
					if($('#ec').val()=="0")
					{
						for (i=1; i<=$('#countFilas').val(); i++)
						{
							if($('#A'+i).val()=='' || $('#C'+i).val()=='')
							{
								$('#A'+i).val("");
								$('#C'+i).val("");
								$('#evaluacion'+i).html("Reprobado");
								$('#E'+i).val("R");
							}
							
							if($('#Ar'+i).val()=='' || $('#Cr'+i).val()=='')
							{
								$('#Ar'+i).val("");
								$('#Cr'+i).val("");
								$('#evaluacionr'+i).html("Reprobado");
								$('#Er'+i).val("R");
							}						
						}
	
								//window.open($('#frmingevacdn').attr('action'),$('#frmingevacdn').serialize());
						$.post($('#frmingevacdn').attr('action'),$('#frmingevacdn').serialize(),function(d){
																						$("#list1").trigger("reloadGrid"); 
																						});
						$("#list1").trigger("reloadGrid");
						$(this).dialog('close');
					}
					else
					{
						vacios=0;
						for (i=1; i<=$('#countFilas').val(); i++)
						{
							if(($('#A'+i).val()=='' && $('#C'+i).val()=='' && $('#Ar'+i).val()=='' && $('#Cr'+i).val()=='') || ($('#A'+i).val()!='' && $('#C'+i).val()=='' && $('#Ar'+i).val()=='' && $('#Cr'+i).val()=='') || ($('#A'+i).val()=='' && $('#C'+i).val()!='' && $('#Ar'+i).val()=='' && $('#Cr'+i).val()=='') || $('#EDT'+i).html()=="Guardar")
							{
								vacios=1;
							}
						}
						
						if(vacios==0)
						{
							$('#sendMsn').html('Espere Mientras se Envia Notificación.');
							$("#sendConf").dialog('open');
							$.post($('#frmingevacdn').attr('action'),$('#frmingevacdn').serialize(),function(d){
																							$("#sendConf").dialog('close');
																							$('#sendMsn').html('');
																							$("#list1").trigger("reloadGrid"); 
																							});
							$("#list1").trigger("reloadGrid");
							$(this).dialog('close');
						}
						else
						{
							$("#datos").dialog('open');
						}
					}
				}	
			}
		});	
	
		$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height:380,
			width: 980,
			modal: true,
			buttons: {
				'Guardar': function() {
					//vacios=0;
					for (i=1; i<=$('#countFilas').val(); i++)
					{
						if($('#A'+i).val()=='' || $('#C'+i).val()=='')
						{
							//vacios=1;
							$('#A'+i).val("");
							$('#C'+i).val("");
							$('#evaluacion'+i).html("Reprobado");
							$('#E'+i).val("R");
							cambiaEstado();
						}
					}
					/*
					if(vacios==0)
					{*/

						if($("#frmcierreeval").valid())
						{
							//window.open($('#frmcierreeval').attr('action')+','+$('#frmcierreeval').serialize());
							$.post($('#frmcierreeval').attr('action'),$('#frmcierreeval').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
 																						   });
							$("#list1").trigger("reloadGrid");
							
							if($('#ec2').val()=="1"){
							  evaCierre2($('#i').val(),$('#r').val(),$('#u').val(),$('#b').val(),$('#t').val(),1);
							}
							
							$(this).dialog('close');
						}
					/*}
					else
					{
						$("#datos").dialog('open');
					}*/
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	
		$("#dialogPre").dialog({
			bgiframe: true,
			autoOpen: false,
			height:380,
			width: 980,
			modal: true,
			buttons: {
				'Guardar': function() {
					vacios=0;
					for (i=1; i<=$('#countFilas').val(); i++)
					{
						if(/*$('#A'+i).val()=='' || */$('#C'+i).val()=='')
						{
							vacios=1;
							//$('#A'+i).val("");
							//$('#C'+i).val("");
							//$('#evaluacion'+i).html("Reprobado");
							//$('#E'+i).val("R");
						}
					}
						
					if(vacios==0)
					{
						/*if($("#frmcierrepre").valid())
						{*/
							$.post($('#frmcierrepre').attr('action'),$('#frmcierrepre').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
 																						   });
							$("#list1").trigger("reloadGrid");
							$(this).dialog('close');
						//}
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
			height: 550,
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
			title: 'Libro de Clases'
		});
	});	

	function update(prog,relator,bloque){
		//alert(prog+' '+relator+' '+bloque);
		documento("libroclases/pdf3.asp?prog="+prog+"&relator="+relator+"&idbloque="+bloque);
		
	}
	
	function documento(arch){
		$("#ifPagina").attr('src',arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
	}
	
	function edtInfoRelCan(id){
		    $('#EDT'+id).html("Editar");
			$('#ASR'+id).show();
			$('#CAR'+id).show();
			$('#Ar'+id).hide();
			$('#Cr'+id).hide();
			$('#Ar'+id).val($('#ASR'+id).html());
			$('#Cr'+id).val($('#CAR'+id).html());
			$('#C_DT'+id).hide();
			cambiaEstadoRel();
	}
	
	function edtInfoRel(id){
		//alert($('#EDT'+id).html());
		if($('#EDT'+id).html()=="Editar")
		{
			$('#EDT'+id).html("Guardar");
			$('#ASR'+id).hide();
			$('#CAR'+id).hide();
			$('#Ar'+id).show();
			$('#Cr'+id).show();
			$('#C_DT'+id).show();
		}
		else
		{
			$('#EDT'+id).html("Editar");
			$('#E_DT'+id).val("1");
			$('#ASR'+id).show();
			$('#CAR'+id).show();
			$('#Ar'+id).hide();
			$('#Cr'+id).hide();
			$('#ASR'+id).html($('#Ar'+id).val());
			$('#CAR'+id).html($('#Cr'+id).val());
			$('#C_DT'+id).hide();
		}

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
	
	function excel(prog,relator){
		window.open("evaluacioncierre/xlsCredencial.asp?prog="+prog+"&relator="+relator,'Informe')
	}
	
	function cambiaEstadoRel()
	{
		for (i=1; i<=$('#countFilas').val(); i++)
		{
		    if($('#txtIdMutual').val()=="33"){
				if($('#Ar'+i).val()!='' && $('#Ar'+i).val()==100 && $('#Cr'+i).val()!='' && $('#Cr'+i).val()>=65)
				{
					if($('#Cr'+i).val()>=65 && $('#Cr'+i).val()<=69){
						$('#evaluacionr'+i).html("Con Obs.");
						$('#Er'+i).val("C");
					}
					else
					{
						$('#evaluacionr'+i).html("Aprobado");
						$('#Er'+i).val("A");
					}
				}
				else
				{
					$('#evaluacionr'+i).html("Reprobado");
					$('#Er'+i).val("R");
				}
		   }
 	           else if ($('#txtIdMutual').val()=="2317" || $('#txtIdMutual').val()=="2319" || $('#txtIdMutual').val()=="2324" || $('#txtIdMutual').val()=="2325" || $('#txtIdMutual').val()=="2332" || $('#txtIdMutual').val()=="2334" || $('#txtIdMutual').val()=="2335" || $('#txtIdMutual').val()=="2318" || $('#txtIdMutual').val()=="2341" || $('#txtIdMutual').val()=="2342" || $('#txtIdMutual').val()=="2343" || $('#txtIdMutual').val()=="2326" || $('#txtIdMutual').val()=="2320" || $('#txtIdMutual').val()=="2321" || $('#txtIdMutual').val()=="2322" || $('#txtIdMutual').val()=="2333" || $('#txtIdMutual').val()=="2330" || $('#txtIdMutual').val()=="2331" || $('#txtIdMutual').val()=="2327" || $('#txtIdMutual').val()=="2328" || $('#txtIdMutual').val()=="82" || $('#txtIdMutual').val()=="2316" || $('#txtIdMutual').val()=="2337" || $('#txtIdMutual').val()=="2338" || $('#txtIdMutual').val()=="2339" || $('#txtIdMutual').val()=="2340"){
				if($('#Ar'+i).val()!='' && $('#Ar'+i).val()==100 && $('#Cr'+i).val()!='' && $('#Cr'+i).val()>=80)
				{
					$('#evaluacionr'+i).html("Aprobado");
					$('#Er'+i).val("A");
				}
				else
				{
					$('#evaluacionr'+i).html("Reprobado");
					$('#Er'+i).val("R");
				}
		   }
		   else
		   {
				if($('#Ar'+i).val()!='' && $('#Ar'+i).val()==100 && $('#Cr'+i).val()!='' && $('#Cr'+i).val()>=70)
				{
					$('#evaluacionr'+i).html("Aprobado");
					$('#Er'+i).val("A");
				}
				else
				{
					$('#evaluacionr'+i).html("Reprobado");
					$('#Er'+i).val("R");
				}
		   }
		}
	}	
	
	/*function cambiaEstado()
	{
		for (i=1; i<=$('#countFilas').val(); i++)
		{
		   if($('#txtIdMutual').val()=="33"){
				if($('#A'+i).val()!='' && $('#A'+i).val()==100 && $('#C'+i).val()!='' && $('#C'+i).val()>=65)
				{
					if($('#C'+i).val()>=65 && $('#C'+i).val()<=69){
						$('#evaluacion'+i).html("Con Obs.");
						$('#E'+i).val("C");
					}
					else
					{
						$('#evaluacion'+i).html("Aprobado");
						$('#E'+i).val("A");
					}
				}
				else
				{
					$('#evaluacion'+i).html("Reprobado");
					$('#E'+i).val("R");
				}
		   }
		   else if ($('#txtIdMutual').val()=="2317" || $('#txtIdMutual').val()=="2319" || $('#txtIdMutual').val()=="2324" || $('#txtIdMutual').val()=="2325" || $('#txtIdMutual').val()=="2332" || $('#txtIdMutual').val()=="2334" || $('#txtIdMutual').val()=="2335" || $('#txtIdMutual').val()=="2318" || $('#txtIdMutual').val()=="2341" || $('#txtIdMutual').val()=="2342" || $('#txtIdMutual').val()=="2343" || $('#txtIdMutual').val()=="2326" || $('#txtIdMutual').val()=="2320" || $('#txtIdMutual').val()=="2321" || $('#txtIdMutual').val()=="2322" || $('#txtIdMutual').val()=="2333" || $('#txtIdMutual').val()=="2330" || $('#txtIdMutual').val()=="2331" || $('#txtIdMutual').val()=="2327" || $('#txtIdMutual').val()=="2328" || $('#txtIdMutual').val()=="82" || $('#txtIdMutual').val()=="2316" || $('#txtIdMutual').val()=="2337" || $('#txtIdMutual').val()=="2338" || $('#txtIdMutual').val()=="2339" || $('#txtIdMutual').val()=="2340"){
			   	if($('#A'+i).val()!='' && $('#A'+i).val()==100 && $('#C'+i).val()!='' && $('#C'+i).val()>=80)
				{
					$('#evaluacion'+i).html("Aprobado");
					$('#E'+i).val("A");
				}
				else
				{
					$('#evaluacion'+i).html("Reprobado");
					$('#E'+i).val("R");
				}
		   }
		   else
		   {
			   	if($('#A'+i).val()!='' && $('#A'+i).val()==100 && $('#C'+i).val()!='' && $('#C'+i).val()>=70)
				{
					$('#evaluacion'+i).html("Aprobado");
					$('#E'+i).val("A");
				}
				else
				{
					$('#evaluacion'+i).html("Reprobado");
					$('#E'+i).val("R");
				}
		   }
		}
	}*/
	function cambiaEstado()
	{
		
		calificacionNecesaria = $('#CalificacionNecesaria').val();
		AsistenciaNecesaria = $('#AsistenciaNecesaria').val();
		console.log(calificacionNecesaria);
		console.log(AsistenciaNecesaria);
		for (i=1; i<=$('#countFilas').val(); i++)
		{		console.log($('#C'+i).val());
			   	if($('#A'+i).val()!='' && parseInt($('#A'+i).val()) >= AsistenciaNecesaria)
				{
					if($('#C'+i).val()!='' && parseInt($('#C'+i).val()) >= calificacionNecesaria){
						$('#evaluacion'+i).html("Aprobado");
						$('#E'+i).val("A");
					}else{
					$('#evaluacion'+i).html("Reprobado");
					$('#E'+i).val("R");
					}
				}
				else
				{
					$('#evaluacion'+i).html("Reprobado");
					$('#E'+i).val("R");
				}
		   
		}
	}
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}
	
	function evaCierre(i,relator,u,b,t){
		//alert(i+' '+relator);
		$.post("evaluacioncierre/frmIngresoEval.asp",
			   {id:i,relator:relator,u:u,b:b,t:t},
			   function(f){
				    $('#dialog').html(f);
					tableToGrid("#mytable");
					cambiaEstado();
		});
		 $('#dialog').dialog('open');
	}
	
	function evaCierre3(i,relator,u,b,t,ec){
		$.post("evaluacioncierre/frmIngresoEvalCdn.asp",
			   {id:i,relator:relator,u:u,b:b,t:t,ec:ec},
			   function(f){
				    $('#dialog').html(f);
					tableToGrid("#mytable");
		});
		 $('#dialog').dialog('open');
	}
	
	function evaCierre2(i,relator,u,b,t,ec){
		//alert(i+' '+relator);
		$.post("evaluacioncierre/frmIngEvaCdn.asp",
			   {id:i,relator:relator,u:u,b:b,t:t,ec:ec},
			   function(f){
				    $('#dialogCDN').html(f);
					tableToGrid("#mytable2");
		});
		 //$('#Cerrar Curso').attr('disabled', true);
		 $('#dialogCDN').dialog('open');
		 //$('#Cerrar Curso').attr('disabled', true);
		 //$('#dialogCDN').dialog({buttons: [id:'myButton', disabled:true, text:'button', ...]});
		 //$('#dialogCDN').dialogButtons('Cerrar Curso', 'disabled');
	}	
	
	function evaCierrePre(i,relator,u,b,t){
		//alert(i+' '+relator);
		$.post("evaluacioncierre/frmIngresoPre.asp",
			   {id:i,relator:relator,u:u,b:b,t:t},
			   function(f){
				    $('#dialogPre').html(f);
					tableToGrid("#mytable");
		});
		 $('#dialogPre').dialog('open');
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
            
				<li class="first"><a href="manejocursoslibroclases.asp">Libro de Clases</a></li>
                <%if(Session("usuarioMutual")="108" or Session("usuarioMutual")="32")then%>
                <li><a href="manejocursosempresas.asp">Empresas</a></li>
                <%end if%>
                 <%if Session("usuarioMutual")="144" then%>
                <li><a href="manejocursoslibroclases_bkp_05-02-2015.asp">Libro de Clases Nuevo</a></li>
                <%end if%>
                <p>&nbsp;</p>
                <p>&nbsp;</p>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
			<h2><em style="text-transform: capitalize;">Libro de Clases</em></h2>
          <table id="list1"></table> 
            <div id="pager1"></div> 
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro de Evaluación de Curso">
</div>
<div id="dialogPre" title="Registro de Pre-Evaluación de Curso">
</div>
<div id="dialogCDN" title="Registro de Evaluación de Curso">
</div>
<div id="datos" title="Registro de Evaluación de Curso">
	<p>Por Favor, Ingresar Toda la Información Requerida y Guarde todos los cambios realizados a las Evaluaciones del Relator.</p>
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="sendConf" title="Enviando">
     <p align="center"><label id="sendMsn" name="sendMsn"></label></br></br><img src="images/loadfbk.gif"/></p>
</div>
</body>
</html>