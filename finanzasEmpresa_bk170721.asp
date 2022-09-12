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
<link rel="stylesheet" type="text/css" href="css/ui.jqgrid.css"/>
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
<script src="js/jquery.jqGrid.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
var cargada=false;
var idc=0;
$(document).ready(function(){					
tabla();
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 470,
			width: 990,
			modal: true,
			buttons: {
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
		});
		
		$("#dialogContacto").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 470,
			width: 990,
			modal: true,
			buttons: {
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
		});
		
		$("#dialog_Confirmar").dialog({
			bgiframe: true,
            autoOpen: false,
			resizable: false,
			height: "auto",
			width: 400,
			modal: true,
			buttons: {
				"Aceptar": function() {
					$.post("finanzasempresa/eliminar_trab.asp",{idc:idc},
					function(d){
						$("#list2").trigger("reloadGrid");                                            
					});
					$(this).dialog('close');
				},
				Cancelar: function() {
					$( this ).dialog( "close" );
				}
			}
		});
		
		$("#dialogFormulario").dialog({
			autoOpen: false,
			bgiframe: true,
			height:400,
			width: 500,
			modal: true,
			buttons: {	
				'Aceptar': function() {
					if($("#frmTrabajador").valid()){
						//window.open($('#frmTrabajador').attr('action')+'?'+$('#frmTrabajador').serialize())
						$.post($('#frmTrabajador').attr('action')+'?'+$('#frmTrabajador').serialize(),function(d){
								$("#list2").trigger("reloadGrid"); 
							});
						$("#list2").trigger("reloadGrid");
						$(this).dialog('close');
					}
				},
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
		});
		
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
		
		$("#dialog_historico").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 350,
			width: 900,
			modal: true,
			buttons: {
				Cerrar: function() {
					$(this).dialog('close');
				}
			}
		});		
		
		$("#dialog_Bloquear").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 300,
			width: 550,
			modal: true,
			buttons: {
				'Aceptar': function() {
						if($("#frmBloqueaEmp").valid())
						{
								//window.open($('#frmFactAnular').attr('action')+'?'+$('#frmFactAnular').serialize())
								$.post($('#frmBloqueaEmp').attr('action')+'?'+$('#frmBloqueaEmp').serialize(),function(d){
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
	});	

	function tabla()
	{
		sUrl='finanzasempresa/listar.asp?campo='+$('[name="rbtn_buscar"]:checked').val()+'&tipo='+$('[name="rbtn_tipo"]:checked').val()+'&texto='+$("#txt_buscar").val();
		if(cargada){
			jQuery("#list1").jqGrid('setGridParam',{url:sUrl});
			jQuery("#list1").trigger("reloadGrid"); 
		}else{
			 cargada=true;
   			 jQuery("#list1").jqGrid({ 
				url:sUrl, 
				datatype: "xml", 
				colNames:['Rut', 'Razón Social', '&nbsp;', '&nbsp;', '&nbsp;', '&nbsp;'], 
				colModel:[
						   {name:'rut',index:'rut', width:100, align:"right"}, 
						   {name:'r_social',index:'r_social', width: 400}, 
						   {align:"right",editable:true, width:20},
						   {align:"right",editable:true, width:20},	
						   {align:"right",editable:true, width:20},					   
						   {align:"right",editable:true, width:20} ], 
				rowNum:300, 
				autowidth: true, 
				rowList:[300,500,1000], 
				pager: jQuery('#pager1'), 
				sortname: 'rut', 
				viewrecords: true, 
				sortorder: "asc", 
				caption:"Listado de Empresas" 
				}); 
			
			jQuery("#list1").jqGrid('navGrid','#pager1',{edit:false,add:false,del:false,search:false,refresh:false});
			jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"",
															  title:"Refrescar", 
															  buttonicon :'ui-icon-refresh', 
															  onClickButton:function(){ 
															     $("#txt_buscar").val("");
																 tabla();
															  } }); 
															  
			jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"Excel",
															  title:"Exportar a Excel", 
															  buttonicon :'ui-icon-script',
															  onClickButton:function(){ 
															     window.open("consultasempresa/xls.asp?c=1",'Informe')
															  } }); 
															  
		  }
	}
		
	function update(i, e){
		//window.open("finanzasempresa/frmEmpresa.asp?id=i");
		$.post("finanzasempresa/frmEmpresa.asp",{id:i,ide:e},
			   function(f){
				   $('#dialog').html(f);
		});
		 $('#dialog').dialog('open');
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
	
	function contacto(i){
		//window.open("finanzasempresa/frmEmpresa.asp?id=i");
		$.post("finanzasempresa/frmContactos.asp",{id:i},
			   function(f){
				   $('#dialogContacto').html(f);
				   grillaContacto17(i);
		});
		$('#dialogContacto').dialog('open');		 
	}
	
	function update_usr(i, x, y){
		//window.open("finanzasempresa/frmEmpresa.asp?id=i");
		$.post("finanzasempresa/frmTrabajador.asp",{idc:i, ide:x, val:y},
			   function(f){
				   $('#dialogFormulario').html(f);				   
				   validarFrmTrab();
				   $('#trComentario').show();
				   if($('#val1').val()=='0'){
						$('#trComentario').hide();
				   }
				});
		$('#dialogFormulario').dialog('open');
	}
	
	function delete_usr(i){
		//$.post("finanzasempresa/eliminar_trab.asp",{idc:i},
		//function(d){
		//	$("#list2").trigger("reloadGrid");							  
		//});	
		idc = i;
		$('#dialog_Confirmar').dialog('open');
	}
	
	function grillaContacto17(i){
					jQuery("#list2").jqGrid({ 
						url:"finanzasempresa/listar2.asp?id="+i, 
						datatype: "xml", 
						colNames:['NOMBRE','EMAIL','FONO','CARGO','COMENTARIOS','',''],
						colModel:[  {name:'NOMBRES',index:'NOMBRES',width:110}, 
									{name:'EMAIL',index:'EMAIL',width:100}, 
									{name:'FONO',index:'FONO',width:50}, 
									{name:'CARGO',index:'CARGO',width:100},
									{name:'COMENTARIO',index:'COMENTARIO',width:100},
									{align:"right",width:20},	
						   			{align:"right",width:20} ],
						rowNum:100, 
						rownumbers: true, 
						rownumWidth: 40, 
						width: "850", 
						pager: jQuery('#pager2'), 
						sortname: 'NOMBRES', 
						viewrecords: true, 
						sortorder: "desc", 
						caption:"Listado Contactos de Contabilidad"
					}); 
					jQuery("#list2").jqGrid('navGrid','#pager2',{edit:false,add:false,del:false,search:false,refresh:false});
					jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
						title:"Refrescar", 
						buttonicon :'ui-icon-refresh', 
						onClickButton:function(){ 
							$("#list2").trigger("reloadGrid");
						} 
					});
					jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
						title:"Agregar contacto", 
						buttonicon :'ui-icon-plus', 
						onClickButton:function(){						
							$.post("finanzasempresa/frmTrabajador.asp",{idc:0, ide:i},							
							function(f){
								$('#dialogFormulario').html(f);								
								validarFrmTrab();
							});
						$('#dialogFormulario').dialog('open');
					}}); 
	}
	
	function validarFrmTrab(){
		$("#frmTrabajador").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txtNombre:{
				required:true
			},
			txtEmail:{
				required:true,
				email:true
			},	
			txtFono:{
                required:true                        
            },		
			txtCargo:{
                required:true                        
            }
		},
		messages:{			
			txtNombre:{
				required:"&bull; Ingrese Nombre del Contacto."
			},
			txtEmail:{
				required:"&bull; Ingrese Email del Contacto.",
				email:"&bull; Formato de correo invalido."
			},
			txtFono:{
				required:"&bull; Ingrese Teléfono del Contacto."
			},
			txtCargo: {
                required:"&bull; Ingrese Cargo del Contacto."                        
            }
		}
	});
	}
	
	function verHist(emp){
		$.post("empresa/frmHistoricos.asp",		  
			   function(f){
				    $('#dialog_historico').html(f);				
					sUrlHist='empresa/listarHist.asp?e='+emp; 

		/*if(carga_hist){
			jQuery("#listHist").jqGrid('setGridParam',{url:sUrlHist});
			jQuery("#listHist").trigger("reloadGrid"); 
		}else{*/
					//carga_hist=true;
					jQuery("#listHist").jqGrid({ 
						scroll: 1, 					
						url:sUrlHist, 
						datatype: "xml", 
						colNames:['Estado', 'Fecha', 'Usuario', 'Descripción'],
						colModel:[
									{name:'EST',index:'EST', width:40}, 
									{name:'FECHA_ESTADO',index:'FECHA_ESTADO', width:70}, 
									{name:'nom_usuario',index:'nom_usuario', width:90}, 
									{name:'DESCRIPCION',index:'DESCRIPCION'}], 
						rowNum:100, 
						rownumbers: true, 
						rownumWidth: 40, 
						width: "850", 
						pager: jQuery('#pagerHist'), 
						sortname: 'eb.FECHA_ESTADO', 
						viewrecords: true, 
						sortorder: "desc", 
						caption:"Listado"
					}); 
		//}
		$('#dialog_historico').dialog('open');
		   });
	 //});	
	}	
	
	function actDes_Empresa(emp,est){
		/*$.post("empresa/actDesEmpresa.asp",{empresa:i,estado:estado},function(d){
			   																		$("#list1").trigger("reloadGrid"); 
																				 });*/
		if(<%=Session("usuarioMutual")%>=='8' || <%=Session("usuarioMutual")%>=='178'){																					  			$.post("empresa/frmBloquear.asp",{u:<%=Session("usuarioMutual")%>,c:1,e:emp,est:est},
			   function(f){
				    $('#dialog_Bloquear').html(f);
					$('#razonBloquear').focus();
					valFrmBloquear();
			});
	    	$('#dialog_Bloquear').dialog('open');		
		}
	}
	
	function valFrmBloquear(){
		$("#frmBloqueaEmp").validate({
		errorContainer: "#messageBox2",
  		errorLabelContainer: "#messageBox2 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			razonBloquear:{
				required:true
			}
		},
		messages:{
			razonBloquear:{
				required:"&bull; Ingrese las Razones por las que bloqueará a la Empresa."
			}
		}
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
			<h2><em style="text-transform: capitalize;">Empresas</em></h2>
           <center><table width="709" border="0">
                    <tr>
                    <td width="650"><div id="buscarContent" style="border:1px solid #000099;display:block" >
					<div>
						<table cellpadding="0" cellspacing="0" style="border:0px;width:100%;">
							<tr>
								<td align="left"><h3><em style="text-transform: capitalize;">Busqueda</em></h3></td>
							</tr>
						</table>
					</div>
					<table border="0" align="center" cellpadding="1" cellspacing="1">
						<tr>
							  <td width="400" align="center">
								
								<input name="rbtn_buscar" id="rbtn_buscar" type="radio" value="RUT" checked="checked" />Por Rut
           				        <input name="rbtn_buscar" id="rbtn_buscar" type="radio" value="R_SOCIAL"/>Por Razón Social
								</td>
						</tr>
						<tr>
							<td><input name="txt_buscar" type="text" id="txt_buscar" size="80"/></td>
							<td><input type="button" class="boton" value="Buscar" id="btnBuscar" onclick="tabla();"/></td>
						</tr>
						<tr>
                          <td align="center">
								<input name="rbtn_tipo" type="radio" value=" = '"/>Es exacta
								<input name="rbtn_tipo" type="radio" value=" LIKE '%" checked="checked"/>Que contienen
							    	  
						  </td>
						</tr>
					</table>
				</div></td>
              </tr>
            </table></center>
            <table width="200" border="0">
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Empresa"></div>
<div id="pantContrasena" title="Cambiar Contraseña"></div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
<div id="dialog_Bloquear" title="Bloquear/Desbloquear Empresa"></div>
<div id="dialog_historico" title="Hist&oacute;rico Bloquear/Desbloquear Empresa"></div>
<div id="dialogContacto" title="Contactos de Contabilidad"></div>
<div id="dialogFormulario" title="Formulario Contactos de Contabilidad"></div>
<div id="dialog_Confirmar" title="Confirmaci&oacute;n">
  <p><span class="ui-icon ui-icon-alert" style="float:left;"></span>
  &nbsp;&nbsp;¿Esta seguro de realizar esta acci&oacute;n?</p>
</div>
</body>
</html>