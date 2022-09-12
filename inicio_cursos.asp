<%
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
		height:200px;
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
</style>
<script type="text/javascript">
var i=1;
var cargada=false;

$(document).ready(function(){					
	$('#txRut').defaultValue('Todos');
	$('#txRutEmpresa').defaultValue('Todas');
	llena_curriculo(0);
	$('#txtFInicio').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
	$('#txtFTermino').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });	
	
	$("#CRTMasivo").dialog({
			autoOpen: false,
			bgiframe: true,
			height:180,
			width: 350,
			modal: true,
			buttons: {
				Aceptar: function() {
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
			title: 'Certificado de Asistencia a Curso de Capacitación'
		});
	});	
	
	
	function tabla()
	{
		if(!cargada){
		//window.open("bhp/listar.asp?trabajador="+$('#Trabajador').val()+"&empresa="+$('#Empresa').val()+"&curso=0");
		jQuery("#list1").jqGrid({ 
		url:'bhp/listar.asp?trabajador=0&empresa=0&curso=0&inicio=&termino=&u='+<%=userID%>, 
		datatype: "xml", 
		colNames:['Rut Emp.','Empresa', 'Rut Trab.', 'Nombre Trabajador', 'Fecha','Cod. Autentificación','EV','EI','&nbsp;'], 
		colModel:[
				   {name:'RUTE',index:'E.RUT', width:30, align:'right'}, 
				   {name:'R_SOCIAL',index:'E.R_SOCIAL', width:85}, 
				   {name:'RUT',index:'T.RUT', width:30, align:'right'}, 
				   {name:'NOMBRES',index:'T.NOMBRES', width:80}, 
				   {name:'FECHA_TERMINO',index:'P.FECHA_TERMINO', width:30, align:'center'}, 
				   {name:'COD_AUTENFICACION',index:'H.COD_AUTENFICACION', width:62, align:'center'},
				   {align:"right", width:8, align:'center'},
				   {align:"right", width:8},
				   {align:"right", width:8}], 
		rowNum:300, 
		autowidth: true, 
		rowList:[300,500,1000], 
		pager: jQuery('#pager1'), 
		sortname: 'E.R_SOCIAL', 
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
													  
	jQuery("#list1").jqGrid('navButtonAdd',"#pager1",{caption:"Certificado Masivo",
													  title:"Certificado Masivo", 
													  buttonicon :'ui-icon-script', 
													  onClickButton:function(){ 
													  if($('#Empresa').val()!="" && $('#Curriculo').val()!=0)
													  {
 									     documento("libroventas/Cert_masivo.asp?emp="+$('#Empresa').val()+"&curso="+$('#Curriculo').val());
													  }
													  else
													  {
								   $("#txtMsn").html("Seleccione Curso e Ingrese Rut de Empresa para generar Certificado Masivo");
														  $("#CRTMasivo").dialog('open');
													  }
													  } }); 													  
													  
		cargada=true;
	}else{
jQuery("#list1").jqGrid('setGridParam',{url:"bhp/listar.asp?trabajador="+$('#Trabajador').val()+"&empresa="+$('#Empresa').val()+"&curso="+$('#Curriculo').val()+'&inicio='+$('#txtFInicio').val()+'&termino='+$('#txtFTermino').val()+'&u='+<%=userID%>}).trigger("reloadGrid")
	}
	}
	
	function llena_curriculo(id_curriculo){
		$("#Curriculo").html("");
		$.get("bhp/curriculo.asp?u="+<%=userID%>,
					function(xml){
						
						$("#Curriculo").append("<option value='0'>Seleccione</option>");
						if(<%=userID%>!='207'){
						$('row',xml).each(function(i) { 
							if(id_curriculo==$(this).find('ID_MUTUAL').text())
								$("#Curriculo").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" selected="selected">'+
																		$(this).find('NOMBRES').text()+ '</option>');
							else
								$("#Curriculo").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" >'+
																		$(this).find('NOMBRES').text()+ '</option>');
						});
							//tabla();
						}
						else
						{
						   $("#Curriculo").append("<option value='82'>(12.37.949-138) - inducción codelco vp proyectos</option>");
							//tabla();
						}
					});
			tabla();
	}
	
	 function lookup(inputString) {
		if(inputString.length <=2) {
			$('#suggestions').hide();
			$("#txtNombre").html("");
		    $("#txtTrabajador").html("");
			$('#Trabajador').val("");
			if(inputString.length <=0)
			{
			tabla();
			}
		} else {
				$.post("bhp/sugTrabajador.asp", {txt: inputString}, function(data){
						if(data.length >0) {
								$('#suggestions').show();
								$('#autoSuggestionsList').html(data);
						}
				});
		}
	}

	function fill(id,rut) {
	   $('#Trabajador').val(id);
	   cargaDatosTrabajador(id);
	   $('#suggestions').hide();
	}

	function cargaDatosTrabajador(id)	
	{
		$("#txRut").val("");
		$("#txtNombre").html("");
		$("#txtTrabajador").html("");
 
		$.get("bhp/datosTrabajador.asp",
						 {id:id},
						 function(xml){
									$('row',xml).each(function(i) { 
									$("#txtNombre").html("Nombre :");
									$("#txtTrabajador").html($(this).find('NOMBRE').text());
									$("#txRut").val($(this).find('RUT').text());
									tabla();
						});
			});
	}

	 function lookup2(inputString) {
		if(inputString.length <=2) {
			$('#suggestions2').hide();
		    $("#txtRazon").html("");
			$("#txtRsocEmpresa").html("");
	        $('#Empresa').val("");
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
	   $('#Empresa').val(id);
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

	function certificados2(programa,empresa,trabajador,relator){
		documento("libroclases/Certificado.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);
	}
	
	function certificados(programa,empresa,trabajador,relator,curso){
		if(curso!='82'){
			documento("libroclases/Certificado.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);
		}else{
			//alert("libroclases/CertificadoCodelco.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);
			documento("libroclases/CertificadoCodelco.asp?prog="+programa+"&empresa="+empresa+"&trabajador="+trabajador+"&relator="+relator);	
		}
	}	
	
	function documento(arch){
		$("#ifPagina").attr('src',arch);
		if(!$('#Doc').dialog('isOpen'))
			$('#Doc').dialog('open');
	}
	
	</script>
</head>
<body>
<%
if(Session("usuarioMutual")="")then
Session.Abandon
Response.Redirect("bhp.asp")
end if
%>
<div id="header">
	<h1><img src="images/logo.png"/></h1>
	<ul>
   <li><a href="bhp.asp" accesskey="5" title="" class="selItem">Cerrar Sesión</a></li>
	</ul>
</div>
<div id="content"  class="bg1">
            <h2><em style="text-transform: capitalize;">Histórico de Capacitaciones</em></h2>
             <table border="0">
              <tr>
                <td width="100">&nbsp;</td>
                <td width="40">&nbsp;</td>
                <td width="60">&nbsp;</td>
                <td width="400">&nbsp;</td>
              </tr>
              <tr>
                <td>Rut Trabajador :</td>
                <td><input id="txRut" name="txRut" type="text" tabindex="1" maxlength="30" size="30" onkeyup="lookup(this.value);"/>
                   <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:1;left:345px">
                        <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
                        <div class="suggestionList" id="autoSuggestionsList">
                          &nbsp;
                        </div>
                   </div><input type="HIDDEN" id="Trabajador" name="Trabajador"/></td>
                   <td><label id="txtNombre" name="txtNombre"></label></td>
                   <td><label id="txtTrabajador" name="txtTrabajador"></label></td>
              </tr>
              <tr>
                <td>Rut Empresa :</td>
                <td><input id="txRutEmpresa" name="txRutEmpresa" type="text" tabindex="2" maxlength="30" size="30" onkeyup="lookup2(this.value);"/>
                   <div class="suggestionsBox2" id="suggestions2" style="display: none;position:absolute;z-index:1;left:345px">
                        <img src="images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
                        <div class="suggestionList2" id="autoSuggestionsList2">
                          &nbsp;
                        </div>
                   </div><input type="hidden" id="Empresa" name="Empresa"/></td>
                    <td><label id="txtRazon" name="txtRazon"></label></td>
                    <td><label id="txtRsocEmpresa" name="txtRsocEmpresa"></label></td>
              </tr>
              <tr>
                <td>Curso :</td>
                <td colspan="3"><select id="Curriculo" name="Curriculo" tabindex="1" onchange="tabla();" style="width:50em;"></select></td>
              </tr>
              <tr>
                <td>Desde :</td>
                <td><input id="txtFInicio" name="txtFInicio" type="text" tabindex="4" maxlength="50" size="12" onchange="tabla();"/></td>
                <td>Hasta :</td>
                <td><input id="txtFTermino" name="txtFTermino" type="text" tabindex="5" maxlength="50" size="12" onchange="tabla();"/></td>
              </tr>               
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <table id="list1"></table> 
            <div id="pager1"></div> 
            <br />
                <table width="300" border="0">
                <tr>
                	<td><strong>Simbologia</strong></td>
                </tr>
                  <tr>
                    <td><table width="100" border="0">
                          <tr>
                            <td colspan="2"><strong>Columna EV</strong></td>
                          </tr>
                          <tr>
                            <td><span style="color:#000000; position:relative;">A<span style="color:#66ff00;top:-1px;left:-1px;position:absolute;">A</span></span></td>
                            <td>Aprobado</td>
                          </tr>
                          <tr>
                            <td><span style="color:#000000; position:relative;">R<span style="color:#ff0000;top:-1px;left:-1px;position:absolute;">R</span></span></td>
                            <td>Reprobado</td>
                          </tr>
                        </table>
                    </td>
                    <td><table width="200" border="0">
                          <tr>
                            <td colspan="2"><strong>Columna EI</strong></td>
                          </tr>
                          <tr>
                            <td><span style="color:#000000; position:relative;">ILB<span style="color:#66ff00;top:-1px;left:-1px;position:absolute;">ILB</span></span></td>
                            <td>Inscripción Liberada</td>
                          </tr>
                          <tr>
                            <td><span style="color:#000000; position:relative;">IPL<span style="color:#ff0000;top:-1px;left:-1px;position:absolute;">IPL</span></span></td>
                            <td>Inscripción Pendiente de Liberar</td>
                          </tr>
                        </table>
                    </td>
                  </tr>
                </table>
		</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="Doc" title="Pagina">
	<iframe style="width:100%;height:100%" id="ifPagina"></iframe>
</div>
<div id="CRTMasivo" title="Atención">
<p><label id="txtMsn"></label></p>
</div>
</body>
</html>
