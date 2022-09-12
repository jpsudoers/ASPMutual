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
<link type="text/css" rel="stylesheet" media="all" href="css/chat.css" />
<script src="js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery.validate.js" type="text/javascript"></script>
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
<script src="js/jquery.jqGrid.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
   $(document).ready(function(){					
	jQuery("#list1").jqGrid({ 
		url:'facturapagos/listar.asp', 
		datatype: "xml", 
		colNames:['Rut', 'Empresa', 'Documento', 'N° Factura', 'Monto', '&nbsp;', '&nbsp;'], 
		colModel:[
				   {name:'rut',index:'E.RUT', width:95}, 
				   {name:'empresa',index:'E.R_SOCIAL', width:260}, 
				   {name:'documento',index:'documento', width:250, align:"right"},
				   {name:'factura',index:'FC.FACTURA', width:100, align:"right"}, 
				   {name:'valor',index:'FC.MONTO', width:90, align:"right"},
				   {align:"right",editable:true, width:20},
				   {align:"right",editable:true, width:20}], 
		rowNum:300, 
		autowidth: true, 
		rowList:[300,500,1000],  
		pager: jQuery('#pager1'), 
		sortname: 'FC.ID_FACTURA', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Facturas por Cobrar" 
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
	
	$("#dialog").dialog({
			bgiframe: true,
			autoOpen: false,
			height: 510,
			width: 925,
			modal: true,
			buttons: {
				'Guardar': function() {
						if($("#frmregpagos").valid())
						{
							$.post($('#frmregpagos').attr('action')+'?'+$('#frmregpagos').serialize(),function(d){
			   																				$("#list1").trigger("reloadGrid"); 
																						   });
							$("#list1").trigger("reloadGrid");
							$(this).dialog('close');
						}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	

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

	function validar(){
		$("#frmregpagos").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			txFolio:{
				required:true,
				number:true
			},
			txDoc:{
				required:true
			},
			txtFechIngreso:{
				required:true
			},
			FormaPago:{
				required:true
			},
			txtMonto:{
				required: true,
				max: $("#Saldo").val()
			}
		},
		messages:{
			txFolio:{
				required:"&bull; Ingrese Número de Folio Factura."
			},
			txDoc:{
				required:"&bull; Ingrese el Número del Documento."
			},
			txtFechIngreso:{
				required:"&bull; Ingrese Fecha de Pago Factura."
			},
			FormaPago:{
				required:"&bull; Seleccione Forma de Pago de la Factura."
			},
			txtMonto:{
				required:"&bull; Ingrese Monto de Pago Factura.",
				max: "&bull; Ingrese Monto Menor o Igual a "+formatCurrency($("#Saldo").val())
			}
		}
	});
    }
	
	function llena_formaPago(forma){
		$("#FormaPago").html("");
		$("#FormaPago").append("<option value=\"\">Seleccione</option>");
		if(forma==1){
			$("#FormaPago").append("<option value=\"1\" selected=\"selected\">Vale Vista</option>");
			$("#FormaPago").append("<option value=\"2\">Depósito Cheques</option>");
			$("#FormaPago").append("<option value=\"3\">Transferencia Electrónica</option>");
			$("#FormaPago").append("<option value=\"5\">Cheque a Fecha</option>");
		}else if(forma==2){
			$("#FormaPago").append("<option value=\"1\">Vale Vista</option>");
			$("#FormaPago").append("<option value=\"2\ selected=\"selected\">Depósito Cheques</option>");
			$("#FormaPago").append("<option value=\"3\">Transferencia Electrónica</option>");
			$("#FormaPago").append("<option value=\"5\">Cheque a Fecha</option>");
		}else if(forma==3){
			$("#FormaPago").append("<option value=\"1\">Vale Vista</option>");
			$("#FormaPago").append("<option value=\"2\">Depósito Cheques</option>");
			$("#FormaPago").append("<option value=\"3\ selected=\"selected\">Transferencia Electrónica</option>");
			$("#FormaPago").append("<option value=\"5\">Cheque a Fecha</option>");
		}else if(forma==5){
			$("#FormaPago").append("<option value=\"1\">Vale Vista</option>");
			$("#FormaPago").append("<option value=\"2\">Depósito Cheques</option>");
			$("#FormaPago").append("<option value=\"3\">Transferencia Electrónica</option>");
			$("#FormaPago").append("<option value=\"5\ selected=\"selected\">Cheque a Fecha</option>");
		}else{
			$("#FormaPago").append("<option value=\"1\">Vale Vista</option>");
			$("#FormaPago").append("<option value=\"2\">Depósito Cheques</option>");
			$("#FormaPago").append("<option value=\"3\">Transferencia Electrónica</option>");
			$("#FormaPago").append("<option value=\"5\">Cheque a Fecha</option>");
		}
		calculaSaldo();
	}

	function val(e) {
    tecla = (document.all) ? e.keyCode : e.which; // 2
    if (tecla==8) {return true;} // 3
    patron =/[A-Za-z\s]/; // 4
    te = String.fromCharCode(tecla); // 5
    return patron.test(te); // 6
	}
	
	var nav4 = window.Event ? true : false;
	function acceptNum(evt){	
	var key = nav4 ? evt.which : evt.keyCode;	
	return (key <= 13 || (key >= 48 && key <= 57));
	}
	
	function Decimales(Numero, Decimales) {
		pot = Math.pow(10,Decimales);
		num = parseInt(Numero * pot) / pot;
		nume = num.toString().split('.');
	
		entero = nume[0];
		decima = nume[1];
	
		if (decima != undefined) {
			fin = Decimales-decima.length; }
		else {
			decima = '';
			fin = Decimales; }
	
		for(i=0;i<fin;i++)
		  decima+=String.fromCharCode(48); 
	
		num=entero+'.'+decima;
		return num;
	}
	
	function update(i){
		$.post("facturapagos/frmregpagos.asp",{id:i},
			   function(f){
				   $('#dialog').html(f);
				   $('#txtFechIngreso').datepicker({firstDay: 1,dateFormat: 'dd-mm-yy' });
				   llena_formaPago(0);
				   grilla($('#idFact').val());
				   validar();
		});
		 $('#dialog').dialog('open');
	}
	
	function calculaSaldo()
	{
		var cancelado=0;
		
		if($('#txtMonto').val()!='')
		{
			cancelado=$('#txtMonto').val();
		}
		
		if(($('#Saldo').val()-cancelado)>=0)
		{
		    $('#lbSaldo').html(formatCurrency($('#Saldo').val()-cancelado));
		}
		else
		{
			$('#lbSaldo').html(formatCurrency(0));
		}
	}
	
	function grilla(id)
	{
		jQuery("#list2").jqGrid({
		scroll:1, 								
		url:'facturapagos/listarPagos.asp?id_factura='+id, 
		datatype: "xml", 
		colNames:['Fecha de Pago','Forma de Pago', 'N° de Documento','Monto Cancelado'], 
		colModel:[
				   {name:'fecha',index:'fecha', width:50, align:"center"},
				   {name:'formapago',index:'formapago', width:70},
				   {name:'comprobante',index:'comprobante', width:50, align:"right"},
				   {name:'monto',index:'monto', width:50, align:"right"}], 
		rowNum:100, 
		rownumbers: true, 
		rownumWidth: 40,  
		autowidth: true, 
		pager: jQuery('#pager2'), 
		sortname: 'id_ingresos', 
		viewrecords: true, 
		sortorder: "asc", 
		caption:"Listado de Pagos" 
		}); 
	
	jQuery("#list2").jqGrid('navGrid','#pager2',{edit:false,add:false,del:false,search:false,refresh:false});
	jQuery("#list2").jqGrid('navButtonAdd',"#pager2",{caption:"",
													  title:"Refrescar", 
													  buttonicon :'ui-icon-refresh', 
													  onClickButton:function(){ 
 														 $("#list2").trigger("reloadGrid");
													  } }); 
																  
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
	
	function createChatBox(chatboxtitle,minimizeChatBox) {
		$(" <div />" ).attr("id","chatbox_"+chatboxtitle)
		.addClass("chatbox")
		.html('<div class="chatboxhead" onclick="javascript:toggleChatBoxGrowth(\''+chatboxtitle+'\')"><div class="chatboxtitle">'+chatboxtitle+'</div><br clear="all"/></div><div class="chatboxcontent"><OL id="selectable"><LI value="wer" class="ui-widget-content ui-tabs ui-widget  ui-corner-all">Item 1</LI><LI value=\"\" class="ui-widget-content ui-tabs ui-widget  ui-corner-all">Item 2</LI></OL></div>')
		.appendTo($( "body" ));
				   
		$("#chatbox_"+chatboxtitle).css('bottom', '0px');
	
		$("#chatbox_"+chatboxtitle).css('right', '20px');
	
		$("#chatbox_"+chatboxtitle).show();
    }

    function toggleChatBoxGrowth(chatboxtitle) {
		if ($('#chatbox_'+chatboxtitle+' .chatboxcontent').css('display') == 'none') {  
			
			var minimizedChatBoxes = new Array();
			
			if ($.cookie('chatbox_minimized')) {
				minimizedChatBoxes = $.cookie('chatbox_minimized').split(/\|/);
			}
	
			var newCookie = '';
	
			for (i=0;i<minimizedChatBoxes.length;i++) {
				if (minimizedChatBoxes[i] != chatboxtitle) {
					newCookie += chatboxtitle+'|';
				}
			}
	
			newCookie = newCookie.slice(0, -1)
	
	
			$.cookie('chatbox_minimized', newCookie);
			$('#chatbox_'+chatboxtitle+' .chatboxcontent').css('display','block');
			$("#chatbox_"+chatboxtitle+" .chatboxcontent").scrollTop($("#chatbox_"+chatboxtitle+" .chatboxcontent")[0].scrollHeight);
		} else {
			
			var newCookie = chatboxtitle;
	
			if ($.cookie('chatbox_minimized')) {
				newCookie += '|'+$.cookie('chatbox_minimized');
			}
	
			$.cookie('chatbox_minimized',newCookie);
			$('#chatbox_'+chatboxtitle+' .chatboxcontent').css('display','none');
		}
  }

jQuery.cookie = function(name, value, options) {
    if (typeof value != 'undefined') { // name and value given, set cookie
        options = options || {};
        if (value === null) {
            value = '';
            options.expires = -1;
        }
        var expires = '';
        if (options.expires && (typeof options.expires == 'number' || options.expires.toUTCString)) {
            var date;
            if (typeof options.expires == 'number') {
                date = new Date();
                date.setTime(date.getTime() + (options.expires * 24 * 60 * 60 * 1000));
            } else {
                date = options.expires;
            }
            expires = '; expires=' + date.toUTCString(); // use expires attribute, max-age is not supported by IE
        }
        // CAUTION: Needed to parenthesize options.path and options.domain
        // in the following expressions, otherwise they evaluate to undefined
        // in the packed version for some reason...
        var path = options.path ? '; path=' + (options.path) : '';
        var domain = options.domain ? '; domain=' + (options.domain) : '';
        var secure = options.secure ? '; secure' : '';
        document.cookie = [name, '=', encodeURIComponent(value), expires, path, domain, secure].join('');
    } else { // only name given, get cookie
        var cookieValue = null;
        if (document.cookie && document.cookie != '') {
            var cookies = document.cookie.split(';');
            for (var i = 0; i < cookies.length; i++) {
                var cookie = jQuery.trim(cookies[i]);
                // Does this cookie string begin with the name we want?
                if (cookie.substring(0, name.length + 1) == (name + '=')) {
                    cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                    break;
                }
            }
        }
        return cookieValue;
    }
};
</script>
<script type="text/javascript">
$(function() {
		$("#selectable").selectable({
			stop: function(){
				//var result = $("#select-result").empty();
				$(".ui-selected", this).each(function(){
					//var index = $("#selectable li").index(this);
					//result.append(" #" + (index + 1));
					alert($("#selectable li").value);
				});
			}
		});
	});
</script>
<STYLE type="text/css">
	#feedback { font-size: 1em; }
	#selectable .ui-selecting { background: #0077A7 url(../images/img3.gif) repeat-x left bottom; }
	#selectable .ui-selected { background: #31576F; color: white; }
	#selectable { list-style-type: none; margin: 0; padding: 0; width: 100%; }
	#selectable li { margin: 3px; padding: 0.4em; font-size: 1em; height: 18px; }
</STYLE>
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
			<h2><em style="text-transform: capitalize;">Registro Pagos</em></h2>
            <table id="list1"></table> 
            <div id="pager1"></div> 
	  </div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Registro Pago Factura">
</div>
<div id="pantContrasena" title="Cambiar Contraseña">
</div>
<div id="mContrasena" title="Cambiar Contraseña">
     <label id="txtmContrasena" name="txtmContrasena"></label>
</div>
</body>
</html>