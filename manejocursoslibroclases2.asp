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
$(document).ready(function(){
				$("#codempresa").append("<option value=\"\">Seleccione</option>");
				$("#codTrabajador").append("<option value=\"\">Seleccione</option>");
				llena_curriculo(0);		   
   
	});	

	function llena_curriculo(id_curriculo){
			$("#codprog").html("");
			$.get("autorizacion/curriculo.asp",
						function(xml){
							$("#codprog").append("<option value=\"\">Seleccione</option>");
							$('row',xml).each(function(i) { 
								if(id_curriculo==$(this).find('ID_MUTUAL').text())
									$("#codprog").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" selected="selected">'+
																			$(this).find('NOMBRES').text()+ '</option>');
								else
									$("#codprog").append('<option value="'+$(this).find('ID_MUTUAL').text()+'" >'+
																			$(this).find('NOMBRES').text()+ '</option>');
							});
						});
	}	   


	function VerInforme()
	{
		validar();
		if($("#frminforme").valid())
		{
		window.open("libroclases/pdf2.asp?prog="+$("#codprog").val(),'Informe')
		}
	}
	
	function validar(){
		$("#frminforme").validate({
		errorContainer: "#messageBox1",
  		errorLabelContainer: "#messageBox1 ul",
		wrapper: "li", 
		debug:true,
		rules:{
			codprog:{
				required:true
			}
		},
		messages:{
			codprog:{
				required:"&bull; Seleccione Código de Programación."
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
        <%end if%>
    <%DATOS.MoveNext
   WEND%>
    <%else%>
		<li><a href="empresas.asp" accesskey="5" title="">Empresas</a></li>
    <%end if%>
	</ul>
</div>
<div id="content">
	<div id="colOne">
  	<h3>Login</h3>
		<div class="bg1">
			Usuario:<strong>Nombre de Usuario</strong>
      <i>Perfil de usuario</i><br />
      <%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%>
      <button OnClick="document.location.href='index.asp';">Cerrar Sesión</button>
		</div>
		<h3>Opciones</h3>
		<div class="bg1">
			<ul>
				<li class="first"><a href="manejocursoslibroclases.asp">Libro de Clases</a></li>
                <li><a href="manejoeva_cierre.asp">Evaluación y Cierre</a></li>
                <li><a href="manejocursoscertificados.asp">Certificados</a></li>
			</ul>
		</div>		
	</div>
	<div id="colTwo">
		<div class="bg2">
        <h2><em style="text-transform: capitalize;">Libro de Clases</em></h2>
        <br />
       <div>
						<table cellpadding="0" cellspacing="0" style="border:0px;width:100%;" class="cabezaMsj">
							<tr class="buscarTitulo">
								<td><h3><em style="text-transform: capitalize;">Informes</em></h3></td>
							</tr>
						</table>
					</div>
			<div id="buscarContent" style="border:1px solid #000099;display:block" >
             <form name="frminforme" id="frminforme" action="" method="post">
					<table width="620" border="0" align="center" cellpadding="1" cellspacing="1">
                   <tr>
    						<td width="150">&nbsp;</td>
                            <td width="200">&nbsp;</td>
                            <td width="50">&nbsp;</td>
                         
                       
   						</tr>
						<tr>
                            <td>Código de Programación :</td>
                            <td colspan="2"><select name="codprog" id="codprog" class="formulario"></select></td>
                           
                            <td rowspan="1"><input type="button" class="boton" value="Ver Informe" onclick="VerInforme();" id="btnBuscar" name="btnBuscar"/></td>
						</tr>
                        <tr>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                           
   					   </tr>
					</table>
				</form> 
                    
				</div>	
                <div id="messageBox1" style="height:100px;overflow:auto;width:300px;"> 
                   	 <ul></ul> 
               		</div>		
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
</body>
</html>
