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
<script src="js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="js/i18n/grid.locale-sp.js" type="text/javascript"></script>
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
	
	sql = "select PERMISO1,PERMISO2,PERMISO3,PERMISO4,PERMISO5,PERMISO6,PERMISO12,PERMISO16 from USUARIOS "
	sql = sql&" where USUARIOS.ID_USUARIO='"&Session("usuarioMutual")&"'"

  DATOS.Open sql,oConn
  opcion=0
  pagina_inicio="finanzas.asp"
   WHILE NOT DATOS.EOF
		if(DATOS("PERMISO1")<>"0")then
			if(opcion=0)then
				pagina_inicio="administracion.asp"
				opcion=1
			end if
		%>
		<li><a href="administracion.asp" accesskey="1" title="">Administraci&oacute;n</a></li>
        <%end if
		if(DATOS("PERMISO2")<>"0")then
		    if(opcion=0)then
				pagina_inicio="operacion.asp"
				opcion=1
			end if
		%>
		<li><a href="operacion.asp" accesskey="2" title="">Operaci&oacute;n</a></li>
        <%end if
		if(DATOS("PERMISO3")<>"0")then
		    if(opcion=0)then
				pagina_inicio="manejocursos.asp"
				opcion=1
			end if
		%>
        <li><a href="manejocursos.asp" accesskey="3" title="">Manejo de Cursos</a></li>
        <%end if
		if(DATOS("PERMISO4")<>"0")then
		    if(opcion=0)then
				pagina_inicio="finanzas.asp"
				opcion=1
			end if
		%>
		<li><a href="finanzas.asp" accesskey="4" title="">Finanzas</a></li>
        <%end if
		if(DATOS("PERMISO6")<>"0")then
		    if(opcion=0)then
				pagina_inicio="consultas.asp"
				opcion=1
			end if
		%>
		<li><a href="consultas.asp" accesskey="5" title="">Consultas</a></li>        
        <%end if
		if(DATOS("PERMISO12")<>"0")then
		    if(opcion=0)then
				pagina_inicio="manuales.asp"
				opcion=1
			end if
		%>
		<li><a href="manuales.asp" accesskey="6" title="">Manuales</a></li>       
        <%end if	
		if(DATOS("PERMISO16")<>"0")then
		    if(opcion=0)then
				pagina_inicio="infocalendario.asp"
				opcion=1
			end if
		%>
		<li><a href="infocalendario.asp" accesskey="6" title="">Manuales</a></li>       
        <%end if			
	 DATOS.MoveNext
	 Response.Redirect(pagina_inicio)
   WEND
   end if
   if(Session("tipoUsuario")="0")then
   Response.Redirect("empresas.asp")
   %>
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
      <%=right("0"&day(now()),2)&"-"&right("0"&month(now()),2)&"-"&year(now)%>
      <button OnClick="document.location.href='index.asp';">Cerrar Sesión</button>
<table border="0">
  <tr>
    <td height="230">&nbsp;</td>
  </tr>
</table>

		</div>
	</div>
	<div id="colTwo">
    	<div class="bg2">
		</div>
	</div>
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
</body>
</html>
