<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title></title>
<meta name="Keywords" content="" />
<meta name="Description" content="" />
<link href="../css/default.css" rel="stylesheet" type="text/css" />
<link href="../css/jquery-ui.css" rel="stylesheet" type="text/css" />
<link href="../css/ui.all.css" rel="stylesheet" type="text/css"/>
<link href="../css/ui.jqgrid.css" rel="stylesheet" type="text/css"/>
<script src="../js/jquery.js" type="text/javascript" charset="utf-8"></script>
<script src="../js/jquery.validate.js" type="text/javascript"></script>
<script src="../js/jquery.defaultvalue.js" type="text/javascript" charset="utf-8"></script>
<script src="../js/jquery-ui.js" type="text/javascript" charset="utf-8"></script>
<script src="../js/ajaxfileupload.js" type="text/javascript" charset="utf-8"></script>
<script src="../js/app/solicitudCredito.js" type="text/javascript" charset="utf-8"></script>
<style text="css/default">
#drop_file_zone {
    background-color: #EEE;
    border: #999 5px dashed;
    width: 290px;
    height: 200px;
    padding: 8px;
    font-size: 18px;
}
#drag_upload_file {
  width:50%;
  margin:0 auto;
}
#drag_upload_file p {
  text-align: center;
}
#drag_upload_file #selectfile {
  display: none;
}

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
		height:150px;
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

    .button {
        font-size: 20px; 
        padding: 5px;
    }

    .fto_adjunta_table
    {
      float: left; 
      border:0px; 
      width:100px; 
      font-size: 11px; 
      font-family: arial;
      text-align: center;
    }

    .fto_adjunta {
      float: left;
      width: 100px;
      height: 100px;
      background-image: url(../images/img_ejemplo.jpg);
      margin: 13px;
    }

    .eliminar_img_doc {
      background-color: #1f2e5e;
      width: 20px;
      height: 20px;
      float: right;
      padding: 7px;
      padding-right: 7px;
      padding-bottom: 7px;
      padding-right: 7px;
      padding-bottom: 7px;
      padding-right: 10px;
      padding-bottom: 10px;
      position: absolute;
    }

</style>
</head>
<body>
<div id="header">
	<h1><img src="../images/logo.png" /></h1>
	<ul>
	</ul>
</div>
<div id="content" class="bg1">
  <h2><em style="text-transform: capitalize;">Solicitud de Crédito</em></h2>
<form name="frmSolicitud" id="frmSolicitud" action="solicitudcreditoAjax.asp" method="post">
	<table cellspacing="0" cellpadding="1" border=0>
    <tr>
    	<td width="105">&nbsp;</td>
      	<td width="200">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="80">&nbsp;</td>
        <td width="200">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="60">&nbsp;</td>
        <td width="200">&nbsp;</td> 
   	</tr>
     <tr>
    	<td colspan="8"><center><h3><em style="text-transform: capitalize;">Datos de la Solicitud</em></h3></center></td>
   	</tr>
     <tr>
    	<td colspan="8">&nbsp;<input type="hidden" id="solId" name="solId" /> <input type="hidden" id="operacion" name="operacion" /></td>
   	</tr>
    <tr>
    	<td colspan="1">Rut Empresa:</td>
      	<td colspan="7"><input id="txRut" name="txRut" type="text" tabindex="1" maxlength="12" size="12" onkeyup="lookup(this.value);"/>
        <div class="suggestionsBox" id="suggestions" style="display: none;position:absolute;z-index:1;left:36%">
              <img src="../images/upArrow.png" style="position: relative; top: -12px; left: 20px;" alt="upArrow" />
              <div class="suggestionList" id="autoSuggestionsList">
                &nbsp;
              </div>
        </div>
        
        </td>
	<tr>
        <td colspan="1">Nombre Empresa:</td>
        <td colspan="7"><input id="txtNombreEmpresa" name="txtNombreEmpresa" type="text" tabindex="2" maxlength="99" size="40" /></td>
    </tr>
     <tr>
    	<td colspan="8">&nbsp;</td>
   	</tr>
    <tr>
    	<td colspan="8"><center>
    	  <h3><em style="text-transform: capitalize;">Documentos</em></h3></center></td>
   	</tr>
	<tr>
        <td colspan="8">&nbsp;</td>
    </tr>
    <tr>
    	<td colspan="4">
		<div id="drop_file_zone" ondrop="upload_file(event,true);" ondragover="return false">
		<div id="drag_upload_file">
			<p>Arrastra archivos</p>
			<p>o</p>
			<p><input type="button" value="Selecciona Archivos" onclick="file_explorer(event);" /></p>
			<!--<input type="file" id="selectfile" name="selectfile" multiple />-->
			<input type="file" name="selectfile" id="selectfile" tabindex="5" maxlength="400" size="20">
		</div>
		</div>
		</td>
    <td colspan="4">
      <div id="documentosContent" id="documentosContent" ></div>
    </td>
   	</tr>
    <tr>
        <td colspan="8">&nbsp;</td>
    </tr>
     <tr>
        <td colspan="8">&nbsp;</td>
    </tr>
	<tr>
        <td colspan="8">&nbsp;</td>
    </tr>
	<tr>
    	<td colspan="4">&nbsp;</td>
        <td align="right"><INPUT TYPE="button" OnClick="window.close();" name="btn_volver" id="btn_volver" class="ui-state-default ui-corner-all button" VALUE="Cerrar" /></td>
        <td>&nbsp;</td>
        <td><INPUT TYPE="button" OnClick="guardar();" id="enviar" name="enviar" class="ui-state-default ui-corner-all button" VALUE="Enviar" /></td>
        <td>&nbsp;</td> 
   	</tr>
</table>
</form>  
<div id="messageBox1" style="height:100px;overflow:auto;width:400px;"> 
  	<ul></ul> 
</div>   
</div>
<div id="footer">
	<p><em style="text-transform: capitalize;">Copyright © 2010 Mutual de Seguridad C.CH.C. Reservados todos los derechos</em></p>
</div>
<div id="dialog" title="Solicitud de Crédito Empresa">
	<p>La Solicitud de Crédito de Empresa fue Ingresada Exitosamente. </p>
</div>
<div id="dialog_espera" title="Solicitud de Inscripción Empresa">
<center>
<table width="200" border="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><center><font size="2"><b>Espere ...</b></font></center></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><center><img id="loading" src="../images/loader.gif"/></center></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</center>
</div>
<div id="rut1" title="Solicitud de Crédito Empresa">
	<p>El Rut de Empresa Ingresado, ya se Encuentra Registrado.</p>
</div>
<div id="datosTerminos" title="Terminos y  Condiciones">
</div>
</body>
</html>
