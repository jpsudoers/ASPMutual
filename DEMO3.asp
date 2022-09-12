<html>
<head>
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
<script src="js/ajaxfileupload.js" type="text/javascript" charset="utf-8"></script>
<script type="text/javascript">
$(document).ready(function(){					
	function subeDoc()
    {
                               $("#loading")
                               .ajaxStart(function(){
                                               $("#loading").html('<img src="../images/loader.gif"/>');
                                               $(this).show();
                               })
                               .ajaxComplete(function(){
                                               $(this).hide();
                               });
							   
                               $.ajaxFileUpload
                               (
                                      {
                                           url:'emp_programacion/subir.asp?',
                                           secureuri:false,
                                           fileElementId:'txtDoc',
                                           elements:'Empresa;programa;EmpMan;Compromiso;txtNum;datostabla;NParticipantes;Valor;ValorTotal;TipoEmpresa;Costo;datosfran',
										   dataType: 'json',
                                           success: function (data, status)
                                           {
											   if(typeof(data.error) != 'undefined')
											   {
												  if(data.error != '')
												  {
													 alert("Problemas "+data.error);
												  }
												  else      
												  {
													//$("#dialog").dialog('close');
													$("#list1").trigger("reloadGrid");
												  }
											   }
                                           },
                                           error: function (data, status, e)
                                           {
                                                 $("#list1").trigger("reloadGrid");
                                           }
                                         }
                               );
                               return false;
    }
	
	$("#dialog").dialog({
			autoOpen: false,
			bgiframe: true,
			height:600,
			width: 1000,
			modal: true,
			overlay: {
				backgroundColor: '#f00',
				opacity: 0.5
			},
			buttons: {
				'Guardar': function() {
					if($("#terminos").attr("checked")) 
					{
						if($("#frmProgEmp").valid())
						{
							//$(this).dialog('close');
							//$("#mensaje").dialog('open');
							subeDoc();
						}
					}
				},
				Cancelar: function() {
					$(this).dialog('close');
				}
			}
		});
	});	
	</script>
</head>
<body>

<form action="funciones/sube.asp" method="post" enctype="multipart/form-data">
<input type="file" name="fichero"><br />
<input type="submit" value="Enviar">
</form>
<div id="dialog" title="Registro de Inscripción">
</div>
</body>
</html>

