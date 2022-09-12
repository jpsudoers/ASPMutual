<!DOCTYPE html>
<html>
<head>
<script src="js/jquery.tooltip.js"></script> 
<script>
$(function() {	
		$("#myform :input").tooltip({
			// place tooltip on the right edge
			position: "center right",
		
			// a little tweaking of the position
			offset: [-2, 10],
		
			// use the built-in fadeIn/fadeOut effect
			effect: "fade",
		
			// custom opacity setting
			opacity: 0.7
		});
		});
</script>
<style>	
	.tooltip {
	background-color:#000;
	border:1px solid #fff;
	padding:10px 15px;
	width:200px;
	display:none;
	color:#fff;
	text-align:left;
	font-size:12px;

	/* outline radius for mozilla/firefox only */
	-moz-box-shadow:0 0 10px #000;
	-webkit-box-shadow:0 0 10px #000;
	}
</style>
</head>
<body>
<form id="myform" action="#">
		<label for="username">Rut : </label>
		<input id="username" title="Formato del rut 16313889-1"/>
</form>
</body>
</html>
