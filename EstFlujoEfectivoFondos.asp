<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Conexion.asp"-->
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="description" content="Free Web tutorials" />
<meta name="keywords" content="HTML,CSS,XML,JavaScript" />
<meta name="author" content="Hege Refsnes" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link href="css/inei.css" type="text/css" rel="stylesheet">

<script type="text/javascript" src="java_script/stmenu.js"></script>
<script type="text/javascript" src="java_script/linq.js"></script>
<script type="text/javascript" src="java_script/jquery-3.1.1.min.js"></script>
<script type="text/javascript" language="JavaScript1.2" src="java_script/funciones.js"></script>

<% Response.expires = 0 
	if Session("tipoAcceso")="" then 
		'Response.redirect "login.html"
	end if
	ruta="imagenes"
	''Response.Write Session("id_usuario")
%>

<script type="text/javascript">

	$(document).ready(function() {
	  $.ajaxSetup({ cache: false });
	});

	function InicializaFiltros()
	{	CargarAnio();	}

	function CargarAnio()
	{
		CargaFiltro("cboAnio",'Filtros.asp?rep=FONDOS_anio','CargarPeriodo');
	}

	function CargarPeriodo()
	{	
		var Anio = document.getElementById("cboAnio").value;
		CargaFiltro("cboTrim",'Filtros.asp?rep=FONDOS_periodo_Anio&anio='+Anio,'');
	}

	function CargarPeriodo()
	{	
		var Anio = document.getElementById("cboAnio").value;
		CargaFiltro("cboTrim",'Filtros.asp?rep=FONDOS_periodo_Anio&anio='+Anio,'');
	}

	function switchMetodo(){
		var Fondo = document.getElementById("cboFondo").value;
		var x = document.getElementById("cboMetodo");

		if(Fondo=="INV"){
			x.style.visibility = 'visible'
		}
		else{
			x.style.visibility = 'hidden'
		}
	}

</script>

<title>.:INEI-DNCN - Sistema de Consultas SMV </title>
<style type="text/css">
body {
	margin-left: 0px;
	margin-right: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
	background-image: url(Imagenes/fdopag.jpg);

}

a.a1 {
	font-family: Verdana, Geneva, sans-serif;
	color:#000000;
	font-weight:bold;
	font-size: 8pt; 	
	}
a.a2 {
	font-family: Verdana, Geneva, sans-serif;
	font-size:8pt; 
	font-weight:bold;
	color: #000000;
	}
a.a3 {
	font-family: Verdana, Geneva, sans-serif;
	font-size:8pt; 
	color: #008BC0;
	}
.combo{
	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
	}

.combo1 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
.combo2 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}

.combo3 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
.combo4 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
.combo5 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
</style>
<body onload="InicializaFiltros();">
<div id="blocker" name="blocker" style="display:none;" ><table width="100%" height="100%" border="1" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><center><img src="Imagenes/progressbar.gif"></img></center></td>
  </tr>
</table>
</div>
<table width="100%" height="72" border="0" cellpadding="0" cellspacing="0" background="<%=ruta%>/CONASEV_fondo.jpg">
  <tr>
  <td>
  	<table width="200" height="72" background="<%=ruta%>/CONASEV_izq.png">
    	<tr><td></td>
    	</tr>
    </table>  </td>
  <td align="right">
  	<table width="210" height="72" background="<%=ruta%>/CONASEV_drch.png">
    	<tr><td></td></tr>
    </table>  </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="10"><strong>
      <script type="text/javascript" language="JavaScript1.2" src="java_script/menu.js"></script>
    </strong></td>
  </tr>          
</table>

<strong>
<br>
<font face="Arial" size='3pt' color='#FFFFFF'>Estado Flujo Efectivo Fondos</font></strong>
<form action="EstFlujoEfectivoFondos.asp" name="fmrEF" method="post" target=_self>

 	<a class="a2">&nbsp;&nbsp;Periodo</a>
 	<select name="cboAnio" id="cboAnio" class="combo2" onchange="CargarPeriodo();"></select>
	<select name="cboTrim" id="cboTrim" class="combo2" onchange="cargaLetra();"></select>

	<A id="detalle" class=a2>&nbsp;&nbsp;Detalle</A>
	<select name="cboDetalle" id="cboDetalle" class="combo2">
		<option value='0'>Por Fondo</option>
		<option value='1'>Por Soc. Adm.</option>
	</select>
	
    <A id="moneda" class=a2>&nbsp;&nbsp;Moneda</A>
	<select name="cboMoneda" id="cboMoneda" class="combo2">
		<option value='Soles'>Soles</option>
		<option value='Dolares'>Dolares</option>
		<option value='Total'>Conversión en soles</option>
	</select>

	<A id="TipFondo" class=a2>&nbsp;&nbsp;Tipo</A>
	<select name="cboFondo" id="cboFondo" class="combo2" onchange="switchMetodo();">
		<!--<option value='MUT'>Mutuos</option>-->
		<option value='INV'>Inversión</option>
		<!--<option value='PAT'>Patrimonio</option>-->
	</select>

    <A id="metodo" class=a2>&nbsp;&nbsp;Método</A>
	<select name="cboMetodo" id="cboMetodo" class="combo2">
		<option value='Directo'>Directo</option>
		<option value='Indirecto'>Indirecto</option>
	</select>

	<a class="a2">&nbsp;&nbsp;</a>
	<button onClick="cargaVariableEF_FONDOS('FE'); return false;" style="border:none;height:21px; width:21px;background: url(imagenes/search.png) no-repeat;" alt="Buscar Consulta"></button>
	<button onClick="ExcelEF_FONDOS('FE'); return false;" style="border:none;height:21px; width:21px;background: url(imagenes/excel.png) no-repeat;" alt="Exportar a Excel"></button>
<br>

<!--<div id="DivVariables" style="overflow:auto;height='400 px'; width='100%'"></div>-->
<div id="DivVariables" style="overflow:scroll;height:420px; width:100%"></div>
  
</FORM>
</body>
</html>

