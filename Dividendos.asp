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
		CargaFiltro("cboAnio",'Filtros.asp?rep=anio_Sup&sup=S','CargarPeriodo');
	}

	function CargarPeriodo()
	{	
		var Anio = document.getElementById("cboAnio").value;
		CargaFiltro("cboTrim",'Filtros.asp?rep=periodo_SupAnio&sup=S&anio='+Anio,'cargaLetra');
	}

	function cargaLetra(){
		var Anio = document.getElementById("cboAnio").value;
		var Periodo = document.getElementById("cboTrim").value;
		CargaFiltro("cboLetra",'Filtros.asp?rep=letra_SupAnioPer&sup=S&anio='+Anio+'&per='+Periodo+'&eeff=04','cargaCodigos');
	}

	function cargaCodigos(){
		var nivel = document.getElementById("cboNivel").value;
		if(nivel==0){
			var c = document.getElementById("cboCodigo");
			c.innerHTML = "";
			var option = document.createElement("option");
			option.text = "--";			
			option.value = "0";
			c.add(option);
		}
		else{
			var Anio = document.getElementById("cboAnio").value;
			var Periodo = document.getElementById("cboTrim").value;
			var letra = document.getElementById("cboLetra").value;

			CargaFiltro("cboCodigo",'Filtros.asp?rep=codigo_SupAnioPerLetrNiv&sup=S&anio='+Anio+'&per='+Periodo+'&letr='+letra+'&niv='+nivel+'&eeff=04','');
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

TABLE.tabla1
{
    BORDER-RIGHT: #75BDF7 1px solid;
    BORDER-TOP: #75BDF7 1px solid;
    BORDER-LEFT: #75BDF7 1px solid;
    BORDER-BOTTOM: #75BDF7 1px solid;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;
	width:100%;
	background:#FFFFFF;
}

TABLE.tabla1 TD
{
    BORDER-RIGHT: #75BDF7 1px solid;
    BORDER-TOP: #75BDF7 1px solid;
    BORDER-LEFT: #75BDF7 1px solid;
    BORDER-BOTTOM: #75BDF7 1px solid;
	
}
TABLE.tabla1 TH
{
    BORDER-RIGHT: #75BDF7 1px solid;
    PADDING-RIGHT: 5px;
    BORDER-TOP: #75BDF7 1px solid;
    PADDING-LEFT: 5px;
	background:#D9E6F4;
    PADDING-BOTTOM: 5px;
    BORDER-LEFT: #75BDF7 1px solid;
    PADDING-TOP: 5px;
    BORDER-BOTTOM: #75BDF7 1px solid;
    HEIGHT: 20px;
	color:#000000;
	font-family:Arial, Helvetica, sans-serif;
	font-size:12px;
	
}

#blocker {
            Z-INDEX: 2000;
            BACKGROUND:#000000; /*: #000; */
            FILTER: alpha(opacity=30);
            LEFT: 0px;
            WIDTH: 100%;
            POSITION: absolute;
            TOP: 0px;
            HEIGHT: 100%;
            opacity: 0.2;
            moz-opacity: 0;
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
	color: #ffffff;
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
<font face="Arial" size='3pt' color='#FFFFFF'>Dividendos</font></strong>
<form action="Dividendos.asp" name="fmrEF" method="post" target=_self>

 	<a class="a2">&nbsp;&nbsp;Periodo</a>
 	<select name="cboAnio" id="cboAnio" class="combo2" onchange="CargarPeriodo();"></select>
	<select name="cboTrim" id="cboTrim" class="combo2" onchange="cargaLetra();"></select>
	
	<A id="letra" class=a2>&nbsp;&nbsp;Categoria</A>
	<select name="cboLetra" id="cboLetra" style="width:35px" class="combo2" onchange="cargaCodigos();">
	</select>

	<A class=a2>&nbsp;&nbsp;Nivel</A>
	<select name="cboNivel" id="cboNivel" class="combo2" onchange="cargaCodigos();">
        <option value="5" selected="selected">Nv AE 101</option>
        <option value="10" >Nv AE 54</option>
        <option value="11" >Nv AE 14</option>
        <option value="6" >Nv SI</option>
        <option value="12" >Tipo Entidad</option>
        <option value="0" >TODOS</option>
    </select>

	<select name="cboCodigo" id="cboCodigo" class="combo2" style="width:320px">
    </select>

    <A id="moneda" class=a2>&nbsp;&nbsp;Moneda</A>
	<select name="cboMoneda" id="cboMoneda" class="combo2">
		<option value='Soles'>Soles</option>
		<option value='Dolares'>Dolares</option>
		<option value='Total'>Conversión en soles</option>
	</select>
	<a class="a2">&nbsp;&nbsp;</a>
	<button onClick="cargaDividendos('S'); return false;" style="border:none;height:21px; width:21px;background: url(imagenes/search.png) no-repeat;" alt="Buscar Consulta"></button>
	<button onClick="ExcelDividendos('S'); return false;" style="border:none;height:21px; width:21px;background: url(imagenes/excel.png) no-repeat;" alt="Exportar a Excel"></button>
<br>

<!--<div id="DivVariables" style="overflow:auto;height='400 px'; width='100%'"></div>-->
<div id="DivVariables" style="overflow:scroll;height:420px; width:100%"></div>
  
</FORM>
</body>
</html>

