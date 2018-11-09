<!--#include file="Conexion.asp"-->
<html>
<head>
<meta name="description" content="Free Web tutorials" />
<meta name="keywords" content="HTML,CSS,XML,JavaScript" />
<meta name="author" content="Hege Refsnes" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<script type="text/javascript" src="java_script/stmenu.js"></script>
<% Response.Expires= 0 
	if Session("tipoAcceso")="" then 
		'Response.Redirect("login.html")
	End if
	ruta="imagenes"	
%>
<title>.:INEI-DNCN - Sistema de Consultas SMV</title>
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
<body >
<div id="blocker" name="blocker" style="display:none;" >
<table width="100%" height="100%" border="1" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><center><img src="Imagenes/progressbar.gif"></img></center></td>
  </tr>
</table>
</div>
<table width="100%" height="72" border="0" cellpadding="0" cellspacing="0" background="<%=ruta%>/CONASEV_fondo.jpg">
  <tr>
    <td>
  	  <table width="200" height="72" background="<%=ruta%>/CONASEV_izq.png" >
    	<tr>
		  <td></td>
		</tr>
      </table>
    </td>
    <td align="right">
  	  <table width="210" height="72" background="<%=ruta%>/CONASEV_drch.png">
        <tr>
		  <td><!--Usuario: <%Response.write Session("id_usuario") %>--></td>
		</tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="10">
	  <script type="text/javascript" language="JavaScript1.2" src="java_script/menu.js"></script>
	  <a href="logoffce.asp" style="font-family:verdana; font-size:10px; color:#000000"></a>
	</td>
  </tr>          
</table>

<div align="center">
  <script type="text/javascript" language="JavaScript1.2" src="java_script/funciones.js"></script>
  <br />
  <strong><font face="Arial" size='3pt' color='FFFFFF'>DIRECTORIO</font></strong></div>
<form action="Directorio.asp" name="fmrEF" method="post" target=_self>
<%		strUbicaTipo=Request("hidTipo")
		strUbicaAnio = Request("hidAnio")
		strUbicaTrim = Request("hidTrim")
%>

  <div name="formulario" id="formulario1">
    <a class="a1">Seleccione:</a> 
	<a class="a2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Año</a>
	<select name="cboAnio" id="cboAnio" class="combo2">
	  <option value="" selected><%Response.Write "[Seleccione]"%></option>
		<%  SQL = "sp_lista_anio"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open SQL , con
			If Not rs.BOF Then rs.MoveFirst
			Do While Not rs.EOF			
		%>
	  <option value="<%=Trim(rs("ANIO_INF"))%>"<%If Trim(CStr(strUbicaAnio)) = Trim(CStr(rs("ANIO_INF"))) Then%>selected<% End If%>><%=Trim(rs("ANIO_INF"))%> </option>
		<%  rs.MoveNext
			Loop
			rs.Close
	       	Set rs = Nothing
		   	SQL=""
		%>
    </select>
	<a class="a2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Periodo</a>
	<select name="cboTrim" id="cboTrim" class="combo2">
	  <option value="" selected><%Response.Write "[Seleccione]"%></option>
		<%  SQL = "sp_lista_Trim"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open SQL , con
			If Not rs.BOF Then rs.MoveFirst
			Do While Not rs.EOF			
		%>
	  <option value="<%=Trim(rs("PERINF_EMP"))%>"<%If Trim(CStr(strUbicaTrim)) = Trim(CStr(rs("PERINF_EMP"))) Then%>selected<% End If%>><%=Trim(rs("PERINF_EMP"))%> </option>
		<%  rs.MoveNext
			Loop
			rs.Close
	       	Set rs = Nothing
		   	SQL=""
		%>
    </select>
	  
	  <a class="a2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tipo: </a>
	  <select name="cboFondo" id="cboFondo" class="combo2" >
        <option value="MUT">Mutuos</option>
        <option value="INV">Inversión</option>
        <option value="PAT">Patrimonios en Fideicomiso</option>
      </select>

	<input type="hidden" name="cboAnio" id="cboAnio" value="0">
    <input type="hidden" name="cboTrim" id="cboTrim" value="0">
	<input type="hidden" name="cboOrden" id="cboOrden" value="0">

	
		&nbsp;&nbsp;
		<button onClick="cargaVariableFONDOS();return false;" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;">
			<img  src="Imagenes/search.png" width="20" height="20" alt="Buscar Consulta" >
		</button>&nbsp;&nbsp;
		<button onClick="ExcelFONDOS();return false;" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;">
			<img  src="imagenes/excel.png" width="20" height="20" alt="Exportar a Excel" >
		</button>&nbsp;&nbsp;
		<button onClick="Refresh(1)" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;">
			<img  src="imagenes/refresh.png" width="20" height="20" alt="Refrescar" >
		</button>
  </div>

<div id="DivVariables" style="overflow:auto;height='420 px'; width='100%'"></div>
  <input type="hidden" name="hidTipo" id="hidTipo" value="<%=strUbicaTipo%>">
  <input type="hidden" name="hidAnio" value="<%=strUbicaAnio%>">
  <input type="hidden" name="hidTrim" value="<%=strUbicaTrim%>">
</FORM>
</body>
</html> 