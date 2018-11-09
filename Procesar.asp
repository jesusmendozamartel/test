<!--#include file="Conexion.asp"-->
<html>
<head>
<meta name="description" content="Free Web tutorials" />
<meta name="keywords" content="HTML,CSS,XML,JavaScript" />
<meta name="author" content="Hege Refsnes" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<script type="text/javascript" src="java_script/stmenu.js"></script>
<% Response.expires = 0 
	if Session("tipoAcceso")="" then 
		response.redirect "login.html"
	end if
	ruta="imagenes"
	'response.write Session("id_usuario")
%>
<title>.:INEI-DNCN - Sistema de Consultas del Gobierno </title>
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
    BORDER-RIGHT: #4786CB 1px solid;
    BORDER-TOP: #4786CB 1px solid;
    BORDER-LEFT: #4786CB 1px solid;
    BORDER-BOTTOM: #4786CB 1px solid;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;
	width:100%;
	background:#FFFFFF;
	 
	
}
TABLE.tabla1 TD
{
    BORDER-RIGHT: #4786CB 1px solid;
    BORDER-TOP: #4786CB 1px solid;
    BORDER-LEFT: #4786CB 1px solid;
    BORDER-BOTTOM: #4786CB 1px solid;
	
}
TABLE.tabla1 TH
{
    BORDER-RIGHT: #4786CB 1px solid;
    PADDING-RIGHT: 5px;
    BORDER-TOP: #4786CB 1px solid;
    PADDING-LEFT: 5px;
    BACKGROUND: url(Imagenes/bar.jpg);
    PADDING-BOTTOM: 5px;
    BORDER-LEFT: #4786CB 1px solid;
    PADDING-TOP: 5px;
    BORDER-BOTTOM: #4786CB 1px solid;
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

</style>

<body >

<div id="blocker" name="blocker" style="display:none;" ><table width="100%" height="100%" border="1" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><center><img src="Imagenes/progressbar.gif"></img></center></td>
  </tr>
</table>
</div>
<table width="100%" height="59" border="0" cellpadding="0" cellspacing="0" background="<%=ruta%>/cabecera_centro.gif">
  <tr>
  <td>
  	<table width="350" height="59" background="<%=ruta%>/cabecera_izquierda.gif">
    	<tr><td></td></tr>
    </table>  </td>
  <td align="right">
  	<table width="175" height="59" background="<%=ruta%>/cabecera_derecha.gif">
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
<script type="text/javascript" language="JavaScript1.2" src="java_script/funciones.js"></script>
<%tipo=request.QueryString("tipo")%>
<%response.Write("<font face='Arial' size='3pt' color='#FFFFFF'>Procesar "& tipo &"</font></strong>")%>
<form action="Procesar.asp" name="fmrSI" method="post" target=_self>
<%		
		strUbicaAnio=Request("hidAnio")
		boton=request("opcion")
		usuario=Session("id_usuario")
		tipo=request.QueryString("tipo")
		
		
%>
 	  
	  <a class="a2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Año</a>
 	   <select name="cboAnio" id="cboAnio" class="combo2">
          <option value="" selected><%Response.Write "[Seleccione]"%></option>
          <% 					 
		  		SQL = "select distinct ano_eje from Gobma_Directorio order by 1 desc"
		        Set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open SQL , con
	       		If Not rs.BOF Then rs.MoveFirst
       			Do While Not rs.EOF			
	 			%>
                <option value="<%=Trim(rs("ano_eje"))%>"<%If Trim(CStr(strUbicaAnio)) = Trim(CStr(rs("ano_eje"))) Then%>selected<% End If%>><%=Trim(rs("ano_eje"))%> </option>
                <%rs.MoveNext
	       		Loop
	       		rs.Close
	       		Set rs = Nothing
		   		SQL=""

		%>
       </select>
	 
	  
	<a class="a2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</a>

    <button onClick="Procesar();" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;"><img  src="imagenes/gear.png" width="20" height="20" alt="Buscar Consulta" ></button>

	<button onClick="Refresh(7)" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;color:#123456;"><img  src="imagenes/refresh.png" width="20" height="20" alt="Refrescar" ></button>

<div id="DivVariables" style="overflow:auto;height='400 px'; width='100%'"></div>
 <input type="hidden" name="hidAnio" id="hidAnio" value="<%=strUbicaAnio%>"> 
 <input type="hidden" name="hidUsuario" id="hidUsuario" value="<%=usuario%>"> 
 <input type="hidden" name="hidTipo" id="hidTipo" value="<%=tipo%>"> 
 </FORM>
  
</body>
</html> 