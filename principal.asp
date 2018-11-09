<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script type="text/javascript" src="java_script/stmenu.js">
</script>


<% Response.expires = 0 
	if Session("tipoAcceso")="" then 
		response.redirect "login.html"
	end if
	ruta="imagenes"
	''response.write Session("id_usuario")
%>
<title>.:INEI-DNCN - Sistema de Consultas Conasev</title>
<style type="text/css">
body {
	margin-left: 0px;
	margin-right: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
</style>
<script type="text/javascript" src="stmenu.js"></script></head>
<body bgcolor="#FFFFFF">
<table width="100%" height="72" border="0" cellpadding="0" cellspacing="0" background="<%=ruta%>/CONASEV_fondo.jpg">
  <tr>
  <td>
  	<table width="200" height="72" background="<%=ruta%>/CONASEV_izq.png">
    	<tr><td></td></tr>
    </table>
  </td>
  <td align="right"> 
  	<table width="210" height="72" background="<%=ruta%>/CONASEV_drch.png">
    	<tr><td></td></tr>
    </table>
  </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="10"><script type="text/javascript" language="JavaScript1.2" src="java_script/menu.js"></script>      <a href="logoffce.asp" style="font-family:verdana; font-size:10px; color:#000000"></a></td>
  </tr>          
</table>
</body>
</html> 