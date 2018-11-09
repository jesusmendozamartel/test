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
  rep =Request.QueryString("rep")
  filt =Request.QueryString("filt")
  des =Request.QueryString("des")
  ''Response.Write Session("id_usuario")
%>


<script type="text/javascript">

  $(document).ready(function() {
    $.ajaxSetup({ cache: false });
  });

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

.combo1 { background-color:ffffff;
  border: 1px solid #008BC0;
  ursor: pointer;
  font-size:9pt;
  color: #008BC0;
}
.combo2 { background-color:ffffff;
  border: 1px solid #008BC0;
  ursor: pointer;
  font-size:9pt;
  color: #008BC0;
}

.combo3 { background-color:ffffff;
  border: 1px solid #008BC0;
  ursor: pointer;
  font-size:9pt;
  color: #008BC0;
}
.combo4 { background-color:ffffff;
  border: 1px solid #008BC0;
  ursor: pointer;
  font-size:9pt;
  color: #008BC0;
}
.combo5 { background-color:ffffff;
  border: 1px solid #008BC0;
  ursor: pointer;
  font-size:9pt;
  color: #008BC0;
}
TABLE.tabla1
{
    BORDER-RIGHT: #DADBDB 1px solid;
    BORDER-TOP: #DADBDB 1px solid;
    BORDER-LEFT: #DADBDB 1px solid;
    BORDER-BOTTOM: #DADBDB 1px solid;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
    font-family:Arial, Geneva, sans-serif;
    font-size:12px;
    color: #83430a;
    width:60%;
    background:#FFFFFF;
  
}
TABLE.tabla1 TD
{
    BORDER-RIGHT: #DADBDB 1px solid;
    BORDER-TOP: #DADBDB 1px solid;
    BORDER-LEFT: #DADBDB 1px solid;
    BORDER-BOTTOM: #DADBDB 1px solid;
  
  border: #CECECE 1px dotted; 
    padding:0.5em;
    vertical-align:middle;  
    background-color:#FFFFFF;
    color:#205596;
    font-family:Arial, Helvetica, sans-serif;
    font-size:12px;
    width:auto; 
}

</style>
<body onload="cargaArchivos(<%=rep%>,<%=filt%>);">
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
<font face="Arial" size='3pt' color='#FFFFFF'><%=des%></font></strong>
<form action="Files.asp" name="fmrEF" method="post" target=_self>

  <table border="1" align="center" cellpadding="0" cellspacing="0" class="tabla1">
  <tbody>
  <tr>
    <td width="241" colspan="2" align="left" valign="middle">

      <% 
        if (filt=1) then
      %>

    <a class=a2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Año</a>

  <select name="cboAnio" id="cboAnio" onchange="cargaArchivos(<%=rep%>,<%=filt%>);">
      <% 
        Set rs = Server.CreateObject("ADODB.Recordset")
      rs.CursorLocation=3
      rs.Open "exec sp_Menu_PDT_ListarAnio '"&rep&"'" , con

        chck="selected='selected'"

      while not rs.eof  
    %>
        <option value='<%=Trim(rs("anio"))%>' <%=chck%> ><%=Trim(rs("anio"))%></option>
      <%
        chck=""
        rs.MoveNext
      wend
      rs.Close
        Set rs = Nothing
    %>
  </select>

      <% 
        end if
      %>
  </tr></tbody></table>

<div id="DivVariables"></div>
<!--<div id="DivVariables" style="overflow:scroll;height:520px; width:100%"></div>-->

</FORM>
</body>
</html>
