<!--#include file="Conexion.asp"-->
<html xmlns:v="urn:schemas-microsoft-com:vml" 
xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 9">
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
<!--[if !mso]>
<style>
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
x\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
<![endif]-->
<!--[if gte mso 9]><xml>
 <o:OfficeDocumentSettings>
  <o:DoNotRelyOnCSS/>
  <o:DoNotUseLongFilenames/>
  <o:DownloadComponents/>
  <o:LocationOfComponents HRef="file:msowc.cab"/>
 </o:OfficeDocumentSettings>
</xml><![endif]-->

<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   <x:ExcelWorksheet>
    <x:Name>Dividendos</x:Name>
    <x:WorksheetOptions>
 
     <x:Print>
      <x:ValidPrinterInfo/>
      <x:PaperSizeIndex>9</x:PaperSizeIndex>
      <x:Scale>85</x:Scale>
      <x:HorizontalResolution>600</x:HorizontalResolution>
      <x:VerticalResolution>600</x:VerticalResolution>
     </x:Print>
     <x:Selected/>

    </x:WorksheetOptions>
   </x:ExcelWorksheet>
  </x:ExcelWorksheets>
  <x:WindowHeight>8070</x:WindowHeight>
  <x:WindowWidth>11580</x:WindowWidth>
  <x:WindowTopX>1</x:WindowTopX>
  <x:WindowTopY>1</x:WindowTopY>
  <x:ProtectStructure>False</x:ProtectStructure>
  <x:ProtectWindows>False</x:ProtectWindows>
 </x:ExcelWorkbook>
  <x:ExcelName>
  <x:Name>Print_Titles</x:Name>
  <x:SheetIndex>1</x:SheetIndex>
  <x:Formula>='reporte'!$1:$4</x:Formula>
 </x:ExcelName>
 </xml><![endif]-->
 
<title>Resultado</title></head>
<body >
<%
	Response.Charset= "ISO-8859-1" 

	annio=Request.QueryString("annio")
	trime=Request.QueryString("trime")
	nivel=Request.QueryString("nivel")
	codigo=Request.QueryString("codigo")
	moneda=Request.QueryString("moneda")
	letra=Request.QueryString("letra")
	detalle=Request.QueryString("detalle")

	letText=Request.QueryString("letText")
	codText=Request.QueryString("codText")

	
	SQL=" exec sp_lista_DividendosCamPat_AnioTriNivMonLetSup '"&annio&"','"&trime&"','"&nivel&"','"&codigo&"','"&moneda&"','"&letra&"','NS'"

	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con 

	NivText=	""
	if nivel =5 then
		NivText="ACTIVIDAD"
	elseif nivel =6 then
		NivText="SECTOR INSTITUCIONAL"
	elseif nivel =10 then
		NivText="ACTIVIDAD 54"
	elseif nivel =11 then
		NivText="ACTIVIDAD 14"
	elseif nivel =12 then
		NivText="TIPO ENTIDAD"
	end if

	monedaText=moneda
	if moneda ="Total" then
		monedaText="Soles (Incluye Dólares convertidos a Soles)"
	end if

	x=rs.Fields.Count-1
	
	if rs.RecordCount=1 then
		Response.Write(rs.RecordCount) ''No se encontraron registros!
		Response.End
	End if
	j=0

	Archivo="Dividendos"
	Titulo="Dividendos SMV - Empresas NO Supervisadas, " &annio&" - "&trime&"<br>"&NivText&": "&codText&" (Categoria "&letText&")<br> Miles de "&monedaText

    Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+"_"+annio+".xls" 
	Response.Charset= "ISO-8859-1" 	
	
	Response.Write("<table ><tr><td colspan='10' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")

	Response.Write("<br>")
	Response.Write("<table>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'><br>Recuento: "&rs.RecordCount-1&"<br></td>")
	Response.Write("</tr>")
	Response.Write("</table>")
	Response.Write("<br>")
	Response.Write("</td></tr>")

	response.write("<table  class='tabla1'  border='1'>")

	for i=0 to x 
		Response.Write("<th bgcolor='#314576' >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr "&bg&">")
	
		for i=k to x
			if (i=10 or i=11) then
				Response.Write("<td align=right>"&FormatNumber(Rs(i),0)&"</td>")
			else
				Response.Write("<td STYLE='vnd.ms-excel.numberformat:@' align=left>"&Rs(i)&"</td>")
			End if
	
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")	
	response.write("</tr></table>")
	Response.ContentType = "application/save"

%>
</body >
</html>
