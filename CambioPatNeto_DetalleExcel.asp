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
    BORDER-RIGHT: #000000 solid;
    BORDER-TOP: #000000 solid;
    BORDER-LEFT: #000000 solid;
    BORDER-BOTTOM: #000000 solid;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;
	width:100%;
	background:#FFFFFF;
}


TABLE.tabla1 TH
{
    BORDER-RIGHT: #000000 solid;
    PADDING-RIGHT: 5px;
    BORDER-TOP: #000000 solid;
    PADDING-LEFT: 5px;
	background:#D9E6F4;
    PADDING-BOTTOM: 5px;
    BORDER-LEFT: #000000 solid;
    PADDING-TOP: 5px;
    BORDER-BOTTOM: #000000 solid;
    HEIGHT: 40px;
	color:#000000;
	font-family:Arial, Helvetica, sans-serif;
	font-size:12px;
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
    <x:Name>Cambio Patrimonio</x:Name>
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
	CodiEnt=Request.QueryString("CodiEnt")
	
	Archivo="CambioPatrimonio"&CodiEnt
	Titulo="Estados de Cambios en el Patrimonio Neto<br>SMV "&annio&"-"&trime
	
	SQL="EXEC sp_lista_DetalleCamPat_AnioCodiEnt "&annio&",'"&trime&"','"&CodiEnt&"'"

	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con 

	x=rs.Fields.Count-1
	
    Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+"_"+annio+".xls" 
	Response.Charset= "ISO-8859-1" 	

	Response.Write("<table ><tr><td colspan='10' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")

	Response.Write("<br>")
	Response.Write("<table>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>Codigo Entidad: "&rs(0)&"</td>")
	Response.Write("</tr>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>RUC: "&rs(1)&"</td>")
	Response.Write("</tr>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>RAZÓN SOCIAL: "&rs(2)&"</td>")
	Response.Write("</tr>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>Ciiu_R4_4d: "&rs(3)&"</td>")
	Response.Write("</tr>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>AE: "&rs(4)&"</td>")
	Response.Write("</tr>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>DESCRIPCIÓN AE: "&rs(5)&"</td>")
	Response.Write("</tr>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>SI: "&rs(6)&"</td>")
	Response.Write("</tr>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>Moneda: "&rs(7)&"</td>")
	Response.Write("</tr>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'>Empresa "&rs(8)&"</td>")
	Response.Write("</tr>")
	Response.Write("</table>")
	Response.Write("<br>")

	Response.Write("</td></tr><tr>")

	response.write("<table  class='tabla1'  border='1'>")

	set rs = rs.NextRecordset
	
	x=rs.Fields.Count-1

	j=0
	for i=0 to x 
		Response.Write("<th bgcolor='#314576' >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr>")
	
		for i=0 to x
			if (i>1) then
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
