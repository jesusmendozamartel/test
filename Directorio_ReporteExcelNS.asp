<!--#include file="Conexion.asp"-->
<html xmlns:v="urn:schemas-microsoft-com:vml" 
xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 9">
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
    <x:Name>Directorio</x:Name>
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

	orden=Request.QueryString("orden") 
	annio=Request.QueryString("annio")
	trime=Request.QueryString("trime")
	
	SQL="exec sp_lista_directorio_x_anio_ns '"&orden& "','"&annio& "','"&trime& "'"

	Archivo="DirectorioNS"
	Titulo="Directorio de Empresas No Supervisadas"
		
	'response.Write(SQL)
	'response.End()


	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con 

	if rs.recordcount=1 then
			Response.Write(rs.recordcount)''No se encontraron registros!
			Response.End
	end if
	j=0

	Response.Charset = "UTF-8"
	''Server.CreateObject("Excel.Application") 
	Response.ContentType = "application/vnd.ms-excel" 
	Response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+"_"+annio+".xls" 
	Response.Charset= "ISO-8859-1" 
	
	Response.Write("<table class='tabla1'><tr><td colspan='6' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&UCase(Titulo)&"   "&annio&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")
	x=rs.Fields.Count-1
	for i=0 to x
	Response.Write("<th bgcolor='#D9E6F4' >"&rs.fields(i).name&"</th>")
	next
	
	Response.Write("</tr>")
	
	while not rs.eof
		Response.Write("<tr >")
		for i=0 to x
			'if (i>=6 and i<=x) then alig="center" else if (i=0) then alig="center" else alig="left" End if End if
			alig="left"
			Response.Write("<td nowrap='nowrap' "&col&bg&" STYLE='vnd.ms-excel.numberformat:@' align="&alig&">"&Rs(i)&"</td>")
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")
	Response.ContentType = "application/save" 

%>
</body >
</html>
