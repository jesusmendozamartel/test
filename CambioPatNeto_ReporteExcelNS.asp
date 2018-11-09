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
    <x:Name>CambioPatNetoNS</x:Name>
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
	sector=Request.QueryString("sector")
	tipo=Request.QueryString("tipo")
	Archivo="CambioPatNetoNs "
	Titulo="Cambio de Estado de Patrimonio Neto de Empresas no Supervisadas "

	'/****Para los titulos***/
	if tipo ="04" or  tipo ="08" or  tipo ="11" or tipo ="18" then 
	SQL2="SELECT desccta_pat FROM dbo.SMVMA_CAMPAT_DESC where anio_inf='"&annio&"' and cod2_emp='"&tipo&"' order by orden"
	else
	SQL2="SELECT desccta_pat FROM dbo.SMVMA_CAMPAT_DESC where anio_inf='"&annio&"' and cod2_emp='' order by orden"
	end if
	Set rs2 = Server.CreateObject("ADODB.Recordset")	
	rs2.CursorLocation=3
    rs2.Open SQL2, con
	X2=cint(RS2.fields.count)-1
	Y2=cint(rs2.RecordCount )-1
	
	dim Tabla2(5000,5000)
	 f=2  
	 c=0
	 while not rs2.eof
	   for c=0 to X2
			 Tabla2(f,c)=rs2(c)

		next
	  rs2.MoveNext
	  f=f+1
	 wend 
	 Tabla2(0,2)=""
 	 Tabla2(1,2)=""
  	
	'/****Para los datos***/
	SQL=" exec sp_lista_campat_ns '"&tipo&"','"&annio&"','"&trime&"'"
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con 
	x=rs.Fields.Count-1
	
	if rs.RecordCount=1 then
		Response.Write(rs.RecordCount) ''No se encontraron registros!
		Response.End
	End if
	
	
	j=0

	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	Response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+"_"+sector+"_"+tipo+"_"+annio+".xls" 
	Response.Write("<table class='tabla1'><tr><td colspan='6' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&UCase(Titulo)&"   "&annio&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")

	
	for c=0 to X2
			if c<=6 then bg="bgcolor='#D9E6F4'"  else bg="" end if
			Response.Write("<tr "&bg&" >")
			for f=0 to Y2+2
				 if isnull(Tabla2(f,c)) then
					response.write("<td >&nbsp;</td>")	 
				 else 
					if c<=6then  centra="align='center'" else if f>3 then  centra="align='right'" else centra="" end if	end if
					if (f=1 ) then Response.Write("<td colspan='12' "&nw&" "&centra&" "&color&" >"&Tabla2(f,c)&"</td>") End if
					if (f>1) then Response.Write("<td colspan='1' "&nw&" "&centra&" "&color&" >"&Tabla2(f,c)&"</td>") End if
					
				End if	
				
			next
			 Response.Write("</tr>")
		
		next

	for i=0 to x 
		Response.Write("<th >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr "&bg&">")
	
		for i=k to x
			if (i=3) then alig="left" else if (i<=2) then alig="center" else alig="right" End if End if
		Response.Write("<td  align="&alig&">"&Rs(i)&"</td>")
	
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")
	Response.ContentType = "application/save" 
%>
