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
    <x:Name>BalGen</x:Name>
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
<title>Resultado</title>
 <style type="text/css">
<!--
TABLE
{
    BORDER-RIGHT: #c0c0c0 1px dotted;
    BORDER-TOP: #c0c0c0 1px dotted ;
    BORDER-LEFT: #c0c0c0 1px dotted ;
    BORDER-BOTTOM: #c0c0c0 1px dotted;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;
	width:100%;
}
TABLE.TD
{
    BORDER-RIGHT: #828282 1px dotted ;
    BORDER-TOP: #828282 1px dotted ;
    BORDER-LEFT: #828282 1px dotted ;
    BORDER-BOTTOM: #828282 1px dotted ;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;

}
TD.titulo
{
	BORDER-RIGHT: #828282 1px dotted;
    BORDER-TOP: #828282 1px dotted;
    background:#E4F2FC;
    BORDER-LEFT: #828282 1px dotted;
    BORDER-BOTTOM: #828282 1px dotted;
	PADDING: 0.5em;
	HEIGHT: auto;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:right;

}
TD.titulo1
{
	BORDER-RIGHT: #828282 1px dotted;
    BORDER-TOP: #828282 1px dotted;
    BORDER-LEFT: #828282 1px dotted;
    BORDER-BOTTOM: #828282 1px dotted;
	PADDING: 0.5em;
	HEIGHT: auto;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:left;


}
TD.act
{
    BORDER-RIGHT: #828282 1px dotted;
    BORDER-TOP: #828282 1px dotted;
    BORDER-LEFT: #828282 1px dotted;
    BORDER-BOTTOM: #828282 1px dotted;
	PADDING: 0.5em;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:center;
	HEIGHT:40px;
	width:80px;	
}
TD.dat
{
	BORDER-RIGHT: #828282 1px dotted;
    BORDER-TOP: #828282 1px dotted;
    BORDER-LEFT: #828282 1px dotted;
    BORDER-BOTTOM: #828282 1px dotted;
	PADDING: 0.5em;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:right;
}
TD.cab
{
	BORDER-RIGHT: #828282 1px dotted ;
    BORDER-TOP: #828282 1px dotted ;
    BORDER-LEFT: #828282 1px dotted ;
    BORDER-BOTTOM: #828282 1px dotted ;
	PADDING: 0.5em;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:center;
}
-->
</style>
</head>
<body >
<%
dim Tabla(5000,5000)
dim Tabla1(5000,2000)
Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	trime=Request.QueryString("trime")
	moneda=Request.QueryString("moneda")
	detalle=Request.QueryString("detalle")
	metodo=""'Método sólo servirá para Flujo Efectivo	
	TipFondo=Request.QueryString("TipFondo")

	detText=Request.QueryString("detText")
	metText=Request.QueryString("metText")
	
	SQL=" exec sp_lista_cuentas_RepAnioTrim_FONDOS '01','"&annio&"','"&trime&"','','"&detalle&"','"&TipFondo&"'"

	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
   	rs.Open SQL, con
	
	monedaText=moneda
	if moneda ="Total" then
		monedaText="Soles (Incluye Dólares convertidos a Soles)"
	end if

	FondoText=""
	if TipFondo ="MUT" then
		FondoText="FONDOS MUTUOS"
	end if
	if TipFondo ="INV" then
		FondoText="FONDOS DE INVERSION"
	end if
	if TipFondo ="PAT" then
		FondoText="FONDOS DE PATRIMONIO"
	end if

	Archivo="BalanceGeneralFondos_"&FondoText
	Titulo="BALANCE GENERAL SMV, "&FondoText&" " &annio&" - "&trime&"<br><br> Miles de "&monedaText

	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+".xls" 
	Response.Charset = "UTF-8"
	Response.Write("<table ><tr><td colspan='10' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")

	 response.write("<table width='50%' border='1' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1'  border='1'>")

	if detalle =0 then		
		response.write("<tr><td colspan='2' rowspan='5' align='center' bgcolor='#E4F2FC'><strong><font size='2pt'>Balance General</font></strong></td><td bgcolor='#E4F2FC' align='right'>Codigo Sociedad Administradora</td></tr>")
		response.write("<tr bgcolor='#E4F2FC'><td align='right'>Codigo Fondo</td></tr>")
		response.write("<tr bgcolor='#E4F2FC'><td align='right'>Fondo</td></tr>")
		'response.write("<tr bgcolor='#E4F2FC'><td align='right'>Ciiu</td></tr>")
		response.write("<tr bgcolor='#E4F2FC'><td align='right'>Sociedad Administradora</td></tr>")
		response.write("<tr bgcolor='#E4F2FC'><td align='right'>Moneda</td></tr>")
		response.write("<tr bgcolor='#94B9FD'><td>NroOrden</td><td align='center'>Cuenta</td><td align='left'>Descripcion</td></tr>")
	elseif detalle =1 then
		response.write("<tr><td colspan='2' align='center' bgcolor='#E4F2FC'><strong><font size='1pt'>Consolidado EEFF</font></strong></td><td bgcolor='#E4F2FC' align='right'>"&NivText&"</td></tr>")	
		response.write("<tr bgcolor='#94B9FD'><td>NroOrden</td><td align='center'>Cuenta</td><td align='left'>Descripcion</td></tr>")
	end if

	while not rs.eof
		response.write("<tr><td align='center'>"&rs(2)&"</td><td>"&rs(0)&"</td><td align='left'>"&rs(3)&"</td></tr>")
    	rs.MoveNext
	wend
	rs.Close
	Set rs=Nothing

	response.write("</table></td>")

	response.write("<td width='76%'  valign='top'><table class='tabla1' border='1'>")

	SQL="EXEC sp_lista_directorio_RepAnioTriMonMet_FONDOS '01','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"
	SQL2=" exec sp_lista_reporteDatos_RepAnioTrimMonMet_FONDOS '01','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"

	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con

	X1=cint(rs.fields.count)-1
	Y1=cint(rs.RecordCount )-1
	 while not rs.eof
	   for j=0 to X1
			 Tabla(i,j)=rs(j)
		next
	  rs.MoveNext
	  i=i+1
	wend 
	rs.Close
	Set rs=Nothing

	format ="style='vnd.ms-excel.numberformat:0000;'"

	z=0
	
	if detalle =0 then
		X1=X1-1 'TODAS LAS CABECERAS MENOS METODO
		z=1
	end if

	for j=z to X1

		format =""

		if detalle=0 then
			if j=0 then
				format ="style='vnd.ms-excel.numberformat:000000;'"
			elseif j=2 then
				format ="style='vnd.ms-excel.numberformat:0000;'"
			end if
		end if

		response.write("<tr>")
		for i=0 to Y1
			if isnull(Tabla(i,j)) then
				dato="&nbsp;"
			else
				dato=Tabla(i,j)
			end if

			if i Mod 2 = 0 then				
				response.write("<td colspan='3' align='center' bgcolor='#FFE7BB' "&format&">"&dato&"</td>")
			else
				response.write("<td colspan='3' align='center' bgcolor='#E3EEF7' "&format&">"&dato&"</td>")
			end if

		next
		 	response.write("</tr>")
	next

	Set rs2 = Server.CreateObject("ADODB.Recordset")	
	rs2.CursorLocation=3
    rs2.Open sql2, con
	X2=cint(RS2.fields.count)-1
	Y2=cint(rs2.RecordCount )-1

	i=0
	 while not rs2.eof
	   for j=0 to X2
		 Tabla1(i,j)=rs2(j)
		next
	  rs2.MoveNext
	  i=i+1
	wend 
	rs2.Close
	Set rs2=Nothing
'	
	for j=0 to X2
	response.write("<tr>")
		for i=0 to Y2
					if isnull(Tabla1(i,j)) then
						dato="&nbsp;"
					else
						dato=Tabla1(i,j)
					end if

			if j=0 then
				response.write("<td bgcolor='#94B9FD' align='center'><strong>"&dato&"</strong></td>")
			else
				if IsNumeric(dato) then
					response.write("<td align='right' >"&Round(dato)&"</td>")	
				else
					response.write("<td align='right' >"&dato&"</td>")
				end if			
			end if
		next
		'Response.Flush
		response.write("</tr>")
	next

	response.write("</table></td></tr></table>")
	
	response.write("</tr></table>")
	
	Response.ContentType = "application/save" 

%>
