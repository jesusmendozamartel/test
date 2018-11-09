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
	TipFondo=Request.QueryString("TipFondo")

	detText=Request.QueryString("detText")
	xdetFondo=Request.QueryString("xdetFondo")	
	
	SQL01=" exec sp_lista_cuentas_RepAnioTrim_FONDOS '01','"&annio&"','"&trime&"','','"&detalle&"','"&TipFondo&"'"
	SQL02=" exec sp_lista_cuentas_RepAnioTrim_FONDOS '02','"&annio&"','"&trime&"','','"&detalle&"','"&TipFondo&"'"
	SQL03=" exec sp_lista_cuentas_RepAnioTrim_FONDOS '03','"&annio&"','"&trime&"','','"&detalle&"','"&TipFondo&"'"


	Set rs01 = Server.CreateObject("ADODB.Recordset")
	rs01.CursorLocation=3
	rs01.Open SQL01, con
	
	Set rs02 = Server.CreateObject("ADODB.Recordset")	
	rs02.CursorLocation=3
	rs02.Open SQL02, con

	Set rs03 = Server.CreateObject("ADODB.Recordset")	
	rs03.CursorLocation=3
	rs03.Open SQL03, con

	monedaText=moneda
	if moneda ="Total" then
		monedaText="Soles (Incluye Dólares convertidos a Soles)"
	end if

	xdetFondo=""
	if TipFondo ="MUT" then
		xdetFondo="DE FONDOS MUTUOS"
	end if
	if TipFondo ="INV" then
		xdetFondo="DE FONDOS DE INVERSIÓN"
	end if
	if TipFondo ="PAT" then
		xdetFondo="DE PATRIMONIO EN FIDEICOMISO"
	end if

	Archivo="ConsolidadoEEFF_Fondos"
	Titulo="CONSOLIDADO DE ESTADOS FINANCIEROS " &xdetFondo& " SMV, " &annio&" - "&trime&"<br>("&detText&")<br> Miles de "&monedaText

	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+".xls" 
	Response.Charset = "UTF-8"
	Response.Write("<table ><tr><td colspan='10' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")

	 response.write("<table width='50%' border='1' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1'  border='1'>")

	if detalle =0 then
		response.write("<tr><td colspan='2' rowspan='5' align='center' bgcolor='#E4F2FC'><strong><font size='2pt'>Consolidado EEFF</font></strong></td><td bgcolor='#E4F2FC' align='right'>Codigo Sociedad Administradora</td></tr>")
		response.write("<tr bgcolor='#E4F2FC'><td align='right'>Código Fondo</td></tr>")
		 response.write("<tr bgcolor='#E4F2FC'><td align='right'>Fondo</td></tr>")
		response.write("<tr bgcolor='#E4F2FC'><td align='right'>Sociedad Administradora</td></tr>")
		 response.write("<tr bgcolor='#E4F2FC'><td align='right'>Moneda</td></tr>")
	elseif detalle =1 then
		response.write("<tr><td colspan='2' align='center' bgcolor='#E4F2FC'><strong><font size='1pt'>Consolidado EEFF</font></strong></td><td bgcolor='#E4F2FC' align='right'>"&NivText&"</td></tr>")
	end if


	response.write("<tr bgcolor='#94B9FD'><td>NroOrden</td><td align='center'>Cuenta</td><td align='left'>Descripcion</td></tr>")

	'CUENTAS BALANCE GENERAL X
	response.write("<tr bgcolor='#F2DCDB'><td></td><td align='center'></td><td align='left'><strong>Balance General "&annio&"</strong></td></tr>")

	while not rs01.eof
		response.write("<tr><td>"&rs01(2)&"</td><td align='center'>"&rs01(0)&"</td><td align='left'>"&rs01(3)&"</td></tr>")
    	rs01.MoveNext
	wend


	'CUENTAS BALANCE GENERAL X-1
	rs01.MoveFirst
	response.write("<tr bgcolor='#F2DCDB'><td></td><td align='center'></td><td align='left'><strong>Balance General "&annio-1&"</strong></td></tr>")

	while not rs01.eof
		response.write("<tr><td>"&rs01(2)&"</td><td align='center'>"&rs01(0)&"</td><td align='left'>"&rs01(3)&"</td></tr>")
    	rs01.MoveNext
	wend

	'CUENTAS BALANCE GENERAL FLUJO
	rs01.MoveFirst
	response.write("<tr bgcolor='#F2DCDB'><td></td><td align='center'></td><td align='left'><strong>Balance General Flujo</strong></td></tr>")

	while not rs01.eof
		response.write("<tr><td>"&rs01(2)&"</td><td align='center'>"&rs01(0)&"</td><td align='left'>"&rs01(3)&"</td></tr>")
    	rs01.MoveNext
	wend

	rs01.Close
	Set rs01=Nothing

	'CUENTAS ESTADO DE GANANCIAS Y PERDIDAS

	rs02.MoveFirst
	response.write("<tr bgcolor='#F2DCDB'><td></td><td align='center'></td><td align='left'><strong>Estado de Resultados</strong></td></tr>")

	while not rs02.eof
		response.write("<tr><td>"&rs02(2)&"</td><td align='center'>"&rs02(0)&"</td><td align='left'>"&rs02(3)&"</td></tr>")
    	rs02.MoveNext
	wend
	rs02.Close
	Set rs02=Nothing

	'CUENTAS FLUJO DE EFECTIVO
	'response.write("<tr bgcolor='#F2DCDB'><td></td><td></td><td align='center'></td><td align='left'><strong>Flujo de Efectivo</strong></td></tr>")

	'while not rs03.eof
	''	response.write("<tr><td>"&rs03(0)&"</td><td>"&rs03(1)&"</td><td align='center'>"&rs03(2)&"</td><td align='left'>"&rs03(3)&"</td></tr>")
    ''	rs03.MoveNext
	'wend
	'rs03.Close
	'Set rs03=Nothing

	response.write("</table></td>")

	response.write("<td width='76%'  valign='top'><table class='tabla1' border='1'>")
	
	SQL="EXEC sp_lista_directorio_RepAnioTriMonMet_FONDOS '00','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"

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
	z=0

	if detalle =0 then
		X1=X1-1 'TODAS LAS CABECERAS MENOS METODO
		z=1
	end if

	format ="style='vnd.ms-excel.numberformat:0000;'"

	for j=z to X1

		if j=0 then
			format ="style='vnd.ms-excel.numberformat:000000;'"
		elseif j=2 then
			format ="style='vnd.ms-excel.numberformat:0000;'"
		else
			format =""
		end if

		response.write("<tr>")
		for i=0 to Y1
			if isnull(Tabla(i,j)) then
				dato="&nbsp;"
			else
				dato=Tabla(i,j)
			end if

			if i Mod 2 = 0 then				
				response.write("<td align='center' bgcolor='#FFE7BB' "&format&">"&dato&"</td>")
			else
				response.write("<td align='center' bgcolor='#E3EEF7' "&format&">"&dato&"</td>")
			end if

		next
		 	response.write("</tr>")
	next
	'SQL03=" exec sp_lista_reporteDatos_Consolidado_RepAnioTrimMonMet_FONDOS '03','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"

	SQL01=" exec sp_lista_reporteDatos_Consolidado_RepAnioTrimMonMet_FONDOS '01','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"
	SQL02=" exec sp_lista_reporteDatos_Consolidado_RepAnioTrimMonMet_FONDOS '02','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"
	'SQL03=" exec sp_lista_reporteDatos_Consolidado_RepAnioTrimMonMet_FONDOS '03','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"

	Set rs01 = Server.CreateObject("ADODB.Recordset")
	rs01.CursorLocation=3
	rs01.Open SQL01, con

	Set rs02 = Server.CreateObject("ADODB.Recordset")
	rs02.CursorLocation=3
	rs02.Open SQL02, con

	Set rs03 = Server.CreateObject("ADODB.Recordset")	
	rs03.CursorLocation=3
	rs03.Open SQL03, con

	'BALANCE GENERAL
	if rs01.fields.count>0 then
		X2=cint(rs01.fields.count)-1
		Y2=cint(rs01.RecordCount )-1

		i=0
		 while not rs01.eof
		   for j=0 to X2
			 Tabla1(i,j)=rs01(j)
			next
		  rs01.MoveNext
		  i=i+1
		wend 

		rs01.Close
		Set rs01=Nothing
		'BALANCE GENERAL X
		for j=0 to X2
			response.write("<tr>")
			for i=0 to Y2

				if isnull(Tabla1(i,j)) then
					dato="&nbsp;"
				else
					dato=Tabla1(i,j)
				end if

				if j=0 then
					response.write("<td bgcolor='#94B9FD' align='center'><strong>&nbsp;</strong></td>")
				else
					if IsNumeric(dato) then
						response.write("<td align='right' >"&Round(dato)&"</td>")	
					else
						response.write("<td align='right' >"&dato&"</td>")
					end if			
				end if
				i=i+2
			next
			response.write("</tr>")

			if j=0 then
				response.write("<tr>")
				for i=0 to Y2/3
					response.write("<td bgcolor='#F2DCDB' align='center'>&nbsp;</td>")
				next
				response.write("</tr>")
			end if

		next

		'BALANCE GENERAL X-1
		response.write("<tr>")
		for i=0 to Y2/3
			response.write("<td bgcolor='#F2DCDB' align='center'>&nbsp;</td>")
		next
		response.write("</tr>")

		for j=1 to X2
			response.write("<tr>")
			for i=1 to Y2

				if isnull(Tabla1(i,j)) then
					dato="123"
				else
					dato=Tabla1(i,j)
				end if

				if j>0 then					
					if IsNumeric(dato) then
						response.write("<td align='right' >"&Round(dato)&"</td>")
					else
						response.write("<td align='right' >"&dato&"</td>")
					end if
				end if
				i=i+2
			next
			response.write("</tr>")

		next

		'BALANCE GENERAL FLUJO
		response.write("<tr>")
		for i=0 to Y2/3
			response.write("<td bgcolor='#F2DCDB' align='center'>&nbsp;</td>")
		next
		response.write("</tr>")
		
		for j=1 to X2
			response.write("<tr>")
			for i=2 to Y2

				if isnull(Tabla1(i,j)) then
					dato="&nbsp;"
				else
					dato=Tabla1(i,j)
				end if

				if j>0 then					
					if IsNumeric(dato) then
						response.write("<td align='right' >"&Round(dato)&"</td>")	
					else
						response.write("<td align='right' >"&dato&"</td>")
					end if
				end if
				i=i+2
			next
			response.write("</tr>")

		next

	'ESTADO DE GANANCIAS Y PERDIDAS
		X2=cint(rs02.fields.count)-1
		'Y2=(cint(rs02.RecordCount )*1.5)-1
		Y2=cint(rs02.RecordCount)-1

		i=0
		 while not rs02.eof
		   for j=0 to X2
			 Tabla1(i,j)=rs02(j)
			next
		  rs02.MoveNext
		  i=i+1
		wend 
		rs02.Close
		Set rs02=Nothing

		response.write("<tr>")
		for i=1 to (Y2+1)/2
			response.write("<td bgcolor='#F2DCDB' align='center'>&nbsp;</td>")
		next
		response.write("</tr>")

		for j=1 to X2
			response.write("<tr>")
			for i=0 to Y2

				if isnull(Tabla1(i,j)) then
					dato="&nbsp;"
				else
					dato=Tabla1(i,j)
				end if

				if i mod 2 =0 then
					if IsNumeric(dato) then
						response.write("<td align='right' >"&Round(dato)&"</td>")	
					else
						response.write("<td align='right' >"&dato&"</td>")
					end if
				end if
			next
			response.write("</tr>")

		next


	'ESTADO DE FLUJO DE EFECTIVO
		
	''	X2=cint(rs03.fields.count)-1
	''	Y2=cint(rs03.RecordCount )-1

	''	i=0
	''	 while not rs03.eof
	''	   for j=0 to X2
	''		 Tabla1(i,j)=rs03(j)
	''		next
	''	  rs03.MoveNext
	''	  i=i+1
	''	wend 

	''	rs03.Close
	''	Set rs03=Nothing

		'FLUJO DE EFECTIVO X
	''	for j=0 to X2
	''		response.write("<tr>")
	''		for i=0 to Y2

	''			if isnull(Tabla1(i,j)) then
	''				dato="&nbsp;"
	''			else
	''				dato=Tabla1(i,j)
	''			end if

	''			if j>0 then
	''				if IsNumeric(dato) then
	''					response.write("<td align='right' >"&FormatNumber(dato,0)&"</td>")	
	''				else
	''					response.write("<td align='right' >"&dato&"</td>")
	''				end if
	''			end if
	''			i=i+2
	''		next
	''		response.write("</tr>")

	''		if j=0 then
	''			response.write("<tr>")
	''			for i=0 to Y2/3
	''				response.write("<td bgcolor='#F2DCDB' align='center'>&nbsp;</td>")
	''			next
	''			response.write("</tr>")
	''		end if
	''	next

	end if
	response.write("</table></td></tr></table>")
	
	response.write("</tr></table>")
	
	Response.ContentType = "application/save" 

%>
