<!--#include file="Conexion.asp"-->
<html xmlns:v="urn:schemas-microsoft-com:vml" 
xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
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
    <x:Name>Directorio de Variables</x:Name>
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
TABLE.tabla1 TD
{
    BORDER-RIGHT: #DADBDB 1px solid;
    BORDER-TOP: #DADBDB 1px solid;
    BORDER-LEFT: #DADBDB 2px solid;
    BORDER-BOTTOM: #DADBDB 1px solid;
	
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
	anio=Request.QueryString("anio")
	moneda=Request.QueryString("moneda")
	agrupa=Request.QueryString("agrupa")
	tipfiltro=Request.QueryString("tipfiltro")
	valfiltro=Request.QueryString("valfiltro")

	anio1=anio-1

	SQL="exec sp_SMVMA_BALGEN_DirectorioVariables_AnioMoneda "&anio&","&moneda&","&agrupa&","&tipfiltro&",'"&valfiltro&"'"
	'response.Write(sql)
	'response.End()

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation=3
   	rs.Open SQL, con

	Archivo="Directorio de Variables"
	Titulo="PRINCIPALES CUENTAS DE LOS ESTADOS FINANCIEROS DE LAS EMPRESAS SUPERVISADAS POR LA SMV, Año " &anio

	Filtro=""
	Agrupacion=""

	if TipFilText<>"" then
		Filtro="Filtrado por "&TipFilText&": "&FilText
	end if

	if AgrText<>"" then
		Agrupacion="Agrupado por "&AgrText
	end if

	Response.Charset = "ISO-8859-1"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+".xls" 
	Response.Charset = "ISO-8859-1"
	Response.Write("<table ><tr><td colspan='6' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>"&Filtro&" / "&Agrupacion&"</td></tr><tr>")

	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1' width='100%' border='0'>")

	if agrupa=1 then
		response.write("<tr bgcolor='#E4F2FC' height='30'><td colspan='6'>Empresas que declaran en "&moneda&"</td>")
	else
		response.write("<tr bgcolor='#E4F2FC' height='30'><td colspan='2'>Empresas que declaran en "&moneda&"</td>")
	end if
	    response.write("<td bgcolor='#E4F2FC' colspan='9' align='center'><strong>Ganancias y Perdidas</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' colspan='9' align='center'><strong>Balance General</strong></td>")
    response.write("</tr>")

	response.write("<tr bgcolor='#E4F2FC' height='30'>")

	if agrupa=1 then
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>RUC</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>RAZON SOCIAL</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Cod_AE</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Descripcion AE</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>CIIU_Rev4_DNCN</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>SI</strong></td>")
	elseif agrupa=2 then
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Cod_AE14</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Descripcion AE14</strong></td>")
	elseif agrupa=3 then
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Cod_AE54</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Descripcion AE54</strong></td>")
	elseif agrupa=4 then
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Cod_AE101</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Descripcion AE101</strong></td>")
	end if

	    response.write("<td bgcolor='#E4F2FC' colspan='3' align='center'><strong>INGRESOS</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' colspan='3' align='center'><strong>GASTOS FINANCIEROS</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' colspan='3' align='center'><strong>UTILIDAD NETA</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' colspan='3' align='center'><strong>ACTIVO TOTAL</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' colspan='3' align='center'><strong>PASIVO TOTAL</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' colspan='3' align='center'><strong>PATRIMONIO</strong></td>")
    response.write("</tr>")



	response.write("<tr bgcolor='#FFE2C6'>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio1&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>Var(%)</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio1&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>Var(%)</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio1&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>Var(%)</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio1&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>Var(%)</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio1&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>Var(%)</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio1&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>"&anio&"</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' align='center'><strong>Var(%)</strong></td>")
    response.write("</tr>")


	Set objFieldsC = rs.Fields

	response.write("<tr bgcolor='#FFE2C6'>")
	if agrupa=1 then
		response.write("<td bgcolor='#D5D9DC' align='center'><strong></strong></td>")
		response.write("<td bgcolor='#D5D9DC' align='center'><strong></strong></td>")
		response.write("<td bgcolor='#D5D9DC' align='center'><strong></strong></td>")
		response.write("<td bgcolor='#D5D9DC' align='center'><strong>TOTAL</strong></td>")
		response.write("<td bgcolor='#D5D9DC' align='center'><strong></strong></td>")
		response.write("<td bgcolor='#D5D9DC' align='center'><strong></strong></td>")
	else
		response.write("<td bgcolor='#D5D9DC' align='center'><strong></strong></td>")
		response.write("<td bgcolor='#D5D9DC' align='center'><strong>TOTAL</strong></td>")
	end if

	i=2

	if agrupa=1 then
		i=6
	end if

	For intLoop = i To (objFieldsC.Count - 1)
	    response.write("<td bgcolor='#D5D9DC' align='center'><strong>"&objFieldsC.Item(intLoop).Name&"</strong></td>")
    Next
    response.write("</tr>")

	rs.MoveFirst
	response.write("<tr>")
    while not rs.eof
	    For intLoop = 0 To (objFieldsC.Count - 1)

			if (agrupa=1 and intLoop<5) or (agrupa<>1 and intLoop=0) then
		    	response.write("<td STYLE='vnd.ms-excel.numberformat:@' align='right'>"&rs(intLoop)&"</td>")
		    else
	    		response.write("<td align='right'>"&rs(intLoop)&"</td>")
			end if
	    Next      	
    	response.write("</tr>")
		rs.MoveNext
	wend
	
	rs.Close
	Set rs=Nothing
	response.write("</table></td></tr></table>")

'	response.Write(SQL)
'	response.Write(SQL2)
'	response.End()

	response.write("</tr></table>")
	Response.ContentType = "application/save"

%>
