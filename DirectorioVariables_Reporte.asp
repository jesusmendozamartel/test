<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,2000)

Response.Charset= "ISO-8859-1" 
	
	anio=Request.QueryString("anio")
	moneda=Request.QueryString("moneda")
	nivel=Request.QueryString("nivel")
	
	anio1=anio-1

	Dim rs
	Set rs=Server.CreateObject("ADODB.RecordSet")	
	rs.CursorLocation=3
	rs.Open "exec sp_SMVMA_BALGEN_DirectorioVariables_AnioMoneda "&anio&","&moneda&","&nivel, Con

	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1' width='100%' border='0'>")

	if nivel=1 then
		response.write("<tr bgcolor='#E4F2FC' height='30'><td colspan='6'></td>")
	else
		response.write("<tr bgcolor='#E4F2FC' height='30'><td colspan='2'></td>")
	end if
	    response.write("<td bgcolor='#E4F2FC' colspan='9' align='center'><strong>Ganancias y Perdidas</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' colspan='9' align='center'><strong>Balance General</strong></td>")
    response.write("</tr>")

	response.write("<tr bgcolor='#E4F2FC' height='30'>")

	if nivel=1 then
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>RUC</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>RAZON SOCIAL</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Cod_AE</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Descripcion AE</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>CIIU_Rev4_DNCN</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>SI</strong></td>")
	elseif nivel=2 then
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Cod_AE14</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Descripcion AE14</strong></td>")	
	elseif nivel=3 then
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Cod_AE54</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Descripcion AE54</strong></td>")	
	elseif nivel=4 then
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Cod_AE101</strong></td>")
	    response.write("<td bgcolor='#E4F2FC' rowspan='2' align='center'><strong>Descripcion AE101</strong></td>")	
	end if

	    response.write("<td bgcolor='#E4F2FC' colspan='3' align='center'><strong>INGRESOS ACTIVIDADES ORDINARIAS</strong></td>")
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
	if nivel=1 then
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

	if nivel=1 then
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
	    	response.write("<td align='left'>"&rs(intLoop)&"</td>")
	    Next      	
    	response.write("</tr>")
		rs.MoveNext
	wend
	
	rs.Close
	Set rs=Nothing
	response.write("</table></td></tr></table>")
%>
