<!--#include file="Conexion.asp"-->
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

	SQL=" exec sp_lista_cuentas_RepAnioTrim_FONDOS '01','"&annio&"','"&trime&"','','"&detalle&"','"&TipFondo&"'"

	'response.Write(sql)
	'response.End()
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
   	rs.Open SQL, con
	
	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1'  border='1'>")

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

	response.write("<td width='76%'  valign='top'><table class='tabla1' border='0'>")

	SQL="EXEC sp_lista_directorio_RepAnioTriMonMet_FONDOS '01','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"
	SQL2=" exec sp_lista_reporteDatos_RepAnioTrimMonMet_FONDOS '01','"&annio&"','"&trime&"','"&moneda&"','',"&detalle&",'"&TipFondo&"'"

	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
    rs.Open sql, con

	X1=cint(RS.fields.count)-1
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

	for j=z to X1
		response.write("<tr>")
		for i=0 to Y1
			if isnull(Tabla(i,j)) then
				dato="&nbsp;"
			else
				dato=Tabla(i,j)
			end if
			
			if i Mod 2 = 0 then
				response.write("<td colspan='3' align='center' bgcolor='#FFE7BB'>"&dato&"</td>")
				
			else
				response.write("<td colspan='3' align='center' bgcolor='#E3EEF7'>"&dato&"</td>")
				
			end if
		next
		 	response.write("</tr>")
	next
'
	Set rs2 = Server.CreateObject("ADODB.Recordset")	
	rs2.CursorLocation=3
    rs2.Open sql2, con
	X2=cint(RS2.fields.count)-1
	Y2=cint(rs2.RecordCount )-1
	'response.write(X2&"-hhh"&Y2)
'    'response.End()
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
					response.write("<td align='right'>"&FormatNumber(dato,0)&"</td>")
				else
					response.write("<td align='right' >"&dato&"</td>")
				end if
				
			end if
		next
		'Response.Flush
		response.write("</tr>")
	next
'	
	response.write("</table></td></tr></table>")

	
%>
