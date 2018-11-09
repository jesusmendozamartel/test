<!--#include file="Conexion.asp"-->
<%
	dim Tabla(5000,5000)
	dim Tabla1(5000,2000)
	Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	trime=Request.QueryString("trime")
	moneda=Request.QueryString("moneda")
	detalle=Request.QueryString("detalle")
	TipFondo=Request.QueryString("TipFondo")

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


	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1'  border='1'>")

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

	response.write("<td width='76%'  valign='top'><table class='tabla1' border='0'>")
	
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

	for j=z to X1
		response.write("<tr>")
		for i=0 to Y1
			if isnull(Tabla(i,j)) then
				dato="&nbsp;"
			else
				dato=Tabla(i,j)
			end if

			if i Mod 2 = 0 then
					response.write("<td align='center' bgcolor='#FFE7BB'>"&dato&"</td>")				
			else
					response.write("<td align='center' bgcolor='#E3EEF7'>"&dato&"</td>")				
			end if

		next
		 	response.write("</tr>")
	next

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
						response.write("<td align='right' >"&FormatNumber(dato,0)&"</td>")	
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
		for j=0 to X2
			response.write("<tr>")
			for i=1 to Y2

				if isnull(Tabla1(i,j)) then
					dato="&nbsp;"
				else
					dato=Tabla1(i,j)
				end if

				if j>0 then					
					if IsNumeric(dato) then
						response.write("<td align='right' >"&FormatNumber(dato,0)&"</td>")	
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

		'BALANCE GENERAL FLUJO
		for j=0 to X2
			response.write("<tr>")
			for i=2 to Y2

				if isnull(Tabla1(i,j)) then
					dato="&nbsp;"
				else
					dato=Tabla1(i,j)
				end if

				if j>0 then					
					if IsNumeric(dato) then
						response.write("<td align='right' >"&FormatNumber(dato,0)&"</td>")	
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
						response.write("<td align='right'>"&FormatNumber(dato,0)&"</td>")
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
	
%>
