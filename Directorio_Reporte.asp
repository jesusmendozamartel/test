<!--#include file="Conexion.asp"-->
<%
	Response.Charset= "ISO-8859-1" 

	orden=Request.QueryString("orden") 
	annio=Request.QueryString("annio")
	trime=Request.QueryString("trime")
	
	SQL="exec sp_lista_directorio_x_anio '"&orden& "','"&annio& "','"&trime& "'"
	'response.Write(SQL)
	'response.End()

	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con 

	if rs.RecordCount=0 then
		response.write("<div align='center'><p style='color:#000';><strong>�No se encontraron datos!</strong></p></div>")
		response.end
	end if

	x=rs.Fields.Count-1
	
	if rs.RecordCount=1 then
		Response.Write(rs.RecordCount) ''No se encontraron registros!
		Response.End
	End if
	j=0

	Response.Write("<table class='tabla1'>")

	for i=0 to x 
		Response.Write("<th >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr "&bg&">")
	
		for i=k to x
			'if (i>=6 and i<=x) then alig="center" else if (i=0) then alig="center" else alig="left" End if End if
			alig="left"
		Response.Write("<td  align="&alig&">"&Rs(i)&"</td>")
	
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")
%>
