<!--#include file="Conexion.asp"-->
<%
	Response.Charset= "ISO-8859-1" 

	annio=Request.QueryString("annio")
	trime=Request.QueryString("trime")
	CodiEnt=Request.QueryString("CodiEnt")
	
	SQL="EXEC sp_lista_DetalleCamPat_AnioCodiEnt "&annio&",'"&trime&"','"&CodiEnt&"'"

	Set rs = Server.CreateObject("ADODB.Recordset")	

	rs.CursorLocation=3
	rs.Open SQL, con 

	Response.Write("<br>")
	Response.Write("<table class='tabla1'>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td align='left'>Codigo Entidad: "&rs(0)&"</td>")
	Response.Write("<td align='left'>RUC: "&rs(1)&"</td>")
	Response.Write("<td align='left'>RAZÓN SOCIAL: "&rs(2)&"</td>")
	Response.Write("<td align='left'>Ciiu_R4_4d: "&rs(3)&"</td>")
	Response.Write("<td align='left'>AE: "&rs(4)&"</td>")
	Response.Write("<td align='left'>DESCRIPCIÓN AE: "&rs(5)&"</td>")
	Response.Write("<td align='left'>SI: "&rs(6)&"</td>")
	Response.Write("<td align='left'>Moneda: "&rs(7)&"</td>")
	Response.Write("<td align='left'>Empresa "&rs(8)&"</td>")
	Response.Write("</tr>")
	Response.Write("</table>")

	Response.Write("<table class='tabla1'>")

	set rs = rs.NextRecordset
	
	x=rs.Fields.Count-1

	j=0
	for i=0 to x 
		Response.Write("<th >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr "&bg&">")
	
		for i=0 to x
			if (i>=6 and i<=x) then alig="left" else if (i=0) then alig="left" else alig="left" End if End if
		Response.Write("<td  align="&alig&">"&rs(i)&"</td>")
	
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")

	rs.close
%>
