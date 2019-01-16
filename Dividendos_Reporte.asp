<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,2000)
Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	trime=Request.QueryString("trime")
	nivel=Request.QueryString("nivel")
	codigo=Request.QueryString("codigo")
	moneda=Request.QueryString("moneda")
	letra=Request.QueryString("letra")
	detalle=Request.QueryString("detalle")

	SQL=" exec sp_lista_DividendosCamPat_AnioTriNivMonLetSup '"&annio&"','"&trime&"','"&nivel&"','"&codigo&"','"&moneda&"','"&letra&"','S'"

	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
   	rs.Open SQL, con
	
	'RESPONSE.Write(SQL)
	'RESPONSE.End()


	Response.Write("<tr class='a1'> Recuento: "&rs.RecordCount-1&"</tr>")
	'rs.close		
	'SQL="exec sp_lista_empresas_x_anio '"&reporte& "','"&annio& "' "
	


'	rs.MoveNext
	'rs.NextRecordset


	x=rs.Fields.Count-1
	
	j=0

	Response.Write("<table class='tabla1'>")

	for i=0 to x 
		Response.Write("<th >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr "&bg&">")
	
		for i=k to x
			if (i>=6 and i<=x) then alig="left" else if (i=0) then alig="left" else alig="left" End if End if
		Response.Write("<td  align="&alig&">"&Rs(i)&"</td>")
	
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")

	rs.close
%>
