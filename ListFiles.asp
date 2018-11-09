<!--#include file="Conexion.asp"-->
<%
Response.Charset= "ISO-8859-1" 
  
  anio =Request.QueryString("anio")
  jerarquia =Request.QueryString("jer")
  filt =Request.QueryString("filt")
  cadena=""

  Dim rs
  Set rs = Server.CreateObject("ADODB.Recordset")
  rs.CursorLocation=3

  if filt=1 then
    rs.Open "exec sp_Menu_PDT_Listar_AnioItem "&anio&",'"&jerarquia&"'", Con
  ElseIf filt=0 then
    rs.Open "exec sp_Menu_PDT_Listar_Item '"&jerarquia&"'", Con
  end if

  response.write("<br>  <table border='1' align='center' cellpadding='0' cellspacing='0' class='tabla1'><tbody>")

  link=""
	while not rs.eof	
		if Len(Trim(rs("Jerarquia")))=4 then

     response.write("<tr><td width='241' colspan='2' align='left' valign='middle'><strong>"&rs(1)&"</strong></td></tr>")
		ElseIf Len(Trim(rs("Jerarquia")))=6 then
      link=rs(2)
      if (rs(7)=1) then
        link="getfile.asp?path="&link
      end if

     response.write("<tr><td width='100' align='center' valign='middle'><a href='getfile.asp?path="&rs(2)&"' target='_blank'><img src="&rs(5)&" width='20' height='20' border='0' align='middle' style='border:none'></a></td><td><a href='"&link&"' target='_blank'>"&rs(1)&"</a></td></tr>")

		end if
  	chck=""
    rs.MoveNext
	wend
	rs.Close
    Set rs = Nothing
    response.write("</tbody></table><BR><BR>")
%>

