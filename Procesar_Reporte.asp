<!--#include file="Conexion.asp"-->
<%
	Response.Charset= "ISO-8859-1" 

	annio=Request.QueryString("annio")
	usuario=Session("id_usuario")
	tipo=Request.QueryString("tipo")

	
	if usuario = "0000000149" then 
		if tipo="SI" THEN 
		SQL = "sp_GobPro_SisInt '" &annio & "'" 
		END IF
	
		if tipo="SU" THEN 
		SQL = "sp_GobPro_Sumaria '" &annio & "'" 
		
		END IF
		
		Set rs = Server.CreateObject("ADODB.Recordset")	
		Rs.Open SQL,con 
		
		response.Write("Datos Procesados!!!")
		
		
	else
		response.Write("Su Usuario no tiene permisos para procesar DATOS")

	end if
	
%>
