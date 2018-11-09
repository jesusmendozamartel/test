<%OPTION EXPLICIT%>
<!--#include file="conexion.asp"-->
<% 
Response.Buffer=true
'Las tres líneas de código siguientes se utilizan para garantizar que esta página no se almacena en la memoria caché del cliente.'
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1

Dim user, pass, aplicacion, pag, codPersona
user=Request.Form("txtUser")
pass=Request.Form("txtPass") 
aplicacion=Request.QueryString("aplicacion")
pag=request.QueryString("pag")

''Response.Write("-"&user&"-"&pass&"-"&aplicacion&"-"&pag&"-"): response.End()
''response.End()
'Comprueba si el identificador de usuario es una cadena vacía. Si está vacío, redirija a Logon.asp.'
'Si no está vacío, conecta a la base de datos y valida al usuario.'
''Response.Write "<form action='login.asp' name='frmuser' method='post'>"
if user <>"" and pass<>""  then
	Dim RS, sql
	Set RS=Server.CreateObject("ADODB.RecordSet")
	''Response.Write("exec dncn_administrativo.dbo.sp_validar_usuario '"&user&"','"&pass&"'"): response.end
	RS.Open "exec dncn_administrativo.dbo.sp_validar_usuario '"&user&"','"&pass&"'", Con
	if not RS.EOF then 
		codPersona=RS("cod_persona")
		Session("id_usuario")=codPersona
		RS.close
		'Si tiene permiso para accesar a esta aplicacion'
		sql="exec dncn_administrativo.dbo.sp_tipo_acceso_aplicacion '"&aplicacion&"','"&codPersona&"'" 
''		Response.Write(sql) : response.end
		RS.Open sql,Con
		if not RS.EOF then 			
			Session("tipoAcceso")=RS("coc_tipoacceso")
			RS.close
			Response.Redirect pag''&"?USER="&user&"&PSW="&pass 
			Response.End 
		End if
	End if
End if
Response.Redirect "logoffce.asp"
Set RS=nothing
Response.End 
''Response.Write "</form>"
%> 