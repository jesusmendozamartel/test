<% Response.Buffer=True
'Las tres l�neas de c�digo siguientes se utilizan para garantizar que esta p�gina no se almacena en la memoria cach� del cliente. 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1
Session.Abandon 
Response.Redirect "login.html" 
Response.End 
%> 