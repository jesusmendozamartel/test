<% Response.Buffer=True
'Las tres líneas de código siguientes se utilizan para garantizar que esta página no se almacena en la memoria caché del cliente. 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1
Session.Abandon 
Response.Redirect "login.html" 
Response.End 
%> 