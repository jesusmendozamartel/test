<% 
Dim Con
Set Con = CreateObject("ADODB.Connection")
Con.ConnectionString = "Provider=SQLOLEDB; Data Source=localhost; Initial Catalog=DNCN_SMV; UID=sa; PWD=123456;"
'Con.ConnectionString = 	"Provider=SQLOLEDB.1;Password=lorsimmampre;Persist Security Info=True;User ID=sa;Initial Catalog=DNCN_SMV;Data Source=suyana"
Con.ConnectionTimeout = 1500000000
Response.Expires = 0
Server.ScriptTimeout = 20000
Con.CommandTimeout = 0 '(especialmente este, en 0=ilimitado )'
Con.Open

%>
