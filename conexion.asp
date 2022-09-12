<%
public conn 
Set conn = Server.CreateObject("ADODB.Connection") 
conn.CommandTimeout = 0
conn.Open("Provider=SQLOLEDB; User ID=sa;Password=Mostwanted123;data source=localhost;Initial Catalog=dbmas") 
'conn.Open("Provider=SQLOLEDB; User ID=uAntof2;Password=mutu4ls3g*2019;data source=192.168.10.11;Initial Catalog=dbmas") 
'"Provider=SQLOLEDB; User ID=sa;Password=SCL.2013.2013;data source=.\SQLEXPRESS;Initial Catalog=dbmas"

%>