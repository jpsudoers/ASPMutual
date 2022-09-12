<%
public conn 
Set conn = Server.CreateObject("ADODB.Connection") 
conn.CommandTimeout = 0
conn.Open("Provider=SQLOLEDB; User ID=User_Auditor;Password=Teatinos258!;data source=192.168.10.11;Initial Catalog=BD_Auditoria")
%>