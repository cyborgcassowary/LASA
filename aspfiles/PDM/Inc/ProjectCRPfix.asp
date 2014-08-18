<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID from pdm.dbo.tbl_crp where project='';")
%>
<%do until rstemp.eof %>
<%=rstemp(0)%>
<%Project=conntemp.execute("Select Project from pdm.dbo.tbl_project where PID='" & rstemp(0) & "';")%>
<%=Project(0)%>

<%Conntemp.Execute("Update pdm.dbo.tbl_crp Set Project='" & Project(0) & "' Where PID=" & rstemp(0) & ";")%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</body>

</html>
