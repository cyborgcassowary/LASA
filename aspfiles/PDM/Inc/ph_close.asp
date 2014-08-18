<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 2</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from ph_close where PID='" & Session("ProjectID") & "';")
%>
<%do until rstemp.eof %>
<%FI=rstemp(2)%>
<%IF FI="1/1/1990" then FI=""%>
<%ER=rstemp(3)%>
<%IF ER="1/1/1990" then ER=""%>
<%PCC=rstemp(4)%>
<%IF PCC="1/1/1990" then PCC=""%>
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
