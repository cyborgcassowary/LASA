<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Get Capacity Values</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from ph_capacity where PID='" & Session("ProjectID") & "';")
%>
<%do until rstemp.eof %>
<%SRF=rstemp(2)%>
<%IF SRF="1/1/1990" then SRF=""%>
<%LOCA=rstemp(3)%>
<%IF LOCA="1/1/1990" then LOCA=""%>
<%ROC=rstemp(4)%>
<%IF ROC="1/1/1990" then ROC=""%>
<%RFA=rstemp(5)%>
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
