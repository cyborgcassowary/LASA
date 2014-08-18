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
set rstemp=conntemp.execute("Select * from ph_indemn where PID='" & Session("ProjectID") & "';")
%>
<%do until rstemp.eof %>
<%TOI=rstemp(2)%>
<%IAR=rstemp(3)%>
<%IF IAR="1/1/1990" then IAR=""%>
<%ICR=rstemp(4)%>
<%IF ICR="1/1/1990" then ICR=""%>
<%IED=rstemp(5)%>
<%IF IED="1/1/1990" then IED=""%>
<%SFF=rstemp(6)%>
<%DODD=rstemp(7)%>
<%IF DODD="1/1/1990" then DODD=""%>
<%ABR=rstemp(8)%>
<%IF ABR="1/1/1990" then ABR=""%>
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
