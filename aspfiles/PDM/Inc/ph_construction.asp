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
set rstemp=conntemp.execute("Select * from ph_construction where PID='" & Session("ProjectID") & "';")
%>
<%do until rstemp.eof %>
<%ACP=rstemp(2)%>
<%IF ACP="1/1/1990" then ACP=""%>
<%CSR=rstemp(3)%>
<%IF CSR="1/1/1990" then CSR=""%>
<%CPN=rstemp(4)%>
<%CPD=rstemp(5)%>
<%IF CPD="1/1/1990" then CPD=""%>
<%ROWR=rstemp(6)%>
<%IF ROWR="1/1/1990" then ROWR=""%>
<%SDA=rstemp(7)%>
<%IF SDA="1/1/1990" then SDA=""%>
<%Copies=rstemp(8)%>
<%CopyDate=rstemp(9)%>
<%If CopyDate="1/1/1990" then CopyDate=""%>
<%HOP_APP=rstemp(10)%>
<%HOP_Permit=rstemp(11)%>
<%HOP_EXP=rstemp(12)%>
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
