<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 3</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from ph_board where PID='" & Session("ProjectID") & "';")
%>
<%do until rstemp.eof %>
<%CEA=rstemp(2)%>
<%IF CEA="1/1/1990" then CEA=""%>
<%FSA=rstemp(3)%>
<%FSR=rstemp(4)%>
<%IF FSR="1/1/1990" then FSR=""%>
<%FSE=rstemp(5)%>
<%IF FSE="1/1/1990" then FSE=""%>
<%BA=rstemp(6)%>
<%IF BA="1/1/1990" then BA=""%>
<%AC=rstemp(7)%>
<%IF AC="1/1/1990" then AC=""%>
<%ROP=rstemp(8)%>
<%IF ROP="1/1/1990" then ROP=""%>
<%ACT=rstemp(9)%>
<%IF ACT="1/1/1990" then ACT=""%>
<%DEP=rstemp(10)%>
<%IF DEP="1/1/1990" then DEP=""%>
<%BAD=rstemp(11)%>
<%IF BAD="1/1/1990" then BAD=""%>
<%EFR=rstemp(12)%>
<%EFRC=rstemp(13)%>
<%IF EFRC="1/1/1990" then EFRC=""%>
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
