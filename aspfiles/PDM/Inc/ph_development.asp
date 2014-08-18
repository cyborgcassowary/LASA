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
set rstemp=conntemp.execute("Select * from ph_development where PID='" & Session("ProjectID") & "';")
%>
<%do until rstemp.eof %>
<%DRF=rstemp(2)%>
<%IF DRF="1/1/1990" then DRF=""%>
<%FPA=rstemp(3)%>
<%IF FPA="1/1/1990" then FPA=""%>
<%PMA=rstemp(4)%>
<%IF PMA="1/1/1990" then PMA=""%>
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
