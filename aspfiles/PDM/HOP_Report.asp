<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>HOP Report</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0" style="border: 2px outset #0000FF">
	<tr>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Project</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>HOP Application #</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>HOP Permit #</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>HOP Expiration</b></td>
	</tr>



<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID,HOP_APP,HOP_Permit,HOP_EXP from ph_construction where HOP_APP <> '';")%>
<%set Project=conntemp.execute("Select Project from tbl_project where PID = " & rstemp(0) & ";")%>
<%do until rstemp.eof %>
	<tr>
		<td><%=Project(0)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(3)%>&nbsp;</td>
	</tr>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</table>
</body>

</html>