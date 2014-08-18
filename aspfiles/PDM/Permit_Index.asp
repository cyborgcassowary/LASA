<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Open Permit Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<div align="center">
<b><font size="3">Open/Closed Permits</font></b>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Project Name</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Open Permits</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Closed Permits</b></td>
		</tr>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select Distinct(PID) from tbl_open;")
%>
<%do until rstemp1.eof %>
<%PID=rstemp1(0)%>
<%Set Permits = Conntemp1.Execute("Select Count(ID) From tbl_open where PID='" & PID & "';")%>
<%Set closed = Conntemp1.Execute("Select Count(ID) From tbl_permit where PID='" & PID & "';")%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project from tbl_project where PID = " & PID & ";")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td><a href="Project_add_3.asp?PID=<%=PID%>"><%=rstemp(0)%></a>&nbsp;</td>
			<td style="text-align: center"><a href="Permit_Open.asp?PID=<%=PID%>"><%=Permits(0)%></a>&nbsp;</td>
			<td style="text-align: center"><%=closed(0)%>&nbsp;</td>
		</tr>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
	</table>
</div>

</body>

</html>