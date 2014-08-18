<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Course Details</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%Desc=Request.QueryString("desc")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=training;"
set rstemp=conntemp.execute("Select Description from tbl_training where ID=" & DESC & ";")
%>
<%do until rstemp.eof %>
<%Desc=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<div align="center">
<caption><b><font size="3"><%=desc%></font></b></caption>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Employee</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Total Hours</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Date</b></td>
		</tr>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select employee,Hours,Dates from tbl_training where Description='" & desc & "';")
%>

<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td>&nbsp;<%=rstemp(0)%></td>
			<td style="text-align: center"><%If rstemp(1) <> "" then response.write rstemp(1) else response.write "N/A"%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
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
</div>
<p><a href="Course_Index.asp">Back</a></p>
</body>

</html>