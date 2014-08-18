<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Course History Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<div align="center">
	<table border="0" cellpadding="0" cellspacing="1" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px" colspan="2"><b>Course Name</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Total Attendees</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Total Hours</b></td>
		</tr>

<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=Training;"
set category=conntemp1.execute("Select Distinct Category from tbl_Training;")
%>
<%do until category.eof %>
<%CatCount=conntemp1.execute("Select Count(Description) as CatCount from (Select Distinct Description from tbl_training where Category='" & Category(0) & "');")%>
<%TotalHours=conntemp1.execute("Select Sum(Hours) from tbl_training where Category='" & category(0) & "';")%>
<tr><td colspan="2" style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><b><%=Category(0)%> - Total Courses <%=CatCount(0)%></b></td>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">&nbsp;</td>
<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><%=TotalHours(0)%></td>
</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select Distinct Description from tbl_training where Category='" & Category(0) & "';")
%>

<%do until rstemp.eof %>
<%Attendees=conntemp.execute("Select Count(employee) from tbl_Training where Description='" & rstemp(0) & "';")%>
<%Hours=conntemp.execute("Select sum(Hours) from tbl_training where Description='" & rstemp(0) & "' and Category='" & Category(0) & "';")%>
<%ID=conntemp.execute("Select ID from tbl_training where Description='" & rstemp(0) & "';")%>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>

<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td width="15">&nbsp;</td>
			<td width="500">&nbsp;<a href="Course_Detail.asp?Desc=<%=ID(0)%>"><%=rstemp(0)%></a></td>
			<td style="text-align: center"><%=Attendees(0)%>&nbsp;</td>
			<td style="text-align: center"><%If Hours(0) <> "" then response.write Hours(0) else response.write "N/A"%>&nbsp;</td>
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
category.movenext
loop

category.close
set category=nothing
conntemp1.close
set conntemp1=nothing
%>
	</table>
</div>
</body>

</html>