<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Employee Training Detail</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%EMP=Request.QueryString("Emp")%>

<div align="center">
<table width="100%">
<tr>
<td align="left"><!--#include file="inc/print.inc"--></td>
<td align="center"><b><font size="3">Training Detail for: <%=Emp%>
	</font></b></td>
<td align="right"><a href="Train_History.asp">Back</a></td>
</tr>
</table>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Category</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Location</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Hours</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Credits</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Cert</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Description</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Dates</b></td>
		</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select Category,Location,Hours,Credits,Cert,Description,Dates from tbl_training where employee='" & emp & "' order by Dates DESC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="font-size: 8pt"><%=rstemp(0)%>&nbsp;</td>
			<td style="text-align: center; font-size:8pt"><%=rstemp(1)%>&nbsp;</td>
			<td style="text-align: center; font-size:8pt"><%=rstemp(2)%>&nbsp;</td>
			<td style="text-align: center; font-size:8pt"><%=rstemp(3)%>&nbsp;</td>
			<td style="text-align: center; font-size:8pt"><%=rstemp(4)%>&nbsp;</td>
			<td style="font-size: 8pt"><%=rstemp(5)%>&nbsp;</td>
			<td style="text-align: center; font-size:8pt"><%=rstemp(6)%>&nbsp;</td>
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
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=training;"
set TotalHours=conntemp.execute("Select Sum(Hours) from tbl_training where Employee='" & emp & "';")
%>
<%Set TotalCredits=conntemp.execute("Select Sum(Int(Credits)) from tbl_training where Employee='" & emp & "';")%>
<%Set TotalCourses=conntemp.execute("Select Count(ID) from tbl_training where Employee='" & emp & "';")%>
<%Set TotalCerts=conntemp.execute("Select Count(ID) from tbl_training where Cert='Y' and employee='" & emp & "';")%>
<tr>
<td colspan="2" style="text-align: right; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px">
<b>Totals</b>&nbsp;</td>
<td style="text-align: center; border-top-style:double; border-top-width:3px"><%=totalHours(0)%>&nbsp;</td>
<td style="text-align: center; border-top-style:double; border-top-width:3px"><%=TotalCredits(0)%>&nbsp;</td>
<td style="text-align: center; border-top-style:double; border-top-width:3px"><%=TotalCerts(0)%>&nbsp;</td>
<td style="text-align: right; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px">
<b>Total Courses</b>&nbsp;</td>
<td style="text-align: center; border-top-style:double; border-top-width:3px"><%=TotalCourses(0)%>&nbsp;</td>
</tr>
<%
conntemp.close
set conntemp=nothing
%>
	</table>
</div>
</body>

</html>