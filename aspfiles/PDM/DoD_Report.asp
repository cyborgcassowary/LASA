<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" style="border: 2px outset #0000FF">
	<tr>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Project&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Deed of Dedication Date&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Months&nbsp;Open</b></td>
		<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<p style="text-align: center"><b>Months Remain/Overdue</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Status&nbsp;</b></td>
	</tr>



<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID,DODD from ph_indemn where DODD <> '1/1/1990';")
%>
<%do until rstemp.eof %>
<%Project=conntemp.execute("Select Project from tbl_Project where PID=" & rstemp(0) & ";")%>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
		<td><a href="Project_Add_3.asp?PID=<%=rstemp(0)%>" target="ibody"><%=project(0)%></a>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
		<%DMonths=DateDiff("m",rstemp(1),FormatDateTime(now,vbshortdate))%>
		<td style="text-align: center"><%=DMonths%>&nbsp;</td>
		<td style="text-align: center"><%If DMonths => 18 then OMonths=DMonths-18 else OMonths=0%><%=ABS(DMonths-18)%></td>
		<td style="text-align: center"><%If DateDiff("m",rstemp(1),formatDateTime(now,vbshortdate)) => 18 then response.write "<font color=red>Overdue</font>" else Response.write "Open"%>&nbsp;</td>
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