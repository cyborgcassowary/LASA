<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Assignment Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<div align="center">
<b><font size="3">Open Assignments</font></b>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px" colspan="4"><b>Assigned To</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Project/Assignment</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Status</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Last Update</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Due Date</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Priority</b></td>
		</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select ID,Assignee,PID,LastUpdate,DueDate,Status,Title,Priority from tbl_assign where C_Date = '1/1/1990' order by DueDate ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<%Today=CDate(formatDateTime(now,vbshortdate))%>
<%If ABS(Int(DateDiff("d",Today,CDate(rstemp(4))))) <= 3 then fntcolor="yellow" else fntcolor=""%>
<%if Today > CDate(rstemp(4)) then fntcolor="red" else fntcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="text-align: center"><a href="Assign_edit.asp?ID=<%=rstemp(0)%>"><img border="0" src="../images/dir_txt.gif" width="16" height="16"></a></td>
			<td style="text-align: center"><a href="Assign_Update.asp?AID=<%=rstemp(0)%>">
			<img border="0" src="../images/dir_htm.gif" width="16" height="16"></a></td>
			<td style="text-align: center">
			<img border="0" src="../images/dir_delete.gif" width="16" height="16"></td>
			<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
			<td><%=rstemp(2)%><br><%=rstemp(6)%></td>
			<td style="text-align: center"><%=rstemp(5)%></td>
			<td style="text-align: center"><%=rstemp(3)%></td>
			<td style="text-align: center"><font color="<%=fntcolor%>"><b><%=rstemp(4)%></b></font></td>
			<td style="text-align: center"><%=rstemp(7)%></td>
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
</body>

</html>