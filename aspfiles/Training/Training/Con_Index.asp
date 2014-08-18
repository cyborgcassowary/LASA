<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Training Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%Access=Request.Cookies("Portal")("TN")%>

<div align="center">
<b><font size="3">Pending Courses/Conferences</font></b><br><br>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="border: 2px outset #0000FF; padding: 0">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Close</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Course Name</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Attendee</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Department</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Course Dates</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Training Category</b></td>
		</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select ID,UserName,Department,C_Name,C_Date,Category from tbl_users Where complete='No';")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="text-align: center"><%If Instr(Access,"W") then GetLink%></td>
			<td style="text-align: center"><a href="Con_Detail.asp?ID=<%=rstemp(0)%>&RP=Con_Index.asp"><%=rstemp(3)%></a>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(4)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(5)%>&nbsp;</td>
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

<%sub GetLink%>
<a href="Train_Close.asp?ID=<%=rstemp(0)%>">
			<img border="0" src="../images/dir_txt.gif" width="16" height="16"></a>
<%end Sub%>
</body>

</html>