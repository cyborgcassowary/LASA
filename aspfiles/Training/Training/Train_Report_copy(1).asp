<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Training Reports</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<p style="text-align: center"><b><font size="4">Training Report</font></b> </p>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Employee&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Department&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Category&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Days&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Hours&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Total Cost&nbsp;</b></td>
		</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select UserName,Department,Category,Days,H_Attend,Total_Cost,ID from tbl_users order by Department ASC;")
%>
<%do until rstemp.eof %>
<%If rstemp(5) <> "" then CourseCost=rstemp(5) else coursecost=0%>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
		<tr bgcolor="<%=bgcolor%>">
			<td><a href="Con_Detail.asp?ID=<%=rstemp(6)%>&RP=Train_report.asp"><%=rstemp(0)%></a>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(3)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(4)%>&nbsp;</td>
			<td style="text-align: center"><%=FormatCurrency(CourseCost)%>&nbsp;</td>
		</tr>
<%CATDays=CATDays+Int(rstemp(3))%>
<%CATHours=CATHours+Int(rstemp(4))%>
<%CATCost=CATCost+Int(CourseCost)%>
		<tr>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">&nbsp;</td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px">&nbsp;</td>
			<td style="text-align: right; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px"><b><%=rstemp(2)%>&nbsp;Totals:</b></td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px"><%=CATDays%>&nbsp;</td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px"><%=CATHours%>&nbsp;</td>
			<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-style:solid; border-top-width:1px; border-bottom-width:1px"><%=FormatCurrency(CourseCost)%>&nbsp;</td>
		</tr>

<%TotalDays=TotalDays+Int(rstemp(3))%>
<%TotalHours=TotalHours+Int(rstemp(4))%>
<%TotalCost=TotalCost+Int(CourseCost)%>
<%
rstemp.movenext
X=0
loop

CATDAys=0
CATHours=0
CATCost=0

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
		<tr>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">&nbsp;</td>
			<td style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">&nbsp;</td>
			<td style="text-align: right; border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">
			<b>&nbsp;Grand Totals:</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-style: double; border-top-width: 3px; border-bottom-width: 1px"><%=TotalDays%>&nbsp;</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-style: double; border-top-width: 3px; border-bottom-width: 1px"><%=TotalHours%>&nbsp;</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-style: double; border-top-width: 3px; border-bottom-width: 1px"><%=formatcurrency(TotalCost)%>&nbsp;</td>
		</tr>
	</table>
</div>

</body>

</html>