<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Training History</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%GMonth=Request.form("GMonth")%>
<%GYear=Request.form("GYear")%>
<%If GMonth="" and GYear="" or Request.Form("B1")="View All" then CView="All Records" else CView=MonthName(GMonth,True) & "-" & GYear%>
<div align="center">
<form method="POST" action="Train_History.asp" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1">
		<tr>
			<td><b><font size="3">Training History - <%=CView%></font></b></td>
			<td style="text-align: right">

<select name="GMonth" size="1">
<%If GMonth="1" then sel="selected" else sel=""%>
<option <%=sel%> value="1">January</option>
<%If GMonth="2" then sel="selected" else sel=""%>
<option <%=sel%> value="2">February</option>
<%If GMonth="3" then sel="selected" else sel=""%>
<option <%=sel%> value="3">March</option>
<%If GMonth="4" then sel="selected" else sel=""%>
<option <%=sel%> value="4">April</option>
<%If GMonth="5" then sel="selected" else sel=""%>
<option <%=sel%> value="5">May</option>
<%If GMonth="6" then sel="selected" else sel=""%>
<option <%=sel%> value="6">June</option>
<%If GMonth="7" then sel="selected" else sel=""%>
<option <%=sel%> value="7">July</option>
<%If GMonth="8" then sel="selected" else sel=""%>
<option <%=sel%> value="8">August</option>
<%If GMonth="9" then sel="selected" else sel=""%>
<option <%=sel%> value="9">September</option>
<%If GMonth="10" then sel="selected" else sel=""%>
<option <%=sel%> value="10">October</option>
<%If GMonth="11" then sel="selected" else sel=""%>
<option <%=sel%> value="11">November</option>
<%If GMonth="12" then sel="selected" else sel=""%>
<option <%=sel%> value="12">December</option>
</select>
<select size="1" name="GYear">
<%For Y=Year(now) to 1990 step -1%>
<%If Int(GYear)=Int(Y) then sel="selected" else sel=""%>
<option <%=sel%>><%=Y%></option>
<%Next%>
</select>
<input type="submit" value="Submit" name="B1" style="font-size: 8pt; font-family: Tahoma"><input type="submit" value="View All" name="B1" style="font-family: Tahoma; font-size: 8pt">
</td>
		</tr>
	</table>
</form>
</div>




<%If Request.form("B1")="Submit" then Monthly%>
<%If Request.form("B1")="View All" or Request.form("B1")="" then ViewAll%>
<%If Request.form("B1")="Print" then Print%>

<%Sub ViewAll%>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Employee</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b># of Courses</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>OnSite</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>OffSite</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b># of Hours</b></td>
		</tr>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=training;"
set rstemp=conntemp.execute("Select Distinct Employee from tbl_training order by Employee ASC;")
%>

<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>

<%Set Courses=conntemp.execute("Select Count(ID) from tbl_Training where employee='" & rstemp(0) & "';")%>
<%Set Hours=conntemp.execute("Select Sum(Hours) from tbl_Training where employee='" & rstemp(0) & "';")%>
<%Set Onsite=conntemp.execute("Select Count(ID) from tbl_Training where Location='On-Site' and employee='" & rstemp(0) & "';")%>
<%Set Offsite=conntemp.execute("Select Count(ID) from tbl_Training where Location='Off-Site' and employee='" & rstemp(0) & "';")%>
		<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td>&nbsp;<a href="Employee_Detail.asp?Emp=<%=rstemp(0)%>"><%=rstemp(0)%></a></td>
			<td align="center"><%=Courses(0)%></td>
			<td align="center"><%=OnSite(0)%></td>
			<td align="center"><%=OffSite(0)%></td>
			<td align="center"><%=Hours(0)%></td>
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
%>
<%Set TotalCourses=conntemp.execute("Select Count(ID) from tbl_Training;")%>
<%Set TotalHours=conntemp.execute("Select Sum(Hours) from tbl_Training;")%>
<%Set TotalOnsite=conntemp.execute("Select Count(ID) from tbl_Training where Location='On-Site';")%>
<%Set TotalOffsite=conntemp.execute("Select Count(ID) from tbl_Training where Location='Off-Site';")%>
<tr>
<td align="right" style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px"><b>Totals</b>&nbsp;</td>
<td align="center" style="border-top-style: double; border-top-width: 3px"><%=Totalcourses(0)%>&nbsp;</td>
<td align="center" style="border-top-style: double; border-top-width: 3px"><%=TotalOnsite(0)%>&nbsp;</td>
<td align="center" style="border-top-style: double; border-top-width: 3px"><%=TotalOffSite(0)%>&nbsp;</td>
<td align="center" style="border-top-style: double; border-top-width: 3px"><%=TotalHours(0)%>&nbsp;</td>
</tr>

<%
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

	</table>
</div>
<!--#include file="inc/print.inc"-->
<%End Sub%>
<%sub Monthly%>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Employee</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b># of Courses</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>OnSite</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>OffSite</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b># of Hours</b></td>
		</tr>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=training;"
set rstemp=conntemp.execute("Select Distinct Employee from tbl_training where Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "' order by Employee ASC;")
%>

<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>

<%Set Courses=conntemp.execute("Select Count(ID) from tbl_Training where employee='" & rstemp(0) & "' and Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "';")%>
<%Set Hours=conntemp.execute("Select Sum(Hours) from tbl_Training where employee='" & rstemp(0) & "' and Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "';")%>
<%Set Onsite=conntemp.execute("Select Count(ID) from tbl_Training where Location='On-Site' and employee='" & rstemp(0) & "' and Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "';")%>
<%Set Offsite=conntemp.execute("Select Count(ID) from tbl_Training where Location='Off-Site' and employee='" & rstemp(0) & "' and Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "';")%>
		<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td>&nbsp;<a href="Employee_Detail.asp?Emp=<%=rstemp(0)%>"><%=rstemp(0)%></a></td>
			<td align="center"><%=Courses(0)%></td>
			<td align="center"><%=OnSite(0)%></td>
			<td align="center"><%=OffSite(0)%></td>
			<td align="center"><%=Hours(0)%></td>
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
%>
<%Set TotalCourses=conntemp.execute("Select Count(ID) from tbl_Training where Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "';")%>
<%Set TotalHours=conntemp.execute("Select Sum(Hours) from tbl_Training where Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "';")%>
<%Set TotalOnsite=conntemp.execute("Select Count(ID) from tbl_Training where Location='On-Site' and Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "';")%>
<%Set TotalOffsite=conntemp.execute("Select Count(ID) from tbl_Training where Location='Off-Site' and Month(Dates)='" & GMonth & "' and Year(Dates)='" & GYear & "';")%>
<tr>
<td align="right" style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px"><b>Totals</b>&nbsp;</td>
<td align="center" style="border-top-style: double; border-top-width: 3px"><%=Totalcourses(0)%>&nbsp;</td>
<td align="center" style="border-top-style: double; border-top-width: 3px"><%=TotalOnsite(0)%>&nbsp;</td>
<td align="center" style="border-top-style: double; border-top-width: 3px"><%=TotalOffSite(0)%>&nbsp;</td>
<td align="center" style="border-top-style: double; border-top-width: 3px"><%=TotalHours(0)%>&nbsp;</td>
</tr>

<%
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

	</table>
</div>
<!--#include file="inc/print.inc"-->
<%Response.end%>
<%end sub%>
</body>

</html>