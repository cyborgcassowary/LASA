<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add New Assignment</title>
<!-- #include file="calctl.inc" -->
<script language="JavaScript" src="overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%AID=Request.querystring("AID")%>

<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<%If Request.form("B1")="Submit" then InsertData%>
<%If Request.form("formerror")="inputerror" then Assignment=Request.form("assignment"):Details=Request.form("details")%>
<%If Request.form("formerror")="inputerror" then duedate=Request.form("Duedate") else Duedate="MM/DD/YYYY"%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Assignee,PID,Title,Assignment,DueDate from tbl_Assign where ID=" & AID & ";")
%>
<%do until rstemp.eof %>
<%Assignee=rstemp(0)%>
<%PID=rstemp(1)%>
<%Assignment=rstemp(2)%>
<%Details=rstemp(3)%>
<%DueDate=rstemp(4)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<form method="POST" action="Assign_Update.asp" name="form1" webbot-action="--WEBBOT-SELF--">

<div align="center">
<br>
<b><font size="3">Edit Assignment
	</font></b>
	<br>
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: right"><b>Assignee:</b></td>
			<td>
			<Select name="Assignee" size="1">
<option>Select Assignee</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select assignee from tbl_users order by assignee ASC;")
%>
<%do until rstemp.eof %>
<%If Assignee=rstemp(0) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(0)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Project:</b></td>
			<td>
			<Select name="Project" size="1">
<option>Select Project</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project from tbl_Project order by Project ASC;")
%>
<%do until rstemp.eof %>
<%If PID=rstemp(0) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(0)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Due Date:</b></td>
			<td><input type="text" name="duedate" value="<%=DueDate%>" size="20">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('form1.duedate');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Priority:</b></td>
			<td>
			<select size="1" name="Priority">
			<option>Low</option>
			<option selected>Normal</option>
			<option>High</option>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Assignment:</b></td>
			<td>
			<input type="text" name="Assignment" size="66" value="<%=assignment%>"></td>
		</tr>
		<tr>
			<td colspan="2">
			<p style="text-align: center"><b>Assignment Details</b></td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: center"><textarea rows="4" name="Details" cols="80"><%=Details%></textarea></td>
		</tr>
		<tr>
			<td colspan="2">
<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
<input type="hidden" name="AID" value="<%=AID%>">
</td>
		</tr>
	</table>
</div>
</form>

<%Sub InsertData%>
<%ID=Request.form("AID")%>
<%Assignee=Request.form("assignee")%>
<%Project=Request.form("Project")%>
<%DueDate=Request.form("DueDate")%>
<%A_Date=FormatDatetime(Now,VbshortDate)%>
<%Assignment=Request.form("Assignment")%>
<%Assignment=Replace(Assignment,"'","&#039")%>
<%Details=Request.form("details")%>
<%Details=Replace(Details,"'","&#039")%>
<%Status="Updated"%>
<%C_Date="1/1/1990"%>
<%Priority=Request.form("Priority")%>
<%If Assignee="Select Assignee" or Project="Select Project" then InputError%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_assign Set Assignee='" & Assignee & "', PID='" & Project & "', DueDate='" & DueDate & "', Status='" & Status & "', Assignment='" & Details & "', Title='" & Assignment & "' Where ID=" & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">Assignment has been posted</p>
<form method="POST" action="assign_index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>
<%Sub InputError%>
<p align="center">Input Error, Please Correct Errors</p>
<%Assignment=Request.form("Assignment")%>
<%Details=Request.form("details")%>
<%duedate=Request.form("duedate")%>
<form method="POST" action="Assign_Update.asp" name="Update" target="ibody">
<input type="hidden" name="Assignment" value="<%=assignment%>">
<input type="hidden" name="details" value="<%=Details%>">
<input type="hidden" name="duedate" value="<%=duedate%>">
<input type="hidden" name="formerror" value="inputerror">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%end Sub%>
</body>

</html>
