<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>On-Site Training Edit</title>
<!-- #include file="inc/calctl.inc" -->
<script language="JavaScript" src="inc/overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="Submit" then UpdateData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<%ID=Request.QueryString("ID")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select * from OnSite where ID=" & ID & ";")
%>
<%do until rstemp.eof %>
<%Subject=rstemp(1)%>
<%C_Date=rstemp(2)%>
<%C_Hours=rstemp(3)%>
<%Attendees=rstemp(4)%>
<%Description=rstemp(5)%>
<%Category=rstemp(6)%>

<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<form method="POST" action="On-site_Edit.asp" name="onsite" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="font-family: Tahoma; font-size: 8pt; border: 2px outset #0000FF">
		<tr>
			<td style="text-align: right"><b>Submission By:</b></td>
			<td><%=Request.Cookies("Portal")("FullName")%>,&nbsp;<%=Request.Cookies("Portal")("Jobtitle")%></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Subject:</b></td>
			<td>
			<input type="text" name="Subject" size="40" value="<%=Subject%>"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Category:</b></td>
			<td>
<Select name="Category" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select Category from tbl_Category order by Category ASC;")
%>
<%do until rstemp.eof %>
<%If Category=rstemp(0) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=Sel%>><%=rstemp(0)%></option>
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
			<td style="text-align: right"><b>Date:</b></td>
			<td><input type="text" name="c_date" size="10" value="<%=C_Date%>">
<a href="javascript:show_calendar('onsite.c_date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Hours:</b></td>
			<td>
			<input type="text" name="C_Hours" size="10" value="<%=C_Hours%>">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: center"><b>Attendees</b></td>
		</tr>
		<tr>
			<td colspan="2">
<table style="font-family: Tahoma; font-size: 8pt">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=SiteSettings;"
set rstemp=conntemp.execute("Select Fullname from users order by Lastname ASC;")
%>
<%do until rstemp.eof %>
<%If X=0 then Response.write"<tr>"%>
<%X=X+1%>
<%If Instr(Attendees,rstemp(0)) then chk="checked" else chk=""%>
<td><input type="checkbox" value="<%=rstemp(0)%>" <%=chk%> name="Attendees"></td><td><%=rstemp(0)%></td>
<%If X=4 then X=0:Response.write"</tr>"%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</table>
			</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: center"><b>Description</b></td>
		</tr>
		<tr>
			<td colspan="2"><textarea rows="4" name="Description" cols="80"><%=Description%></textarea></td>
		</tr>
		<tr>
			<td colspan="2">
<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>
		</tr>
	</table>
</div>
<input type="hidden" name="ID" value="<%=ID%>">
</form>
</body>

<%Sub UpdateData%>
<%ID=Request.form("ID")%>
<%Subject=Request.form("Subject")%>
<%Subject=Replace(Subject,"'","&#039")%>
<%Category=Request.form("Category")%>
<%C_Date=Request.form("C_Date")%>
<%C_Hours=Request.form("C_Hours")%>
<%Attendees=Request.form("Attendees")%>
<%SubmitBy=Request.Cookies("Portal")("FullName")%>
<%Description=Request.form("Description")%>
<%Description=Replace(Description,"'","&#039")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=Training;"%>
<%RS = Conn.Execute("Update OnSite Set Subject='" & Subject & "', C_Date='" & C_Date & "', C_Hours='" & C_Hours & "', Attendees='" & Attendees & "', Description='" & Description & "', Category='" & Category & "', SubmitBy='" & SubmitBy & "' Where ID=" & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">Record Updated.</p>
<form method="POST" action="On-Site_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 2000); 
    	</script>
<%response.end%>
<%End Sub%>

<%Sub Cancel%>
<form method="POST" action="On-Site_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 20); 
    	</script>
<%response.end%>
<%End sub%>
</html>