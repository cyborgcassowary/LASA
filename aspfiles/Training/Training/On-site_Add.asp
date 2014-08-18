<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>On-site Training</title>
<!-- #include file="inc/calctl.inc" -->
<script language="JavaScript" src="inc/overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="inc/Train_Check.inc"-->
<!--#include file="inc/Write_Check.inc"-->
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<%If Request.form("B1")="Submit" and Request.Form("C_Date") <> "MM/DD/YYYY" then Insertdata%>

<%If Request.form("B1")="Submit" and Request.form("C_Date")="MM/DD/YYYY" then ErrorMsg%>
<form method="POST" action="On-site_Add.asp" name="onsite" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="font-family: Tahoma; font-size: 8pt; border: 2px outset #0000FF">
		<tr>
			<td style="text-align: right"><b>Submission By:</b></td>
			<td><%=Request.Cookies("Portal")("FullName")%>,&nbsp;<%=Request.Cookies("Portal")("Jobtitle")%></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Subject:</b></td>
			<td><input type="text" name="Subject" size="60"></td>
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
<option value="<%=rstemp(0)%>"><%=rstemp(0)%></option>
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
			<td style="text-align: right"><b>Location</b></td>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td width="15">
						<input type="radio" value="On-Site" name="Location" checked></td>
						<td>On-Site</td>
						<td width="15">
						<input type="radio" value="Off-Site" name="Location"></td>
						<td>Off-Site</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Date:</b></td>
			<td><input type="text" name="C_Date" size="10" value="MM/DD/YYYY" readonly>
<a href="javascript:show_calendar('onsite.C_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Hours:</b></td>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4">
					<tr>
						<td width="42"><select size="1" name="Hours">
<%For I = 00 to 40%>
<option><%=I%></option>
<%Next%>
			</select></td>
						<td style="font-family: Tahoma; font-size: 8pt; text-align: center" width="51">
						<b>Minutes</b></td>
						<td><select size="1" name="Minutes">
						<option>.00</option>
						<option>.25</option>
						<option>.50</option>
						<option>.75</option>
						</select></td>
					</tr>
				</table>
			</div>
			</td>
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
set rstemp=conntemp.execute("Select FullName from users order by Lastname ASC;")
%>
<%do until rstemp.eof %>
<%If X=0 then Response.write"<tr>"%>
<%X=X+1%>
<td><input type="checkbox" value="<%=rstemp(0)%>" name="Attendees"></td><td><%=rstemp(0)%></td>
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
			<td colspan="2">
<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></td>
		</tr>
	</table>
</div>
</form>
<%Sub InsertData%>
<%Subject=Request.form("Subject")%>
<%Subject=Replace(Subject,"'","&#039")%>
<%Category=Request.form("Category")%>
<%C_Date=Request.form("C_Date")%>
<%Hours=Request.form("Hours")%>
<%Minutes=Request.form("Minutes")%>
<%C_Hours=Hours & Minutes%>
<%Attendees=Request.form("Attendees")%>
<%SubmitBy=Request.Cookies("Portal")("FullName")%>
<%Location=Request.form("Location")%>
<%Emp=Split(Attendees,",")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=Training;"%>
<%set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=SiteSettings;"%>
<%For I = 0 to Ubound(EMP)%>
<%Set Dept=conntemp.execute("Select Dept, TrainName from users where FullName='" & Trim(Emp(I)) & "';")%>
<%RS = Conn.Execute("Insert Into tbl_training (Employee,Dept,Category,Location,Hours,Description,Dates,Submitby) Values ('" & Dept(1) & "','" & Dept(0) & "','" & Category & "','" & Location & "','" & C_Hours & "','" & Subject & "','" & C_Date & "','" & SubmitBy & "')")%>
<%Next%>
<%
set Dept=nothing
set RS=nothing
conn.close
conntemp.close
Set conn=nothing
set conntemp=nothing
%>

<%End Sub%>

<%Sub ErrorMsg%>
<% 
    msg = "Invalid Date Entered, Please Correct." 
    Response.Write("<" & "script>alert('" & msg & "');") 
    Response.Write("<" & "/script>") 
%>
<%End Sub%>


</body>

</html>