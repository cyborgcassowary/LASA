<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Close Formal Training</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="inc/Train_Check.inc"-->
<!--#include file="inc/Write_Check.inc"-->
<%ID=Request.QueryString("ID")%>
<%If Request.form("B1")="Accept" then UpdateData%>
<%If Request.form("B1")="Close" then CloseCourse%>
<%If Request.form("B1")="Cancel" then Cancel%>
<form method="POST" action="Train_Close.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="600" id="table1" style="border: 2px outset #0000FF; font-size:8pt">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select * from tbl_users where ID=" & ID & ";")
%>
<%do until rstemp.eof %>
<%EndDate=DateAdd("d",rstemp(22),rstemp(6))%>
<%If Days <> 1 then Attend_Date=rstemp(6) & " - " & Enddate else Attend_Date=rstemp(6)%>
		<tr>
			<td colspan="4">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2" style="font-size: 8pt">
					<tr>
						<td width="146" style="text-align: right"><b>Employee:</b></td>
						<td>&nbsp;<%=rstemp(1)%>&nbsp;</td>
						<td width="50"><b>Department:</b></td>
						<td>&nbsp;<%=rstemp(2)%>&nbsp;</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Category:</b></td>
			<td width="450" colspan="3">&nbsp;<%=rstemp(16)%></td>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Course Name:</b></td>
			<td width="450" colspan="3">&nbsp;<%=rstemp(3)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Course Sponsor:</b></td>
			<td width="450" colspan="3">&nbsp;<%=rstemp(4)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Course Location:</b></td>
			<td width="192">&nbsp;<%=rstemp(5)%>&nbsp;</td>
			<td width="65" style="text-align: right"><b>City/State:</b></td>
			<td width="186">&nbsp;<%=rstemp(19)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Dates Attended:</b></td>
			<td width="450" colspan="3">&nbsp;<%=Attend_Date%></td>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Approved Program ID:</b></td>
			<td width="450" colspan="3">&nbsp;<%=rstemp(25)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Contact Credit:</b></td>
			<td width="450" colspan="3">&nbsp;<%=rstemp(24)%>&nbsp;</td>
			<%If rstemp(24)="Yes" then Cert="Y" else Cert="N"%>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Contact Hours:</b></td>
			<td width="450" colspan="3">&nbsp;<%=rstemp(26)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="146" style="text-align: right"><b>Attendance Hours:</b></td>
			<td width="450" colspan="3">&nbsp;<%=rstemp(17)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="146">&nbsp;</td>
			<td width="450" colspan="3">&nbsp;</td>
		</tr>
		<tr>
			<td width="596" colspan="4"><b>Please choose one of the following 
			options:</b></td>
		</tr>
		<tr>
			<td width="146">&nbsp;</td>
			<td width="450" colspan="3">&nbsp;</td>
		</tr>
				<tr>
			<td width="596" colspan="4">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3" style="font-size: 8pt">
					<tr>
						<td style="text-align: center"><b>Completed Course</b></td>
						<td style="text-align: center"><b>Did not Attend</b></td>
						<td style="text-align: center"><b>Cancel making changes</b></td>
					</tr>
					<tr>
						<td style="text-align: center">
						<input type="submit" value="Accept" name="B1"></td>
						<td style="text-align: center">
						<input type="submit" value="Close" name="B1"></td>
						<td style="text-align: center">
						<input type="submit" value="Cancel" name="B1"></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>

		<tr>
			<td width="596" colspan="4">&nbsp;</td>
		</tr>

	</table>
</div>
<input type="hidden" name="ID" value="<%=ID%>">
<input type="hidden" name="Employee" value="<%=rstemp(1)%>">
<input type="hidden" name="Dept" value="<%=rstemp(2)%>">
<input type="hidden" name="Category" value="<%=rstemp(16)%>">
<input type="hidden" name="Location" value="Off-Site">
<input type="hidden" name="Hours" value="<%=rstemp(17)%>">
<input type="hidden" name="Credits" value="<%=rstemp(26)%>">
<input type="hidden" name="Cert" value="<%=cert%>">
<input type="hidden" name="Description" value="<%=rstemp(3)%>">
<input type="hidden" name="dates" value="<%=rstemp(6)%>">
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>		


</form>

<%Sub Updatedata%>
<%ID=Request.form("ID")%>
<%Employee=Request.form("employee")%>
<%Dept=Request.form("Dept")%>
<%Category=Request.form("Category")%>
<%Location=Request.form("Location")%>
<%Hours=Request.form("Hours")%>
<%If Hours="" then Hours="0"%>
<%Credits=Request.form("Credits")%>
<%If Credits="" then Credits=0%>
<%Cert=Request.form("Cert")%>
<%Description=Request.form("Description")%>
<%Dates=Request.form("Dates")%>
<%submitBy=Request.Cookies("Portal")("FullName")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Sitesettings;"
set rstemp=conntemp.execute("Select TrainName from Users where Fullname='" & employee & "';")
%>
<%TrainName=rstemp(0)%>
<%
rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=training;"%>
<%RS = Conn.Execute("Update tbl_users set Complete='Yes' Where ID=" & ID & ";")%>
<%Upd = Conn.Execute("Insert Into tbl_training (Employee,Dept,Category,Location,Hours,Credits,Cert,Description,Dates,SubmitBy) Values ('" & TrainName & "','" & Dept & "','" & Category & "','" & Location & "','" & Hours & "','" & Credits & "','" & Cert & "','" & Description & "','" & Dates & "','" & Submitby & "')")%>
<%
set RS=nothing
set Upd=Nothing
conn.close
Set conn=nothing
%>
<p align="center">Item has been Updated.</p>
<form method="POST" action="Con_Index.asp" name="Update">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 2000); 
    	</script>
<%Response.end%>
<%End sub%>

<%Sub CloseCourse%>
<%ID=Request.form("ID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=training;"%>
<%RS = Conn.Execute("Update tbl_users set Complete='Did Not Attend' Where ID=" & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">Item has been Closed.</p>
<form method="POST" action="Con_Index.asp" name="Update">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 2000); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<form method="POST" action="Con_Index.asp" name="Update">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>
