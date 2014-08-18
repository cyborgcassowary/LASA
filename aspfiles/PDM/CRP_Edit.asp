<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Capacity Reserve Projects</title>
<!-- #include file="calctl.inc" -->
<script language="JavaScript" src="overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="Cancel" then Cancel%>
<%If Request.form("B1")="Submit" then UpdateData%>

<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<%PID=Request.QueryString("PID")%>

<form method="POST" name="frm1" action="CRP_Edit.asp" webbot-action="--WEBBOT-SELF--">

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project, Res_Type,Res_Amount from tbl_Project where PID=" & PID & ";")
%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select ID,Status from tbl_CRP where PID='" & PID & "';")
%>
<%do until rstemp1.eof %>
<%CRP_ID=rstemp1(0)%>
<%ChkStatus=rstemp1(1)%>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
<%If CInt(CRP_ID) <> 0 then GetData%>


<%If Session("Reserve_Date")="1/1/1900" then Session("Reserve_Date")=""Else%>
<%If Session("Expire_Date")="1/1/1900" then Session("Expire_Date")=""%>
<%IF Session("Renew_Date")="1/1/1900" then Session("Renew_Date")=""%>
<%IF Session("Expire_Notice")="1/1/1900" then Session("Expire_Notice")=""%>
<%If Session("Capacity_Close")="1/1/1900" then Session("Capacity_Close")=""%>

<%do until rstemp.eof %>

<div align="center">
<table border="1" id="table1" cellspacing="0" cellpadding="0">
	<tr>
		<td style="text-align: right"><b>Project Name:</b></td>
		<td><%=rstemp(0)%></td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Reserve Type:</b></td>
		<td><%=rstemp(1)%>&nbsp;</td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Reserve Amount:</b></td>
		<td><%=rstemp(2)%>&nbsp;</td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Reserve Date:</b></td>
		<td><input type="text" name="Reserve_Date" value="<%=Session("Reserve_Date")%>">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('frm1.Reserve_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Expire Date:</b></td>
		<td><input type="text" name="Expire_Date" value="<%=Session("Expire_Date")%>" >
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('frm1.Expire_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Date Renewed:</b></td>
		<td><input type="text" name="Renew_Date" value="<%=Session("Renew_Date")%>" >
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('frm1.Renew_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Expiration Notice Sent:</b></td>
		<td><input type="text" name="Expire_Notice" value="<%=Session("Expire_Notice")%>" >
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('frm1.Expire_Notice');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td style="text-align: right"><b>#<%=rstemp(1)%>'s Reserved </b> </td>
		<td><input type="text" name="Res_Amount" size="20" value="<%=Session("Res_Amount")%>">&nbsp;</td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Status:</b></td>
		<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table2">
			<tr>
				<td><%If Session("Expire_Date") <> "" then EDate=Session("Expire_Date") else EDate="1/1/1900"%>
				<%If CDate(EDate) => CDate(FormatDateTime(Now,VbshortDate)) then
						 Status="Active"
						  else
						 Status="Expired"
						  end if%>
					<%If chkStatus = "Closed" then
				 		chk="checked"
				 		Status="Closed"
				 	 else
				 	  	chk=""
				 	 End if%>
				 	 
				    <%=Status%></td>
				<td width="27">
				<input type="checkbox" name="ChkStatus" value="Closed" <%=chk%>></td>
				<td width="105">Closed</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td valign="top">
		<p style="text-align: right"><b>Notes:</b></td>
		<td><textarea rows="6" name="Notes" cols="60"><%=Session("Notes")%></textarea></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1">&nbsp;</td>
	</tr>
</table>
</div>
<input type="hidden" name="PID" value="<%=PID%>">
<Input type="hidden" name="res_Type" value="<%=rstemp(1)%>">
<input type="hidden" name="Record_Type" value="<%=Session("Record_type")%>">
<input type="hidden" name="Project" value="<%=rstemp(0)%>">
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</form>

<%Sub UpdateData%>
<%PID=Request.form("PID")%>
<%Project=Replace(Request.form("Project"),"'","&#039")%>
<%Reserve_Date=Request.form("Reserve_Date")%>
<%Expire_Date=Request.form("Expire_Date")%>
<%Renew_date=Request.form("Renew_date")%>
<%Expire_Notice=Request.form("Expire_notice")%>
<%Res_Amount=Request.form("Res_Amount")%>
<%Res_Type=Request.form("Res_type")%>
<%Capacity_Close=Request.form("Capacity_close")%>
<%If Reserve_Date="" or Reserve_Date="MM/DD/YYYY" then Reserve_Date="01/01/1900"%>
<%If Expire_Date="" or Reserve_Date="MM/DD/YYYY" then Expire_Date="01/01/1900"%>
<%If Renew_Date="" or Reserve_Date="MM/DD/YYYY" then Renew_Date="01/01/1900"%>
<%If Expire_Notice="" or Reserve_Date="MM/DD/YYYY" then Expire_Notice="01/01/1900"%>
<%If Cdate(Expire_Date) => CDate(FormatDateTime(Now,VbShortDate)) then Status="Active" else Status="Expired"%>
<%If Request.form("ChkStatus") = "Closed" then Status="Closed"%>
<%Notes=Request.form("Notes")%>
<%Notes=Replace(Notes,"'","&#039")%>

<%SQLInsert1="Insert Into tbl_CRP (PID,Reserve_Date,Expire_Date,Renew_Date,Expire_Notice,Res_Amount,Res_Type,Status,Notes,Project)"%>
<%SQLInsert=SQLInsert1 & " Values ('" & PID & "','" & Reserve_Date & "','" & Expire_Date & "','" & Renew_Date & "','" & Expire_Notice & "','" & Res_Amount & "','" & Res_type & "','" & Status & "','" & Notes & "','" & Project & "')"%>

<%SQLUpdate="Update tbl_CRP Set Reserve_Date='" & Reserve_Date & "', Expire_date='" & Expire_Date & "', renew_Date='" & Renew_Date & "', Expire_Notice='" & Expire_Notice & "', res_Amount='" & res_amount & "', Res_Type='" & Res_Type & "', Status='" & Status & "', Notes='" & Notes & "', Project='" & Project & "' Where PID='" & PID & "';"%>

<%If Session("Record_Type")="Old" then SQLStmt=SQLUpdate else SQLStmt=SQLInsert%>
<%=SQLStmt%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute(SQLStmt)%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%Session("Reserve_Date")=""%>
<%Session("Expire_Date")=""%>
<%Session("Renew_date")=""%>
<%Session("Expire_Notice")=""%>
<%Session("Res_Amount")=""%>
<%Session("Res_Type")=""%>
<%Session("Status")=""%>
<%Session("Notes")=""%>
<%session("Record_type")=""%>
<p align="center"><b>Record has been Updated!</b></p>
	<script language="JavaScript">
		setTimeout("self.close();", 5000);
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<%Session("Reserve_Date")=""%>
<%Session("Expire_Date")=""%>
<%Session("Renew_date")=""%>
<%Session("Expire_Notice")=""%>
<%Session("Res_Amount")=""%>
<%Session("Res_Type")=""%>
<%Session("Status")=""%>
<%Session("Notes")=""%>
<%session("Record_type")=""%>

	<script language="JavaScript">
		setTimeout("self.close();", 20);
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub GetData%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set CRP=conntemp1.execute("Select * from tbl_CRP where PID='" & PID & "';")
%>
<%do until CRP.eof %>
<%Session("Reserve_Date")=CRP(2)%>
<%Session("Expire_Date")=CRP(3)%>
<%Session("Renew_date")=CRP(4)%>
<%Session("Expire_Notice")=CRP(5)%>
<%Session("Res_Amount")=CRP(6)%>
<%Session("Res_Type")=CRP(7)%>
<%Session("Status")=CRP(8)%>
<%Session("Notes")=CRP(9)%>
<%
CRP.movenext
loop

CRP.close
set CRP=nothing
conntemp1.close
set conntemp1=nothing
%>
<%Session("Record_Type")="Old"%>
<%End Sub%>

</body>

</html>