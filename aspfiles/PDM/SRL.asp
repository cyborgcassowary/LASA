<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Service Request EscrowFile</title>
<!-- #include file="calctl.inc" -->
<script language="JavaScript" src="overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<%PID=Request.QueryString("PID")%>
<%If Request.form("B1")="Cancel" then Cancel%>
<%If Request.form("B1")="Submit" then UpdateData%>


<form method="POST" name="frm1" action="SRL.asp" webbot-action="--WEBBOT-SELF--">
<table border="1" width="100%" cellspacing="1">
<caption><b>Service Request EscrowFile Worksheet</b></caption>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project,Developer,EscrowFile from tbl_Project where PID=" & PID & ";")
%>
<%Set Dev=conntemp.execute("Select D_Name,D_Number from tbl_Developer where DID=" & rstemp(1) & ";")%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select ID from tbl_SRL where PID='" & PID & "';")
%>
<%do until rstemp1.eof %>
<%SRL_ID=rstemp1(0)%>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
<%If IsEmpty(SRL_ID)=False then GetData%>
<%If Session("Open_Date")="1/1/1900" then Session("Open_Date")=""%>
<%If Session("Request_Rec")="1/1/1900" then Session("Request_rec")=""%>
<%IF Session("Review_Rec")="1/1/1900" then Session("Review_Rec")=""%>
<%do until rstemp.eof %>
	<tr>
		<td style="text-align: right"><b>EscrowFile #&nbsp;</b></td>
		<td><%=rstemp(2)%>&nbsp;</td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Date Opened&nbsp;</b></td>
		<td><input type="text" name="Open_Date" value="<%=Session("Open_Date")%>">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('frm1.Open_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Project Name&nbsp;</b></td>
		<td><%=rstemp(0)%>&nbsp;</td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Contact & Phone #&nbsp;</b></td>
		<td><%=Dev(0)%> - <%=Dev(1)%>&nbsp;</td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Comments&nbsp;</b></td>
		<td><input type="text" name="comments" size="30" value="<%=session("comments")%>">&nbsp;</td>
	</tr>
	<tr>
		<td style="text-align: right"><b>File Created?&nbsp;</b></td>
		<td>
		<table border="0" width="100%" cellspacing="1" cellpadding="0" id="table1">
			<tr>
				<td width="15">
				<%If Session("FileCreated")="Yes" then chk="checked" else chk=""%>
				<input type="radio" value="Yes" name="FileCreated" <%=chk%>></td>
				<td width="30">Yes</td>
				<td width="15">
				<%If Session("FileCreated")="No" then chk="checked" else chk=""%>
				<input type="radio" value="No" name="FileCreated" <%=chk%>></td>
				<td>No</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Service Request Received&nbsp;</b></td>
		<td><input type="text" name="Request_Rec" value="<%=Session("Request_Rec")%>">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('frm1.Request_Rec');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td style="text-align: right"><b>Development Review Received&nbsp;</b></td>
		<td><input type="text" name="Review_Rec" value="<%=Session("Review_Rec")%>">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('frm1.Review_Rec');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td style="text-align: right">
		<input type="hidden" name="EscrowFile" value="<%=rstemp(2)%>">
		<input type="hidden" name="PID" value="<%=PID%>">
		<input type="hidden" name="CID" value="<%=rstemp(1)%>">
		<input type="hidden" name="Record_Type" value="<%=Session("Record_type")%>">
		<input type="submit" value="Submit" name="B1">&nbsp;</td>
		<td><input type="submit" value="Cancel" name="B1">&nbsp;</td>
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
</form>

<%Sub UpdateData%>
<%EscrowFile=Request.form("EscrowFile")%>
<%Open_Date=Request.form("Open_Date")%>
<%PID=Request.form("PID")%>
<%CID=Request.form("CID")%>
<%comments=Request.form("comments")%>
<%FileCreated=request.form("FileCreated")%>
<%Request_rec=Request.form("Request_Rec")%>
<%Review_rec=Request.form("Review_Rec")%>
<%If Open_Date="" or Open_date="MM/DD/YYYY" then Open_Date="1/1/1900"%>
<%If Request_Rec="" or Request_Rec="MM/DD/YYYY" then Request_Rec="1/1/1900"%>
<%If Review_Rec="" or Review_Rec="MM/DD/YYYY" then Review_Rec="1/1/1900"%>
<%comments=Replace(comments,"'","&#039")%>

<%SQLUpdate="Update tbl_SRL Set Comments='" & Comments & "', FileCreated='" & FileCreated & "', Request_Rec='" & Request_Rec & "', Review_Rec='" & Review_rec & "' Where PID='" & PID & "';"%>
<%SQLInsert1="Insert Into tbl_SRL (EscrowFile,Open_date,PID,CID,comments,FileCreated,Request_Rec,Review_Rec)"%>
<%SQLInsert=SQLInsert1 & " Values ('"&EscrowFile&"','"&Open_date&"','"&PID&"','"&CID&"','"&comments&"','"&FileCreated&"','"&Request_Rec&"','"&Review_Rec&"')"%>

<%If Session("Record_Type")="Old" then SQLStmt=SQLUpdate else SQLStmt=SQLInsert%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute(SqlStmt)%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%Session("Record_Type")=""%>
<p align="center"><b>Record has been Updated!</b></p>
	<script language="JavaScript">
		setTimeout("self.close();", 2000);
    	</script>
<%Response.end%>
<%end sub%>

<%sub Getdata%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp.execute("Select * from tbl_SRL where PID='" & PID & "';")
%>
<%do until rstemp1.eof %>
<%Session("Open_Date")=rstemp1(2)%>
<%Session("comments")=rstemp1(5)%>
<%Session("filecreated")=rstemp1(6)%>
<%Session("Request_rec")=rstemp1(7)%>
<%Session("Review_Rec")=rstemp1(8)%>

<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing

%>
<%Session("Record_Type")="Old"%>

<%End Sub%>

<%Sub Cancel%>
<%Session("Open_Date")=""%>
<%Session("Comments")=""%>
<%Session("FileCreated")=""%>
<%Session("Request_Rec")=""%>
<%Session("Review_Rec")=""%>
	<script language="JavaScript">
		setTimeout("self.close();", 30);
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>