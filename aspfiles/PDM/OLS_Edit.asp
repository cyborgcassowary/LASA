<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>On-Lot Septic Editor</title>
<Head>
<!-- #include file="calctl.inc" -->
<script language="JavaScript" src="overlib_mini.js"></script>
<style type="text/css">
.style1 {
	text-align: right;
}
</style>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="Cancel" then Cancel%>
<%If Request.form("B1")="Submit" and Request.form("SDATE")="MM/DD/YYYY" then Cancel%>
<%If Request.form("B1")="Submit" and Request.form("TAG")="IT" then UpdateData%>
<%If Request.form("B1")="Submit" and Request.form("TAG") <> "IT" then InsertData%>

<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<%PID=Request.QueryString("PID")%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project,GISID from tbl_Project where PID=" & PID & ";")
%>

<%do until rstemp.eof %>
<%Project=rstemp(0)%>
<%GISID=rstemp(1)%>
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
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select SDATE,Permit,Status,Plant,EDU,BuildDate from tbl_OLS where PID='" & PID & "';")
%>
<%do until rstemp.eof %>
<%SDATE=rstemp(0)%>
<%Permit=rstemp(1)%>
<%Status=rstemp(2)%>
<%Plant=rstemp(3)%>
<%EDU=rstemp(4)%>
<%BDate=rstemp(5)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%If SDATE="" then SDATE="MM/DD/YYYY" Else TAG="IT"%>
<%If BDate="" then BDate="MM/DD/YYYY"%>
<form method="POST" name="OLS" action="OLS_Edit.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
<b><font size="4">On-Lot Septic Editor</font></b>
<table border="0" cellspacing="0" cellpadding="0" id="table1">
	<tr>
		<td><b>Project Name:</b></td>
		<td><%=Project%></td>
	</tr>
	<tr>
		<td><b>Original Tax Parcel#</b></td>
		<td><input type="text" name="GISID" size="20" value="<%=GISID%>"></td>
	</tr>
	<tr>
		<td class="style1"><strong>Build Date</strong></td>
		<td><input type="text" name="BDate" value="<%=BDATE%>">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('OLS.BDate');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td><b>Out of Service Date:</b></td>
		<td><input type="text" name="SDate" value="<%=SDATE%>">
<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('OLS.SDate');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></td>
	</tr>
	<tr>
		<td style="text-align: right">
<b>Treatment Plant</b></td>
		<td><select size="1" name="Plant">
		<%if Plant="LASA" then chk="selected" else chk=""%>
		<option <%=chk%>>LASA</option>
		<%If Plant="Lancaster" then chk="selected" else chk=""%>
		<option <%=chk%>>Lancaster</option>
		</select></td>
	</tr>
	<tr>
		<td style="text-align: right">
<b>EDU's</b></td>
		<td><input type="text" name="EDU" size="20" value="<%=edu%>"></td>
	</tr>
	<tr>
		<td style="text-align: right">
<b>Permit #</b></td>
		<td><input type="text" name="Permit" size="20" value="<%=Permit%>"></td>
	</tr>
	<tr>
		<td style="text-align: right">
<b>Status</b></td>
		<td>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table2">
			<tr>
				<td><%if Status="Out of Service" then sel="checked" else sel=""%>
				<input type="radio" value="Out of Service" name="Status" <%=sel%>></td>
				<td>&nbsp;Out of Service</td>
				<td><%If Status="Connected" then sel="checked" else sel=""%>
				<input type="radio" value="Connected" name="Status" <%=sel%>></td>
				<td>&nbsp;Connected</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1">
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="TAG" value="<%=TAG%>">
</td>
		<td>&nbsp;</td>
	</tr>
</table>
</div>
</form>

<%Sub InsertData%>
<%PID=Request.form("PID")%>
<%SDate=Request.form("SDate")%>
<%GISID=Request.form("GISID")%>
<%Permit=Request.form("Permit")%>
<%Status=Request.form("Status")%>
<%Plant=Request.form("Plant")%>
<%EDU=Request.form("EDU")%>
<%BDate=Request.form("BDate")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Insert Into tbl_OLS (PID,SDATE,GISID,Permit,Status,Plant,EDU,BuildDate) Values ('" & PID & "','" & SDATE & "','" & GISID & "','" & Permit & "','" & Status & "','" & Plant & "','" & EDU & "','" & BDate & "')")%>
<%RS = Conn.Execute("Update tbl_Project Set GISID='" & GISID & "' Where PID=" & PID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p>Record Updated!</p>
	<script language="JavaScript">
		setTimeout("self.close();", 2000);
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub UpdateData%>
<%PID=Request.form("PID")%>
<%SDate=Request.form("SDate")%>
<%GISID=Request.form("GISID")%>
<%Permit=Request.form("Permit")%>
<%Status=Request.form("Status")%>
<%Plant=Request.form("Plant")%>
<%EDU=Request.form("EDU")%>
<%Bdate=Request.form("BDate")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_OLS Set GISID='" & GISID & "', SDATE='" & SDATE & "', Permit='" & Permit & "', Status='" & Status & "', Plant='" & Plant & "', EDU='" & EDU & "', BuildDate='" & BDate & "' where PID='" & PID & "';")%>
<%RS = Conn.Execute("Update tbl_Project Set GISID='" & GISID & "' Where PID=" & PID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p>Record Updated!</p>
	<script language="JavaScript">
		setTimeout("self.close();", 2000);
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
	<script language="JavaScript">
		setTimeout("self.close();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>
