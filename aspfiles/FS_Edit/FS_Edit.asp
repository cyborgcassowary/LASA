<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Edit File</title>
<!-- #include file="inc/calctl.inc" -->
<script language="JavaScript" src="inc/overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%If Request.form("B1")="Submit" then UpdateData%>
<%If Request.form("B1")="Cancel" then Cancel%>

<%ID=Request.QueryString("ID")%>
<%FN=Request.QueryString("FN")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select * from tbl_Files where ID=" & ID & ";")
%>
<%do until rstemp.eof %>
<%Category=rstemp(1)%>
<%FileName=rstemp(2)%>
<%Location=rstemp(3)%>
<%Include=rstemp(4)%>
<%Archive_Date=rstemp(5)%>
<%Destroy_Date=rstemp(6)%>
<%Owner=rstemp(7)%>
<%Notes=rstemp(8)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<div align="center">
<b><font size="4">Update Existing File</font></b>
<form method="POST" action="FS_Edit.asp" name="sample" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellspacing="0" id="table1" style="border-style: solid; border-width: 1px; padding: 1px" width="424">
		<tr>
			<td style="text-align: right" width="113"><b>File Category</b></td>
			<td colspan="2">
			<Select name="Category" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Name from tbl_Category order by name ASC;")
%>
<%do until rstemp.eof %>
<%If Category=rstemp(0) then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(0)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select>
			</td>
		</tr>
		<tr>
			<td style="text-align: right" width="113"><b>File Name</b></td>
			<td colspan="2">
			<input type="text" name="FileName" size="60" autocomplete="off" value="<%=FileName%>"></td>
		</tr>
		<tr>
			<td style="text-align: right" width="113"><b>Location</b></td>
			<td colspan="2">
<Select name="Location" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Location,ID from tbl_Links order by Location ASC;")
%>
<%do until rstemp.eof %>
<%If Location=rstemp(0) then sel="selected" else sel=""%>
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
			<td height="31" style="text-align: right" width="113"><b>Included Forms</b></td>
			<td height="31" colspan="2">
<table width="100%">
<tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Name from tbl_includes order by Name ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If Instr(include,rstemp(0)) then chk="checked" else chk=""%>
<td width="15"><input type="checkbox" name="include" value="<%=rstemp(0)%>" <%=chk%>></td><td><%=rstemp(0)%></td>
<%If X/3=Int(X/3) then response.write "</tr>"%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</tr>
</table>			</td>
		</tr>
		<tr>
			<td style="text-align: right" width="113"><b>Archive Date</b></td>
			<td colspan="2">
			<input type="text" name="Archive_Date" size="20" style="font-family: Tahoma; font-size: 8pt" value="<%=Archive_Date%>">  
	<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('sample.Archive_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
	<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
		</tr>
		<tr>
			<td style="text-align: right" width="113"><b>Destroy Date</b></td>
			<td colspan="2">
			<input type="text" name="Destroy_Date" size="20" style="font-family: Tahoma; font-size: 8pt" value="<%=Destroy_Date%>">  
	<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('sample.Destroy_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
	<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
		</tr>
		<tr>
			<td style="text-align: right" width="113"><b>Owner</b></td>
			<td colspan="2">
			<input type="text" name="Owner" size="60" autocomplete="off" value="<%=Owner%>"></td>
		</tr>
		<tr>
			<td colspan="3">
			<p style="text-align: center"><b>Notes</b></td>
		</tr>
		<tr>
			<td colspan="3"><textarea rows="4" name="Notes" cols="80"><%=Notes%></textarea></td>
		</tr>
		<tr>
			<td width="113" nowrap>

<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>
			<td width="25"></td>
			<td width="281" nowrap><a href="New_Cab_index.asp">File Index</a></td>
		</tr>
	</table>
<input type="hidden" name="ID" value="<%=ID%>">
<input type="hidden" name="FN" value="<%=FN%>">
</form>
</div>
<%Sub UpdateData%>
<%ID=Request.form("ID")%>
<%Category=Request.form("Category")%>
<%FileName=Request.form("FileName")%>
<%FileName=Replace(FileName,"'","&#039")%>
<%Location=Request.form("Location")%>
<%Include=Request.form("Include")%>
<%Archive_Date=Request.form("Archive_Date")%>
<%Destroy_Date=Request.form("Destroy_date")%>
<%Owner=Request.form("Owner")%>
<%Notes=Request.form("Notes")%>
<%Notes=Replace(Notes,"'","&#039")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Update tbl_Files Set Category='" & Category & "', FileName='" & FileName & "', Location='" & Location & "', Include='" & Include & "', Archive_Date='" & Archive_Date & "', Destroy_Date='" & Destroy_Date & "', Owner='" & Owner & "', Notes='" & Notes & "' Where ID=" & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">File has been Updated in the database.</p>
<form method="POST" action="New_Cab_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<%FN=Request.Form("FN")%>
<%If FN = "" then FN="New_Cab_index.asp"%>
<%Response.Cookies("FS").Expires=Date()-1%>
<form method="POST" action="<%=FN%>" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>
