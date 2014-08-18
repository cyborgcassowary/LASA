<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add New File</title>
<!-- #include file="inc/calctl.inc" -->
<script language="JavaScript" src="inc/overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%Location=Request.QueryString("Location")%>
<%If Location="" then Location=0%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Location from tbl_links where ID=" & Location & ";")
%>
<%do until rstemp.eof %>
<%Loc=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<%If Request.Cookies("FS")("Location") <> "" then Loc=Request.Cookies("FS")("Location")%>

<%If Request.form("B1")="Submit" then InsertData%>
<%If Request.form("B1")="Cancel" then Cancel%>

<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
<div align="center">
<b><font size="4">Add New File</font></b>
<form method="POST" action="FS_Add.asp" name="sample" webbot-action="--WEBBOT-SELF--">
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
<option value="<%=rstemp(0)%>"><%=rstemp(0)%></option>
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
			<input type="text" name="FileName" size="60" autocomplete="off"></td>
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
<%If Loc=rstemp(0) then sel="selected" else sel=""%>
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
<td width="15"><input type="checkbox" name="include" value="<%=rstemp(0)%>"></td><td><%=rstemp(0)%></td>
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
		<tr><%If Request.Cookies("FS")("Archive_Date") > "" then Arc_Date=Request.Cookies("FS")("Archive_Date") else Arc_Date="MM/DD/YYYY"%>
			<td style="text-align: right" width="113"><b>Archive Date</b></td>
			<td colspan="2"><input type="text" name="Archive_Date" size="20" style="font-family: Tahoma; font-size: 8pt" value="<%=Arc_Date%>">  
	<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('sample.Archive_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
	<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
		</tr>
		<tr><%If Request.Cookies("FS")("Destroy_Date") > "" then Des_Date=Request.Cookies("FS")("Destroy_Date") else Des_Date="MM/DD/YYYY"%>
			<td style="text-align: right" width="113"><b>Destroy Date</b></td>
			<td colspan="2">
			<input type="text" name="Destroy_Date" size="20" style="font-family: Tahoma; font-size: 8pt" value="1/1/1980">  
	<a href="javascript:ggPosX=10;ggPosY=10;show_calendar('sample.Destroy_Date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
	<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a></td>
		</tr>
		<tr>
			<td style="text-align: right" width="113"><b>Owner</b></td>
			<td colspan="2"><input type="text" name="Owner" size="60" autocomplete="off"></td>
		</tr>
		<tr>
			<td colspan="3">
			<p style="text-align: center"><b>Notes</b></td>
		</tr>
		<tr>
			<td colspan="3"><textarea rows="4" name="Notes" cols="80"></textarea></td>
		</tr>
		<tr>
			<td width="113" nowrap>

<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>
			<td width="25"><input type="checkbox" name="SaveData" value="ON"></td>
			<td width="281" nowrap><font size="1">Check to save <b>Location</b> 
			and <b>Dates</b> for additional Records.</font></td>
		</tr>
		<tr>
			<td width="113" nowrap style="text-align: center">

<a href="Cabinet_Index.asp">File Index</a></td>
			<td width="25">&nbsp;</td>
			<td width="281" nowrap>&nbsp;</td>
		</tr>
	</table>
</form>
</div>

<%Sub InsertData%>
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
<%RS = Conn.Execute("Insert Into tbl_Files (Category,FileName,Location,Include,Archive_Date,Destroy_Date,Owner,Notes) Values ('" & Category & "','" & FileName & "','" & Location & "','" & Include & "','" & Archive_Date & "','" & Destroy_Date & "','" & Owner & "','" & Notes & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%If Request.form("SaveData")="ON" then
	Response.Cookies("FS")("Location")=Location
	Response.Cookies("FS")("Archive_Date")=Archive_Date
	Response.Cookies("FS")("Destroy_Date")=Destroy_date
		else
	Response.Cookies("FS")("Location")=""
	Response.Cookies("FS")("Archive_Date")=""
	Response.Cookies("FS")("Destroy_Date")=""
End If%>
<p align="center">File has been added to the database.</p>
<form method="POST" action="FS_Add.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<%Response.Cookies("FS").Expires=Date()-1%>
<form method="POST" action="Cabinet_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>


</body>

</html>