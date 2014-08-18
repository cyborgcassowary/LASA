<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Link Locations</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%If Request.form("B1")="Submit" then Drawer_Location%>
<%If Request.form("B2")="Submit" then Box_Location%>
<%If Request.form("B3")="Submit" then BookCase_Location%>

<div align="center">
<b><font size="4">Link Storage Locations
	</font></b>
	<table border="0" cellpadding="0" cellspacing="0" width="80%">
		<tr>
		<td colspan="3" align="center"><b>Please Select a type of Location to Link</b></td>
		</tr>
		<tr>
			<td align="left" width="33%"><a href="Location_Link.asp?SubRoutine=Drawer">Cabinet/Drawer</a></td>
			<td align="center" width="33%" rowspan="3">Back to <a href="Location_add.asp">
			&quot;Add New Storage Location</a>&quot;</td>
			<td align="right" width="34%" rowspan="3">&nbsp;</td>
		</tr>
		<tr>
			<td align="left" width="33%"><a href="Location_link.asp?Subroutine=BookCase">BookCase/Shelf</a></td>
		</tr>
		<tr>
			<td align="left" width="33%"><a href="Location_Link.asp?SubRoutine=Box">Room/Box</a></td>
		</tr>
	</table>
</div>

<%Loc_Type=Request.QueryString("Subroutine")%>
<%Select Case Loc_type
	Case "Box"
			Box_Sub
	Case "Drawer"
			Drawer_Sub
	Case "BookCase"
			BookCase_Sub
	Case Else
			Drawer_Sub
End Select%>


<%Sub Drawer_Sub%>
<%Dim tbl(4)%>
<%tbl(1)="tbl_Building"%>
<%tbl(2)="tbl_Room"%>
<%tbl(3)="tbl_Cabinet"%>
<%tbl(4)="tbl_Drawer"%>
<div align="center">Currently Linking a "<font color="#FF0000"><b>Cabinet/Drawer</b></font>" Location<br>
<font size="1">Choose 1 from Each Line, then click "Submit" to create the Storage Location.</font>
<form method="POST" action="Location_Link.asp" webbot-action="--WEBBOT-SELF--">
<%for I = 1 to 4%>
<%Location_Name=Replace(tbl(I),"tbl_","")%>

<table width="80%">
	<tr>
		<td><b><%=Location_Name%></b></td>
<table width="80%" cellspacing="0" cellpadding="0" style="border-style: solid; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
<tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Name from " & tbl(I) & " order by name ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<td width="10"><input type="radio" value="<%=rstemp(0)%>" name="<%=tbl(I)%>"></td><td><%=rstemp(0)%>&nbsp;</td>
<%If X/5=Int(X/5) then response.write "</tr>"%>

<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</tr>
</table>
</tr>
</table>
<%X=0%>
<%next%>
<input type="submit" value="Submit" name="B1">
</form>

<table width="80%">
<tr><td><b>Current Links</b>&nbsp;(Click to edit)</td></tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select ID,Location from tbl_links where Location_Type='Drawer';")
%>
<%do until rstemp.eof %>
	<tr><td><a href="Link_Edit.asp?ID=<%=rstemp(0)%>"><%=rstemp(1)%></a></td></tr>
	
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</table>
</div>
<%End Sub%>

<%Sub BookCase_Sub%>
<%Dim tbl(4)%>
<%tbl(1)="tbl_Building"%>
<%tbl(2)="tbl_Room"%>
<%tbl(3)="tbl_BookCase"%>
<%tbl(4)="tbl_Shelf"%>
<div align="center">Currently Linking a "<font color="#FF0000"><b>BookCase/Shelf</b></font>" Location<br>
<font size="1">Choose 1 from Each Line, then click "Submit" to create the Storage Location.</font>
<form method="POST" action="Location_Link.asp" webbot-action="--WEBBOT-SELF--">
<%for I = 1 to 4%>
<%Location_Name=Replace(tbl(I),"tbl_","")%>

<table width="80%">
	<tr>
		<td><b><%=Location_Name%></b></td>
<table width="80%" cellspacing="0" cellpadding="0" style="border-style: solid; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
<tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Name from " & tbl(I) & " order by name ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<td width="10"><input type="radio" value="<%=rstemp(0)%>" name="<%=tbl(I)%>"></td><td><%=rstemp(0)%>&nbsp;</td>
<%If X/5=Int(X/5) then response.write "</tr>"%>

<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</tr>
</table>
</tr>
</table>
<%X=0%>
<%next%>
<input type="submit" value="Submit" name="B3">
</form>

<table width="80%">
<tr><td><b>Current Links</b>&nbsp;(Click to edit)</td></tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select ID,Location from tbl_links where Location_Type='BookCase';")
%>
<%do until rstemp.eof %>
	<tr><td><a href="Link_Edit.asp?ID=<%=rstemp(0)%>"><%=rstemp(1)%></a></td></tr>
	
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</table>
</div>
<%End Sub%>



<%Sub Box_Sub%>
<%Dim tbl(3)%>
<%tbl(1)="tbl_Building"%>
<%tbl(2)="tbl_Room"%>
<%tbl(3)="tbl_Box"%>
<div align="center">Currently Linking a "<font color="#FF0000"><b>Room/BOX</b></font>" Location<br>
<font size="1">Choose 1 from Each Line, then click "Submit" to create the Storage Location.</font>
<form method="POST" action="Location_Link.asp" webbot-action="--WEBBOT-SELF--">
<%for I = 1 to 3%>
<%Location_Name=Replace(tbl(I),"tbl_","")%>
<table width="80%">
	<tr>
		<td><b><%=Location_Name%></b></td>
<table width="80%" border="0" cellspacing="0" cellpadding="0" style="border-style: solid; border-width: 1px; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">

<tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Name from " & tbl(I) & " order by Name ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<td width="10"><input type="radio" value="<%=rstemp(0)%>" name="<%=tbl(I)%>"></td><td><%=rstemp(0)%>&nbsp;</td>
<%If X/5=Int(X/5) then response.write "</tr>"%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</tr>
</table>
</tr>
</table>
<%X=0%>
<%next%>
<input type="submit" value="Submit" name="B2">
</form>
<table width="80%">
<tr><td><b>Current Links</b>&nbsp;(Click to edit)</td></tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select ID,Location from tbl_links where Location_Type='Box';")
%>
<%do until rstemp.eof %>
	<tr><td><a href="Link_Edit.asp?ID=<%=rstemp(0)%>"><%=rstemp(1)%></a></td></tr>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</table>
</div>

<%End Sub%>

<%Sub Drawer_Location%>
<%Building=Request.Form("tbl_Building")%>
<%Room=Request.form("tbl_Room")%>
<%Cabinet=Request.form("tbl_Cabinet")%>
<%Drawer=Request.form("tbl_Drawer")%>
<%If Building = "" or Room = "" or Cabinet="" or Drawer="" then Error%>
<%Location=Building & "," & Room & "," & Cabinet & "," & Drawer%>
<p align="center"><b>Location Link "<font color="#FF0000"><%=Location%></font>" has been Created</b>.</p>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Insert Into tbl_links (Location,Location_Type) Values ('" & Location & "','Drawer')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<form method="POST" action="Location_Link.asp" name="Update" target="ibody">
<input type="hidden" name="Subroutine" value="Drawer">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 1500); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>

<%End Sub%>

<%Sub BookCase_Location%>
<%Building=Request.Form("tbl_Building")%>
<%Room=Request.form("tbl_Room")%>
<%BookCase=Request.form("tbl_BookCase")%>
<%Shelf=Request.form("tbl_Shelf")%>
<%If Building = "" or Room = "" or BookCase = "" or Shelf = "" then Error%>
<%Location=Building & "," & Room & "," & BookCase & "," & Shelf%>
<p align="center"><b>Location Link "<font color="#FF0000"><%=Location%></font>" has been Created</b>.</p>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Insert Into tbl_links (Location,Location_Type) Values ('" & Location & "','BookCase')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<form method="POST" action="Location_Link.asp" name="Update" target="ibody">
<input type="hidden" name="Subroutine" value="Drawer">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 1500); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>

<%End Sub%>

<%Sub Box_Location%>
<%Building=Request.Form("tbl_Building")%>
<%Room=Request.form("tbl_Room")%>
<%Box=Request.form("tbl_Box")%>
<%If building="" or Room="" or Box="" then Error%>
<%Location=Building & "," & Room & "," & Box%>
<p align="center"><b>Location Link "<font color="#FF0000"><%=Location%></font>" has been Created</b>.</p>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Insert Into tbl_links (Location,Location_Type) Values ('" & Location & "','Box')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<form method="POST" action="Location_Link.asp" name="Update" target="ibody">
	<input type="hidden" name="Subroutine" value="Box">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 1500); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Error%>
<p align="center"><b><font size="3">You did not complete the Location Link<br>No Location has been Created.</font></b></p>
<form method="POST" action="Location_link.asp" name="Update" target="ibody">
        <input type="hidden" name="Refresh" Value="Submit">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>