<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Edit Link</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%If Request.form("B1")="Update" then Drawer_Location%>
<%If Request.form("B2")="Update" then Box_Location%>
<%If Request.form("B3")="Update" then BookCase_Location%>
<%If Request.form("B1")="Cancel" then Cancel%>
<%ID=Request.QueryString("ID")%>

<div align="center">
<b><font size="4">Edit Link Storage Locations
	</font></b>
</div>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select Location,Location_Type from tbl_links where ID=" & ID & ";")
%>
<%do until rstemp.eof %>
<%Location=rstemp(0)%>
<%Location_Type=rstemp(1)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<%Select Case Location_Type
	Case "Drawer"
			Drawer_Sub
	Case "Box"
			Box_Sub
	Case "BookCase"
			BookCase_Sub
End Select%>


<%Sub Drawer_Sub%>
<%Dim tbl(4)%>
<%tbl(0)="tbl_Building"%>
<%tbl(1)="tbl_Room"%>
<%tbl(2)="tbl_Cabinet"%>
<%tbl(3)="tbl_Drawer"%>
<%LocArray=Split(Location,",")%>
<div align="center">
<font size="1">Choose 1 from Each Line, then click "Update" to edit the Storage Location.</font>
<form method="POST" action="Link_Edit.asp" webbot-action="--WEBBOT-SELF--">
<%for I = 0 to 3%>
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
<td width="10">
<%If LocArray(I)=rstemp(0) then
	sel="checked"
	fntcolor="red"
		else
	sel=""
	fntcolor=""
end If%>
<input type="radio" value="<%=rstemp(0)%>" name="<%=tbl(I)%>" <%=sel%>></td><td><font color="<%=fntcolor%>"><%=rstemp(0)%></font>&nbsp;</td>
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
<input type="hidden" name="ID" value="<%=ID%>">
<input type="submit" value="Update" name="B1">
<input type="submit" value="Cancel" name="B1"></form>
<%End Sub%>

<%Sub BookCase_Sub%>
<%Dim tbl(4)%>
<%tbl(0)="tbl_Building"%>
<%tbl(1)="tbl_Room"%>
<%tbl(2)="tbl_BookCase"%>
<%tbl(3)="tbl_Shelf"%>
<%LocArray=Split(Location,",")%>
<div align="center">
<font size="1">Choose 1 from Each Line, then click "Update" to edit the Storage Location.</font>
<form method="POST" action="Link_Edit.asp" webbot-action="--WEBBOT-SELF--">
<%for I = 0 to 3%>
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
<td width="10">
<%If LocArray(I)=rstemp(0) then
	sel="checked"
	fntcolor="red"
		else
	sel=""
	fntcolor=""
end If%>
<input type="radio" value="<%=rstemp(0)%>" name="<%=tbl(I)%>" <%=sel%>></td><td><font color="<%=fntcolor%>"><%=rstemp(0)%></font>&nbsp;</td>
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
<input type="hidden" name="ID" value="<%=ID%>">
<input type="submit" value="Update" name="B3">
<input type="submit" value="Cancel" name="B1"></form>
<%End Sub%>

<%Sub Box_Sub%>
<%Dim tbl(3)%>
<%tbl(0)="tbl_Building"%>
<%tbl(1)="tbl_Room"%>
<%tbl(2)="tbl_Box"%>
<%LocArray=Split(Location,",")%>
<div align="center">
<font size="1">Choose 1 from Each Line, then click "Submit" to create the Storage Location.</font>
<form method="POST" action="Link_Edit.asp" webbot-action="--WEBBOT-SELF--">
<%for I = 0 to 2%>
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
<%If LocArray(I)=rstemp(0) then
	sel="checked"
	fntcolor="red"
		else
	sel=""
	fntcolor=""
end If%>
<td width="10"><input type="radio" value="<%=rstemp(0)%>" name="<%=tbl(I)%>" <%=sel%>></td><td><font color="<%=fntcolor%>"><%=rstemp(0)%></font>&nbsp;</td>
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
<input type="hidden" name="ID" value="<%=ID%>">
<input type="submit" value="Update" name="B2">
<input type="submit" value="Cancel" name="B1"></form>
<%End Sub%>

<%Sub Drawer_Location%>
<%ID=Request.form("ID")%>
<%Building=Request.Form("tbl_Building")%>
<%Room=Request.form("tbl_Room")%>
<%Cabinet=Request.form("tbl_Cabinet")%>
<%Drawer=Request.form("tbl_Drawer")%>
<%Location=Building & "," & Room & "," & Cabinet & "," & Drawer%>
<p align="center"><b>Location Link "<font color="#FF0000"><%=Location%></font>" has been Updated</b>.</p>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Update tbl_Links Set Location='" & Location & "' Where ID=" & ID & ";")%>
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
<%ID=Request.form("ID")%>
<%Building=Request.Form("tbl_Building")%>
<%Room=Request.form("tbl_Room")%>
<%BookCase=Request.form("tbl_bookCase")%>
<%Shelf=Request.form("tbl_Shelf")%>
<%Location=Building & "," & Room & "," & BookCase & "," & Shelf%>
<p align="center"><b>Location Link "<font color="#FF0000"><%=Location%></font>" has been Updated</b>.</p>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Update tbl_Links Set Location='" & Location & "' Where ID=" & ID & ";")%>
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
<%ID=Request.form("ID")%>
<%Building=Request.Form("tbl_Building")%>
<%Room=Request.form("tbl_Room")%>
<%Box=Request.form("tbl_Box")%>
<%Location=Building & "," & Room & "," & Box%>
<p align="center"><b>Location Link "<font color="#FF0000"><%=Location%></font>" has been Updated</b>.</p>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Update tbl_Links Set Location='" & Location & "' Where ID=" & ID & ";")%>
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

<%Sub Cancel%>
<form method="POST" action="Location_link.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>