<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add New File Storage Location</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="Submit" then InsertData%>
<form method="POST" action="Location_Add.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
<b><font size="4">Add New Storage Location</font></b>
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: right"><b>Location Name:&nbsp;</b></td>
			<td>
			<input type="text" name="Location_Name" size="30" autocomplete="off" maxlength="30"><font size="1">Max 
			Length 30 Characters</font></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Location Type:&nbsp;</b></td>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
					<tr>
						<td width="10">
						<input type="radio" value="tbl_Building" name="R1"></td>
						<td>&nbsp;Building</td>
						<td width="10">
						<input type="radio" value="tbl_Room" name="R1"></td>
						<td>&nbsp;Room</td>
						<td width="10">
						<input type="radio" value="tbl_Cabinet" name="R1"></td>
						<td>&nbsp;Cabinet</td>
						<td width="10">
						<input type="radio" value="tbl_Drawer" name="R1"></td>
						<td>&nbsp;Drawer</td>
						<td width="10">
						<input type="radio" value="tbl_Box" name="R1"></td>
						<td>&nbsp;Box</td>
						<td>
						<input type="radio" value="tbl_BookCase" name="R1"></td>
						<td>BookCase</td>
						<td>
						<input type="radio" value="tbl_Shelf" name="R1"></td>
						<td>Shelf</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: center">
<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></td>
			<td style="text-align: center"><a href="Location_Link.asp">Link 
			Storage Locations</a></td>
		</tr>
	</table>
</div>
</form>

<%Sub InsertData%>
<%Location_Name=Request.form("Location_Name")%>
<%Location_Name=replace(Location_Name,"'","&#039")%>
<%tbl=Request.form("R1")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Insert Into " & tbl & " (Name) Values ('" & Location_Name & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">Location has been created in the database.</p>
<form method="POST" action="Location_Add.asp" name="building" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.building.submit();", 2000);
    	</script>
<%Response.end%>
<%End sub%>



</body>

</html>