<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add Category or Included Document</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%If Request.form("B1")="Submit" and Request.form("R1") <> "" then InsertData%>

<form method="POST" action="CATINC_Add.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
<b><font size="4">Add New File Category or Included Document
	</font></b><br><br>
	<table border="0" cellpadding="0" cellspacing="0" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: right"><b>Name</b></td>
			<td><input type="text" name="Name" size="20"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Type</b></td>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" style="font-size: 8pt">
					<tr>
						<td width="15"><input type="radio" value="tbl_Category" name="R1"></td>
						<td>File Category</td>
					</tr>
					<tr>
						<td width="15"><input type="radio" name="R1" value="tbl_Includes"></td>
						<td>Included Document</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td>
&nbsp;</td>
			<td>
<input type="submit" value="Submit" name="B1"></td>
		</tr>
	</table>
</div>
</form>

<%sub InsertData%>
<%tbl=Request.form("R1")%>
<%Name=Request.form("Name")%>
<%Name=Replace(Name,"'","&#039")%>
<%If Name="" then NoData%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=FS;"%>
<%RS = Conn.Execute("Insert Into " & tbl & " (Name) Values ('" & Name & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">Item has been added to the database.</p>
<form method="POST" action="CATINC_add.asp" name="Update" target="ibody">

</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 1500); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>
<%End Sub%>
<%Sub NoData%>
<p align="center">You did not enter a value for the Name of this item.<br>Database has not been updated.</p>
<form method="POST" action="CATINC_add.asp" name="Update" target="ibody">

</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); <!--// Time in Milliseconds (1000 = 1second)-->
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>
