<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add New Facility Fee</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="Submit" then InsertData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<form method="POST" action="Facility_Add.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
<b><font size="3">Add New Facility Fee</font></b>
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: right"><b>Facility Name</b></td>
			<td><input type="text" name="Facility" size="20"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Fee Amount</b></td>
			<td><input type="text" name="Fee" size="20"></td>
		</tr>
		<tr>
			<td>
&nbsp;</td>
			<td style="text-align: right">
<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>
		</tr>
	</table>
</div>
</form>

<%Sub InsertData%>
<%Facility=Request.form("Facility")%>
<%Fee=Request.form("Fee")%>
<%Facility=Replace(Facility,"'","&#039")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Insert Into tbl_Facility (FacilityName,Fee) Values ('" & Facility & "','" & Fee & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">Facility has been added to the database</p>
<form method="POST" action="Facility_Add.asp" name="Update" target="ibody">

</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 2000); 
    	</script>
<%Response.end%>
<%end Sub%>

<%Sub Cancel%>
<form method="POST" action="Project_index.asp" name="Update" target="ibody">

</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%End Sub%>

</body>

</html>
