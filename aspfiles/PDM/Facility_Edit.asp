<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%ID=Request.QueryString("ID")%>
<%Facility=Request.QueryString("Facility")%>
<%Fee=Request.QueryString("Fee")%>
<%If Request.form("B1")="Submit" then UpdateData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<form method="POST" action="Facility_Edit.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
<b><font size="3">Edit Facility Fee</font></b>
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: right"><b>Facility Name</b></td>
			<td>
			<input type="text" name="Facility" size="20" value="<%=Facility%>"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Fee Amount</b></td>
			<td><input type="text" name="Fee" size="20" value="<%=Fee%>"></td>
		</tr>
		<tr>
			<td>
&nbsp;</td>
			<td style="text-align: right">
<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>
		</tr>
	</table>
</div>
<input type="hidden" name="ID" value="<%=ID%>">
</form>

<%Sub UpdateData%>
<%ID=Request.form("ID")%>
<%Facility=Request.form("Facility")%>
<%Fee=Request.form("Fee")%>
<%Facility=Replace(Facility,"'","&#039")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Facility Set FacilityName='" & Facility & "', Fee='" & Fee & "' Where ID=" & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center">Facility has been updated in the database</p>
<form method="POST" action="Facility_Index.asp" name="Update" target="ibody">

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
