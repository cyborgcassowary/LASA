<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Delete OLS Record</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%ID=Request.QueryString("ID")%>
<%If Request.form("B1")="Yes" then DeleteRecord%>
<%If Request.form("B1")="No" then CloseWindow%>
<form method="POST" action="OLS_Delete.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td>Are you sure you want to delete this record?</td>
		<td><input type="submit" value="Yes" name="B1"></td>
		<td><input type="Submit" value="No" name="B1"></td>
	</tr>
</table>
</div>
<input type="hidden" name="ID" value="<%=ID%>">
</form>

<%Sub DeleteRecord%>
<%ID=Request.form("ID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("DELETE FROM tbl_OLS where ID = " & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center"><b>Record has been Deleted</b></p>
<form method="POST" action="OLS_Report.asp" name="Update" target="ibody">
        <input type="hidden" name="Refresh" Value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
		setTimeout("self.close();", 2000);
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub CloseWindow%>
<form method="POST" action="OLS_Report.asp" name="Update" target="ibody">
        <input type="hidden" name="Refresh" Value="Updated">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
		setTimeout("self.close();", 100); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>
