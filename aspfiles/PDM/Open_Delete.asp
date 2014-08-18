<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Confirm Open Permit Delete</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%OID=Request.QueryString("OID")%>
<%PID=Request.Querystring("PID")%>
<%If Request.form("B1")="Yes" then DeleteRecord%>
<%If Request.form("B1")="No" then Cancel%>
<form method="POST" action="Open_Delete.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table1" height="150">
		<tr>
			<td style="text-align: center" colspan="4"><b><font size="3">Are you sure you want to delete this Open Permit?</font></b>
			</td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: center"><b><font size="3">This 
			will permanently delete this open permit with no chance of recovery.</font></b></td>
		</tr>
		<tr>
			<td width="15">&nbsp;</td>
			<td width="49%" valign="bottom">
<input type="submit" value="Yes" name="B1"></td>
			<td style="text-align: right" width="40%" valign="bottom">
			<input type="submit" value="No" name="B1"></td>
			<td style="text-align: right" width="15">&nbsp;</td>
		</tr>
	</table>
</div>
<input type="hidden" name="OID" value="<%=OID%>">
<input type="hidden" name="PID" value="<%=PID%>">
</form>

<%Sub DeleteRecord%>
<%OID=Request.form("OID")%>
<%PID=Request.form("PID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("DELETE FROM tbl_Open where ID = " & OID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center"><b>Open Permit has been Deleted.</b></p>
<form method="POST" action="Permit_open.asp" name="Update" target="ibody">
        <input type="hidden" name="Refresh" Value="Updated">
        <input type="hidden" name="PID" Value="<%=PID%>">
        <input type="hidden" name="OID" value="<%=OID%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
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