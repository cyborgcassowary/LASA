<html>
<%ID=Request.querystring("ID")%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Viewing Notes for dbID#<%=ID%></title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="Update" then UpdateData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Notes,Permit from tbl_Permit where ID=" & ID & ";")
%>
<%do until rstemp.eof %>
<form method="POST" action="Permit_Notes.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="420" id="table1">
		<tr>
		<td colspan="4"><b><font size="1">History</font></b></td>
		</tr>
		<tr>
			<td colspan="4"><textarea rows="15" name="P_Notes" cols="80" readonly><%=rstemp(0)%></textarea></td>
		</tr>
		<tr>
		<td colspan="4"><b><font size="1">New Note</font></b></td>
		</tr>
		<tr>
			<td colspan="4"><textarea rows="15" name="C_Notes" cols="80"></textarea></td>
		</tr>
		<tr>
			<td width="83">

<input type="submit" value="Update" name="B1"></td>
			<td width="121"><b>Notes for Permit #</b></td>
			<td width="162">&nbsp;<%=rstemp(1)%></td>
			<td style="text-align: right" width="54">
			<input type="submit" value="Cancel" name="B1"></td>
		</tr>
	</table>
</div>
<input type="hidden" name="ID" value="<%=ID%>">
</form>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

<%Sub UpdateData%>
<%P_Notes=Request.form("P_Notes")%>
<%P_Notes=replace(P_Notes,"'","&#039")%>
<%C_Notes=Request.form("C_Notes")%>
<%C_Notes=Replace(C_Notes,"'","&#039")%>
<%Notes=P_Notes & VBCRLF & formatDateTime(Now,vbshortDate) & " - " & Request.Cookies("Portal")("FullName") & VBCRLF & C_Notes%>
<%ID=Request.form("ID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Permit Set Notes = '" & Notes & "' Where ID=" & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%End Sub%>

<%Sub Cancel%>
	<script language="JavaScript">
		setTimeout("self.close();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>
