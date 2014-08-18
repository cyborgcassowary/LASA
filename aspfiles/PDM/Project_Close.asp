<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Close/Re-Open Project</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%If Request.form("B1")="Submit" then UpdateData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<form method="POST" action="Project_Close.asp" webbot-action="--WEBBOT-SELF--">

<div align="center">
<b><font size="3">Close or Re-Open a Project</font>
</b><br>
<font color="#FF0000">Closing a project will also delete all Open Permits
	for that Project </font>
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="font-family: Tahoma; font-size: 8pt; border: 2px outset #0000FF">
		<tr>
			<td><input type="radio" value="Closed" checked name="Status"></td>
			<td>Close Project</td>
			<td><input type="radio" name="Status" value="Open"></td>
			<td>Re-Open Project</td>
		</tr>
		<tr>
			<td colspan="4">
<Select name="PID" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID,Project,Status from tbl_Project order by Project ASC;")
%>
<%do until rstemp.eof %>
<option value="<%=rstemp(0)%>"><%=rstemp(1)%> (<%=rstemp(2)%>)</option>
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
			<td colspan="4">
<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>
		</tr>
	</table>
</div>
</form>
<%Sub UpdateData%>
<%Status=Request.form("Status")%>
<%PID=Request.form("PID")%>
<%CloseDate=FormatDateTime(Now,VbShortDate)%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Project Set Status='" & Status & "', LastUpdate='" & CloseDate & "' Where PID=" & PID & ";")%>
<%RS = Conn.Execute("DELETE FROM tbl_Open where PID = '" & PID & "';")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%If Status="Open" then Suffix="ed" else Suffix=""%>
<p align="center">Project has been <%=Status%><%=Suffix%>.</p>
<form method="POST" action="Project_Close.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000);
    	</script>
<%Response.end%>
<%end Sub%>
<%Sub Cancel%>
<form method="POST" action="Project_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
    	</script>
<%response.end%>
<%End Sub%>
</body>

</html>