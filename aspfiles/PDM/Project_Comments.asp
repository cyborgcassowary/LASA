<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Project Comments</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%If Request.form("B1")="Cancel" then Cancel%>
<%PID=Request.QueryString("PID")%>

<%If Request.form("B1")="Submit" then InsertData%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project,Escrowfile,Notes from tbl_Project where PID=" & PID & ";")
%>
<%do until rstemp.eof %>
<%Project=rstemp(0)%>
<%Escrowfile=rstemp(1)%>
<%Notes=rstemp(2)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<form method="POST" action="Project_Comments.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1">
	<tr>
	<td align="center">
	<b><%=Project%> - <%=Escrowfile%></b>
	</td>
	</tr>
		<tr>
			<td><textarea rows="15" name="Notes" cols="80" tabindex="1"><%=Notes%></textarea></td>
		</tr>
		<tr>
			<td>
<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>
		</tr>
	</table>
</div>
<input type="hidden" name="PID" value="<%=PID%>">
<input type="hidden" name="OldNotes" value="<%=Notes%>">
</form>

<%Sub InsertData%>
<%PID=Request.form("PID")%>
<%OldNotes=Request.form("OldNotes")%>
<%Notes=Request.form("Notes")%>
<%If Len(OldNotes)+2=Len(Notes) or Len(Notes)=2 then Cancel%>
<%Notes=Replace(Notes,"'","&#039")%>
<%LastUpdate=FormatDatetime(Now)%>
<%If Len(Notes) > 2 then Notes=Notes & vbcrlf & "Comments by:" & Request.Cookies("Portal")("FullName") & " " & LastUpdate & vbcrlf & "------------------------------------------------------------" & vbcrlf%>

<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Project Set Notes='" & Notes & "', LastUpdate='" & LastUpdate & "' Where PID=" & PID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center"><b>Comments have been added.</b></p>
<form method="POST" action="Project_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 2000);
    	</script>
<%Response.end%>
<%End Sub%>


<%Sub Cancel%>
<form method="POST" action="Project_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>