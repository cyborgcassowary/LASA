<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Update Assignments</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%ID=Request.QueryString("ID")%>
<%If ID="" then ID=Request.Form("ID")%>
<%If Request.form("B1")="Submit" then InsertData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<div align="center">
<form method="POST" action="Assign_popup.asp" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_assign where ID=" & ID & ";")
%>
<%do until rstemp.eof %>
<%session("detail")=rstemp(5)%>
<%session("notes")=rstemp(6)%>
		<tr>
			<td style="text-align: right" width="13%" nowrap><b>Assigned To:</b></td>
			<td width="87%" colspan="3"><%=rstemp(1)%></td>
		</tr>
		<tr>
			<td style="text-align: right" width="13%"><b>Project:</b></td>
			<td width="87%" colspan="3"><%=rstemp(2)%></td>
		</tr>
		<tr>
			<td style="text-align: right" width="13%"><b>Due Date:</b></td>
			<td width="87%" colspan="3"><%=rstemp(9)%></td>
		</tr>
		<tr>
			<td colspan="4"><%=rstemp(4)%></td>
		</tr>
		<tr>
			<td colspan="4">
			<iframe frameborder="1" height="100" width="100%" name="inotes" src="idetail.asp" marginwidth="1" marginheight="1"></iframe></td>
		</tr>
		<tr>
			<td style="text-align: right" width="13%"><b>Status:</b></td>
			<td width="29%"><select size="1" name="Status">
			<%If rstemp(8)="New" then sel="selected" else sel=""%>
			<option <%=sel%>>New</option>
			<%If rstemp(8)="In Progress" then sel="selected" else sel=""%>
			<option <%=sel%>>In Progress</option>
			<%If rstemp(8)="Waiting on Someone Else" then sel="selected" else sel=""%>
			<option <%=sel%>>Waiting on Someone Else</option>
			<%If rstemp(8)="Deferred" then sel="selected" else sel=""%>
			<option <%=sel%>>Deferred</option>
			<%If rstemp(8)="Complete" then sel="selected" else sel=""%>
			<option <%=sel%>>Complete</option>
			</select></td>
			<td width="29%" style="text-align: right"><b>Priority:</b></td>
			<td width="29%">
			<select size="1" name="Priority">
			<%if rstemp(11)="Low" then sel="selected" else sel=""%>
			<option <%=sel%>>Low</option>
			<%if rstemp(11)="Normal" then sel="selected" else sel=""%>
			<option <%=sel%>>Normal</option>
			<%if rstemp(11)="High" then sel="selected" else sel=""%>
			<option <%=sel%>>High</option>
			</select></td>
		</tr>
		<tr>
			<td colspan="4"><%if Session("Notes") <> "" then GetIframe%>

		</tr>
		<tr>
			<td colspan="4" style="text-align: center">
			<textarea rows="8" name="Notes" cols="120"></textarea></td>
		</tr>
		<tr>
			<td colspan="4">
	<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>
		</tr>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
	</table>
	<input type="hidden" name="ID" value="<%=ID%>">
	<input type="hidden" name="Winder" value="<%=Winder%>">
</form>
</div>

<%Sub InsertData%>
<%ID=Request.form("ID")%>
<%Notes=Request.form("Notes")%>
<%If Session("Notes") <> "" then LineBreak="<br><br>" else LineBreak=""%>
<%Notes=Session("Notes") & Linebreak & Notes & "<br>" & formatDateTime(now,vbshortdate)%>
<%Notes=Replace(Notes,"'","&#039")%>
<%Priority=Request.form("Priority")%>
<%Status=Request.form("Status")%>
<%LastUpdate=FormatDatetime(Now,VbShortDate)%>
<%If Status="Complete" then C_Date=FormatDatetime(Now,VbshortDate) Else C_Date="01/01/1990"%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Assign Set Notes='" & Notes & "', Priority='" & Priority & "', Status='" & Status & "', LastUpdate='" & LastUpdate & "', C_Date='" & C_Date & "' Where ID=" & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>

<p align="center"><b>Assignment has been Updated.</b><br><br>This Window will close on it's own.</p>
<form method="POST" action="Assign_popup.asp" name="Update" target="ibody">
        <input type="hidden" name="ID" Value="<%=ID%>">
        </form>
	<script language="JavaScript">
		setTimeout("self.close();", 3000); 
    	</script>

<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<form method="POST" action="Assign_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("self.close();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub GetIframe%>
<iframe frameborder="1" height="100" width="100%" name="inotes" src="inotes.asp" marginwidth="1" marginheight="1"></iframe></td>
<%End Sub%>

</body>

</html>