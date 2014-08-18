<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Delete Project</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%if Request.form("B1")="Submit" then ProjectDelete%>
<form method="POST" action="Project_Delete.asp" webbot-action="--WEBBOT-SELF--">
<p style="text-align: center"><b><font size="3">Delete an Incomplete Project</font></b><br>You can only delete projects that have not closed Permits.<br>
<Select name="PID" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select PID,Project,escrowfile from tbl_Project order by Project,PID;")
%>
<%do until rstemp.eof %>
<option value="<%=rstemp(0)%>"><%=rstemp(1)%> - <%=rstemp(2)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select><input type="submit" value="Submit" name="B1">
</p>
</form>

<%sub ProjectDelete%>
<%PID=Request.form("PID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%Set Closed_Permit = Conn.Execute("Select Count(ID) From tbl_Permit where PID='" & PID & "';")%>
<%Set Open_Permit = Conn.Execute("Select Count(ID) From tbl_Open where PID='" & PID & "';")%>

<%CP=Closed_Permit(0)%><br>
<%OP=Open_Permit(0)%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%If CP > 0 then Response.write"This Project has Closed Permits and cannot be deleted!":Response.end%>

<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
Deleting Project...<br>
<%For I = 1 to 1000000%><%Next%>
<%RS = Conn.Execute("DELETE FROM tbl_project where PID = " & PID & ";")%>
Deleting Board Info...<br>
<%For I = 1 to 1000000%><%Next%>
<%RS = Conn.Execute("DELETE FROM ph_Board where PID = '" & PID & "';")%>
Deleting Capacity Info...<br>
<%For I = 1 to 1000000%><%Next%>
<%RS = Conn.Execute("DELETE FROM ph_Capacity where PID = '" & PID & "';")%>
Deleting Closing Info...<br>
<%For I = 1 to 1000000%><%Next%>
<%RS = Conn.Execute("DELETE FROM ph_close where PID = '" & PID & "';")%>
Deleting Construction Info...<br>
<%For I = 1 to 1000000%><%Next%>
<%RS = Conn.Execute("DELETE FROM ph_Construction where PID = '" & PID & "';")%>
Deleting Development Info...<br>
<%For I = 1 to 1000000%><%Next%>
<%RS = Conn.Execute("DELETE FROM ph_development where PID = '" & PID & "';")%>
Deleting Indemnity Info...<br>
<%For I = 1 to 1000000%><%Next%>
<%RS = Conn.Execute("DELETE FROM ph_indemn where PID = '" & PID & "';")%>
Deleting <%=OP%> Open Permits...<br>
<%For I = 1 to 1000000%><%Next%>
<%RS = Conn.Execute("DELETE FROM tbl_open where PID = '" & PID & "';")%>
Project has been deleted!<br>

<%
set RS=nothing
conn.close
Set conn=nothing
%>
<form method="POST" action="Project_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000);
    	</script>
<%Response.end%>
<%End sub%>
</body>

</html>
