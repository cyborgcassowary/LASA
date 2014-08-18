<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Fix Broken Project</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>


<p>This program will repair incompleted tables.<br>It should only be used when Project has been created,<br>but updates cannot be applied to the phases.</p>

<form method="POST" action="Project_Fix.asp" webbot-action="--WEBBOT-SELF--"><b>Enter Project ID to Repair</b>
<input type="text" name="PID" size="3"><input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1">
</form>

<%If Request.form("B1")="Submit" then FixTables%>
<%If Request.form("B1")="Cancel" then Cancel%>

<%Sub FixTables%>
<%PID=Request.form("PID")%>
<%Dim Table(7)%>
<%Table(1)="ph_board"%>
<%Table(2)="ph_capacity"%>
<%Table(3)="ph_close"%>
<%table(4)="ph_construction"%>
<%table(5)="ph_development"%>
<%table(6)="ph_indemn"%>


<%for I = 1 to 6%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%PIDChk=conn.execute("Select PID from " & table(I) & " where PID='" & PID & "';")%>
<%If PIDChk(0) <> PID then Status="Fixed" else Status="Does not need fixed"%>
<%=Table(I)%>:<%=PIDChk(0)%>:<%=Status%><br>
<%If PIDChk(0) <> PID then RS = Conn.Execute("Insert Into " & Table(I) & " (PID) Values ('" & PID & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<%next%>
<p align="center"><b>Repair Complete!</b></p>
	<script language="JavaScript">
		setTimeout("self.close();", 3000);
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