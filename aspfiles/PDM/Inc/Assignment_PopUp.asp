<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Assignment Past Due</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%TDate=FormatDateTime(Now,VbShortDate)%>
<%If Request.form("B1")="Close" then CloseWindow%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Count(ID) from tbl_Assign where Status <> 'Complete' and DueDate < '" & TDate & "';")
%>
<%DoDD=conntemp.execute("Select Count(PID) from ph_indemn where DateDiff(m,DODD,'" & TDate & "') >= 18 and DODD <> '1/1/1990';")%>
<div align="center">
Today is <%=FormatDateTime(TDate,vbLongDate)%><hr>
There are currently <%=rstemp(0)%> tasks overdue.<br>
<a href="../Assign_Index.asp" target="ibody">Click here</a> to view all open tasks.
<%If DoDD(0) > 0 then DoDDMessage%>
<hr>
<form method="POST" action="Assignment_PopUp.asp" webbot-action="--WEBBOT-SELF--">
	
	<p>Check this box to ignore this message until tomorrow<br>
	<input type="checkbox" name="C1" value="ON"><input type="submit" value="Close" name="B1"></p>
</form>
</div>


<%Sub DoDDMessage%>
<hr>
There are currently <%=DoDD(0)%> Overdue Deeds of Dedication.<br>
<a href="../Dod_report.asp" target="ibody">Click Here</a> to view All Deeds of Dedication
<%End Sub%>
<%Sub CloseWindow%>

<%TDate=FormatDateTime(Now,VbShortDate)%>
<%If Request.form("C1")="ON" then Response.cookies("PDM")("Assignment")=TDate%>
	<script language="JavaScript">
		setTimeout("self.close();", 30);
    	</script>
<%End Sub%>
</body>

</html>
