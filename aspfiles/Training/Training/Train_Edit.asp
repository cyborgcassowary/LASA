<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Training Editor</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="inc/Train_Check.inc"-->
<!--#include file="inc/Write_Check.inc"-->
<%If Request.form("B1")="Submit" then InsertData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<%ID=Request.QueryString("ID")%>
<b><%=ID%></b>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select CID from tbl_JobTitle where Jobtitle='" & ID & "';")
%>
<%do until rstemp.eof %>
<%Suggested=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<form method="POST" action="Train_Edit.asp" webbot-action="--WEBBOT-SELF--">
<table style="border-style:solid; border-width:1px; font-family: Tahoma; font-size: 8pt; padding-left:4px; padding-right:4px; padding-top:1px; padding-bottom:1px" cellspacing="0" cellpadding="0">
<tr>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">Category</td>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">Suggested Training</td>
</tr>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select Category from tbl_Category;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>">
<td><%=rstemp(0)%>&nbsp;</td>
<%If Int(Instr(Suggested,rstemp(0))) > 0 then chk="checked" else chk=""%>
<td style="text-align: center"><input type="checkbox" name="CID" value="<%=rstemp(0)%>" <%=chk%>></td>
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
<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1">
<input type="hidden" name="ID" value="<%=ID%>"%>
</form>

<%Sub InsertData%>
<%CID=Request.form("CID")%>
<%ID=Request.form("ID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=Training;"%>
<%RS = Conn.Execute("Update tbl_Jobtitle Set CID='" & CID & "' Where JobTitle='" & ID & "';")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<form method="POST" action="train_grid.asp" name="Update">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<form method="POST" action="train_grid.asp" name="Update">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
    	</script>
<%Response.end%>
<%End Sub%>

</body>

</html>