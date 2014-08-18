<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Developer Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->

<div align="center">
<b><font size="3">Developer Index</font></b><hr width="80%">
	<table border="0" cellpadding="0" cellspacing="0" width="80%" id="table1" style="border: 2px outset #0000FF">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Developer order by D_Name ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2 = Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
		<tr bgcolor="<%=bgcolor%>">
			<td><a href="D_Edit.asp?ID=<%=rstemp(0)%>"><b><%=rstemp(1)%></b></a></td>
			<td align="right"><%=rstemp(7)%></td>
		</tr>
		<tr bgcolor="<%=bgcolor%>">
			<td><%=rstemp(2)%></td>
			<td align="right"><%=rstemp(6)%></td>
		</tr>
		<tr bgcolor="<%=bgcolor%>">
			<td><%=rstemp(3)%>, <%=rstemp(4)%>&nbsp;&nbsp;<%=rstemp(5)%></td>
			<td>&nbsp;</td>
		</tr>
		<tr bgcolor="<%=bgcolor%>">
			<td colspan="2">&nbsp;</td>
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
</div>
</body>

</html>