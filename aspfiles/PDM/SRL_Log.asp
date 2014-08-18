<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<table border="0" width="100%" cellspacing="0" style="font-family: Tahoma; font-size: 8pt; border: 2px outset #0000FF" cellpadding="0">
	<tr>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>EscrowFile&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Open Date&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Project Name&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Contact & Phone#&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Comments&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>File Created&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Service Request Received&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Development Review Received&nbsp;</b></td>
	</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_SRL order by EscrowFile ASC;")
%>

<%do until rstemp.eof %>
<%Set Project=conntemp.execute("Select Project from tbl_Project where PID=" & rstemp(3) & ";")%>
<%Set Dev=conntemp.execute("Select D_Name,D_Number from tbl_Developer where DID=" & rstemp(4) & ";")%>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
		<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
		<td style="text-align: left"><%=Project(0)%>&nbsp;</td>
		<td style="text-align: left"><%=Dev(0)%> - <%=Dev(1)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(5)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(6)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(7)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(8)%>&nbsp;</td>
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
</body>

</html>