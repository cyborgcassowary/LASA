<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="inc/print.inc"-->
<div align="center">
<table border="0" width="80%" cellspacing="0" cellpadding="0" style="border: 3px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
	<tr>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Project</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>IDU/EDU</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Total Reserved</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Customer Type</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Parcel ID</b></td>
	</tr>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Project,Res_Type,Res_Amount,User_Type,GISID from tbl_Project where GISID <> '';")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">		<td><%=rstemp(0)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(3)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(4)%>&nbsp;</td>
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