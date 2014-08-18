<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>On-Site Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="80%" style="font-family: Tahoma; font-size: 8pt; border: 2px outset #0000FF">&nbsp;</td>
			<tr>
			<td  style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" colspan="2">&nbsp;</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Submitted By</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Title</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Category</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Date</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Hours</b></td>
		</tr>


<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select Subject,C_Date,C_Hours,Category,SubmitBy,ID from OnSite order by C_Date ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
		<tr bgcolor="<%=Bgcolor%>">
			<td align="center"><a href="On-Site_Edit.asp?ID=<%=rstemp(5)%>">
			<img src="../images/dir_misc.gif" title="Edit" border="0"></a></td>
			<td align="center"><a href="On-Site_Details.asp?ID=<%=rstemp(5)%>">
			<img src="../images/dir_dir.gif" title="View Details" border="0"></a></td>
			<td align="center"><%=rstemp(4)%></td>
			<td align="center"><%=rstemp(0)%></td>
			<td align="center"><%=rstemp(3)%></td>
			<td align="center"><%=rstemp(1)%></td>
			<td align="center"><%=rstemp(2)%></td>
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