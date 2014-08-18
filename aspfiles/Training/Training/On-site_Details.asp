<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>On-Site Training Details</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%ID=Request.QueryString("ID")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select * from OnSite where ID=" & ID & ";")
%>
<%do until rstemp.eof %>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="80%" style="font-family: Tahoma; font-size: 8pt; border: 2px outset #0000FF">
		<tr>
			<td width="16%" style="text-align: right"><b>Submitted By:</b></td>
			<td width="1%">&nbsp;</td>
			<td width="82%"><%=rstemp(7)%></td>
		</tr>
		<tr>
			<td width="16%" style="text-align: right"><b>Training Subject:</b></td>
			<td width="1%">&nbsp;</td>
			<td width="82%"><%=rstemp(1)%></td>
		</tr>
		<tr>
			<td width="16%" style="text-align: right"><b>Cateogry:</b></td>
			<td width="1%">&nbsp;</td>
			<td width="82%"><%=rstemp(6)%></td>
		</tr>
		<tr>
			<td width="16%" style="text-align: right"><b>Date:</b></td>
			<td width="1%">&nbsp;</td>
			<td width="82%"><%=rstemp(2)%></td>
		</tr>
		<tr>
			<td width="16%" style="text-align: right; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b># of Hours:</b></td>
			<td width="1%" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">&nbsp;</td>
			<td width="82%" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><%=rstemp(3)%></td>
		</tr>
		<tr>
			<td colspan="3" style="text-align: center"><b>Description</b></td>
		</tr>
		<tr>
			<td width="99%" colspan="3" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><%=rstemp(5)%></td>
		</tr>
		<tr>
			<td colspan="3" style="text-align: center"><b>Attendees</b></td>
		</tr>
		<tr>
			<td colspan="3"><%=rstemp(4)%></td>
		</tr>
	</table>
</div>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<p style="text-align: center"><a href="On-Site_Index.asp">Back to On-Site 
Training Index</a></p>
</body>

</html>
