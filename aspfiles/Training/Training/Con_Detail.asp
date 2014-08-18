<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Training Detail</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%ID=Request.QueryString("ID")%>
<%RP=Request.QueryString("RP")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select * from tbl_users where ID=" & ID & ";")
%>
<%do until rstemp.eof %>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="70%" id="table1" style="border: 2px outset #000080">
		<tr>
			<td width="23%" style="text-align: right" nowrap><b>Attendee:</b></td>
			<td colspan="2"><%=rstemp(1)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%" style="text-align: right" nowrap><b>Department:</b></td>
			<td colspan="2"><%=rstemp(2)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%" style="text-align: right" nowrap><b>Course Category:</b></td>
			<td colspan="2"><%=rstemp(16)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%" style="text-align: right" nowrap><b>Course Name:</b></td>
			<td colspan="2"><%=rstemp(3)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%" style="text-align: right" nowrap><b>Course Sponsor:</b></td>
			<td colspan="2"><%=rstemp(4)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%" style="text-align: right" nowrap><b>Course Location:</b></td>
			<td colspan="2"><%=rstemp(5)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%" style="text-align: right" nowrap><b>CEU Program ID</b></td>
			<td colspan="2"><%=rstemp(25)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%" style="text-align: right" nowrap><b>Date &amp; Time:</b></td>
			<td colspan="2"><%=rstemp(6)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%" style="text-align: right" nowrap>
			<p style="text-align: right"><b>Costs:</b></td>
			<td width="17%" style="text-align: right"><b>Registration:</b></td>
			<td style="text-align: justify" width="60%"><%=FormatCurrency(rstemp(9))%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%">&nbsp;</td>
			<td width="17%" style="text-align: right"><b>Travel:</b></td>
			<td style="text-align: justify" width="60%"><%=FormatCurrency(rstemp(10))%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%">&nbsp;</td>
			<td width="17%" style="text-align: right"><b>Lodging:</b></td>
			<td style="text-align: justify" width="60%"><%=FormatCurrency(rstemp(11))%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%">&nbsp;</td>
			<td width="17%" style="text-align: right"><b>Meals:</b></td>
			<td style="text-align: justify" width="60%"><%=FormatCurrency(rstemp(12))%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%">&nbsp;</td>
			<td width="17%" style="text-align: right"><b>Other:</b></td>
			<td style="text-align: justify" width="60%"><%=FormatCurrency(rstemp(13))%>&nbsp;</td>
		</tr>
		<tr>
			<td width="23%">&nbsp;</td>
			<td width="17%" style="text-align: right"><b>Total:</b></td>
			<td style="text-align: justify" width="60%"><%=FormatCurrency(Int(rstemp(9)) + Int(rstemp(10)) + Int(rstemp(11)) + Int(rstemp(12)) + Int(rstemp(13)))%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="3">
			<p style="text-align: center"><b>Course Description</b></td>
		</tr>
		<tr>
			<td colspan="3"><%=rstemp(7)%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="3">
			<p style="text-align: center"><b>Reason for Taking Course</b></td>
		</tr>
		<tr>
			<td colspan="3"><%=rstemp(8)%>&nbsp;</td>
		</tr>
	</table>
	<a href="<%=RP%>">Back to Index</a>
	- <a href="RFTC.asp?Req_Date=<%=rstemp(15)%>">Re-Print Form</a></div>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

</body>

</html>