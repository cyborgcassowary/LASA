<%Response.buffer=True%>
<%response.ContentType="application/vnd.ms-excel"%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Yearly Report Export</title>

</head>

<body>
<%IYear=Request.QueryString("IYear")%>

<table width="100%" cellspacing="0" cellpadding="0" style="border: 3px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
	<tr>
		<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">&nbsp;</td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Address</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>City, State</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Development</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Taps/GPD</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Issue Date</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Municipality</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Permit#</b></td>
	</tr>

<%SDate="01/01/" & IYear%>
<%EDate="12/31/" & IYear%>

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select P_Number,P_Street,P_CityState,Taps,I_Date,Municipality,Permit,Development,EST_GPD from TBL_Permit where I_Date Between '" & SDate & "' and '" & Edate & "';")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<%if rstemp(8) <> "" then GPD="/" & rstemp(8) else GPD=""%>
<tr bgcolor="<%=bgcolor%>">
		<td align="center"><%=X%>&nbsp;</td>
		<td><%=rstemp(0)%>&nbsp;<%=rstemp(1)%></td>
		<td align="center"><%=rstemp(2)%>&nbsp;</td>
		<td align="center"><%=rstemp(7)%>&nbsp;</td>
		<td align="center"><%=rstemp(3)%>/<%=GPD%></td>
		<td align="center"><%=rstemp(4)%>&nbsp;</td>
		<td align="center"><%=rstemp(5)%>&nbsp;</td>
		<td align="center"><%=rstemp(6)%>&nbsp;</td>
		
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
