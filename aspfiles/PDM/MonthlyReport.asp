<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Permits Issued/Month</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%Dim MDays(12)%>
<%MDays(1)=31%>
<%MDays(2)=28%>
<%MDays(3)=31%>
<%MDays(4)=30%>
<%MDays(5)=31%>
<%MDays(6)=30%>
<%MDays(7)=31%>
<%MDays(8)=31%>
<%MDays(9)=30%>
<%MDays(10)=31%>
<%MDays(11)=30%>
<%MDays(12)=31%>

<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table1" style="border: 2px outset #0000FF">
	<tr>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Year</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Jan</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Feb</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Mar</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Apr</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>May</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>June</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>July</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Aug</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Sept</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Oct</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Nov</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Dec</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Total</b></td>
	</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
%>
<%set TP=conntemp.execute("Select Count(ID) from tbl_Permit;")%>
<%For YYYY = 1974 to Year(Now)%>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
<td style="text-align: center"><%=YYYY%></td>	

<%For MM = 1 to 12%>
<%if MM > 10 then MM= "0" & MM%>

<%If (YYYY/4) = Int(YYYY/4) and MM = 2 then Days=29 else Days=MDays(MM)%>

<%SDate=MM & "/01/" & YYYY%>
<%EDate=MM & "/" & Days & "/" & YYYY%>

<%set Permits=conntemp.execute("Select Count(ID) from tbl_permit where I_Date Between '" & SDate & "' and '" & Edate & "';")%>
<td style="text-align: center"><%=Permits(0)%>&nbsp;</td>
<%Total=Total+Permits(0)%>
<%YTotal=YTotal+Permits(0)%>
<%Next%>
<td style="text-align: center"><%=Total%></td>
<%Total=0%>
</tr>


<%Next%>

	<b>Total Permits:</b><%=TP(0)%>  <b>Permits with correct Issue Date:</b><%=YTotal%>
<%
conntemp.close
set conntemp=nothing
%>
</table>	
</body>

</html>