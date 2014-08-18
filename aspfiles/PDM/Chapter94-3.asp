<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Chapter 94 PumpStation Report</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%For I = 1 to 3%>
<%If I = 1 then CType="R":Customer="Residential"%>
<%If I = 2 then CType="C":Customer="Commercial"%>
<%If I = 3 then CType="I":Customer="Industrial"%>
<table border="0" width="100%" cellpadding="0" style="border:2px outset #0000FF; border-collapse: collapse" id="table1">
<% 
set Munitemp=server.createobject("adodb.connection")
munitemp.open "DSN=PDM;"
set muni=munitemp.execute("Select PS_Name,Station_Number from tbl_PS;")
%>
<%do until muni.eof %>
	<tr>
		<td colspan="2" style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><%=muni(0)%>&nbsp;(<%=Muni(1)%>)</td>
		<td style="border-left-width: 1px; border-right-style: solid; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><%=Customer%>&nbsp;<%=Int(year(now))-1%></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="75">
		<b>Account&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="75">
		<b>Permit&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Development&nbsp;</b></td>
		<td colspan="2" style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Address</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="50">
		<b>PS&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="50">
		<b>Taps&nbsp;</b></td>
		
	</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Account,Permit,Development,P_Number,P_Street,DS_Pump,Taps from tbl_Permit where DS_Pump ='" & muni(1) & "' and Right(I_date,4) = Int(Year(Now))-1 and C_Type='" & CType & "' order by Development ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
		<td style="text-align: center" width="75"><%=rstemp(0)%>&nbsp;</td>
		<td style="text-align: center" width="75"><%=rstemp(1)%>&nbsp;</td>
		<td nowrap><%=rstemp(2)%>&nbsp;</td>
		<td colspan="2"><%=rstemp(3)%>&nbsp;<%=rstemp(4)%></td>
		<td style="text-align: center" width="50"><%=rstemp(5)%>&nbsp;</td>
		<td style="text-align: center" width="50"><%=rstemp(6)%>&nbsp;</td>
	</tr>
<%If rstemp(6) <> "" then TotalTaps=TotalTaps+Int(rstemp(6))%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<tr>
<td colspan="6" style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px; text-align:right" height="18">&nbsp;<b>Total Units:</b></td>
<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px" width="50"><%=TotalTaps%>&nbsp;</td>
</tr>
<%TotalTaps=0%>
<%
muni.movenext
loop

muni.close
set muni=nothing
munitemp.close
set munitemp=nothing
%>
</table>
<%Next%>
</body>

</html>