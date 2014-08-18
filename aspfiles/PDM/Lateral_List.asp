<%Response.buffer=False%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td style="text-align: center"><a href="Lateral_List.asp?Data=Yes">Lateral Station Data</a></td>
			<td style="text-align: center"><a href="Lateral_List.asp?Data=No">No Lateral Station Data</a></td>
		</tr>
	</table>
</div>

<%If Request.QueryString("Data")="" then Data="Yes" else Data=Request.QueryString("Data")%>

<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			&nbsp;</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Station</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Depth&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Length&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>DownStream&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Upstream&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Parcel ID&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Permit#&nbsp;</b></td>
		</tr>

<%If Data="Yes" then
	SqlStmt="Select Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,ParcelID,Permit from tbl_Permit where IsNull(Lat_Sta)=FALSE order by Permit ASC;"
	Else
	SqlStmt="Select Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,ParcelID,Permit from tbl_Permit where IsNull(Lat_Sta) order by Permit ASC;"
End if%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute(SqlStmt)
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="text-align: center"><%=X%></td>
			<td style="text-align: center"><%=rstemp(0)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(1)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(3)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(4)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(5)%>&nbsp;</td>
			<td style="text-align: center"><a href="Lateral_Edit.asp?Permit=<%=rstemp(6)%>&RetPage=Lateral_list.asp"><%=rstemp(6)%></a>&nbsp;</td>
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