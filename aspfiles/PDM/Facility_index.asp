<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Facility Fee Index</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<div align="center">
<table style="border: 2px outset #0000FF" width="80%">
<tr>
<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Facility Name</b></td>
<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Fee</b></td>
</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Facility order by FacilityName ASC;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>">
<td><a href="Facility_Edit.asp?ID=<%=rstemp(0)%>&Facility=<%=rstemp(1)%>&Fee=<%=rstemp(2)%>"><%=rstemp(1)%></a></td>
<td align="center">$<%=rstemp(2)%>
</td>
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