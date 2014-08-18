<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>On-Lot Septic Report</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<script language="javascript" type="text/javascript">
<!--
function OLS_Delete(id)
{
window.open('OLS_Delete.asp?ID=' + id +'','', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=0,width=500,height=300,left = 250,top = 200');
}
//-->
</script>
<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0" style="border: 3px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
<caption><b><font size="4">On-Lot Septic Report</font></b></caption>
	<tr>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">&nbsp;</td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><b>Project&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"><b>Tax Parcel ID&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Status Change Date&nbsp;</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Septic Status</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Permit #</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Plant</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>EDU's</b></td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<strong>Build Date</strong></td>
	</tr>
	
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_OLS order by SDATE DESC;")
%>
<%do until rstemp.eof %>
<%If rstemp(1)="0" then
	GetAddress
	Else
	GetProject
End If%>


<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
		<td><a href="javascript:OLS_Delete(<%=rstemp(0)%>)"><img border="0" src="../images/dir_delete.gif" width="16" height="16"></a></td>
		<td><%=Session("Project")%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(3)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(5)%>&nbsp;</td>
		<td style="text-align: center"><a href="Permit_Detail.asp?OLS=<%=rstemp(4)%>"><%=rstemp(4)%></a>&nbsp;</td>
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
<table width="100%">
<tr>
<td align="right"><!--#include file="inc/print.inc"--></td>
</tr>
</table>
<%Sub GetAddress%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select P_Number,P_Street,P_CityState from tbl_Permit where Permit=" & rstemp(4) & ";")
%>
<%do until rstemp1.eof %>
<%Session("Project")=rstemp1(0) & " " & rstemp1(1) & " " & rstemp1(2)%>

<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
<%End Sub%>

<%Sub GetProject%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select Project from tbl_Project where PID=" & rstemp(1) & ";")
%>
<%do until rstemp1.eof %>
<%Session("Project")=rstemp1(0)%>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
<%End Sub%>


</body>

</html>