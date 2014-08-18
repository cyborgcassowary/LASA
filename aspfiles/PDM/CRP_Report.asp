<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Capacity Reserve Report as of <%=formatDateTime(Now,VbshortDate)%></title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="inc/print.inc"-->
<%If Len(Month(now)) = 1 then MZ="0" else MZ=""%>
<%If Len(Day(now)) = 1 then DZ="0" else DZ=""%>
<%Today=Year(now) & "-" & MZ & Month(Now) & "-" & DZ & Day(now)%>
<%=Today%>

<table width="100%" cellspacing="0" cellpadding="0" style="font-family: Tahoma; font-size: 8pt; border: 2px outset #0000FF">
	<tr>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table1">
			<tr>
				<td>
				<p style="text-align: right"><a href="CRP_Report.asp?FieldName=Project&SO=ASC">
				<img border="0" src="../images/ASC.gif" width="18" height="15"></a></td>
				<td>
				<p style="text-align: center"><b>Project</b></td>
				<td><a href="CRP_Report.asp?FieldName=Project&SO=DESC">
				<img border="0" src="../images/DESC.gif" width="18" height="15" align="left"></a></td>
			</tr>
		</table>
		</td>
		<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
		<b>Municipality</b>&nbsp;</td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px">
		<b>Original&nbsp;Units</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>
		Original Reserve Date&nbsp;</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table2">
			<tr>
				<td><a href="CRP_Report.asp?FieldName=Expire_Date&SO=ASC">
				<img border="0" src="../images/ASC.gif" width="18" height="15" align="right"></a></td>
				<td><b>Expire Date</b></td>
				<td><a href="CRP_Report.asp?FieldName=Expire_Date&SO=DESC">
				<img border="0" src="../images/DESC.gif" width="18" height="15"></a></td>
			</tr>
		</table>
		</td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>Expire Notice&nbsp;</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px"><b>
		Reserved Units&nbsp;</b></td>
		<td style="text-align: center; border-left-width:1px; border-right-width:1px; border-top-width:1px; border-bottom-style:solid; border-bottom-width:1px">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table3">
			<tr>
				<td><a href="CRP_Report.asp?FieldName=Status&SO=ASC">
				<img border="0" src="../images/ASC.gif" width="18" height="15" align="right"></a></td>
				<td><b>Status</b></td>
				<td><a href="CRP_Report.asp?FieldName=Status&SO=DESC">
				<img border="0" src="../images/DESC.gif" width="18" height="15"></a></td>
			</tr>
		</table>
		</td>
	</tr>
<%If Request.QueryString("Fieldname")="" then FieldName="Project" else FieldName=Request.QueryString("FieldName")%>
<%If Request.QueryString("SO")="" then SO="ASC" else SO=Request.QueryString("SO")%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_CRP order by '" & FieldName & "' " & SO & ";")
%>

<%do until rstemp.eof %>
<%set Project=conntemp.execute("Select Project,Res_Type,res_Amount,Municipality from tbl_project where PID=" & rstemp(1) & ";")%>
<%Set Muni=conntemp.execute("Select Municipality from tbl_Municipality where MID=" & Project(3) & ";")%>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<%Set IDUE=conntemp.execute("Select Sum(res_Amount) from tbl_crp where Res_Type='IDU' and Status='Expired';")%>
<%Set IDUC=conntemp.execute("Select Sum(res_Amount) from tbl_crp where Res_Type='IDU' and Status='Active';")%>
<%Set GPDE=conntemp.execute("Select Sum(res_Amount) from tbl_crp where Res_Type='GPD' and Status='Expired';")%>
<%Set GPDC=conntemp.execute("Select Sum(res_Amount) from tbl_crp where Res_Type='GPD' and Status='Active';")%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
		<td><%=Project(0)%>&nbsp;</td>
		<td><%=Muni(0)%>&nbsp;</td>
		<td style="text-align: center"><%=Project(2)%>&nbsp;<%=Project(1)%>'s</td>
		<td style="text-align: center"><%If rstemp(2)="1/1/1990" then response.write "----" else response.write rstemp(2)%>&nbsp;</td>
		<td style="text-align: center"><%If rstemp(3)="1/1/1900" then response.write "----" else response.write rstemp(3)%>&nbsp;</td>
		<td style="text-align: center"><%If rstemp(5)="1/1/1900" then response.write "----" else response.write rstemp(5)%>&nbsp;</td>
		<td style="text-align: center"><%=rstemp(6)%>&nbsp;<%=rstemp(7)%>'s</td>
		<td style="text-align: center"><%=rstemp(8)%>&nbsp;</td>
	</tr>

<%
rstemp.movenext
loop%>
<tr>
<td colspan="7" style="border-left-width: 1px; border-right-width: 1px; border-top-style: solid; border-top-width: 1px; border-bottom-width: 1px">
IDU Expired:<%=IDUE(0)%><br>
IDU Active:<%=IDUC(0)%><br>
GPD Expired:<%=GPDE(0)%><br>
GPD Active:<%=GPDC(0)%><br>
</td>
</tr>
<%rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

</table>

</body>

</html>