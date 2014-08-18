<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Training Requirement Grid</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="inc/Train_Check.inc"-->
<%Access=Request.Cookies("Portal")("TN")%>
<%Dim Category(40)%>
<%Dim JobTitle(40)%>
<%Dim Suggest(40)%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select Category from tbl_Category;")
%>
<%do until rstemp.eof %>
<%Y=Y+1%>
<%Category(Y)=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%X=0%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select JobTitle,CID from tbl_Jobtitle;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%Jobtitle(X)=rstemp(0)%>
<%Suggest(X)=rstemp(1)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%Page=Request.QueryString("Page")%>
<div align="center">
<b><font size="3">Suggested Training Grid</font></b>
</div>
<table width="100%">
<tr>
<%If Page="Page1" then fntcolor="#FF0000" else fntcolor=""%>
<td align="center"><a style="color: <%=fntcolor%>" href="Train_Grid.asp?Page=Page1">Page 1</a></td>
<%If Page="Page2" then fntcolor="#FF0000" else fntcolor=""%>
<td align="center"><a style="color: <%=fntcolor%>" href="Train_Grid.asp?Page=Page2">Page 2</a></td>
<%If Page="Page3" then fntcolor="#FF0000" else fntcolor=""%>
<td align="center"><a style="color: <%=fntcolor%>" href="Train_Grid.asp?Page=Page3">Page 3</a></td>
</tr>
</table>

<%If Page="" then Page1%>
<%If Page="Page1" then Page1%>
<%If Page="Page2" then Page2%>
<%If Page="Page3" then Page3%>

<%Sub Page1%>
<table width="100%" cellspacing="0" cellpadding="0" style="border-style:solid; border-width:1px; padding:1px; font-family: Tahoma; font-size: 8pt" border="0">
<tr>
<td align="center" valign="middle" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">&nbsp;</td>
<%For I = 1 to 10%>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:center; border-left-style:solid">
<a style="font-size: 8pt" href="Train_Edit.asp?ID=<%=Jobtitle(I)%>"><%=Jobtitle(I)%></a></td>
<%Next%>
</tr>
<%For I= 1 to Y%>
<%If I/2=Int(I/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
<td><%=Category(I)%></td>
<%For Q= 1 to 10%>
<%If Int(Instr(Suggest(Q),Category(I))) > 0 then Req="X" else Req="-"%>
<td align="center" style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px"><%=Req%>&nbsp;</td>
<%next%>
</tr>
<%next%>
</table>
<%End Sub%>
<%Sub Page2%>
<table width="100%" cellspacing="0" cellpadding="0" style="border-style:solid; border-width:1px; padding:1px; font-family: Tahoma; font-size: 8pt" border="0">
<tr>
<td align="center" valign="middle" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">&nbsp;</td>
<%For I = 11 to 20%>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:center" nowrap>
<a style="font-size: 8pt" href="Train_Edit.asp?ID=<%=Jobtitle(I)%>"><%=Jobtitle(I)%></a></td>
<%Next%>
</tr>
<%For I= 1 to Y%>
<%If I/2=Int(I/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
<td><%=Category(I)%></td>
<%For Q= 11 to 20%>
<%If Int(Instr(Suggest(Q),Category(I))) > 0 then Req="X" else Req="-"%>
<td align="center" style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px"><%=Req%>&nbsp;</td>
<%next%>
</tr>
<%next%>
</table>
<%End Sub%>

<%Sub Page3%>
<table width="100%" cellspacing="0" cellpadding="0" style="border-style:solid; border-width:1px; padding:1px; font-family: Tahoma; font-size: 8pt" border="0">
<tr>
<td align="center" valign="middle" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">&nbsp;</td>
<%For I = 21 to 30%>
<td style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; text-align:center" nowrap>
<a style="font-size: 8pt" href="Train_Edit.asp?ID=<%=Jobtitle(I)%>"><%=Jobtitle(I)%></a></td>
<%Next%>
</tr>
<%For I= 1 to Y%>
<%If I/2=Int(I/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
<td><%=Category(I)%></td>
<%For Q= 21 to 30%>
<%If Int(Instr(Suggest(Q),Category(I))) > 0 then Req="X" else Req="-"%>
<td align="center" style="border-left-style: solid; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-width: 1px"><%=Req%>&nbsp;</td>
<%next%>
</tr>
<%next%>
</table>
<%End Sub%>
</body>

</html>