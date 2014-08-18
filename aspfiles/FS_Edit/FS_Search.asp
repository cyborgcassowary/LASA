<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>File Search</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<form method="POST" action="FS_Search.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td><b>Search:</b></td>
			<td><input type="text" name="SearchString" size="30"></td>
			<td><input type="radio" value="FileName" name="Field"></td>
			<td>File Name</td>
			<td><input type="radio" value="Notes" name="Field"></td>
			<td>Notes</td>
			<td>&nbsp;<input type="submit" value="Submit" name="B1"></td>
		</tr>
	</table>
</div>
</form>
<%If Request.form("B1")="Submit" and Len(Request.form("SearchString")) => 3 and Request.form("Field") <> "" then
	SearchResults
		else
	ErrorMsg
End If%>

<%Sub SearchResults%>
<%Field=Request.form("Field")%>
<%SearchString=Request.form("SearchString")%>

<div align="center">
<b><font size="4">Search Results
	</font></b>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>File Name</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Location&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Owner&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Category&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Archive Date&nbsp;</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
			<b>Destroy Date&nbsp;</b></td>
		</tr>


<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select ID,FileName,Location,Owner,Category,Archive_Date,Destroy_date from tbl_Files where " & Field & " Like '%" & SearchString & "%';")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td><a href="FS_Edit.asp?ID=<%=rstemp(0)%>"><%=rstemp(1)%></a>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(2)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(3)%>&nbsp;</td>
			<td style="text-align: center"><%=rstemp(4)%>&nbsp;</td>
			<td style="text-align: center" width="100"><%=rstemp(5)%>&nbsp;</td>
			<td style="text-align: center" width="100"><%=rstemp(6)%>&nbsp;</td>
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
<%End Sub%>

<%Sub ErrorMsg%>
<p align="center"><font color="#FF0000">Three Character minimum must be entered to search.<br>You must choose a field in which to Search.</font></p>
<%End Sub%>
</body>

</html>