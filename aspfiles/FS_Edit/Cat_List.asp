<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<table border="0" width="100%" height="100%" cellspacing="0" cellpadding="0" id="table1">
<tr><td colspan="2" height="20" style="text-align: center"><b>Category Listing</b></td></tr>
	<tr>
		<td width="200" valign="top"><br><% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=FS;"
set rstemp=conntemp.execute("Select name from tbl_category;")
%>
<%do until rstemp.eof %>
<a href="File_List.asp?Cat=<%=rstemp(0)%>" target="flist"><%=rstemp(0)%></a><br>

<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%></td>
		<td><iframe frameborder="no" height="100%" width="100%" src="file_list.asp" name="flist"></iframe></td>
	</tr>
</table>





</body>

</html>