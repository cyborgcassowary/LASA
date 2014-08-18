<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=FS;"
set rstemp1=conntemp1.execute("Select ID,Filename from tbl_Files where Category='" & Request.QueryString("CAT") & "';")
%>
<%do until rstemp1.eof %>
<a href="FS_Edit.asp?ID=<%=rstemp1(0)%>&FN=Cat_List.asp" target="ibody"><%=rstemp1(1)%></a><br>
<%
rstemp1.movenext
loop

rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>
</body>

</html>
