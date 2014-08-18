<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>

<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td width="33%">&nbsp;</td>
    <td width="33%">Municipality&nbsp;</td>
    <td width="34%"># of Permits&nbsp;</td>
  </tr>


<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select * from tbl_Municipality;")
%>
<%do until rstemp.eof %>
<%set muni=conntemp.execute("Select count(id) from tbl_permit where municipality='" & rstemp(1) &"';")
%>
  <tr>
    <td width="33%"><%=rstemp(0)%>&nbsp;</td>
    <td width="33%"><%=rstemp(1)%>&nbsp;</td>
    <td width="34%"><%=muni(0)%>&nbsp;</td>
  </tr>
<%X=X+Muni(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<tr>
<td>&nbsp;</td>
<td>Total Counted Permits</td>
<td><%=x%>&nbsp;</td>
</tr>
</table>

</body>

</html>