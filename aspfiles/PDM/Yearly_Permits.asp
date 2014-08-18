<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%For I = 1974 to Year(now)%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%Set RS = Conn.Execute("Select Count(ID) From tbl_Permit where Right(I_Date,4) = '" & I & "';")%>
<%=I%>:<%=RS(0)%><br>
<%Total=Total+RS(0)%>
<%RS=Conn.Close%>
<%next%>
<%=Total%>
</body>

</html>
