<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>



<form name="the_form">
<select name="choose_category" onChange="swapOptions(window.document.the_form.choose_category.options[selectedIndex].text);">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select EID,E_Name from tbl_Engineer order by E_Name ASC;")
%>
<%do until rstemp.eof %>
<option value="<%=rstemp(0)%>"><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</select>
<select name="the_examples">
<%for X=1 to 5%>
<%Response.write "<option>"%>
<%Next%>
</select>
</form>

<%Dim Contact_Array(100)%>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select E_Contact from tbl_Engineer;")
%>
<%do until rstemp.eof %>
<%X=X+1%>
<%Contact_Array(X)=rstemp(0)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%for I = 1 to X%>
<%Dim TempArray
TempArray=Split(Contact_Array(I),",")
%>
<%For Q = 0 to Ubound(TempArray)%>
<%NewContact=NewContact & "'" & TempArray(Q) & "',"%>
<%Next%>
<%Next%>
<%Newcontact=Left(NewContact,Len(Newcontact)-1)%>
<script language = "JavaScript">

<!-- hide me

var dogs = new Array(<%=NewContact%>);
var fish = new Array("trout", "mackerel", "bass");
var birds = new Array("robin", "hummingbird", "crow");

function swapOptions(the_array_name)
{
	var numbers_select = window.document.the_form.the_examples;
	var the_array = eval(the_array_name);
	setOptionText(window.document.the_form.the_examples, the_array);
}

function setOptionText(the_select, the_array)
{
	for (loop=0; loop < the_select.options.length; loop++)
	{
		the_select.options[loop].text = the_array[loop];
	}
}

// show me -->

</script>
</body>

</html>
