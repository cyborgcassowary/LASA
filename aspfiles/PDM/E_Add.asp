<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Add New Engineer</title>
<script language="JavaScript">
<!--
function closewinder()
{
 self.close();
}
//-->
</script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->
<%If Request.form("B1")="Submit" Then InsertData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<%HomeState="PA"%>
<div align="center">
<b><font size="3">Add New Engineer</font></b><hr width="80%">
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.E_Zip.value.length > 5)
  {
    alert("Please enter at most 5 characters in the \"Zip Code\" field.");
    theForm.E_Zip.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.E_Zip.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Zip Code\" field.");
    theForm.E_Zip.focus();
    return (false);
  }

  var checkOK = "0123456789-()-";
  var checkStr = theForm.E_Number.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only digit and \"()-\" characters in the \"Phone Number\" field.");
    theForm.E_Number.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="E_Add.asp" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellpadding="0" cellspacing="0" id="table1">
		<tr>
			<td style="text-align: right"><b>Name</b></td>
			<td colspan="3">
			<input type="text" name="E_Name" size="51"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Address</b></td>
			<td colspan="3"><input type="text" name="E_Street" size="51"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>City, State, Zip</b></td>
			<td><input type="text" name="E_City" size="20"></td>
			<td><Select name="E_State" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=SiteSettings;"
set rstemp=conntemp.execute("Select * from States;")
%>
<%do until rstemp.eof %>
<%If rstemp(0)=HomeState then Sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>" <%=sel%>><%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Zip Code" s-data-type="Integer" s-number-separators="x" i-maximum-length="5" --><input type="text" name="E_Zip" size="5" maxlength="5"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Phone Number</b></td>
			<td colspan="3">
			<!--webbot bot="Validation" s-display-name="Phone Number" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars="()-" --><input type="text" name="E_Number" size="20"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Contact Person</b></td>
			<td colspan="3"><input type="text" name="E_Contact" size="51"></td>
		</tr>
		<tr>
			<td>
<input type="submit" value="Submit" name="B1"><input type="submit" value="Cancel" name="B1"></td>

			<td colspan="3">
			<p style="text-align: center"><b><font size="1">Separate contact 
			persons by a comma</font></b></td>
		</tr>
	</table>
</div>
</form>

<%Sub InsertData%>
<%E_Name=Request.form("E_Name")%>
<%E_Street=Request.form("E_Street")%>
<%E_City=Request.form("E_city")%>
<%E_State=Request.form("E_State")%>
<%E_Zip=Request.form("E_Zip")%>
<%E_Number=Request.form("E_Number")%>
<%E_Contact=Request.form("E_Contact")%>
<%E_Name=Replace(E_Name,"'","&#039")%>
<%E_Street=Replace(E_Street,"'","&#039")%>
<%E_City=Replace(E_City,"'","&#039")%>
<%E_Contact=Replace(E_Contact,"'","&#039")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Insert Into tbl_Engineer (E_Name,E_Street,E_City,E_State,E_Zip,E_Number,E_Contact) Values ('" & E_Name & "','" & E_Street & "','" & E_City & "','" & E_State & "','" & E_Zip & "','" & E_Number & "','" & E_Contact & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center"><b>Engineer has been added to the database.</b></p>
<form method="POST" action="Project_Add.asp" name="Update" target="ibody">
<input type="hidden" name="E_Name" value="<%=E_Name%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<div align="center">
<a href="javascript:closewinder()">Close</a>
</div>
<%Response.end%>
<%End Sub%>
<%Sub Cancel%>
<form method="POST" action="Project_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>
