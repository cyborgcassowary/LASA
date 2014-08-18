<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Edit Developer</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->
<%ID=Request.QueryString("ID")%>

<%If Request.form("B1")="Update" then UpdateData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<% 
set conntemp1=server.createobject("adodb.connection")
conntemp1.open "DSN=PDM;"
set rstemp1=conntemp1.execute("Select * from tbl_Developer where DID=" & ID & ";")
%>
<div align="center">
<b><font size="3">Update Existing Developer</font></b><hr width="80%">
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.D_Name.value == "")
  {
    alert("Please enter a value for the \"Name\" field.");
    theForm.D_Name.focus();
    return (false);
  }

  if (theForm.D_Zip.value.length > 5)
  {
    alert("Please enter at most 5 characters in the \"Zip Code\" field.");
    theForm.D_Zip.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.D_Zip.value;
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
    theForm.D_Zip.focus();
    return (false);
  }

  var checkOK = "0123456789-()-";
  var checkStr = theForm.D_Number.value;
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
    theForm.D_Number.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="D_Edit.asp" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellpadding="0" cellspacing="0" id="table1">
		<tr>
			<td style="text-align: right"><b>Name</b></td>
			<td colspan="3">
			<!--webbot bot="Validation" s-display-name="Name" s-data-type="String" b-value-required="TRUE" --><input type="text" name="D_Name" size="51" value="<%=rstemp1(1)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Address</b></td>
			<td colspan="3"><input type="text" name="D_Street" size="51" value="<%=rstemp1(2)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>City, State, Zip</b></td>
			<td><input type="text" name="D_City" size="20" value="<%=rstemp1(3)%>"></td>
			<td><Select name="D_State" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=SiteSettings;"
set rstemp=conntemp.execute("Select * from States;")
%>
<%do until rstemp.eof %>
<%If rstemp(0)=rstemp1(4) then Sel="selected" else sel=""%>
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
			<!--webbot bot="Validation" s-display-name="Zip Code" s-data-type="Integer" s-number-separators="x" i-maximum-length="5" --><input type="text" name="D_Zip" size="5" maxlength="5" value="<%=rstemp1(5)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Phone Number</b></td>
			<td colspan="3">
			<!--webbot bot="Validation" s-display-name="Phone Number" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars="()-" --><input type="text" name="D_Number" size="20" value="<%=rstemp1(6)%>"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Contact Person</b></td>
			<td colspan="3"><input type="text" name="D_Contact" size="51" value="<%=rstemp1(7)%>"></td>
		</tr>
		<tr>
			<td>
<input type="submit" value="Update" name="B1"><input type="submit" value="Cancel" name="B1">
<input type="hidden" name="DID" value="<%=ID%>">
</td>
			<td colspan="3">
			<p style="text-align: center"><b><font size="1">Separate contact 
			persons by a comma</font></b></td>
		</tr>
	</table>
</div>
</form>
<%
rstemp1.close
set rstemp1=nothing
conntemp1.close
set conntemp1=nothing
%>

<%Sub UpdateData%>
<%DID=Request.form("DID")%>
<%D_Name=Request.form("D_Name")%>
<%D_Street=Request.form("D_Street")%>
<%D_City=Request.form("D_city")%>
<%D_State=Request.form("D_State")%>
<%D_Zip=Request.form("D_Zip")%>
<%D_Number=Request.form("D_Number")%>
<%D_Contact=Request.form("D_Contact")%>
<%D_Name=Replace(D_Name,"'","&#039")%>
<%D_Street=Replace(D_Street,"'","&#039")%>
<%D_City=Replace(D_City,"'","&#039")%>
<%D_Contact=Replace(D_Contact,"'","&#039")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Developer Set D_Name='" & D_Name & "', D_Street='" & D_Street & "', D_City='" & D_City & "', D_State='" & D_State & "', D_Zip='" & D_Zip & "', D_Number='" & D_Number & "', D_Contact='" & D_Contact & "' Where DID=" & DID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center"><b>Developer has been updated in the database.</b></p>
<form method="POST" action="D_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 3000); 
    	</script>
<%Response.end%>
<%End Sub%>
<%Sub Cancel%>
<form method="POST" action="D_Index.asp" name="Update" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End Sub%>

</body>

</html>
