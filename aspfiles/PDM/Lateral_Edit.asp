<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Lateral Edit</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<%If Request.QueryString("Permit")<>"" then GoEdit%>
<%If Request.form("B1")="Update" then UpdateData%>
<%If Request.form("B1")="Cancel" then Cancel%>

<form method="POST" action="Lateral_Edit.asp" webbot-action="--WEBBOT-SELF--">
<div align="center">
<b><font size="3">Lateral Info Edit</font></b>
<table style="font-size: 8pt">
<tr>
<td colspan="2">Enter Permit Number to Edit</td>
</tr>
<tr>
	<td><input type="text" name="Permit" size="20" autocomplete="off"></td>
	<td><input type="submit" value="Submit" name="B1"></td>
</tr>
</table>
</div>
</form>
<%If Request.form("B1")="Submit" and Request.form("Permit") <> "" then GoEdit%>

<%Sub GoEdit%>
<%Permit=Request.form("Permit")%>
<%If Permit="" then
		Permit=Request.QueryString("Permit")
		RetPage=Request.QueryString("RetPage")
	Else
		RetPage="Lateral_Edit.asp"
End if%>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form2_Validator(theForm)
{

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789-";
  var checkStr = theForm.Contract.value;
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
    alert("Please enter only letter and digit characters in the \"Contract\" field.");
    theForm.Contract.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789-";
  var checkStr = theForm.Page.value;
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
    alert("Please enter only letter and digit characters in the \"Page\" field.");
    theForm.Page.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="Lateral_Edit.asp" onsubmit="return FrontPage_Form2_Validator(this)" language="JavaScript" name="FrontPage_Form2" webbot-action="--WEBBOT-SELF--">
<%set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,ParcelID,ID,Contract,Map_Page from tbl_Permit where Permit=" & Permit & ";")
%>
<%do until rstemp.eof %>
<div align="center">
<table cellspacing="0" cellpadding="0" style="border: 2px outset #0000FF" width="215">
<caption>Lateral Information for Permit# <%=Permit%></caption>
<tr>
<td style="font-size: 8pt; text-align: right" colspan="2">Lateral Station</td>
<td colspan="2"><input type="text" name="Lat_Sta" size="20" value="<%=rstemp(0)%>" autocomplete="off"></td>
</tr>
<tr>
<td style="font-size: 8pt; text-align: right" colspan="2">Lateral Depth</td>
<td colspan="2"><input type="text" name="Lateral_D" size="20" value="<%=rstemp(1)%>" autocomplete="off"></td>
</tr>
<tr>
<td style="font-size: 8pt; text-align: right" colspan="2">Lateral Length</td>
<td colspan="2"><input type="text" name="Lateral_L" size="20" value="<%=rstemp(2)%>" autocomplete="off"></td>
</tr>
<tr>
<td style="font-size: 8pt; text-align: right" colspan="2">Downstream Manhole</td>
<td colspan="2"><input type="text" name="DS_Manhole" size="20" value="<%=rstemp(3)%>" autocomplete="off"></td>
</tr>
<tr>
<td style="font-size: 8pt; text-align: right" colspan="2">Upstream Manhole</td>
<td colspan="2"><input type="text" name="US_Manhole" size="20" value="<%=rstemp(4)%>" autocomplete="off"></td>
</tr>
<tr>
<td style="font-size: 8pt; text-align: right" colspan="2">Tax Parcel ID</td>
<td colspan="2"><input type="text" name="ParcelID" size="20" value="<%=rstemp(5)%>" autocomplete="off"></td>
</tr>
<tr>
<td style="text-align: right" width="84"><font size="1">Contract#</font></td>
<td width="19">
<!--webbot bot="Validation" s-display-name="Contract" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" -->
<input type="text" name="Contract" size="5" value="<%=rstemp(7)%>"></td>
<td width="70" style="text-align: right"><font size="1">Page</font></td>
<td width="38">
<!--webbot bot="Validation" s-display-name="Page" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" -->
<input type="text" name="Page" size="6" value="<%=rstemp(8)%>"></td>
</tr>
<tr>
<td colspan="4"><input type="submit" value="Update" name="B1"><input type="submit" value="Cancel" name="B1">
<input type="hidden" name="ID" value="<%=rstemp(6)%>">
<input type="hidden" name="RetPage" value="<%=RetPage%>">
<input type="hidden" name="Permit" value="<%=Permit%>">
</td>
</tr>
</table>
</div>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</form>
<%response.end%>
<%End Sub%>

<%Sub UpdateData%>
<%Permit=Request.form("Permit")%>
<%Lat_Sta=Request.form("Lat_Sta")%>
<%Lateral_D=Request.form("Lateral_D")%>
<%Lateral_L=Request.form("Lateral_L")%>
<%DS_Manhole=Request.form("DS_Manhole")%>
<%US_Manhole=Request.form("US_Manhole")%>
<%ParcelID=Request.form("ParcelID")%>
<%Lat_Sta=Replace(Lat_Sta,"'","''")%>
<%Lateral_D=replace(Lateral_D,"'","''")%>
<%Lateral_L=replace(lateral_L,"'","''")%>
<%Contract=Request.form("Contract")%>
<%Page=Request.form("Page")%>
<%ID=Request.form("ID")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Permit Set Lat_Sta='" & Lat_Sta & "', Lateral_D='" & Lateral_D & "', Lateral_L='" & Lateral_L & "', DS_Manhole='" & DS_Manhole & "', US_Manhole='" & US_Manhole & "', ParcelID='" & ParcelID & "', Contract='" & Contract & "', Map_Page='" & Page & "' Where ID=" & ID & ";")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<a align="center">Permit# <%=Permit%> has been updated.</a>
<%RetPage=request.form("retPage")%>
<form method="POST" action="<%=RetPage%>" name="Update">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 2000); 
    	</script>
<%Response.end%>
<%End Sub%>

<%Sub Cancel%>
<%RetPage=request.form("retPage")%>
<form method="POST" action="<%=RetPage%>" name="Update">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<%Response.end%>
<%End sub%>
</body>

</html>