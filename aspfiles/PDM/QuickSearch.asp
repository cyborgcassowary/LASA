<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>PDM Search</title>

<style type="text/css">
.style1 {
	margin-bottom: 0px;
}
</style>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>

<p style="text-align: center"><b><font size="3">Permit Quick Search</font></b> </p>

<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.SearchField.selectedIndex < 0)
  {
    alert("Please select one of the \"SearchField\" options.");
    theForm.SearchField.focus();
    return (false);
  }

  if (theForm.SearchStrg.value == "")
  {
    alert("Please enter a value for the \"Search Value\" field.");
    theForm.SearchStrg.focus();
    return (false);
  }

  if (theForm.SearchStrg.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Search Value\" field.");
    theForm.SearchStrg.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789- \t\r\n\f";
  var checkStr = theForm.SearchStrg.value;
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
    alert("Please enter only letter, digit and whitespace characters in the \"Search Value\" field.");
    theForm.SearchStrg.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="QuickSearch.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1" webbot-action="--WEBBOT-SELF--">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" id="table1">
		<tr>
			<td><b><font size="1">Search Field</font></b></td>
			<td><b><font size="1">Search Value</font></b></td>
		</tr>
		<tr>
			<td><!--webbot bot="Validation" b-value-required="TRUE" -->
			<select size="1" name="SearchField" class="style1">
<option value="Permit">Permit Number</option>
<option value="P_Street">Permit Street</option>
<option value="Owner">Owner Name</option>
<option value="O_Street">Owner Street</option>
<option value="App_Name">Applicant Name</option>
<option value="App_Address">Applicant Street</option>
<option>Development</option>
<option value="ParcelID">Tax Parcel ID</option>
<option value="Account">Account</option>
</select>&nbsp;</td>
			<td>
			<!--webbot bot="Validation" s-display-name="Search Value" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" b-value-required="TRUE" i-minimum-length="3" --><input type="text" name="SearchStrg" size="20" authocomplete="off">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="2" style="text-align: right">
<input type="submit" value="Submit" name="B1"></td>
		</tr>
	</table>
</div>
</form>
<%If Request.form("B1")="Submit" then GetData%>


<%sub GetData%>

<%SearchStrg=Request.form("SearchStrg")%>
<%SearchField=Request.form("SearchField")%>
<%Select Case SearchField
		Case "P_Street"
			F_Name="Permit Street"
			SqlStmt="Select ID,Account,Permit,P_Number,P_Street,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=12
		Case "Permit"
			F_Name="Permit Number"
			SqlStmt="Select ID,Account,Permit,P_Number,P_Street,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=12
		Case "Owner"
			F_Name="Owner Name"
			SqlStmt="Select ID,Account,Permit,Owner,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=11
		Case "O_Street"
			F_Name="Owner Street"
			SqlStmt="Select ID,Account,Permit,O_Number,O_Street,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=12
		Case "App_Name"
			F_Name="Applicant Name"
			SqlStmt="Select ID,Account,Permit,App_Name,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=11
		Case "App_Address"
			F_Name="Applicant Address"
			SqlStmt="Select ID,Account,Permit,App_Address,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=12
		Case "Development"
			F_Name="Development"
			Sqlstmt="Select ID,Account,Permit,Development,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=11
		Case "ParcelID"
			F_Name="ParcelID"
			Sqlstmt="Select ID,Account,Permit,ParcelID,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=11
		Case "Account"
			F_Name="Development"
			Sqlstmt="Select ID,Account,Permit,Development,Contract,Map_Page,Lat_Sta,Lateral_D,Lateral_L,DS_Manhole,US_Manhole,Inspection_Status from tbl_Permit where " & SearchField & " like '%" & SearchStrg & "%';"
			F=11
End Select%>

<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="120" rowspan="2">
			<b>Account#</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="120" rowspan="2">
			<b>Permit#</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; width: 200px;" rowspan="2">
			&nbsp;</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px" width="120" rowspan="2">
			<b>Asbuilt Link</b></td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; " colspan="6">
			<b>Lateral Info</b></td>
		</tr>
		<tr>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; font-size:8pt; width: 100px;">
			Station</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; font-size:8pt; width: 100px;">
			Depth</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; font-size:8pt; width: 120px;">
			Length</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; font-size:8pt; width: 130px;">
			Downstream</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; font-size:8pt; width: 130px;">
			Upstream</td>
			<td style="text-align: center; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px; font-size:8pt; width: 130px;">
			Ready for Inspection</td>
		</tr>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute(SqlStmt)
%>
<%do until rstemp.eof %>
<%If F=5 then LastVal=rstemp(3) & " " & rstemp(4) else LastVal=rstemp(3)%>
<%Select Case F
		Case "12"
			LastVal=rstemp(3) & " " & rstemp(4)
			Contract=rstemp(5)
			MapPage=rstemp(6)
			Lat_Sta=rstemp(7)
			Lateral_D=rstemp(8)
			Lateral_L=rstemp(9)
			DS_Manhole=rstemp(10)
			US_Manhole=rstemp(11)
			Inspection_Status=rstemp(12)
		Case "11"
			LastVal=rstemp(3)
			Contract=rstemp(4)
			MapPage=rstemp(5)
			Lat_Sta=rstemp(6)
			Lateral_D=rstemp(7)
			Lateral_L=rstemp(8)
			DS_Manhole=rstemp(9)
			US_Manhole=rstemp(10)
			Inspection_Status=rstemp(11)
End Select%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<%DisplayLink=Contract & "-" & MapPage%>
<%If Instr(Contract,"WA") then Contract=Replace(Contract,"WA","")%>
<%FileCheck="D:\Shared\Asbuilts\" & Contract & "\" & MapPage & ".tif"%>

<%
dim fs
set fs=Server.CreateObject("Scripting.FileSystemObject")%>
<%Link="//intranet/asbuilts/" & Contract & "/" & MapPage & ".tif"%>
<%NoLink="<font color=darkred><b>" & DisplayLink & " </b></font>"%>
<%X=X+1%>
<%If X/2=Int(X/2) then bgcolor="#C0C0C0" else bgcolor=""%>
<tr bgcolor="<%=bgcolor%>" onMouseOver="this.bgColor='white';" onMouseOut="this.bgColor='<%=bgcolor%>';">
			<td style="text-align: center" width="120"><a href="Permit_Detail.asp?X=<%=rstemp(0)%>"><%=rstemp(1)%></a>&nbsp;</td>
			<td style="text-align: center" width="120"><a href="Lateral_Edit.asp?Permit=<%=rstemp(2)%>&RetPage=QuickSearch.asp"><%=rstemp(2)%></a>&nbsp;</td>
			<td style="text-align: center; width: 200px;"><%=LastVal%>&nbsp;</td>
			<td style="text-align: center" width="120">
			<%if fs.FileExists(Filecheck)=true then Response.write "<b><a href=":Response.write link:Response.write " target=_blank>":Response.write DisplayLink:Response.write" </a></b>" else Response.write NoLink%>
			<%Set FS=nothing%>
			</td>
			<td style="text-align: center; width: 100px;"><%=Lat_Sta%>&nbsp;</td>
			<td style="text-align: center; width: 100px;"><%=Lateral_D%>&nbsp;</td>
			<td style="text-align: center; width: 120px;"><%=Lateral_L%>&nbsp;</td>
			<td style="text-align: center; width: 130px;"><%=DS_Manhole%>&nbsp;</td>
			<td style="text-align: center; width: 130px;"><%=US_Manhole%>&nbsp;</td>
			<td style="text-align: center; width: 130px;"><%=Inspection_Status%>&nbsp;</td>
			<td style="text-align: center; width: 130px;">&nbsp;</td>
		</tr>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
<%end Sub%>
	</table>
</div>

</body>

</html>