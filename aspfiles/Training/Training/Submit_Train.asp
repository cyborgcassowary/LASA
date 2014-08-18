<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Request for Training Form</title>
<!-- #include file="inc/calctl.inc" -->
<script language="JavaScript" src="inc/overlib_mini.js"></script>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<%If Request.form("B1")="Submit" and Request.Form("C_Date") <> "MM/DD/YYYY" then Insertdata%>
<%If Request.form("B1")="Submit" and Request.form("C_Date")="MM/DD/YYYY" then ErrorMsg%>
<%If Request.form("B1")="Cancel" then Cancel%>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.Category.selectedIndex == 0)
  {
    alert("The first \"Category\" option is not a valid selection.  Please choose one of the other options.");
    theForm.Category.focus();
    return (false);
  }

  var checkOK = "0123456789-.,";
  var checkStr = theForm.Reg.value;
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
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch == "," && decPoints != 0)
    {
      validGroups = false;
      break;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Registration\" field.");
    theForm.Reg.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"Reg\" field.");
    theForm.Reg.focus();
    return (false);
  }

  var checkOK = "0123456789-.,";
  var checkStr = theForm.Travel.value;
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
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch == "," && decPoints != 0)
    {
      validGroups = false;
      break;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Travel\" field.");
    theForm.Travel.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"Travel\" field.");
    theForm.Travel.focus();
    return (false);
  }

  var checkOK = "0123456789-.,";
  var checkStr = theForm.Lodge.value;
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
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch == "," && decPoints != 0)
    {
      validGroups = false;
      break;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Lodging\" field.");
    theForm.Lodge.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"Lodge\" field.");
    theForm.Lodge.focus();
    return (false);
  }

  var checkOK = "0123456789-.,";
  var checkStr = theForm.Meals.value;
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
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch == "," && decPoints != 0)
    {
      validGroups = false;
      break;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Meals\" field.");
    theForm.Meals.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"Meals\" field.");
    theForm.Meals.focus();
    return (false);
  }

  var checkOK = "0123456789-.,";
  var checkStr = theForm.Other.value;
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
    if (ch == ".")
    {
      allNum += ".";
      decPoints++;
    }
    else if (ch == "," && decPoints != 0)
    {
      validGroups = false;
      break;
    }
    else if (ch != ",")
      allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Other\" field.");
    theForm.Other.focus();
    return (false);
  }

  if (decPoints > 1 || !validGroups)
  {
    alert("Please enter a valid number in the \"Other\" field.");
    theForm.Other.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="Submit_Train.asp" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" webbot-action="--WEBBOT-SELF--">
<div align="center">
<div id="overDiv" style="position:absolute; visibility:hidden; z-index:1000;"></div>
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border:2px outset #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">
		<tr>
			<td style="text-align: right"><b>Employee Name:</b></td>
			<td colspan="3">
			<Select name="Employee" size="1">
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=SiteSettings;"
set rstemp=conntemp.execute("Select FullName,Jobtitle from Users order by LastName ASC;")
%>
<%do until rstemp.eof %>
<%If rstemp(0)=Request.Cookies("Portal")("Fullname") then sel="selected" else sel=""%>
<option value="<%=rstemp(0)%>,<%=rstemp(1)%>" <%=sel%>><%=rstemp(0)%>,&nbsp;<%=rstemp(1)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select>
</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Department:</b></td>
			<td colspan="3"><select size="1" name="Department">
			<option>Administration</option>
			<option>Laboratory</option>
			<option>Maintenance</option>
			<option>Treatment Plant</option>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Category:</b></td>
			<td colspan="3">
			<!--webbot bot="Validation" s-display-name="Category" b-disallow-first-item="TRUE" --><Select name="Category" size="1">
<option>Select Category</option>
<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=Training;"
set rstemp=conntemp.execute("Select Category from tbl_Category;")
%>
<%do until rstemp.eof %>
<option value="<%=rstemp(0)%>"><%=rstemp(0)%></option>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>
</Select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Course Name:</b></td>
			<td colspan="3"><input type="text" name="C_Name" size="60"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Course Location:</b></td>
			<td colspan="3"><input type="text" name="C_Location" size="60"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>City, State</b></td>
			<td colspan="3"><input type="text" name="CityState" size="60"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Course Sponsor</b></td>
			<td colspan="3"><input type="text" name="C_Sponsor" size="60"></td>
		</tr>
		<tr>
			<td style="text-align: right" title="Continuing Education Unit"><b>
			Contact Hours</b></td>
			<td colspan="3">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table4">
					<tr>
						<td width="21">
						<input type="radio" value="No" name="CEU_Credit" checked></td>
						<td><b>No</b></td>
						<td width="20">
						<input type="radio" value="Yes" name="CEU_Credit"></td>
						<td width="219"><b>Yes</b></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Approved Program ID</b></td>
			<td colspan="3">
			<input type="text" name="Program_ID" size="40"></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Date</b></td>
			<td colspan="3">
			<input type="text" name="c_date" size="10" value="MM/DD/YYYY" readonly>
<a href="javascript:show_calendar('FrontPage_Form1.c_date');" onMouseOver="window.status='Date Picker'; overlib('Click here to choose a date from a one month pop-up calendar.'); return true;" onMouseOut="window.status=''; nd(); return true;">
<img src="../images/show-calendar.gif" width=20 height=19 border=0 align="top"></a>
			<b><font size="1" color="#FF0000">Required Field</font></b></td>
		</tr>
		<tr>
			<td style="text-align: right"><b># of Days</b></td>
			<td style="text-align: left">
			<input type="text" name="Days" size="10"></td>
			<td colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3">
					<tr>
						<td style="text-align: right"><b>Contact Hours</b></td>
						<td width="40">
						<input type="text" name="Contact_Hours" size="10"></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: right">&nbsp;</td>
			<td style="text-align: right">&nbsp;</td>
			<td colspan="2">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table6">
					<tr>
						<td style="text-align: right"><b>Attendance Hours</b></td>
						<td width="40">
						<input type="text" name="A_Hours" size="10"></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Costs:</b></td>
			<td style="text-align: right"><b>Registration:</b></td>
			<td colspan="2">
			&nbsp;<!--webbot bot="Validation" s-display-name="Registration" s-data-type="Number" s-number-separators=",." --><input type="text" name="Reg" size="6"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td style="text-align: right"><b>Travel:</b></td>
			<td colspan="2">
			&nbsp;<!--webbot bot="Validation" s-display-name="Travel" s-data-type="Number" s-number-separators=",." --><input type="text" name="Travel" size="6"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td style="text-align: right"><b>Lodging:</b></td>
			<td colspan="2">
			&nbsp;<!--webbot bot="Validation" s-display-name="Lodging" s-data-type="Number" s-number-separators=",." --><input type="text" name="Lodge" size="6"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td style="text-align: right"><b>Meals:</b></td>
			<td colspan="2">
			&nbsp;<!--webbot bot="Validation" s-display-name="Meals" s-data-type="Number" s-number-separators=",." --><input type="text" name="Meals" size="6"></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td style="text-align: right"><b>Other:</b></td>
			<td>
			&nbsp;<!--webbot bot="Validation" s-display-name="Other" s-data-type="Number" s-number-separators=",." --><input type="text" name="Other" size="6"></td>
			<td><input type="text" name="Other_Desc" size="30"></td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: center"><b>Course Description</b></td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: center">
			<textarea rows="5" name="C_Description" cols="80"></textarea></td>
		</tr>
		<tr>
			<td colspan="4">
			<p style="text-align: center"><b>Reason for Taking Course</b></td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: center">
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table5">
					<tr>
						<td width="6%" style="text-align: center">
						<input type="radio" name="Reason" value="Required by Supervisor"></td>
						<td colspan="2">Required by Supervisor</td>
					</tr>
					<tr>
						<td width="6%" style="text-align: center">
						<input type="radio" name="Reason" value="Self-Improvement"></td>
						<td colspan="2">Self-Improvement</td>
					</tr>
					<tr>
						<td width="6%" style="text-align: center">
						<input type="radio" name="Reason" value="Expand Knowledge"></td>
						<td colspan="2">Expand Knowledge</td>
					</tr>
					<tr>
						<td width="6%" style="text-align: center">
						<input type="radio" name="Reason" value="Learn Something New"></td>
						<td colspan="2">Learn Something New</td>
					</tr>
					<tr>
						<td width="6%" style="text-align: center">
						<input type="radio" name="Reason" value="Earn Credit Hours"></td>
						<td colspan="2">Earn Credit Hours</td>
					</tr>
					<tr>
						<td width="6%" style="text-align: center">
						<input type="radio" name="Reason" value="Other"></td>
						<td width="13%">Other</td>
						<td width="81%">
						<input type="text" name="Other_Reason" size="60"></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td colspan="4" style="text-align: left">
<div align="center">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2">
		<tr>
			<td width="48"><input type="hidden" name="Debug" value="0">
<input type="submit" value="Submit" name="B1"></td>
			<td width="58">&nbsp;</td>
			<td><font size="1">Once you complete this form, you will be asked to print<br>a paper copy to submit to your manager.</font></td>
		</tr>
	</table>
</div>
			</td>
		</tr>
	</table>
</div>
</form>

<%Sub InsertData%>
<%EmpInfo=Split(Request.form("Employee"),",")%>
<%Employee=EmpInfo(0)%>
<%JobTitle=EmpInfo(1)%>
<%Department=Request.form("Department")%>
<%C_Name=Request.form("C_name")%>
<%C_Name=Replace(C_Name,"'","&#039")%>
<%C_Sponsor=Request.form("C_sponsor")%>
<%C_Sponsor=Replace(C_Sponsor,"'","&#039")%>
<%C_location=Request.form("C_Location")%>
<%C_Location=Replace(C_Location,"'","&#039")%>
<%C_Date=Request.form("C_Date")%>
<%C_Date=Replace(C_DAte,"'","&#039")%>
<%C_description=Request.form("C_Description")%>
<%C_Description=Replace(C_Description,"'","&#039")%>
<%If Request.form("Reason")="Other" and Request.form("Other_Reason")<>"" then Reason=Request.form("Reason") & "-" & Request.form("Other_Reason") else Reason=Request.form("Reason")%>
<%Reason=Replace(Reason,"'","&#039")%>
<%Reg=Request.form("Reg")%>
<%If Reg="" then Reg=0%>
<%Travel=Request.form("Travel")%>
<%If Travel="" then Travel=0%>
<%Lodge=Request.form("Lodge")%>
<%If Lodge="" then Lodge=0%>
<%meals=Request.form("meals")%>
<%If meals="" then Meals=0%>
<%Other=Request.form("Other")%>
<%If Other="" then Other=0%>
<%Other_Desc=Request.form("Other_Desc")%>
<%Other_Desc=Replace(Other_Desc,"'","&#039")%>
<%Req_date=FormatDateTime(now)%>
<%Category=Request.form("Category")%>
<%CityState=Replace(Request.Form("CityState"),"'","&#039")%>
<%Complete="No"%>
<%Days=Request.form("Days")%>
<%If Days="" then Days=1%>
<%A_Hours=Request.form("A_Hours")%>
<%Contact_Hours=Request.form("Contact_Hours")%>
<%CEU_Credit=Request.form("CEU_Credit")%>
<%Program_ID=Request.form("Program_ID")%>
<%If Request.form("Debug")="1" then
Response.write "Employee:"
Response.write Employee
Response.write "<br>"
Response.write "Jobtitle:"
Response.write Jobtitle
Response.write "<br>"
Response.write "Department:"
Response.write Department
Response.write "<br>"
Response.write "Course Name:"
Response.write C_Name
Response.write "<br>"
Response.write "Course Sponsor:"
Response.write C_Sponsor
Response.write "<br>"
Response.write "Course Location:"
Response.write C_Location
Response.write "<br>"
Response.write "Course Date:"
Response.write C_Date
Response.write "<br>"
Response.write "Course Description:"
Response.write C_Desc
Response.write "<br>"
Response.write "Reason:"
Response.write Reason
Response.write "<br>"
Response.write "Registration:"
Response.write Reg
Response.write "<br>"
Response.write "Travel:"
Response.write Travel
Response.write "<br>"
Response.write "Meals:"
Response.write Meals
Response.write "<br>"
Response.write "Other:"
Response.write Other
Response.write "<br>"
Response.write "Other Description:"
Response.write Other_Desc
Response.write "<br>"
Response.write "Category:"
Response.write Category
Response.write "<br>"
Response.write "City/State:"
Response.write CityState
Response.write "<br>"
Response.write "Days:"
Response.write Days
Response.write "<br>"
Response.write "Contact Hours:"
Response.write Contact_Hours
Response.write "<br>"
Response.write "Attendance Hours:"
Response.write A_Hours
Response.write "<br>"
Response.write "Contact Hours Credit:"
Response.write CEU_Credit
Response.write "<br>"
Response.write "Approved Program ID:"
Response.write Program_ID
Response.write "<br>"
response.end
End If%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=Training;"%>
<%RS = Conn.Execute("Insert Into tbl_users (UserName,Department,C_name,C_sponsor,C_Location,C_date,C_description,Reason,Reg,Travel,Lodge,Meals,Other,Other_Desc,Req_Date,Category,CityState,Complete,A_Hours,Days,Jobtitle,CEU_Credit,Program_ID,Contact_Hours) Values ('" & Employee & "','" & Department & "','" & C_Name & "','" & C_Sponsor & "','" & C_Location & "','" & C_Date & "','" & C_Description & "','" & Reason & "','" & Reg & "','" & Travel & "','" & Lodge & "','" & Meals & "','" & Other & "','" & Other_Desc & "','" & Req_Date & "','" & Category & "','" & Citystate & "','" & Complete & "','" & A_Hours & "','" & Days & "','" & Jobtitle & "','" & CEU_Credit & "','" & Program_ID & "','" & Contact_Hours & "')")%>

<%
set RS=nothing
conn.close
Set conn=nothing
%>

<form method="POST" action="RFTC.asp" name="Update" target="_blank">
        <input type="hidden" name="Req_Date" Value="<%=Req_Date%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30); 
    	</script>
<p align="center">Course has been entered into the Database.</p>
<form method="POST" action="Con_Index.asp" name="Refresh" target="ibody">
</form>
	<script language="JavaScript">
		setTimeout("document.Refresh.submit();", 3000); 
    	</script>

<%Response.end%>
<%end Sub%>

<%Sub ErrorMsg%>
<% 
    msg = "Invalid Date Entered, Please Correct." 
    Response.Write("<" & "script>alert('" & msg & "');") 
    Response.Write("<" & "/script>") 
%>
<%End Sub%>
</body>

</html>