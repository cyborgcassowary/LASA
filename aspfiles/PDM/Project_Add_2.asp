<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Project Add Step 2</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->
<!--#include file="inc/Project_Queries.asp"-->

<%If Request.form("B1")="Submit" then InsertData%>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.D_Contact.selectedIndex == 0)
  {
    alert("The first \"Developer Contact\" option is not a valid selection.  Please choose one of the other options.");
    theForm.D_Contact.focus();
    return (false);
  }

  if (theForm.E_Contact.selectedIndex == 0)
  {
    alert("The first \"Engineering Contact\" option is not a valid selection.  Please choose one of the other options.");
    theForm.E_Contact.focus();
    return (false);
  }

  var radioSelected = false;
  for (i = 0;  i < theForm.P_Contact.length;  i++)
  {
    if (theForm.P_Contact[i].checked)
        radioSelected = true;
  }
  if (!radioSelected)
  {
    alert("Please select one of the \"Primary Contact\" options.");
    return (false);
  }

  var radioSelected = false;
  for (i = 0;  i < theForm.Res_Type.length;  i++)
  {
    if (theForm.Res_Type[i].checked)
        radioSelected = true;
  }
  if (!radioSelected)
  {
    alert("Please select one of the \"Reservation Type\" options.");
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="Project_Add_2.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1" webbot-action="--WEBBOT-SELF--">
<div align="center">
<b><font size="3">Start a New Project (Step 2)
	</font></b>
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF">
		<tr>
			<td style="text-align: right"><b><%=D_Name%> - Developer Contact</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Developer Contact" b-disallow-first-item="TRUE" --><select size="1" name="D_Contact">
			<option>Select Contact</option>
			<%If IsArray(D_Array) <> "True" then Q=0 else Q=UBound(D_Array)%>
			<%For I = 0 to Q%>
			<option><%=D_Array(I)%></option>
			<%Next%>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b><%=E_Name%> - Engineer Contact</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Engineering Contact" b-disallow-first-item="TRUE" --><select size="1" name="E_Contact">
			<option>Select Contact</option>
			<%If IsArray(E_Array) <> "True" then Q=0 else Q=UBound(E_Array)%>
			<%For I = 0 to Q%>
			<option><%=E_Array(I)%></option>
			<%next%>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Primary Contact</b></td>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2" style="font-size: 8pt">
					<tr>
						<td style="text-align: center">
						<!--webbot bot="Validation" s-display-name="Primary Contact" b-value-required="TRUE" --><input type="radio" value="Developer" name="P_Contact"></td>
						<td>Developer</td>
						<td style="text-align: center">
						<input type="radio" name="P_Contact" value="Engineer"></td>
						<td>Engineer</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Reservation Type</b></td>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3" style="font-size: 8pt">
					<tr>
						<td style="text-align: center">
						<!--webbot bot="Validation" s-display-name="Reservation Type" b-value-required="TRUE" --><input type="radio" name="Res_Type" value="GPD"></td>
						<td>GPD</td>
						<td style="text-align: center">
						<input type="radio" name="Res_Type" value="IDU"></td>
						<td>IDU</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Reserved Amount</b></td>
			<td><input type="text" name="Res_Amount" size="4"></td>
		</tr>
		<tr>
			<td style="text-align: right">&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td style="text-align: right">
<input type="submit" value="Submit" name="B1"></td>
		</tr>
	</table>
</div>
</form>

<%Sub InsertData%>
<%D_Contact=Request.form("D_Contact")%>
<%D_Contact=Replace(D_Contact,"'","&#039")%>
<%E_Contact=Request.form("E_Contact")%>
<%E_Contact=Replace(E_Contact,"'","&#039")%>
<%P_Contact=Request.form("P_contact")%>
<%Res_Type=Request.form("Res_Type")%>
<%Res_Amount=Request.form("Res_Amount")%>
<%set conn=server.createobject("adodb.connection")
conn.open "DSN=PDM;"%>
<%RS = Conn.Execute("Update tbl_Project Set Res_Type='" & Res_Type & "', Res_Amount='" & Res_Amount & "', Contact_To='" & P_Contact & "', D_Contact='" & D_Contact & "', E_Contact='" & E_Contact & "' Where PID=" & Session("ProjectID") & ";")%>
<%RS = Conn.Execute("Insert Into ph_Capacity (PID) Values ('" & Session("ProjectID") & "')")%>
<%RS = Conn.Execute("Insert Into ph_development (PID) Values ('" & Session("ProjectID") & "')")%>
<%RS = Conn.Execute("Insert Into ph_Construction (PID) Values ('" & Session("ProjectID") & "')")%>
<%RS = Conn.Execute("Insert Into ph_Close (PID) Values ('" & Session("ProjectID") & "')")%>
<%RS = Conn.Execute("Insert Into ph_board (PID) Values ('" & Session("ProjectID") & "')")%>
<%RS = Conn.Execute("Insert Into ph_Indemn (PID) Values ('" & Session("ProjectID") & "')")%>
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<div align="center">
Project has been updated in the Database<br>
<a href="Project_Add_3.asp">Proceed</a> to Step 3
</div>
<%response.end%>
<%End Sub%>

</body>

</html>