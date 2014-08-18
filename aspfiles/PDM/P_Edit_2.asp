<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Update Project Header(2)</title>
<!--mstheme--><link rel="stylesheet" type="text/css" href="../_themes/lasa-portal/lasa1011.css"><meta name="Microsoft Theme" content="lasa-portal 1011, default">
</head>

<body>
<!--#include file="PDM_Check.inc"-->
<%If Request.form("B1")="Update" then UpdateData%>
<%If Request.form("B1")="Cancel" then Cancel%>
<!--#include file="inc/Update_Query.asp"-->

<% 
set conntemp=server.createobject("adodb.connection")
conntemp.open "DSN=PDM;"
set rstemp=conntemp.execute("Select D_contact,E_contact,contact_To,Res_Type,Res_Amount from tbl_Project where PID=" & PID & ";")
%>
<%do until rstemp.eof %>
<%D_Contact=rstemp(0)%>
<%E_Contact=rstemp(1)%>
<%P_Contact=rstemp(2)%>
<%Res_Type=rstemp(3)%>
<%Res_Amount=Rstemp(4)%>
<%
rstemp.movenext
loop

rstemp.close
set rstemp=nothing
conntemp.close
set conntemp=nothing
%>

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
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="P_Edit_2.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1" webbot-action="--WEBBOT-SELF--">
<div align="center">
<b><font size="3">Update Information
	</font></b>
	<table border="0" cellpadding="0" cellspacing="0" id="table1" style="border: 2px outset #0000FF; font-size:8pt">
		<tr>
			<td style="text-align: right" nowrap><b><%=D_Name%> - Developer Contact</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Developer Contact" b-disallow-first-item="TRUE" --><select size="1" name="D_Contact">
			<option>Select Contact</option>
			<%If IsArray(D_Array) <> "True" then Q=0 else Q=UBound(D_Array)%>
			<%For I = 0 to Q%>
			<%If Trim(D_Contact)=Trim(D_Array(I)) then sel="selected" else sel=""%>
			<option <%=sel%>><%=D_Array(I)%></option>
			<%Next%>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right" nowrap><b><%=E_Name%> - Engineer Contact</b></td>
			<td>
			<!--webbot bot="Validation" s-display-name="Engineering Contact" b-disallow-first-item="TRUE" --><select size="1" name="E_Contact">
			<option>Select Contact</option>
			<%If IsArray(E_Array) <> "True" then Q=0 else Q=UBound(E_Array)%>
			<%For I = 0 to Q%>
			<%If Trim(E_Contact)=Trim(E_Array(I)) then sel="selected" else sel=""%>
			<option <%=sel%>><%=E_Array(I)%></option>
			<%next%>
			</select></td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Primary Contact</b></td>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table2" style="font-size: 8pt">
					<tr><td><table cellspacing="0" cellpadding="0" width="100%">
						<td style="text-align: center">
						<%If P_Contact="Developer" then chk="checked" else chk=""%>&nbsp;&nbsp;<!--webbot bot="Validation" s-display-name="Primary Contact" b-value-required="TRUE" --><input type="radio" value="Developer" name="P_Contact" <%=chk%>></td>
						<td>Developer</td>
						<td style="text-align: center">
						<%If P_Contact="Engineer" then chk="checked" else chk=""%>
						<input type="radio" name="P_Contact" value="Engineer" <%=chk%>></td>
						<td>Engineer</td>
					</table></td></tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Reservation Type</b></td>
			<td>
			<div align="center">
				<table border="0" cellpadding="0" cellspacing="0" width="100%" id="table3" style="font-size: 8pt">
					<tr><td><table cellspacing="0" cellpadding="0" width="100%">
						<td style="text-align: center" width="34">
						<%if Res_Type="GPD" then chk="checked" else chk=""%>&nbsp;&nbsp;<!--webbot bot="Validation" s-display-name="Reservation Type" b-value-required="TRUE" --><input type="radio" name="Res_Type" value="GPD" <%=chk%>></td>
						<td>GPD</td>
						<td style="text-align: center">
						<%If Res_Type="IDU" then chk="checked" else chk=""%>
						<input type="radio" name="Res_Type" value="IDU" <%=chk%>></td>
						<td>IDU</td>
					</table></td></tr>
				</table>
			</div>
			</td>
		</tr>
		<tr>
			<td style="text-align: right"><b>Reserved Amount</b></td>
			<td>
			<input type="text" name="Res_Amount" size="4" value="<%=Res_Amount%>"></td>
		</tr>
		<tr>
			<td style="text-align: right">&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td style="text-align: right">
<input type="submit" value="Update" name="B1"><input type="submit" value="Cancel" name="B1"></td>
		</tr>
	</table>
</div>
<input type="hidden" name="PID" value="<%=PID%>">
</form>

<%Sub UpdateData%>
<%PID=Request.form("PID")%>
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
<%
set RS=nothing
conn.close
Set conn=nothing
%>
<p align="center"><b>Information has been updated</b></p>
<form method="POST" action="Project_Add_3.asp" name="Update" target="ibody">
        <input type="hidden" name="Refresh" Value="Updated">
        <input type="hidden" name="PID" Value="<%=PID%>">
</form>
	<script language="JavaScript">
		setTimeout("document.Update.submit();", 30);
		setTimeout("self.close();", 3000); 
    	</script>
<%response.end%>
<%End Sub%>
<%Sub Cancel%>
	   <script language="JavaScript">
		setTimeout("self.close();", 30);
    	</script>
<%Response.end%>
<%End Sub%>
</body>

</html>
