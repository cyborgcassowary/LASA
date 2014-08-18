<%Response.buffer=True%>
<%response.ContentType="application/vnd.ms-excel"%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">


<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 11">

<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>Brian Burgoon</o:Author>
  <o:LastAuthor>Brian Burgoon</o:LastAuthor>
  <o:LastPrinted>2006-09-11T16:21:29Z</o:LastPrinted>
  <o:Created>2006-03-01T16:17:28Z</o:Created>
  <o:LastSaved>2006-09-11T16:22:00Z</o:LastSaved>
  <o:Company>Lancaster Area Sewer Authority</o:Company>
  <o:Version>11.8036</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<!--#include file="inc/RFTC_Header.inc"-->
<!--#include file="inc/Excel_XML.inc"-->
</head>

<body link=blue vlink=purple>
<!--#include file="inc/RFTC_Data.asp"-->
<table x:str border=0 cellpadding=0 cellspacing=0 width=640 style='border-collapse:
 collapse;table-layout:fixed;width:480pt'>
 <col width=64 span=10 style='width:48pt'>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 rowspan=3 height=51 class=xl39 width=128 style='height:38.25pt;
  width:96pt'><img src="http://intranet/images/lasa-logo.gif" width="134" height="52"></td>
  <td colspan=7 rowspan=2 class=xl38 width=448 style='width:336pt'>REQUEST FOR
  TRAINING/CONFERENCE</td>
  <td width=64 style='width:48pt'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 style='height:12.75pt'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=7 height=17 class=xl40 style='height:12.75pt'><%=Req_Date%></td>
  <td></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=9 class=xl41 style='height:12.75pt;mso-ignore:colspan'></td>
  <td class=xl42></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Employee Name:</td>
  <td colspan=3 class=xl48><%=C5%></td>
  <td colspan=2 class=xl49>Job Title:</td>
  <td colspan=3 class=xl48><%=H5%></td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 colspan=9 class=xl41 style='height:3.75pt;mso-ignore:colspan'></td>
  <td class=xl42></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Department:</td>
  <td colspan=3 class=xl48><%=C7%></td>
  <td colspan=2 class=xl50>Course Category:</td>
  <td colspan=3 class=xl51><%=H7%></td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 colspan=9 class=xl41 style='height:3.75pt;mso-ignore:colspan'></td>
  <td class=xl42></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Course Name:</td>
  <td colspan=8 class=xl48><%=C9%></td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 colspan=9 class=xl41 style='height:3.75pt;mso-ignore:colspan'></td>
  <td class=xl42></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Course Sponsor:</td>
  <td colspan=8 class=xl48><%=C11%></td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 colspan=9 class=xl41 style='height:3.75pt;mso-ignore:colspan'></td>
  <td class=xl42></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Course Location:</td>
  <td colspan=3 class=xl48><%=C13%></td>
  <td colspan=2 class=xl50>City, State:</td>
  <td colspan=3 class=xl51><%=H13%></td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 colspan=9 class=xl41 style='height:3.75pt;mso-ignore:colspan'></td>
  <td class=xl42></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Course Date(s):</td>
  <td colspan=2 class=xl53><%=C15%></td>
  <td class=xl78>&nbsp;</td>
  <td class=xl55>&nbsp;</td>
  <td colspan=2 class=xl52>Attendance Hours:</td>
  <td colspan=2 class=xl55><%=I15%></td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 colspan=2 class=xl41 style='height:3.75pt;mso-ignore:colspan'></td>
  <td colspan=8 class=xl44 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Approved Program
  ID:</td>
  <td colspan=3 class=xl55><%=C17%></td>
  <td class=xl44></td>
  <td colspan=2 class=xl52>Contact Hours:</td>
  <td colspan=2 class=xl55><%=I17%></td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 colspan=2 class=xl41 style='height:3.75pt;mso-ignore:colspan'></td>
  <td colspan=8 class=xl44 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Reason for Course:</td>
  <td colspan=8 class=xl55><%=C19%></td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 colspan=9 class=xl41 style='height:3.75pt;mso-ignore:colspan'></td>
  <td class=xl42></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl56 style='height:12.75pt'>Course Description:</td>
  <td colspan=8 rowspan=4 class=xl57 style='border-right:.5pt solid black;
  border-bottom:.5pt solid black'><%=C21%></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=2 class=xl45 style='height:12.75pt;mso-ignore:colspan'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=2 class=xl46 style='height:12.75pt;mso-ignore:colspan'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=2 class=xl41 style='height:12.75pt;mso-ignore:colspan'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=2 class=xl41 style='height:12.75pt;mso-ignore:colspan'></td>
  <td colspan=8 class=xl44 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl66 style='height:12.75pt'>&nbsp;</td>
  <td class=xl67>&nbsp;</td>
  <td class=xl68>&nbsp;</td>
  <td class=xl68>&nbsp;</td>
  <td class=xl68>&nbsp;</td>
  <td class=xl68>&nbsp;</td>
  <td class=xl68>&nbsp;</td>
  <td class=xl68>&nbsp;</td>
  <td class=xl68>&nbsp;</td>
  <td class=xl69>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl70 style='height:12.75pt'>Course Costs:</td>
  <td colspan=2 class=xl49>Registration:</td>
  <td class=xl48><%=FormatCurrency(E26)%></td>
  <td colspan=4 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 class=xl72 style='height:3.75pt'>&nbsp;</td>
  <td class=xl43></td>
  <td class=xl41></td>
  <td colspan=2 class=xl43 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl47 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl73 style='height:12.75pt'>&nbsp;</td>
  <td class=xl41></td>
  <td colspan=2 class=xl49>Travel:</td>
  <td class=xl48><%=FormatCurrency(E28)%></td>
  <td colspan=4 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 class=xl73 style='height:3.75pt'>&nbsp;</td>
  <td colspan=2 class=xl41 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl43 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl47 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl73 style='height:12.75pt'>&nbsp;</td>
  <td class=xl41></td>
  <td colspan=2 class=xl49>Lodging:</td>
  <td class=xl48><%=FormatCurrency(E30)%></td>
  <td colspan=4 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 class=xl73 style='height:3.75pt'>&nbsp;</td>
  <td colspan=2 class=xl41 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl43 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl47 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl73 style='height:12.75pt'>&nbsp;</td>
  <td class=xl41></td>
  <td colspan=2 class=xl49>Meals:</td>
  <td class=xl48><%=FormatCurrency(E32)%></td>
  <td colspan=4 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 class=xl73 style='height:3.75pt'>&nbsp;</td>
  <td colspan=2 class=xl41 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl43 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl47 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl73 style='height:12.75pt'>&nbsp;</td>
  <td class=xl41></td>
  <td colspan=2 class=xl49>Other:</td>
  <td class=xl48><%=FormatCurrency(E34)%></td>
  <td class=xl49>Desc:</td>
  <td colspan=3 class=xl48><%=G34%></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=5 style='mso-height-source:userset;height:3.75pt'>
  <td height=5 class=xl73 style='height:3.75pt'>&nbsp;</td>
  <td colspan=2 class=xl41 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl43 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl47 style='mso-ignore:colspan'></td>
  <td class=xl41></td>
  <td class=xl42></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl73 style='height:12.75pt'>&nbsp;</td>
  <td class=xl41></td>
  <td colspan=2 class=xl49>Total:</td>
  <td class=xl48><%=FormatCurrency(E36)%></td>
  <td colspan=4 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl71>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl74 style='height:12.75pt'>&nbsp;</td>
  <td class=xl48>&nbsp;</td>
  <td class=xl48>&nbsp;</td>
  <td class=xl48>&nbsp;</td>
  <td class=xl75>&nbsp;</td>
  <td class=xl76>&nbsp;</td>
  <td class=xl48>&nbsp;</td>
  <td class=xl48>&nbsp;</td>
  <td class=xl48>&nbsp;</td>
  <td class=xl77>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=4 class=xl41 style='height:12.75pt;mso-ignore:colspan'></td>
  <td class=xl43></td>
  <td colspan=4 class=xl41 style='mso-ignore:colspan'></td>
  <td class=xl42></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl49 style='height:12.75pt'>Employee Signature:</td>
  <td colspan=8 class=xl29>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=2 class=xl25 style='height:12.75pt;mso-ignore:colspan'></td>
  <td class=xl24></td>
  <td colspan=5 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl84>Date</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=2 class=xl25 style='height:12.75pt;mso-ignore:colspan'></td>
  <td class=xl24></td>
  <td colspan=5 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl31 style='mso-ignore:colspan'></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=10 rowspan=2 height=34 class=xl32 width=640 style='height:25.5pt;
  width:480pt'>IF THIS COURSE IS FOR CONTACT HOURS CREDIT,<br>YOU MUST PROVIDE PROOF OF COMPLETION TO YOUR MANAGER</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 colspan=9 class=xl24 style='height:12.75pt;mso-ignore:colspan'></td>
  <td></td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=10 height=17 class=xl35 style='border-right:.5pt solid black;
  height:12.75pt'>Management Use Only</td>
 </tr>
 <tr height=7 style='mso-height-source:userset;height:5.25pt'>
  <td height=7 class=xl26 style='height:5.25pt'>&nbsp;</td>
  <td colspan=8 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl79 style='height:12.75pt'>Training Budget:</td>
  <td colspan=2 class=xl29>&nbsp;</td>
  <td colspan=5 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=7 style='mso-height-source:userset;height:5.25pt'>
  <td height=7 class=xl26 style='height:5.25pt'>&nbsp;</td>
  <td colspan=8 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl79 style='height:12.75pt'>Expenses (YTD):</td>
  <td colspan=2 class=xl29>&nbsp;</td>
  <td colspan=5 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=7 style='mso-height-source:userset;height:5.25pt'>
  <td height=7 class=xl26 style='height:5.25pt'>&nbsp;</td>
  <td colspan=8 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl79 style='height:12.75pt'>Remaining Budget:</td>
  <td colspan=2 class=xl29>&nbsp;</td>
  <td colspan=5 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl26 style='height:12.75pt'>&nbsp;</td>
  <td colspan=8 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 rowspan=2 height=34 class=xl81 width=128 style='height:25.5pt;
  width:96pt'>Approved by<br>
    Supervisor:</td>
  <td colspan=7 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=8 height=17 class=xl29 style='border-right:.5pt solid black;
  height:12.75pt'>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl26 style='height:12.75pt'>&nbsp;</td>
  <td colspan=7 class=xl24 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl85 style='border-right:.5pt solid black'>Date</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl26 style='height:12.75pt'>&nbsp;</td>
  <td colspan=8 class=xl24 style='mso-ignore:colspan'></td>
  <td class=xl27>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td colspan=2 height=17 class=xl79 style='height:12.75pt'>Executive Director:</td>
  <td colspan=8 class=xl29 style='border-right:.5pt solid black'>&nbsp;</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl26 style='height:12.75pt'>&nbsp;</td>
  <td colspan=7 class=xl24 style='mso-ignore:colspan'></td>
  <td colspan=2 class=xl85 style='border-right:.5pt solid black'>Date</td>
 </tr>
 <tr height=17 style='height:12.75pt'>
  <td height=17 class=xl28 style='height:12.75pt'>&nbsp;</td>
  <td class=xl29>&nbsp;</td>
  <td class=xl29>&nbsp;</td>
  <td class=xl29>&nbsp;</td>
  <td class=xl29>&nbsp;</td>
  <td class=xl29>&nbsp;</td>
  <td class=xl29>&nbsp;</td>
  <td class=xl29>&nbsp;</td>
  <td class=xl29>&nbsp;</td>
  <td class=xl30>&nbsp;</td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
  <td width=64 style='width:48pt'></td>
 </tr>
 <![endif]>
</table>

</body>

</html>
