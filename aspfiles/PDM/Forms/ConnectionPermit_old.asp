<%Response.buffer=True%>
<%response.ContentType="application/vnd.ms-excel"%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 11">
<link rel=Edit-Time-Data href="cp_files/editdata.mso">
<link rel=OLE-Object-Data href="cp_files/oledata.mso">
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>Brian Burgoon</o:Author>
  <o:LastAuthor>Brian Burgoon</o:LastAuthor>
  <o:LastPrinted>2006-06-21T12:07:08Z</o:LastPrinted>
  <o:Created>2006-02-01T19:58:05Z</o:Created>
  <o:LastSaved>2006-06-21T16:25:28Z</o:LastSaved>
  <o:Company>Lancaster Area Sewer Authority</o:Company>
  <o:Version>11.6568</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<!--#include file="Excel_Style.inc"-->
</head>

<body link=blue vlink=purple>
<!--#include file="Excel_Data.asp"-->
<table x:str border=0 cellpadding=0 cellspacing=0 width=704 style='border-collapse:
 collapse;table-layout:fixed;width:528pt'>
 <col class=xl24 width=64 span=11 style='width:48pt'>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=2 height=17 class=xl32 width=128 style='height:12.75pt;
  width:96pt'><a name="Print_Area">LASA Account No.</a></td>
  <td class=xl28 width=64 style='width:48pt'><%=C1%></td>
  <td colspan=4 class=xl24 width=256 style='width:192pt'></td>
  <td colspan=3 class=xl32 width=192 style='width:144pt'>Connection Permit No.</td>
  <td class=xl28 width=64 style='width:48pt'><%=K1%></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl24 style='height:12.75pt'></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl52 style='height:12.75pt'>LANCASTER AREA
  SEWER AUTHORITY</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl35 style='height:12.75pt'>APPLICATION FOR
  CONNECTION TO SEWER SYSTEM</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl35 style='height:12.75pt'>(717)299-4843</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td height=17 class=xl25 colspan=2 style='height:12.75pt;mso-ignore:colspan'>Permit
  Data</td>
  <td colspan=9 class=xl24></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=3 height=17 class=xl34 style='height:12.75pt'>Address of Property
  Served</td>
  <td colspan=8 class=xl45><%=D7%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=3 height=20 class=xl34 style='height:15.0pt'>Development Name</td>
  <td colspan=8 class=xl45><%=D8%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=3 height=20 class=xl34 style='height:15.0pt'>Names of Owner of
  Property</td>
  <td colspan=8 class=xl45><%=D9%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=3 height=20 class=xl34 style='height:15.0pt'>Address of Property
  Owner/s</td>
  <td colspan=8 class=xl45><%=D10%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=3 height=20 class=xl34 style='height:15.0pt'>and/or Billing
  Address</td>
  <td colspan=8 class=xl45><%=D11%></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 rowspan=4 height=68 class=xl44 width=704 style='height:51.0pt;
  width:528pt'>Permit expires twelve (12) months from the issued date appearing
  below.<span style='mso-spacerun:yes'>  </span>If the connection to LASA's
  sewer is not made by that time, the permit may be renewed at the rates then
  in effect.<span style='mso-spacerun:yes'>  </span>Credit will be given for
  the amount paid to LASA for this permit.</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=4 height=17 class=xl24 style='height:12.75pt'></td>
  <td class=xl26>Initial</td>
  <td class=xl24></td>
  <td colspan=2 class=xl26>Date</td>
  <td colspan=3 class=xl24></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl32 style='height:15.0pt'>Check One</td>
  <td class=xl24></td>
  <td class=xl27>Residential</td>
  <td class=xl28><%=E17%></td>
  <td class=xl24></td>
  <td class=xl27>Commercial</td>
  <td class=xl28><%=H17%></td>
  <td class=xl24></td>
  <td class=xl27>Industrial</td>
  <td class=xl28><%=K17%></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl24 style='height:12.75pt'></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=8 height=17 class=xl34 style='height:12.75pt'
  x:str="Commercial/Industrial Users your Tapping fee was based on an estimated water consumption of ">Commercial/Industrial
  Users your Tapping fee was based on an estimated water consumption of<span
  style='mso-spacerun:yes'> </span></td>
  <td class=xl28><%=I19%></td>
  <td colspan=2 class=xl37>gallons per month.</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl35 style='height:12.75pt'>This account will
  be reviewed within one year and annually thereafter and additional Tapping
  Fees may be charged.</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl35 style='height:12.75pt'></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=3 height=20 class=xl34 style='height:15.0pt'>Water Utility
  Serving Property</td>
  <td colspan=3 class=xl45><%=D22%></td>
  <td class=xl24></td>
  <td colspan=2 class=xl34>Number of IDUs</td>
  <td class=xl28><%=J22%></td>
  <td class=xl24></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=11 height=20 class=xl34 style='height:15.0pt'></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=6 height=17 class=xl34 style='height:12.75pt'>In consideration of
  the granting of this Permit, the undersigned agrees:</td>
  <td colspan=5 class=xl27></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td rowspan=7 height=127 class=xl24 style='height:95.25pt'></td>
  <td class=xl29>1.</td>
  <td colspan=5 class=xl50
  x:str="To accept and abide by all the provisions of the Ordinance of ">To
  accept and abide by all the provisions of the Ordinance of<span
  style='mso-spacerun:yes'> </span></td>
  <td colspan=4 class=xl45><%=H25%></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td rowspan=3 height=51 class=xl24 style='height:38.25pt'></td>
  <td colspan=9 rowspan=3 class=xl48 width=576 style='width:432pt'>concerning
  the sanitary sewer system constructed by LANCASTER AREA SEWER AUTHORITY,
  including subsequent amendments thereto or revisions thereof; the Rules and
  Regulations of LANCASTER AREA SEWER AUTHORITY and all other applicable
  Ordinances, Resolutions and Rules and Regulations now in effect or which may
  hereafter be adopted.</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl29 style='height:15.0pt'>2.</td>
  <td colspan=9 class=xl50>To maintain the service line at no expense to
  LANCASTER AREA SEWER AUTHORITY.</td>
 </tr>
 <tr height=19 style='mso-height-source:userset;height:14.25pt'>
  <td height=19 class=xl29 style='height:14.25pt'>3.</td>
  <td colspan=9 rowspan=2 class=xl48 width=576 style='width:432pt'>To notify
  LANCASTER AREA SEWER AUTHORITY when the service line is ready for inspection
  and connection to the lateral sewer, <font class="font9">which notification
  shall be given before any portion of the work is covered.</font></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td height=17 class=xl24 style='height:12.75pt'></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl24 style='height:12.75pt'></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl24 style='height:12.75pt'></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24 style='height:15.0pt'></td>
  <td colspan=4 class=xl30>&nbsp;</td>
  <td class=xl24></td>
  <td colspan=2 class=xl34>Date Applied</td>
  <td colspan=3 class=xl51><%=I34%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24 style='height:15.0pt'></td>
  <td colspan=4 class=xl41>APPLICANT SIGNATURE</td>
  <td class=xl24></td>
  <td colspan=2 class=xl34>Date Issued</td>
  <td colspan=3 class=xl40><%=I35%></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl24 style='height:12.75pt'></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=6 height=17 class=xl24 style='height:12.75pt'></td>
  <td colspan=2 class=xl34>Facility Fee</td>
  <td colspan=3 class=xl33><%=formatcurrency(I37)%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24 style='height:15.0pt'></td>
  <td colspan=4 class=xl30><%=B38%></td>
  <td class=xl24></td>
  <td colspan=2 class=xl34>Tapping Fee</td>
  <td colspan=3 class=xl51><%=Formatcurrency(I38)%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24 style='height:15.0pt'></td>
  <td colspan=4 class=xl41>PRINT NAME</td>
  <td class=xl24></td>
  <td colspan=2 class=xl34>Connection Fee</td>
  <td colspan=3 class=xl40><%=FormatCurrency(I39)%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24 style='height:15.0pt'></td>
  <td colspan=4 class=xl53><%=B43%></td>
  <td class=xl24></td>
  <td colspan=2 class=xl34>Inspection Fee</td>
  <td colspan=3 class=xl40><%=FormatCurrency(I40)%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24 style='height:15.0pt'></td>
  <td colspan=4 class=xl30><%=B43b%></td>
  <td class=xl24></td>
  <td colspan=2 class=xl47>Total Fees Paid</td>
  <td colspan=3 class=xl40><%=FormatCurrency(Int(I37)+Int(I38)+Int(I39)+Int(I40))%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl24 style='height:15.0pt'></td>
  <td colspan=4 class=xl41>ADDRESS OF APPLICANT</td>
  <td class=xl24></td>
  <td colspan=2 class=xl34>Check #</td>
  <td colspan=3 class=xl40><%=I42%></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=6 height=20 class=xl24 style='height:15.0pt'></td>
  <td colspan=2 class=xl34>Prepared By</td>
  <td colspan=3 class=xl40><%=I43%></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td height=17 class=xl24 style='height:12.75pt'></td>
  <td colspan=4 class=xl30><%=B46%></td>
  <td colspan=6 class=xl24></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td height=17 class=xl24 style='height:12.75pt'></td>
  <td colspan=4 class=xl41>PHONE NUMBER</td>
  <td colspan=3 class=xl24></td>
  <td colspan=3 class=xl45>&nbsp;</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=8 height=17 class=xl24 style='height:12.75pt'></td>
  <td colspan=3 class=xl43>Inspected and Approved by Date</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl24 style='height:12.75pt'></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=6 height=17 class=xl34 style='height:12.75pt'></td>
  <td colspan=2 class=xl34>Reviewed</td>
  <td class=xl30>&nbsp;</td>
  <td colspan=2 class=xl27>Engineer</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl34 style='height:15.0pt'></td>
  <td class=xl24></td>
  <td colspan=3 class=xl24></td>
  <td colspan=2 class=xl34>Reviewed</td>
  <td class=xl30>&nbsp;</td>
  <td colspan=2 class=xl27>Maint. Supervisor</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl34 style='height:15.0pt'></td>
  <td class=xl31></td>
  <td colspan=8 class=xl24></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl34 style='height:15.0pt'></td>
  <td class=xl31></td>
  <td colspan=8 class=xl24></td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl34 style='height:15.0pt'></td>
  <td class=xl31></td>
  <td colspan=4 class=xl24></td>
  <td colspan=4 class=xl30>&nbsp;</td>
 </tr>
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td colspan=2 height=20 class=xl34 style='height:15.0pt'></td>
  <td class=xl31></td>
  <td colspan=4 class=xl38></td>
  <td colspan=4 class=xl46>Permit Approved</td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl24 style='height:12.75pt'></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td colspan=11 height=17 class=xl39 style='height:12.75pt'>Original Pink and
  Yellow copies <font class="font10">must</font><font class="font9"> be on site
  at the time of the inspection.<span style='mso-spacerun:yes'>  </span>There
  will be a $5.00 charge to replace lost copies.</font></td>
 </tr>
 <tr height=17 style='mso-height-source:userset;height:12.75pt'>
  <td height=17 colspan=11 class=xl24 style='height:12.75pt;mso-ignore:colspan'></td>
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
  <td width=64 style='width:48pt'></td>
 </tr>
 <![endif]>
</table>

</body>

</html>
