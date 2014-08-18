                  <script Language="Javascript">  
	  var printflag;
	  function printit(){
		var version = navigator.appVersion;
		if (version.indexOf("Macintosh",0) > 1) 
		{	
			if(version.indexOf("MSIE",0) > 1) 
			{	
				alert("Please press Command P to print.");
			}
			else 
			{
				window.print();
			}
		}
		else
		{
			if (NS) {
				window.print() ;  
			} else {
				if ( printflag != 1) {
					printflag = 1;
			   		var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
					document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
					WebBrowser1.ExecWB(6, 2);//Use a 1 vs. a 2 for a prompting dialog box    WebBrowser1.outerHTML = "";  
				}
				else {
					alert("Please press Control P to print.");
				}
			}
		}
		return;
	  }
</script>
<script Language="Javascript">  
	var NS = (navigator.appName == "Netscape");
	//var VERSION = parseInt(navigator.appVersion);
	
	
	var version2 = navigator.appVersion;
	if (version2.indexOf("Macintosh",0) > 1) 
	{	
		if(version2.indexOf("MSIE",0) > 1)
		{	
		
			document.write('<td colspan="3" align="right"><A HREF="javascript:printit()">Print</A></td>');
		}
		else 
		{
		
			document.write('<td colspan="3" align="right"><A HREF="javascript:printit()" onClick="printit()">Print</A></td>');
		}
	}
	else
	{
		document.write('<td colspan="3" align="right"><A HREF="javascript:printit()">Print</A></td>');
	}
	
</script>
