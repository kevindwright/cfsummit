<!---* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * 
	FILE NAME						report_page.cfm
	AUTHOR							kwright@wng.com
	PROJECT  						Inspections - Reporting Page
	PAGE DESCRIPTION
		Report page for Unit Inspections
				
	DEVELOPMENT HISTORY
	DATE		WHO		VERSION		DESCRIPTION
	========	===		=======		===============================
	02-10-20	KDW		1.0			Original Development
* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * --->

<!--- Declare Variables --->
	
<cfif isdefined("Form.rdoRptType")>
	<cfset Variables.sRptType = #Form.rdoRptType#>
<cfelse>
	<cfset Variables.sRptType = "1">
</cfif>



<table width="75%" cellpadding="0" cellspacing="0" border="0" class="outsidebordernoside" align="center">
	<tr>
		<td>
			<table width="75%" cellpadding="5" cellspacing="0" border="0">		
				<tr>
					<td class="navbottom">
						Inspection Report
					</td>
                    <td class="navbottom"><a href="javascript:window.history.back()" class="backButton"></a></td>
				</tr>
				<tr>
					<td class="navbar" align="center">
					<form action="" name="frmInspection" method="post">
                        <table width="100%" cellpadding="3" cellspacing="0" border="0">
                            
							<tr>
								<td align="right" valign="top" class="navbar">View:</td>
								<td align="left" height="20" class="navbar outsidebordernoside" width="40%">
									<input type="radio" value="1" name="rdoRptType" <cfif #Variables.sRptType# EQ 1>checked</cfif> />Detail</br>
									<input type="radio" value="2" name="rdoRptType" <cfif #Variables.sRptType# EQ 2>checked</cfif>/>Summary</br>
									<input type="radio" value="3" name="rdoRptType" <cfif #Variables.sRptType# EQ 3>checked</cfif>/>Worksheet
                        		</td>
								<td  colspan="2" width="40%">&nbsp;</td>
							</tr>
							<tr><td height="1" colspan="4"></td></tr>
							
							<tr>
								<td align="center"></td>
								<td align="center">
									<input type="submit" name="submit" value="View Report" class="button">
								</td>
								<td  colspan="2" width="40%">&nbsp;</td>
							</tr>
                            <tr>
                        		<td colspan="4" align="center" height="20" class="navbar">
                            		<div id="myDesc"></div>
                        		</td>
                    
							</tr>
						</table>
						</form>
						<td></td>
					</td>	
				</tr>
			</table>	
		</td>
	</tr>

	<tr>
		<td>
	
			<cfif isdefined("FORM.Submit") OR isdefined("FORM.bSubmit")>

            	<cfset sUrl = "#getPageContext().getRequest().getRequestURI()#"  />
                <cfset ext = "\.(cfm?.*|[^.]+)"   />
                <cfset sPath = trim(sUrl) />
                <cfset sEndDir = reFind("/[^/]+#ext#$", sPath) />
                <cfset basePath =  left(sPath, sEndDir) />


				<cfif #FORM.rdoRptType# EQ 1>
                    <cfhttp method="get" url="http://#CGI.SERVER_NAME#/#basePath#/csv/inspectionDetail.csv" name="get_Details">
					<cfinclude template="07_export/report_page_detail.cfm" />

				<cfelseif #FORM.rdoRptType# EQ 2>
                    <cfhttp method="get" url="http://#CGI.SERVER_NAME#/#basePath#/csv/inspectionSummary.csv" name="get_Summary">
                    <cfinclude template="07_export/report_page_summary.cfm" />
                    
				<cfelse>
                    <!--- This is the file from the demo (look in the '07_DOWNLOAD' folder)--->
					<cfhttp method="get" url="http://#CGI.SERVER_NAME#/#basePath#/csv/inspectionData.csv" name="get_Summary">
                    <cfinclude template="07_export/report_page_export_worksheet.cfm" />
					
				</cfif>

			</cfif>

		</td>
	</tr>
</table>

	
</script>