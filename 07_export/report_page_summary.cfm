		
	<tr>
		<td align="right" style="padding:5px;">
	
	<cfif #get_Summary.recordcount# GT 0>

		</td>
	</tr>
		<tr>
			<td>

					<cfoutput query="get_Summary" group="PropNum">
						<table width="100%" cellpadding="2" cellspacing="0" border="0">
							<tr>
								<td class="navbottom" colspan="1">
									<span style="margin-left: 10px;"><strong>#PropNum# - #PropName#</strong></span>
								</td>
								<td class="navbottom" colspan="4">
									<a href="javascript:submitExport('excel')"><image src="./images/btn_sq_excel.gif" /></a>
									<a href="javascript:submitExport('word')"><image src="./images/btn_sq_pdf.gif" /></a>
								</td>
							</tr>
							<cfoutput group="MainCat">
								<tr>
									<td class="navbar" style="border-top: solid 1px;"  colspan="5" valign="top">
										<span style="margin-left: 25px;"><strong>#MainCat#</strong></span>
									</td>
								</tr>
								<cfoutput group="item_Name">
									<tr>
										<td></td>
										<td colspan="4" valign="top">
											
											<strong>#item_Name#</strong><br/>
											<cfoutput group="keyName">
												<span style="margin-left: 20px;">#keyName#</span>
												<ul style="margin-top: 0px;">	
													<cfoutput group="keyValue">
														<cfif #WO_keyName# EQ 'Deficiency'>
															<li><span><font color="red">#keyValue# (#iCount#) - Deficiency</font></span></li>
														<cfelseif #WO_keyName# EQ 'Work Order'>
															<li><span><font color="red"><strong>#keyValue# (#iCount#) - Work Order</strong></font></span></li>
														<cfelse>
															<li><span>#keyValue# (#iCount#)</span></li>
														</cfif>
													</cfoutput>
												</ul>
											</cfoutput>
														
										</td>
									</tr>
									
								</cfoutput>
							</cfoutput>
							
						</table>
					</cfoutput>	
						

<cfoutput>				
	<form name="frmExport" id="frmExport" action="/index.cfm?sPageInclude=happyCo/report_page.cfm" method="post">
		<!--- <cfif #Variables.bDeficiency# EQ true><input type="hidden" name="cbxDeficiency" value="1" /></cfif>
		<cfif #Variables.bWorkOrder# EQ true><input type="hidden" name="cbxWorkOrder" value="1" /></cfif>	 --->
		<input type="hidden" name="rdoRptType" value="#Variables.sRptType#" />
		<input type="hidden" name="Inspection" value="#Variables.sInspection#" />
		<input type="hidden" name="PropList" value="#Variables.sProperty#" />
		<input type="hidden" name="export" id="export" value="1" >
		<input type="hidden" name="exportType" id="exportType" value="" >
		<input type="hidden" name="bSubmit" value="export" >
		
	</form>
</cfoutput>				
				
   
<script language="javascript">
	
function submitExport(sType) {
	document.getElementById("exportType").value = sType; 
	document.getElementById("frmExport").submit(); 
}
							
</script>
				
<cfelse>


		<table width="100%" cellpadding="2" cellspacing="0" border="0">
			<tr>
				<td colspan="6" class="textImportant" align="center">
					<p>Your search returned zero records.</p>
				</td>	
			</tr>
		</table>

</cfif>


				
				
				