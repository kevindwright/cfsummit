
  <link rel="stylesheet" href="styles.css">
	<table width="75%">
		<tr>
			<td align="right" style="padding:5px;">
		
			<cfif #get_Details.recordcount# GT 0>

			</td>
		</tr>
		<tr>
			<td>

				<cfoutput query="get_Details" group="PropNum">

					<table width="100%" cellpadding="2" cellspacing="0" border="0">
						<tr>
							<td>
								<table width="100%" cellpadding="2" cellspacing="0" border="0">
									<tr>
										<td class="navbottom" colspan="1">
											<strong>#PropNum# - #PropName#</strong>
										</td>
										<td class="navbottom" colspan="4">
											
										</td>
									</tr>
									<tr>
										<td class="navbar" width="25%" valign="top">
											<strong>Category / Item</strong>
										</td>
										<td  class="navbar" width:="25%">

										</td>

										<td  class="navbar" width:="25%">
											<strong>Comments</strong>
										</td>
										<!--- <td  class="navbar" width:="25%">
											<strong>Photos</strong>
										</td> --->
									</tr>

									<cfset iCounter = 0>
									<cfoutput group="unit_name">
											<cfset iCounter = iCounter + 1>
											<tr>
												<td colspan="5" class="RuleBlue"></td>
											</tr>
											<tr bgcolor="###iif(iCounter MOD 2,DE('ffffff'),DE('efefef'))#">
												<td colspan="5">
													<strong>Unit ## #unit_name#</strong>
												</td>
											</tr>
											<cfoutput group="sectionName">
												<tr bgcolor="###iif(iCounter MOD 2,DE('ffffff'),DE('efefef'))#">
													<td colspan="5" class="navbar">
														<strong>#sectionName#</strong>
													</td>
												</tr>

												<cfoutput group="item_name">
													<tr bgcolor="###iif(iCounter MOD 2,DE('ffffff'),DE('efefef'))#" >
														<td valign="top">
															#item_name#
														</td>
														<td valign="top">
															#keyName# / #WO_key#=#WO_Value#
														</td>
														<td valign="top">
															#notes#
														</td>
														<!--- loop through for list of links --->
														<!--- <td valign="top" nowrap >
															<cfset pCounter = 1 />
															<cfoutput group="photo">
																<cfif #photo# NEQ "">
																	<a href="https://manage.inspections.com/api/v3/content/photos/#photo#" target="_blank">photo#pCounter#</a></br>
																	<cfset pCounter = pCounter + 1 />
																</cfif>
															</cfoutput>
														</td> --->
													</tr>
												</cfoutput>
											</cfoutput>
											<tr>
												<td colspan="5" class="RuleGrey"></td>
											</tr>
									</cfoutput>
								</table>
							</td>
						</tr>
					</table>
				</cfoutput>
		
			</td>	
		</tr>
	</table>			
		<!--- <cfoutput>				
			<form name="frmExport" id="frmExport" action="/index.cfm?sPageInclude=happyCo/report_page.cfm" method="post">
				<input type="hidden" name="rdoRptType" value="#Variables.sRptType#" />
				<input type="hidden" name="Inspection" value="#Variables.sInspection#" />
				<input type="hidden" name="PropList" value="#Variables.sProperty#" />
				<input type="hidden" name="export" value="1" >
				<input type="hidden" name="bSubmit" value="export" >
			</form>
		</cfoutput>				 --->
			

		<script language="javascript">

		function submitExport() {
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