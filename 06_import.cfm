<!---
	Copyright (c) 2012 by Western National Group, LLC.
	FILE NAME	voucher_import.cfm
	AUTHOR		kwright@wng.com
	PROJECT  	InSite - Electronic Invoicing
	DESCRIPTION	Allows user to upload a file(s) to create dept vouchers
	DEVELOPMENT HISTORY
	DATE		WHO		VERSION		DESCRIPTION
	=======		=====	========	===============================
	04-15-12	KDW		1.0			Original Development
--->


<LINK REL="StyleSheet" HREF="styles.css" type="text/css">



<!--- Upload the file that has just been submitted --->
<cfif IsDefined("FORM.fileName") and FORM.fileName NEQ "">

<!--- set validation values --->
<cfset bValid = 1>  
<cfset bIsValid = 1>
         
<cfset sMessage = ''>
<cfset sLatestVersion = 'v8.4'>

<!--- set CORP or AP --->
<cfset variables.dept = #FORM.dept#>

<!--- Default variable values --->
<cfparam name="fileName" default="" type="string">
<cfset iInvoiceTotal = 0>
<cfset iCountTotal = 0>
<cfset iVoucherCount = 0>


<!--- ***************************************************************     
                            FUNCTIONS                                     
      *************************************************************** --->


<!--- VALIDATION --->
<cffunction name="validateCHRG_results">
<cfargument name="QryName" type="string" required="yes">
	<cfoutput query="#QryName#">
    
    	<!--- validate number of charges --->
		<cfif #COL_4# NEQ '0' AND #COL_4# NEQ ''>
            <cfif #IsValid ("integer", COL_4)# EQ 0>
            	<cfset bValid = 0>
            </cfif>
        </cfif>
        
        <!--- validate charge amount--->
		<cfif #COL_5# NEQ '0' AND #COL_5# NEQ ''>
            <cfif #IsValid ("float", COL_5)# EQ 0>
            	<cfset bValid = 0>
            </cfif>
        </cfif>
        
        <!--- validate total charge --->
		<cfif #COL_6# NEQ '0' AND #COL_6# NEQ ''>
            <cfif #IsValid ("float", COL_6)# EQ 0>
            	<cfset bValid = 0>
            </cfif>
        </cfif>

    	<!--- validate account code --->
		<cfif #COL_6# NEQ '0' AND #COL_6# NEQ ''>
            <cfif #COL_7# EQ 0 OR #COL_7# EQ "">
            	<cfset bValid = 0>
            </cfif>
        </cfif>

    </cfoutput>
</cffunction>

<!--- OUTPUT --->      
<cffunction name="outputCHRG_results">
<cfargument name="QryName" type="string" required="yes">
    <cfsavecontent variable="sHTML">
    <tr>
    	<td colspan="6" style="font-weight:bold"><cfoutput>#QryName#</cfoutput></td>
    </tr>
	<cfset iGroupTotal = 0> 
    <cfset iGroupCount = 0> 
    <cfset iGroupRow = 0>
	<cfoutput query="#QryName#">
        <cfset bIsValid = 1>
		<cfif #COL_6# NEQ '0' AND #COL_6# NEQ ''>
			<!--- validate number of charges --->
            <cfif #COL_4# NEQ '0' AND #COL_4# NEQ ''>
                <cfif #IsValid ("integer", COL_4)# EQ 0>
                    <cfset bIsValid = 0>
                </cfif>
			</cfif>
            <!--- validate charge amount--->
            <cfif #COL_5# NEQ '0' AND #COL_5# NEQ ''>
                <cfif #IsValid ("float", COL_5)# EQ 0>
                    <cfset bIsValid = 0>
                </cfif>
            </cfif>
            <!--- validate total charge --->
            <cfif #COL_6# NEQ '0' AND #COL_6# NEQ ''>
                <cfif #IsValid ("float", COL_6)# EQ 0>
                    <cfset bIsValid = 0>
                </cfif>
            </cfif>
            <!--- validate account code --->
            <cfif #COL_6# NEQ '0' AND #COL_6# NEQ ''>
                <cfif #COL_7# EQ 0 OR #COL_7# EQ "">
                    <cfset bIsValid = 0>
                </cfif>
            </cfif>
        	<cfif bIsValid EQ 0>
            	<tr bgcolor="##FF9999">
                
            <cfelse>
            	<tr bgcolor="#IIF(iGroupRow MOD 2, DE('FFFFFF'), DE('F2F3F7'))#">
            </cfif>
            
                <td align="right" valign="top">#COL_1#</td>
                <td align="right" valign="top">#COL_2#</td>
                <td valign="top">#sDescription.col_1#-<br/>#COL_3#</td>
                <td align="right" valign="top">#COL_4#</td>
                <td align="right" valign="top"><cfif #IsValid ("float", COL_5)# EQ 1>#dollarformat(replace(COL_5,",",""))#<cfelse>#COL_5#</cfif></td>
                <td align="right" valign="top">#COL_6#</td>
                <td align="right" valign="top">#COL_7#</td>
            </tr>
            <cfif #IsValid ("float", COL_6)# EQ 1>
                <cfset iGroupTotal = iGroupTotal + #replace(COL_6,",","")#>
            <cfelse>
                <cfset iGroupTotal = iGroupTotal>
            </cfif>
            
            <cfif #IsValid ("float", COL_4)# EQ 1>
                <cfset iGroupCount = iGroupCount + #COL_4#>
            <cfelse>
                <cfset iGroupCount = iGroupCount>
            </cfif>



            <cfset iVoucherCount = iVoucherCount + 1>
            <cfset iGroupRow = iGroupRow + 1>
        </cfif>
    </cfoutput>
    <tr bgcolor="#DEDEDE">
        <td colspan="2" align="right"></td><td align="right" style="font-weight:bold">Total <cfoutput>#QryName#</cfoutput></td><td align="right" style="font-weight:bold"><cfoutput>#iGroupCount#</cfoutput></td><td align="right"></td><td align="right" style="font-weight:bold"><cfoutput>#dollarformat(iGroupTotal)#</cfoutput></td><td>&nbsp;</td>
    </tr>
    </cfsavecontent>
    
    <cfif iGroupTotal GT 0>
    	<cfset iInvoiceTotal = iInvoiceTotal + iGroupTotal>
        <cfset iCountTotal = iCountTotal + iGroupCount>
        
    	<cfoutput>#sHTML#</cfoutput>
    </cfif>
    
</cffunction>  

	<!--- create unique filename --->
	<cfset variables.sTimeStamp = toString(DateFormat(now(),"yyyymmdd")) & '_' & toString(TimeFormat(now(),"HHmmss"))>
	<cfset variables.sFileNameDynamic = variables.sTimeStamp & ".xlsx">

	<!--- set the directory path --->
	<cfset sPath = ExpandPath( "./") >
    <cfset sFilePath = sPath & "tmpfiles\">

	<cffile action="UPLOAD" filefield="FORM.fileName" destination="#Variables.sFilePath##variables.sFileNameDynamic#" result="upload" accept="application/octet-stream, application/vnd.ms-excel, application/vnd.ms-excel.sheet.macroenabled.12, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" nameconflict="OVERWRITE" attributes="normal">
	<cfset theFile = upload.serverDirectory & "/" & upload.serverFile>



	<!--- set the query beginning / end rows --->
      <cfspreadsheet action="read" src="#theFile#" query="E10_rows" sheet="2" rows="5" columns="4,5">
      <cfspreadsheet action="read" src="#theFile#" query="EOM_rows" sheet="2" rows="6" columns="4,5">
      <cfspreadsheet action="read" src="#theFile#" query="WNPLLC_rows" sheet="2" rows="7" columns="4,5">
      <cfspreadsheet action="read" src="#theFile#" query="WNPM_rows" sheet="2" rows="8" columns="4,5">
      <cfspreadsheet action="read" src="#theFile#" query="MAG_rows" sheet="2" rows="9" columns="4,5">
      <cfspreadsheet action="read" src="#theFile#" query="WNC_rows" sheet="2" rows="10" columns="4,5">
      <cfspreadsheet action="read" src="#theFile#" query="RGS_rows" sheet="2" rows="11" columns="4,5">
      <cfset theCols = "1,2,3,7,9,11,12">
        
    <!--- create query objects --->
            
     <!--- E10 ---> 
     <cfspreadsheet 
        action="read" 
        src="#theFile#" 
        query="E10" 
        sheet="1" 
        rows="#E10_rows.col_1#-#E10_rows.col_2#" 
        columns="#theCols#">
     
     <!--- EOM --->   
     <cfspreadsheet 
        action="read" 
        src="#theFile#" 
        query="EOM" 
        sheet="1" 
        rows="#EOM_rows.col_1#-#EOM_rows.col_2#" 
        columns="#theCols#">
        
    <!--- WNP_CORP --->   
    <cfspreadsheet 
        action="read" 
        src="#theFile#" 
        query="WNPLLC" 
        sheet="1" 
        rows="#WNPLLC_rows.col_1#-#WNPLLC_rows.col_2#"
        columns="#theCols#">
    
    <!--- WNPM --->   
    <cfspreadsheet 
        action="read" 
        src="#theFile#" 
        query="WNPM" 
        sheet="1" 
        rows="#WNPM_rows.col_1#-#WNPM_rows.col_2#"
        columns="#theCols#">
        
    <!--- MAG --->   
    <cfspreadsheet 
        action="read" 
        src="#theFile#" 
        query="MAG" 
        sheet="1" 
        rows="#MAG_rows.col_1#-#MAG_rows.col_2#"
        columns="#theCols#">
        
    <!--- WNC --->   
    <cfspreadsheet 
        action="read" 
        src="#theFile#" 
        query="WNC" 
        sheet="1" 
        rows="#WNC_rows.col_1#-#WNC_rows.col_2#"
        columns="#theCols#">
    
    <!--- RGS --->   
    <cfspreadsheet 
        action="read" 
        src="#theFile#" 
        query="RGS" 
        sheet="1" 
        rows="#RGS_rows.col_1#-#RGS_rows.col_2#"
        columns="#theCols#">
    

	<cfspreadsheet action="read" src="#theFile#" query="versionCheck" sheet="1" rows="3" columns="14">
    <cfspreadsheet action="read" src="#theFile#" query="dtInvoice" sheet="1" rows="3" columns="2">
    <cfspreadsheet action="read" src="#theFile#" query="sInvoiceNum" sheet="1" rows="4" columns="2">
    <cfspreadsheet action="read" src="#theFile#" query="iVendorNum" sheet="1" rows="5" columns="2">
    <cfspreadsheet action="read" src="#theFile#" query="sVendorName" sheet="1" rows="6" columns="2">
    <cfspreadsheet action="read" src="#theFile#" query="sDescription" sheet="1" rows="7" columns="2">
    
    <cfif versionCheck.col_1 NEQ #sLatestVersion#>
    	<cfset sMessage = sMessage & '<li>Invoice Template (#versionCheck.col_1#) is an incorrect version.<br/>Please get the updated version (#sLatestVersion#) before proceeding.</li>'>
    </cfif>
	
	<cfif dtInvoice.col_1 EQ ''>
    	<cfset sMessage = sMessage & '<li>Invoice date must be included.</li>'>
    </cfif>
    
    <cfif sInvoiceNum.col_1 EQ ''>
    	<cfset sMessage = sMessage & '<li>Invoice number must be included.</li>'>
    </cfif>
    
    <cfif len(sInvoiceNum.col_1) GTE 20>
    	<cfset sMessage = sMessage & '<li>Invoice number must be 20 characters or less.</li>'>
    </cfif>
    
    <cfif iVendorNum.col_1 EQ ''>
    	<cfset sMessage = sMessage & '<li>Vendor Number must be included.</li>'>
    </cfif>
    
    <cfif sVendorName.col_1 EQ ''>
    	<cfset sMessage = sMessage & '<li>Vendor Name must be included.</li>'>
    </cfif>
    
    <cfif sDescription.col_1 EQ ''>
    	<cfset sMessage = sMessage & '<li>Description must be included.</li>'>
    </cfif>

	<cfif variables.dept EQ 'ap'>
        <!--- IMPORT PROPERTY CHARGES --->
        <cfset validateCharges = validateCHRG_results('E10')>
        <cfset validateCharges = validateCHRG_results('EOM')>
    <cfelse>
        <!--- IMPORT INTERDEPARTMENT CHARGES --->
        <cfset validateCharges = validateCHRG_results('WNPLLC')>
        <cfset validateCharges = validateCHRG_results('WNPM')>
        <cfset validateCharges = validateCHRG_results('MAG')>
        <cfset validateCharges = validateCHRG_results('WNC')>
        <cfset validateCharges = validateCHRG_results('RGS')>
    </cfif>
	
    <cfif bValid EQ 0>
    	<cfset sMessage = sMessage & '<li>HighlightedRows may have one or more issues<ul>'>
        <cfset sMessage = sMessage & '<li>Number of Charges must be numeric value</li>'>
        <cfset sMessage = sMessage & '<li>Charge Amount is invalid</li>'>
        <cfset sMessage = sMessage & '<li>Account Code missing</li>'>
        <cfset sMessage = sMessage & '</ul></li>'>
    </cfif>

    

<table width="70%" align="center">
    <cfif sMessage NEQ ''>
        <tr>
            <td>

                <style>
                table#errorbox{
                    background:#FFC0C0;
                    border:1px solid #C00000;
                    font-size:16px;
                    margin-bottom:18px;
                    padding:25px;
                    position:relative;
                    text-align:left;
                    width:100%; 
                    height:100%; 
                }
                img#alert{
                    width:100px;
                }
                </style>
                
                <cfoutput>
                    <p>&nbsp;</p> 

                    <table id="errorbox">
                        <tr>
                            <td width="25%" align="center">
                                <img id="alert" src="alert.png" />
                            </td>
                            <td width="75%" align="left">
                                <strong>There is a problem with the uploaded spreadsheet.<br/>
                                Please correct the following issues;</strong><br/>
                                <ul>#sMessage#</ul>
                            </td>
                        </tr>
                    </table> 

                    <!--- <div id="errorbox">
                        <div>
                            <div style="float:left">
                                <img src="alert.png" style="margin-right:40px;" />
                            </div>
                            <div style="float:left">
                                There is a problem with the uploaded spreadsheet.<br/>
                                Please correct the following issues;<br/>
                                <ul>#sMessage#</ul>
                            </div>
                        </div>
                    </div>         --->
                    <p>&nbsp;</p>
                </cfoutput>   
            </td>
        </tr>
    </cfif>
        <tr>
            <td>    
           

            <cfoutput> 
                <p>&nbsp;</p> 
                <table width="35%" bgcolor="##DEDEDE">
                    <tr>
                        <td align="right"><strong>Invoice ##:</strong></td>
                        <td align="left" style="padding-left:20px;"> #sInvoiceNum.col_1#</td>
                    </tr>
                    <tr>
                        <td align="right"><strong>Invoice Date:</strong></td>
                        <td align="left" style="padding-left:20px;"> #dtInvoice.col_1#</td>
                    </tr>
                    <tr>
                        <td align="right"><strong>Vendor ##:</strong></td>
                        <td align="left" style="padding-left:20px;"> #iVendorNum.col_1#</td>
                    </tr>
                    <tr>
                        <td align="right"><strong>Vendor Name:</strong></td>
                        <td align="left" style="padding-left:20px;"> #sVendorName.col_1#</td>
                    </tr>
                </table>
                <p>&nbsp;</p>
            </cfoutput>


            <table width="100%">
                <tr>
                    <th>Building</th><th>Property</th><th>Description</th><th># of Charges</th><th>Charge Amt.</th><th>Tot. Charge</th><th>Account Code</th>
                </tr>
                <cfif variables.dept EQ 'ap'>
                    <!--- IMPORT PROPERTY CHARGES --->
                    <cfset outPutHTML = outputCHRG_results('E10')>
                    <cfset outPutHTML = outputCHRG_results('EOM')>
                <cfelse>
                    <!--- IMPORT INTERDEPARTMENT CHARGES --->
                    <cfset outPutHTML = outputCHRG_results('WNPLLC')>
                    <cfset outPutHTML = outputCHRG_results('WNPM')>
                    <cfset outPutHTML = outputCHRG_results('MAG')>
                    <cfset outPutHTML = outputCHRG_results('WNC')>
                    <cfset outPutHTML = outputCHRG_results('RGS')>
                </cfif>

                <tr>
                    <td colspan="6">&nbsp;</td>
                </tr>
                <tr bgcolor="#DEDEDE">
                    <th></th><th style="font-weight:bold; text-align:right">Invoice Total</th><th style="font-weight:bold; text-align:right"><cfoutput>#iVoucherCount# Vouchers</cfoutput></th><th  style="font-weight:bold; text-align:right"><cfoutput>#iCountTotal#</cfoutput></th><th>&nbsp;</th><th style="font-weight:bold; text-align:right"><cfoutput>#dollarformat(iInvoiceTotal)#</cfoutput></th><th>&nbsp;</th>
                </tr>
                <tr>
                    <td colspan="6">&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="6" align="right">
                    
                    <cfif sMessage EQ ''>
                        <!--- <form name="" id="" action="create_vouchers.cfm" method="post">
                            <input type="hidden" name="docName" id="docName" value="<cfoutput>#theFile#</cfoutput>">
                            <input type="hidden" name="dept" id="dept" value="<cfoutput>#variables.dept#</cfoutput>">
                            <input type="submit" name="submit" id="submit" value="Create Vouchers" class="button">
                        </form> --->
                        <input type="button" name="btnNext" id="btnNext" value="Create Vouchers" class="button" onclick="javascript:document.location='07_export.cfm'" />
                    <cfelse>
                        <input type="button" name="submit" id="submit" value="Back" class="button" onclick="javascript:document.location='06_import.cfm?dept=ap'" />
                    </cfif>
                    </td>
                </tr>
            </table>

        </td>
    </tr>
</table>

<cfelse>


<cfset variables.dept = #URL.dept#>

<script language="JavaScript1.2">
<!--
function CheckForm(){
	
	// Allowed file types
	var FileExt = new Array();
	FileExt[0]=".xls";
	var FileName = document.getElementById('fileName').value.toLowerCase();
	var found = 0;
	for(i=0;i<FileExt.length;i++){
		if(FileName.indexOf(FileExt[i]) > 0)
		found = 1;
	}
	if(found == 0){
		alert('The file type you have selected is not allowed.\nPlease choose only an XLS file.\n');
		return false;
	}
}


function CheckFormCSV(){
	
	// Allowed file types
	var FileExt = new Array();
	FileExt[0]=".csv";
	var FileName = document.getElementById('fileNameCSV').value.toLowerCase();
	var found = 0;
	for(i=0;i<FileExt.length;i++){
		if(FileName.indexOf(FileExt[i]) > 0)
		found = 1;
	}
	if(found == 0){
		alert('The file type you have selected is not allowed.\nPlease choose only an CSV file.\n');
		return false;
	}
}
//-->
</script>


    <table width="90%" align="left" class="outsideborder" cellpadding="5" cellspacing="0">
        <tr><td colspan="2" class="headertopbottom">CREATE <cfif #variables.dept# EQ 'ap'>PROPERTY<cfelse>DEPARTMENT</cfif> PAYMENT VOUCHERS <cfif #variables.dept# EQ 'ap'>(AP)<cfelse>(CORP)</cfif></td></tr>
        <tr><td colspan="2" class="navbottom">UPLOAD SPREADSHEET FILE</td></tr>
        <tr>
            
            <td class="navbar" width="*" valign="top">
                <form action="" method="post" enctype="multipart/form-data" name="frmUpload" id="frmUpload" onSubmit="return CheckForm()">
                <input type="File" name="fileName" class="fmenu">
                <input type="hidden" name="dept" id="dept" value="<cfoutput>#variables.dept#</cfoutput>">
                &nbsp;&nbsp;&nbsp;&nbsp;
                <input type="Submit" name="btnUpload" value="Upload File" class="button">
                </form>
            </td>
            <td class="navhdr" valign="top" align="right">&nbsp;</td>
        </tr>
    </table>



</cfif>