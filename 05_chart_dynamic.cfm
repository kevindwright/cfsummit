
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">

    <title>CFSUMMIT 2022 - Spreadsheet Magic with ColdFusion</title>  
    <meta name="description" content="CFSUMMIT 2022 - Spreadsheet Magic with CFSpreadsheet">
    <meta name="author" content="Kevin D. Wright - Kinetic InterActive">

    <link href="css/default.css" rel="stylesheet" type="text/css" />
</head>

    <body>


	<cfif isdefined('FORM.submit')>
    
 
        
        <cfset sUrl = "#getPageContext().getRequest().getRequestURI()#"  />
        <cfset ext = "\.(cfm?.*|[^.]+)"   />
        <cfset sPath = trim(sUrl) />
        <cfset sEndDir = reFind("/[^/]+#ext#$", sPath) />
        <cfset basePath =  left(sPath, sEndDir) />

        <cfhttp method="get" url="http://#CGI.SERVER_NAME#/#basePath#/csv/salesChart.csv" name="csvData">


		<!--- Read spreadsheet ---> 
		<cfspreadsheet action="read" src="template_files/DynamicChartDemo.xlsx" name="sObj" /> 

		<cfscript>
        
            spreadsheetSetActiveSheet(sObj, "Sheet1");
            
            spreadsheetSetCellValue(sObj, "Bonuses", 1, 2)

            for(i=1;i<=csvData.recordCount;i++){
                spreadsheetSetCellValue(sObj, "#csvData.Associate[i]#", i+1, 1);
                spreadsheetSetCellValue(sObj, #csvData.sales[i]#, i+1, 2)
            }
			
			wb = sObj.getWorkBook();
			
			startCellColRef="A";
			startCellRowRef="2";
			endCellColRef="A";
			endCellRowRef="#(csvData.recordCount+1)#";
			sheetName = "Sheet1";

            NameReference = wb.getName("Associates");
            referenceString = sheetName&"!$"&startCellColRef&"$"&startCellRowRef&":$"&endCellColRef&"$"&endCellRowRef;
			NameReference.setRefersToFormula(referenceString);
			
			startCellColRef="B";
			startCellRowRef="2";
			endCellColRef="B";
			endCellRowRef="#(csvData.recordCount+1)#";
			sheetName = "Sheet1";

            NameReference = wb.getName("Sales");
            referenceString = sheetName&"!$"&startCellColRef&"$"&startCellRowRef&":$"&endCellColRef&"$"&endCellRowRef;
			NameReference.setRefersToFormula(referenceString);
			
            
            spreadsheetSetActiveSheet(sObj, 'Sheet1');
        
        </cfscript>

        
        <cfheader name="Content-Disposition" value="attachment; filename=chart.xlsx" />
        <cfcontent type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" variable="#SpreadSheetReadBinary(sObj)#" />
        
    </cfif>
    
    <cfform name="frmSales" action="" method="post">
        <cfinput name="submit" type="submit" value="Read File" class="btn" /><br/><br/>
        <a href="06_import.cfm?dept=ap">Next</a>
    </cfform>
    
    </body>
</html>