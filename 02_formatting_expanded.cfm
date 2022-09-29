<!doctype html>

<html lang="en">
<head>
  <meta charset="utf-8">

  <title>CFSUMMIT 2022 - Spreadsheet Magic with CFSpreadsheet</title>
  <meta name="description" content="CFSUMMIT 2022 - Spreadsheet Magic with CFSpreadsheet">
  <meta name="author" content="Kevin D. Wright - Kinetic InterActive">

  <link rel="stylesheet" href="css/default.css">

</head>

    <body>
    
	<cfset sUrl = "#getPageContext().getRequest().getRequestURI()#"  />
			<cfset ext = "\.(cfm?.*|[^.]+)"   />
			<cfset sPath = trim(sUrl) />
			<cfset sEndDir = reFind("/[^/]+#ext#$", sPath) />
			<cfset basePath =  left(sPath, sEndDir) />

    <cfhttp method="get" url="http://#CGI.SERVER_NAME#/#basePath#/csv/sales.csv" name="csvData">


	<cfif isdefined('FORM.submit')>
    
		<cfscript>
			format1=StructNew(); 
			format1.font="serif"; 
			format1.fontsize="12"; 
			format1.color="white"; 
			format1.bold="true"; 
			format1.alignment="center";
			format1.fgcolor = "maroon"; 
        </cfscript>
        

   		<cffunction name="formatCell">
            <cfargument name="iValue">
            <cfargument name="iRow">
            <cfargument name="iColumn">
            <cfargument name="nThreshold">
        
       		<cfif #iValue# GTE #nThreshold#>
                <cfset SpreadsheetFormatCell(sObj,format1,iRow,iColumn)>
        	</cfif> 
        </cffunction>
        
        
        
        <cfset iLimit = 9000 >
    
        <cfset sObj = SpreadsheetNew("Sales") /> 

        <cfset spreadsheetaddrow(sObj,"SalesRep,District,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec")> 
        
		<!---<cfset SpreadsheetAddRows(sObj,csvData) />--->
        
        <cfoutput query="csvData" >
        	<cfset spreadsheetaddrow(sObj,"#SalesRep#,#District#,#Jan#,#Feb#,#Mar#,#Apr#,#May#,#Jun#,#Jul#,#Aug#,#Sep#,#Oct#,#Nov#,#Dec#")>
            <cfset formatCell(#Jan#,(#currentRow#+1),3,#iLimit#)>
            <cfset formatCell(#Feb#,(#currentRow#+1),4,#iLimit#)>
            <cfset formatCell(#Mar#,(#currentRow#+1),5,#iLimit#)>
            <cfset formatCell(#Apr#,(#currentRow#+1),6,#iLimit#)>
            <cfset formatCell(#May#,(#currentRow#+1),7,#iLimit#)>
            <cfset formatCell(#Jun#,(#currentRow#+1),8,#iLimit#)>
            <cfset formatCell(#Jul#,(#currentRow#+1),9,#iLimit#)>
            <cfset formatCell(#Aug#,(#currentRow#+1),10,#iLimit#)>
            <cfset formatCell(#Sep#,(#currentRow#+1),11,#iLimit#)>
            <cfset formatCell(#Oct#,(#currentRow#+1),12,#iLimit#)>
            <cfset formatCell(#Nov#,(#currentRow#+1),13,#iLimit#)>
            <cfset formatCell(#Dec#,(#currentRow#+1),14,#iLimit#)>
        </cfoutput>
        
 <!---      <cfscript>
        	for (i=1;i LTE csvData.recordcount; i=i+1){
        		
		   		sValues = csvData['SalesRep'][i]&','&csvData['District'][i]&','&csvData['Jan'][i]&','&csvData['Feb'][i]&','&csvData['Mar'][i]&','&csvData['Apr'][i]&','&csvData['May'][i]&','&csvData['Jun'][i]&','&csvData['Jul'][i]&','&csvData['Aug'][i]&','&csvData['Sep'][i]&','&csvData['Oct'][i]&','&csvData['Nov'][i]&','&csvData['Dec'][i];
        		spreadsheetAddRow(sObj,sValues);
		   
				formatCells(csvData['Jan'][i],i,3,iLimit);
				formatCells(csvData['Feb'][i],i,4,iLimit);
				formatCells(csvData['Mar'][i],i,5,iLimit);
				formatCells(csvData['Apr'][i],i,6,iLimit);
				formatCells(csvData['May'][i],i,7,iLimit);
				formatCells(csvData['Jun'][i],i,8,iLimit);
				formatCells(csvData['Jul'][i],i,9,iLimit);
				formatCells(csvData['Aug'][i],i,10,iLimit);
				formatCells(csvData['Sep'][i],i,11,iLimit);
				formatCells(csvData['Oct'][i],i,12,iLimit);
				formatCells(csvData['Nov'][i],i,13,iLimit);
				formatCells(csvData['Dec'][i],i,14,iLimit);
			}
			
			function formatCells(iValue,iRow,iColumn,nThreshold) { 
				if(iValue > nThreshold){
					SpreadsheetFormatCell(sObj,format1,iRow,iColumn);
				}
			} 
        </cfscript> --->
		
        <cfscript>
			format2 = StructNew();
			format2.dataformat = "$######,####0.00";
		</cfscript>
        
  		<cfset SpreadsheetFormatColumn(sObj,format2,14)>
        <cfset SpreadsheetFormatRow(sObj, {bold=TRUE, alignment="center", fgcolor="grey_50_percent", color="white"}, 1)> 
        
		<cfset format3 = StructNew() />
        <cfset format3.bold = "true" />
        <cfset format3.bottomborder = "thin" />
        <cfset format3.bottombordercolor = "black" />
        <cfset format3.fgcolor = "light_yellow" />
        
        <cfset SpreadsheetFormatCellRange(sObj, format3, 4, 1, 4, 11) />
   <!---<cfset SpreadsheetFormatCellRange(object, format, startRow, startColumn, endRow, endColumn) />--->


  
        <cfspreadsheet 
        	action="write" 
            overwrite="true"
            filename = "#ExpandPath( 'tmpfiles\' )#sales.xls"  
            name="sObj" />
         
            
		<cfheader name="Content-Disposition" value="inline; filename=sales.xls">
		<cfcontent type="application/vnd.ms-excel" file="#ExpandPath( 'tmpfiles\' )#sales.xls">
    
    </cfif>
    
	<table id="demo" class="display" style="width:100%;" border=0>
		<thead>
			<tr>
				<th>Sales Rep</th>
				<th>District</th>
				<th>Jan</th>
				<th>Feb</th>
				<th>Mar</th>
				<th>Apr</th>
				<th>May</th>
				<th>Jun</th>
				<th>Jul</th>
				<th>Aug</th>
				<th>Sep</th>
				<th>Oct</th>
				<th>Nov</th>
				<th>Dec</th>
			</tr>
		</thead>
		<tbody>
			<cfoutput query="csvData">
			<tr>
				<td>#SalesRep#</td>
				<td>#District#</td>
				<td>#Jan#</td>
				<td>#Feb#</td>				
				<td>#Mar#</td>
				<td>#Apr#</td>
				<td>#May#</td>
				<td>#Jun#</td>
				<td>#Jul#</td>				
				<td>#Aug#</td>
				<td>#Sep#</td>
				<td>#Oct#</td>
				<td>#Nov#</td>
				<td>#Dec#</td>
			</tr>            
			</cfoutput>
		</tbody>
	</table>

	<form name="frmSales" action="" method="post">
        <input name="submit" type="Submit" value="Export To Excel" class="btn" /><br/><br/>
    </form>
			
	<a href="03_formatting_conditional.cfm">Next</a>	

    <link href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" rel="stylesheet" type="text/css" media="screen" />	
	<script src="https://code.jquery.com/jquery-1.12.4.js" type="text/javascript"></script>
	<script src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js" type="text/javascript"></script>
	<script language="JavaScript">
	$(document).ready(function() {
	    $('#demo').DataTable();
	} );
	</script>
			
    </body>
</html>