<!doctype html>

<html lang="en">
<head>
  <meta charset="utf-8">

  <title>CFSUMMIT 2022 - Spreadsheet Magic with CFSpreadsheet</title>
  <meta name="description" content="CFSUMMIT 2022 - Spreadsheet Magic with CFSpreadsheet">
  <meta name="author" content="Kevin D. Wright - Kinetic InterActive">

<!---   <link rel="stylesheet" href="css/default.css"> --->

</head>

    <body>
    
	<cfset sUrl = "#getPageContext().getRequest().getRequestURI()#"  />
                <cfset ext = "\.(cfm?.*|[^.]+)"   />
                <cfset sPath = trim(sUrl) />
                <cfset sEndDir = reFind("/[^/]+#ext#$", sPath) />
                <cfset basePath =  left(sPath, sEndDir) />

    <cfhttp method="get" url="http://#CGI.SERVER_NAME#/#basePath#/csv/sales.csv" name="csvData">

	<cfif isdefined('FORM.submit')>
    
        <cfset sObj = SpreadsheetNew("Sales") /> 
        
        <cfset spreadsheetaddrow(sObj,"SalesRep,District,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec")> 

        <cfset SpreadsheetSetCellValue(sObj,'Total',1,15) >
        <cfset SpreadsheetSetCellValue(sObj,'Avg',1,16) >
        
        <cfset SpreadsheetFormatRow(sObj, {bold=TRUE, alignment="center", fgcolor="grey_50_percent", color="white"}, 1)> 



        
        <cfoutput query="csvData" >
        	<cfset spreadsheetaddrow(sObj,"#SalesRep#,#District#,#Jan#,#Feb#,#Mar#,#Apr#,#May#,#Jun#,#Jul#,#Aug#,#Sep#,#Oct#,#Nov#,#Dec#")>
            <cfset spreadsheetSetCellFormula(sObj, "SUM(C#currentRow+1#:N#currentRow+1#)", #currentRow#+1, 15)>
            <cfset spreadsheetSetCellFormula(sObj, "O#currentRow+1#/12", #currentRow#+1, 16)>
        </cfoutput>

		<cfset rowDataStart=2>
        <cfset rowDataEnd=csvData.recordCount+1>
        
		<cfset spreadsheetSetCellValue(sObj, "Total:", rowDataEnd+1, 2)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(C#rowDataStart#:C#rowDataEnd#)", rowDataEnd+1, 3)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(D#rowDataStart#:D#rowDataEnd#)", rowDataEnd+1, 4)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(E#rowDataStart#:E#rowDataEnd#)", rowDataEnd+1, 5)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(F#rowDataStart#:F#rowDataEnd#)", rowDataEnd+1, 6)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(G#rowDataStart#:G#rowDataEnd#)", rowDataEnd+1, 7)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(H#rowDataStart#:H#rowDataEnd#)", rowDataEnd+1, 8)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(I#rowDataStart#:I#rowDataEnd#)", rowDataEnd+1, 9)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(J#rowDataStart#:J#rowDataEnd#)", rowDataEnd+1, 10)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(K#rowDataStart#:K#rowDataEnd#)", rowDataEnd+1, 11)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(L#rowDataStart#:L#rowDataEnd#)", rowDataEnd+1, 12)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(M#rowDataStart#:M#rowDataEnd#)", rowDataEnd+1, 13)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(N#rowDataStart#:N#rowDataEnd#)", rowDataEnd+1, 14)>
        <cfset spreadsheetSetCellFormula(sObj, "SUM(M#rowDataStart#:M#rowDataEnd#)", rowDataEnd+1, 15)>
        
        <cfloop from="2" to="#rowDataEnd+1#" index="x">
			<cfset SpreadsheetFormatRow(sObj,{dataformat="$######,####0.00"},x)>
        </cfloop>
        
        <cfset SpreadsheetFormatRow(sObj, {bold=TRUE}, rowDataEnd+1)> 
        <cfset SpreadsheetFormatCell(sObj, {bold=TRUE, alignment="center", fgcolor="red", color="white"}, rowDataEnd+1,15)>
									  
        <cfspreadsheet 
        	action="write" 
            overwrite="true"
            filename = "#ExpandPath( 'tmpfiles\' )#sales.xls"  
            name="sObj" />
         
        <cfheader name="Content-Disposition" value="inline; filename=sales.xls">
		<cfcontent type="application/vnd.ms-excel" file="#ExpandPath( 'tmpFiles\' )#sales.xls">
    
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
			
	<a href="02_formatting_expanded.cfm">Next</a>	

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