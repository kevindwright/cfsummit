
<cfscript>
get_PropInfo = queryNew("building,Descr,address,city,state,postalCode,regional",
                        "integer,varchar,varchar,varchar,varchar,varchar,varchar",
                        {
                            "building":8657,
                            "descr":"Coco Village",
                            "address":"1200 Bichon Frise",
                            "city":"Barker",
                            "state":"CA",
                            "postalCode":"92807",
                            "regional":"Kevin D. Wright"
                        });
                        				
</cfscript>

    
    <cfif #get_Summary.recordcount# GT 0>

        <cfset theFile = "#GetDirectoryFromPath(GetCurrentTemplatePath())#worksheet.xls" />
        

        <cfspreadsheet action="read" src="#theFile#" query="endrow" sheet="3" rows="1" columns="15">
        <cfspreadsheet action="read" src="#theFile#" query="ss_schema" headerRow="1" sheet="3" rows="2-#((endrow.col_1)-1)#" columns="1-14">

        <cfscript>
            objSpreadSheet = SpreadSheetRead("#GetDirectoryFromPath(GetCurrentTemplatePath())#/worksheet.xls");
        </cfscript>

        <cfif get_PropInfo.recordcount GT 0 >
            <cfscript>
                SpreadsheetSetActiveSheet (objSpreadSheet, "UBU Summary pg 1");
                SpreadsheetSetCellValue(objSpreadSheet, #get_PropInfo.descr#, 2, 15);
                SpreadsheetSetCellValue(objSpreadSheet, #dateformat(now(),"mm/dd/yyyy")#, 3, 15);
                SpreadsheetSetCellValue(objSpreadSheet, #get_PropInfo.regional#, 4, 15);
            </cfscript>
        </cfif>

        <cfoutput query="get_Summary">

            <cfquery dbtype="query" name="cell_schema">
                SELECT SS_Sheet, SS_Cell, SS_Col, SS_Row
                FROM ss_schema WHERE SectionVal = '#Trim(get_Summary.mainCat)#'
                        AND ItemVal = '#Trim(get_Summary.item_name)#'
                        AND keyNameVal = '#Trim(get_Summary.keyName)#'
                        AND KeyVal = '#Trim(get_Summary.keyValue)#' 
            </cfquery>

            <cfif cell_schema.recordcount GT 0 >
                <cfscript>
                    if (cell_schema.SS_Sheet == 1){
                        sheetName = "UBU Summary pg 1";
                    }
                    else{
                        sheetName = "UBU Summary pg 2";
                    }
                    SpreadsheetSetActiveSheet (objSpreadSheet, "#SheetName#");
                    SpreadsheetSetCellValue(objSpreadSheet, #iCount#, #cell_schema.SS_Row#, #cell_schema.SS_Col#,"numeric");
                </cfscript>

            </cfif>

        </cfoutput>

        <cfset cookie.userlogin = "">

        <cfif cookie.userlogin EQ "kwright"> 
            <cfscript>
                SpreadSheetCreateSheet (objSpreadSheet, "raw_data");
                SpreadsheetSetActiveSheet (objSpreadSheet, "raw_data");
                SpreadsheetAddRows(objSpreadSheet, get_Summary);
            </cfscript>
        <cfelse>
            <cfscript>
                SpreadsheetRemoveSheet (objSpreadSheet, "SchemaSheet");
                //SpreadsheetRemoveSheet (objSpreadSheet, "raw_data");
            </cfscript> 
        </cfif>

        <cfscript>
            wb = objSpreadSheet.getWorkBook();
            wb.setForceFormulaRecalculation(true);
        </cfscript>

        <cfset sID = createUUID() />
        <cfset spreadsheetwrite(objSpreadSheet, "#GetDirectoryFromPath(GetCurrentTemplatePath())#/tmp/#sID#.xls", "", true, false)>

        <cfspreadsheet 
            action="read"
            name="sObj"
            src = "#GetDirectoryFromPath(GetCurrentTemplatePath())#/tmp/#sID#.xls"/>

        <cfheader name="Content-Disposition" value="attachment; filename=summary_excel.xls">
        <cfcontent type="application/vnd.ms-excel" variable="#SpreadSheetReadBinary(sObj)#">


    <cfelse>
        
        <table width="100%" cellpadding="2" cellspacing="0" border="0">
			<tr>
				<td colspan="6" class="textImportant" align="center">
					<p>Your search returned zero records.</p>
				</td>	
			</tr>
		</table>

    </cfif>




				
				
				