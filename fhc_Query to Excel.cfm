<!--- This is an example of a CFM you would write for a new report --->

<cfquery 
        name="excelquery" 
        datasource="soniswebp">
        SELECT TOP 10 * FROM name 
        WHERE camp_cod in (select camp_cod from rpt_camp where token = 
          <cfqueryparam cfsqltype="CF_SQL_CHAR" value="#session.token#">)
        ORDER BY soc_sec
</cfquery>


<cfinvoke component="CFC.query_to_excel" method="query_to_excel">
  <cfinvokeargument name = "excel_query" value = #excelquery#>
  <cfinvokeargument name = "report_name" value = 'test'>
</cfinvoke>