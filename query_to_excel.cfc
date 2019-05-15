<cfcomponent>
  <cffunction name="query_to_excel" output="yes" returntype="Any">
    <!---Function Arguments--->
    <cfargument name="excel_query" required="yes" type="query" default="" hint="The query you want output as a spreadsheet">
    <cfargument name="report_name" required="yes" type="string" default="" hint="The name of this report">

    <cfset variables.excel_query = arguments.excel_query>
    <cfset variables.report_name = arguments.report_name>
    <cfquery name="batch_dir_query" datasource="soniswebp" maxrows="1">
      select batch_dir from webopts
    </cfquery>
    <CFIF right(trim(batch_dir_query.batch_dir),1) EQ '/' OR  right(trim(batch_dir_query.batch_dir),1) EQ '\'>
        <CFSET batch_dir = '#trim(batch_dir_query.batch_dir)#'>
    <CFELSE>
        <CFSET batch_dir = '#trim(batch_dir_query.batch_dir)#\'>
    </CFIF>
    <cfset login_id = trim(#session.login_id#)>
    <cfset save_path = #batch_dir# & #login_id# & "_" & #report_name# & ".xls">
    <cfset file_name = #login_id# & "_" & #report_name# & ".xls">

    <cfspreadsheet action = "write" filename=#save_path# query="excel_query" overwrite="true">
    <!--- If the report is run in debug mode, save it in batch directory                                         --->
    <!--- We still get debugging information/mime doesn't change/ no about:blank + code halt  on cfcontent tag   --->
    <!--- http://www.bennadel.com/blog/1939-using-cfcontent-s-variable-attribute-stops-all-further-processing.htm--->
    <!--- Otherwise, stream the file to browser/delete it                                                        --->
    <cfheader name="Content-Disposition" value="attachment;filename=#file_name#">
    <cfcontent type="application/vnd.ms-excel" file="#save_path#" deletefile="Yes">

  </cffunction>
</cfcomponent>
