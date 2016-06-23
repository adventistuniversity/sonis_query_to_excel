<cfcomponent>
  <cffunction name="query_to_excel" output="yes" returntype="Any">
    <!---Function Arguments--->
    <cfargument name="excel_query" required="yes" type="query" default="" hint="The query you want output as a spreadsheet">
    <cfargument name="report_name" required="yes" type="string" default="" hint="The name of this report">

     <cfif IsDebugMode()>
      <cfoutput>
        <p>You are currently in debug mode.<br>The report file was not generated.<br>Please review the debugging information below.<br>If it is necessary to generate a file, remove your ip-address from the debug list and re-run the report.</p>
      </cfoutput>

      <!--- We're not going to generate the file if we're in debug mode. --->
      <!--- <cfhtmlhead text='<meta http-equiv="refresh" content="0; url=#client.weburl#Batch/#login_id#_#report_name#.xls" />'> --->
      <cfelse>
        <!--- Query to get the Batch directory used by Sonis, stored in the webopts table     --->
        <!--- Used to build a save path and a file name for the spreadsheet that is generated --->

      <cfset variables.excel_query = arguments.excel_query>
      <cfset variables.report_name = arguments.report_name>
      <cfquery name="batch_dir_query" datasource="soniswebp" maxrows="1">
        select batch_dir from webopts
      </cfquery>
      <cfset batch_dir = trim(#batch_dir_query.batch_dir#) & "\">
      <cfset login_id = trim(#session.login_id#)>
      <cfset save_path = #batch_dir# & #login_id# & "_" & #report_name# & ".xls">
      <cfset file_name = #login_id# & "_" & #report_name# & ".xls">

      <cfspreadsheet action = "write" filename=#save_path# query="excel_query" overwrite="true">

      <!--- If the report is run in debug mode, only debugging info is output                                      --->
      <!--- We still get debugging information/mime doesn't change/ no about:blank + code halt  on cfcontent tag   --->
      <!--- http://www.bennadel.com/blog/1939-using-cfcontent-s-variable-attribute-stops-all-further-processing.htm--->
      <!--- Otherwise, stream the file to browser/delete it                                                        --->

          <cfheader name="Content-Disposition" value="attachment;filename=#file_name#">
          <cfcontent type="application/vnd.ms-excel" file="#save_path#" deletefile="Yes">
      </cfif>
  </cffunction>
</cfcomponent>
