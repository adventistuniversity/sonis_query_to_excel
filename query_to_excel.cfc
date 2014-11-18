<cfcomponent>
  <cffunction name="query_to_excel" output="yes" returntype="Any">
    <!---Function Arguments--->
    <cfargument name="excel_query" required="yes" type="query" default="" hint="The query you want output as a spreadsheet">
    <cfargument name="report_name" required="yes" type="string" default="" hint="The name of this report">
    
    <cfset variables.excel_query = arguments.excel_query>
    <cfset variables.report_name = arguments.report_name>

    <cfset save_path = #session.sonis_main# & "Batch\" & #session.login_id# & "_" & #report_name# & ".xls">
    <cfspreadsheet 
            action = "write" 
            filename=#save_path#
            query="excel_query" 
            overwrite="true">

    <cfoutput>
      <p>If your report doesn't download automatically, please 
      <a href="#client.weburl#Batch/#session.login_id#_#report_name#.xls">click here</a>.</p>
    </cfoutput>
          
    <cfhtmlhead
      text='<meta http-equiv="refresh" content="0; url=#client.weburl#Batch/#session.login_id#_#report_name#.xls" />'>

      <cfreturn TRUE>

  </cffunction>
</cfcomponent>