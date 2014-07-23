redmine-import-tasks
====================

Import tasks plugin for Redmine 2.2+


Create issues in Redmine  from Excel Sheet/SpreadSheet (Estimation Sheet).

###Installing Issue Importer Plugin
To install Issue Importer plugin do the following steps:

* Go to the {Redmine-Root}/plugins directory in terminal.
* Exceute command : git clone https://github.com/aspiresoftware/redmine-import-tasks.git issue_importer_xls
* After successful cloning restart Redmine Application to see the plugin in action

#####Please note that the plugin uses Roo gem to read & parse  Excel Sheet that depends on Ruby 1.9+ .
So ,plugin only works with Redmine running on ruby 1.9+

###Plugin Configuration
* This plugin uses default configuration to create issues from Excel/SpreadSheet
* You can also download the Excel Template generated as per your plugin configuration and then just fill the Issues in Excel sheet and upload it 

*  To change configuration go to Administration -> Plugins and configure the Issue Importer Plugin as per your Excel Sheet fields
