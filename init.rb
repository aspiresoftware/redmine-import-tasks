Redmine::Plugin.register :issue_importer_xls do
  name 'Issue Importer Xls plugin'
  author 'Aspire Software Solutions'
  description 'Import Excel Sheet to create Redmine issues'
  url 'http://aspiresoftware.co.in'
  author_url 'http://aspiresoftware.co.in'
  version '0.0.1'


  permission :excel_sheet, { :excel_sheet => [:index, :upload_sheet] }, :public => true
  # menu :project_menu, :polls, { :controller => 'polls', :action => 'index' }, :caption => 'Polls', :after => :activity, :param => :project_id

  # menu :application_menu, :issue_importer_xls, { :controller => 'excel_sheet', :action => 'index' }, 
  # 			:caption => 'Import Issues' ,:last => true
  menu :project_menu, :excel_sheet, { :controller => 'excel_sheet', :action => 'index' }, 
  			:caption => 'Import Issues' ,:last => true

  settings :default => {'task_column' => 0,'save_task_as' => 2,'average_hour_column' => 3 ,'start_date_column' => 6,
                        'end_date_column' => 7 ,'asignee_name_column' => 9 ,'task_description_column' => 4}, :partial => 'settings/issue_importer_setting'
        
end
