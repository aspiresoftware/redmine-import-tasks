# Plugin's routes
# See: http://guides.rubyonrails.org/routing.html
get '/projects/:project_id/import_issue', :to => 'excel_sheet#index'
get '/import_issue', :to => 'excel_sheet#index'
post '/upload_sheet', :to => 'excel_sheet#upload_sheet'
match '/generate_excel_sheet', :to => 'excel_sheet#generate_excel_sheet', via: [:get, :post]
match "/settings/plugin/public/uploads/exports/Redmine_Sample_Issue_Sheet.xls" ,:to=>'excel_sheet#export_excel_sheet', via: [:get, :post]
match "/public/uploads/exports/Redmine_Sample_Issue_Sheet.xls",:to=>'excel_sheet#render_excel_sheet', via: [:get, :post]
# resources :excel_sheet do
#    get 'generate_excel_sheet', on: :generate_excel_sheet
#  end
