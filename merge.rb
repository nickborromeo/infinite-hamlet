require 'sinatra'
require './lib/report'

set :public_folder, 'reports'

get "/" do
  erb :index 
end

get "/support" do
  erb :support
end

post "/support" do
   
  report = Report.new
  s_report = report.generate_support_report(params['support-report'][:tempfile])
  
  "Process Finished"
  send_file s_report, :filename => "merged_support_cases_#{Time.new}.xls"
end

get "/customer-service" do
  erb :customer_service
end

post "/customer-service" do

end
