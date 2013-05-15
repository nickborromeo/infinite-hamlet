require 'sinatra'
require './lib/report'

set :public_folder, 'public'

get "/" do
  erb :index 
end

get "/support" do
  erb :support
end

post "/support" do
   
  report = Report.new
  s_report = report.generate_support_report(params['support-report'][:tempfile])
  
  send_file s_report, :filename => "merged_support_cases_#{Time.new}.xls"
end

get "/customer-service" do
  erb :customer_service
end

post "/customer-service" do

  report = Report.new
  cs_report = report.generate_customer_service_report(params['customer-service-report'][:tempfile])

  send_file cs_report, :filename => "merged_customer_service_cases_#{Time.new}.xls"
end


