require 'spreadsheet'

class Report

  def generate_support_report(support_report)

    s_report = Spreadsheet.open(support_report)

    support_stats = s_report.worksheet(0)
    support_cases = s_report.worksheet(1)

    merge_book = Spreadsheet::Workbook.new
    merge_sheet = merge_book.create_worksheet #=> "Name of Sheet"
    created_time = Time.new

    #Generate the Headers
    merge_sheet.row(0).push("Case Number")
    merge_sheet.row(0).push("Type")
    merge_sheet.row(0).push("Bug ID")
    merge_sheet.row(0).push("Product Family")
    merge_sheet.row(0).push("Product")
    merge_sheet.row(0).push("Product Version")
    merge_sheet.row(0).push("Account Name")
    merge_sheet.row(0).push("Contact Name")
    merge_sheet.row(0).push("Subject")
    merge_sheet.row(0).push("Case Owner")
    merge_sheet.row(0).push("Opened Date")
    merge_sheet.row(0).push("Last Date Modified")
    merge_sheet.row(0).push("Age")
    merge_sheet.row(0).push("Status")
    merge_sheet.row(0).push("Case Origin")
    merge_sheet.row(0).push("Region")
    merge_sheet.row(0).push("Created By")
    merge_sheet.row(0).push("Case Record Type")
    merge_sheet.row(0).push("Initial response time")
    merge_sheet.row(0).push("Handling of support case")
    merge_sheet.row(0).push("Resolution of support case")
    merge_sheet.row(0).push("Overall satisfaction with support case")
    merge_sheet.row(0).push("Product quality")
    merge_sheet.row(0).push("Product features")
    merge_sheet.row(0).push("Product usability")
    merge_sheet.row(0).push("Overall satisfaction with product")
    merge_sheet.row(0).push("Open-Ended Response")
    #
    merge_book.write "./reports/merged_support_cases_#{created_time}.xls"

    count = 1 #start with row after the headers
    line_number = 0

    support_cases.each do |sc_row|
      support_stats.each do |cssc_row|
        if (cssc_row[8] == sc_row[0].to_i)
          merge_sheet.row(count).push(sc_row[0])
          merge_sheet.row(count).push(sc_row[1])
          merge_sheet.row(count).push(sc_row[2])
          merge_sheet.row(count).push(sc_row[3])
          merge_sheet.row(count).push(sc_row[4])
          merge_sheet.row(count).push(sc_row[5])
          merge_sheet.row(count).push(sc_row[6])
          merge_sheet.row(count).push(sc_row[7])
          merge_sheet.row(count).push(sc_row[8])
          merge_sheet.row(count).push(sc_row[9])
          merge_sheet.row(count).push(sc_row[10])
          merge_sheet.row(count).push(sc_row[11])
          merge_sheet.row(count).push(sc_row[12])
          merge_sheet.row(count).push(sc_row[13])
          merge_sheet.row(count).push(sc_row[14])
          merge_sheet.row(count).push(sc_row[15])
          merge_sheet.row(count).push(sc_row[16])
          merge_sheet.row(count).push(sc_row[17])
          merge_sheet.row(count).push(cssc_row[9])
          merge_sheet.row(count).push(cssc_row[10])
          merge_sheet.row(count).push(cssc_row[11])
          merge_sheet.row(count).push(cssc_row[12])
          merge_sheet.row(count).push(cssc_row[13])
          merge_sheet.row(count).push(cssc_row[14])
          merge_sheet.row(count).push(cssc_row[15])
          merge_sheet.row(count).push(cssc_row[16])
          merge_sheet.row(count).push(cssc_row[17])
          count += 1
          puts "Merging matched entry"
          merge_book.write "./reports/merged_support_cases_#{created_time}.xls"
        end
      end
    end

    return "./reports/merged_support_cases_#{created_time}.xls"
  end

  def generate_customer_service_report(cs_cases)


    cs_report = Spreadsheet.open(cs_cases)

    customer_service_stats = cs_report.worksheet(0)
    customer_service_cases = cs_report.worksheet(1)

    merge_book = Spreadsheet::Workbook.new
    merge_sheet = merge_book.create_worksheet #=> "Name of Sheet"
    created_time = Time.new

    #Generate the Headers
    merge_sheet.row(0).push("Case Number")
    merge_sheet.row(0).push("Type")
    merge_sheet.row(0).push("Bug ID")
    merge_sheet.row(0).push("Product Family")
    merge_sheet.row(0).push("Product")
    merge_sheet.row(0).push("Product Version")
    merge_sheet.row(0).push("Account Name")
    merge_sheet.row(0).push("Contact Name")
    merge_sheet.row(0).push("Subject")
    merge_sheet.row(0).push("Case Owner")
    merge_sheet.row(0).push("Opened Date")
    merge_sheet.row(0).push("Case Last Modified Date")
    merge_sheet.row(0).push("Age")
    merge_sheet.row(0).push("Status")
    merge_sheet.row(0).push("Case Origin")
    merge_sheet.row(0).push("Region")
    merge_sheet.row(0).push("Created By")
    merge_sheet.row(0).push("Case Record Type")
    merge_sheet.row(0).push("Response Time")
    merge_sheet.row(0).push("Communication")
    merge_sheet.row(0).push("Resolution")
    merge_sheet.row(0).push("Overall satisfaction")
    merge_sheet.row(0).push("Open-Ended Response")
    merge_sheet.row(0).push("How did you contact support?")
    
    merge_book.write "./reports/merged_customer_service_cases_#{created_time}.xls"
    
    count = 1 #start with row after the headers
    customer_service_stats.each do |customer_sat_row|
      customer_service_cases.each do |service_row|
        if (customer_sat_row[0] == service_row[0])
          merge_sheet.row(count).push(service_row[0])
          merge_sheet.row(count).push(service_row[1])
          merge_sheet.row(count).push(service_row[2])
          merge_sheet.row(count).push(service_row[3])
          merge_sheet.row(count).push(service_row[4])
          merge_sheet.row(count).push(service_row[5])
          merge_sheet.row(count).push(service_row[6])
          merge_sheet.row(count).push(service_row[7])
          merge_sheet.row(count).push(service_row[8])
          merge_sheet.row(count).push(service_row[9])
          merge_sheet.row(count).push(service_row[10])
          merge_sheet.row(count).push(service_row[11])
          merge_sheet.row(count).push(service_row[12])
          merge_sheet.row(count).push(service_row[13])
          merge_sheet.row(count).push(service_row[14])
          merge_sheet.row(count).push(service_row[15])
          merge_sheet.row(count).push(service_row[16])
          merge_sheet.row(count).push(service_row[17])
          merge_sheet.row(count).push(customer_sat_row[1])
          merge_sheet.row(count).push(customer_sat_row[2])
          merge_sheet.row(count).push(customer_sat_row[3])
          merge_sheet.row(count).push(customer_sat_row[4])
          merge_sheet.row(count).push(customer_sat_row[5])
          merge_sheet.row(count).push(customer_sat_row[6])

          count += 1
          merge_book.write "./reports/merged_customer_service_cases_#{created_time}.xls"
        end
      end
    end

    return "./reports/merged_customer_service_cases_#{created_time}.xls"
  end
end
