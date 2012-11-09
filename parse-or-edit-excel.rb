require 'rubygems'
require 'roo'

output_file = Google.new("0Aqok6876FoYmdDFQbHpiQXpndzhYZlJ3SFM0ZDkzUGc", "<your-gmail-id>", "<your-gmail-password>")
output_file.default_sheet = output_file.sheets.first

xl= Libreoffice.new("employee-list.ods")
xl.default_sheet = xl.sheets.first
(xl.first_row).upto(xl.last_row) do |line|
  emp_id = xl.cell(line,'A')
  first_name = xl.cell(line,'B')
  last_name  = xl.cell(line,'C')
  start_date  = xl.cell(line,'D')
  dept  = xl.cell(line,'E')

  if start_date
    puts "#{start_date}\t#{last_name}\t#{emp_id}\t#{dept}"
  end
  {'A' => emp_id, 'B' => first_name, 'C' => last_name, 'D' => start_date, 'E' => dept}.each do |cell, value|
    output_file.set_value(line, cell, value )
  end
end


