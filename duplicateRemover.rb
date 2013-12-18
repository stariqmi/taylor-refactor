require 'rubygems'
require 'spreadsheet'
require 'similar_text'

def get_street_number address
	return address[/\d+/]
end

def checkUniqueness address, unique_rows
	if address.nil?
		puts "\tNo address available."
		return false
	else
		addr_sn = get_street_number address
		puts "\tStreet #: #{addr_sn}"
		unique_rows.each do |ur|
			ex_addr_sn = get_street_number ur[0]
			if addr_sn == ex_addr_sn
				puts "\tStreet # matched."
				addr = address.split(" ")[1..-1].join(" ")
				ex_addr = ur[0].split(" ")[1..-1].join(" ")
				if addr.similar(ex_addr) > 80
					puts "\t#{address} is similar to #{ur[0]}"
					return false
				end
			end
		end
	end

	puts "\tUniqueness confirmed."
	return true
end

source_xls = Spreadsheet.open 'MLS Cash Buyers/main_properties.xls'
source_sheet = source_xls.worksheet 0
all_headers = source_sheet.row(0)
headers = all_headers[0...4] << "Same" << all_headers[5] << "# Owns"

unique_rows = [] 

source_sheet.each 1 do |row|
	info = row[0..4] << row[5]
	puts "================================================================="
	puts "Checking '#{info[0]}' for uniqueness:"
	unique_rows << info if checkUniqueness info[0], unique_rows
end

unique_book = Spreadsheet::Workbook.new
unique_sheet = unique_book.create_worksheet :name => 'unique'
unique_sheet.update_row 0, *headers

row_num = 1

unique_rows.each do |ur|
	puts "Writing to new xls doc : #{ur.inspect}"
	unique_sheet.update_row row_num, *ur
	row_num += 1
end

unique_book.write 'MLS Cash Buyers/unique_records.xls'