require 'rubygems'
require 'spreadsheet'
require 'similar_text'

source_xls = Spreadsheet.open 'MLS Cash Buyers/main_properties.xls'
source_sheet = source_xls.worksheet 0
all_headers = source_sheet.row(0)
headers = all_headers[0...4] << "Same" << all_headers[5] << "# Owns"

def checkUniqueness row, unique_rows
	unique_rows.each do |ur|
		check = row[0] || ""
		puts "Comparing '#{check}' with  '#{ur[0]}'"
		if check == ""
			puts "No mailing address ... "
			return false
		elsif check.similar(ur[0]) > 65
			puts "#{ur[0]} is similar to #{check}"
			return false
		end
	end
	puts "Uniqueness confirmed.\n"
	return true
end

unique_rows = [] 

source_sheet.each 1 do |row|
	info = row[0..4] << row[5]
	puts "\n\tChecking '#{info[0]}' for uniqueness ... "
	unique_rows << info if checkUniqueness info, unique_rows
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