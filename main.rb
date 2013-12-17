require_relative 'county'
require 'spreadsheet'

Dir.chdir 'MLS Cash Buyers' # Navigate to the main directory holding all county folders

mlsDir = Dir.pwd    # Variable holding current directory
properties = []
xls_headers = []

Dir.entries(mlsDir).each do |dir|   # Loop through all valid county folders
    if dir.include? 'County'  # If the folder is a County folder
        puts "\nEntering the '#{dir}' folder ... \n\n"
        Dir.chdir "#{mlsDir}/#{dir}"  # Navigate to the County folder in iteration
        county = County.new Dir.pwd # Create a new County object with the county directory
        county.filterProperties  # Call to county instance method, collects all properties common in xls and pdf
        # county.print    # Call to county instance method, prints all properties collected above
        county.createXLS    # Call to county instance method, creates an xls using the properties collected above
        xls_headers = county.mainRows[0]
        properties.concat county.mainRows[1..-1]   # adds above collected properties to global properties array
        Dir.chdir "../"
    end
end

properties.insert 0, xls_headers


# Writing all properties in the "properties" array to main/central xls file
book = Spreadsheet::Workbook.new
sheet = book.create_worksheet :name => "filtered"
i = 0
properties.each do |property|
    sheet.update_row i, *property
    i += 1
end
book.write 'main_properties.xls'