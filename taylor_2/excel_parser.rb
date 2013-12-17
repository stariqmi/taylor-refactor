require 'rubygems'
require 'simple_xlsx_reader'

class EXCEL_Parser
  
    attr_accessor :properties, :headers

    def initialize filename
        excelDoc = SimpleXlsxReader.open filename   # SimpleXlsReader is used to open xlsx doc
        sheet = findSheetByName "Tax Import", excelDoc # Call to helper function to find the Tax Import sheet
        @properties = getProperties sheet # Call to helper function to get all properties from the sheet
        all_headers = @properties.slice! 0
        all_headers.delete_at 2
        @headers = all_headers
    end

    def findSheetByName name, excelDoc
        excelDoc.sheets.each do |sheet|
          return sheet if sheet.name == name
        end
        excelDoc.sheets[0]
    end

    def getProperties sheet
        rows = []
        sheet.rows.each do |row|
            rows << row if isNotEmpty row
        end
        rows 
    end

    def isNotEmpty row
        check = false
        row.each do |field|
          return true unless field.nil?
        end
        check
    end
end