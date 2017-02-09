require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'

##################
# Define Variables
##################
book_path = 'Sample.xls'
task = 'concatenate'
order = 'rows'
separator = ','

book = Spreadsheet.open book_path

# # We can either access all the worksheets in a workbook...
# book.worksheets
# # ...or access them by index or name (encoded in your client_encoding).
# sheet1 = book.worksheet 0
# sheet2 = book.worksheet 'Sheet1'

# # Now you can either iterate over all rows that contain some data. A call to Worksheet.each without arguments will omit empty rows at the beginning of the worksheet:
# sheet1.each do |row|
#   # do something interesting with a row
# end

# # Or you can access rows directly, by their index (0-based):
# 
# row = sheet1.row(3)
# # To access the values stored in a row, treat the row like an array.
# 
# row[0]
# # This will return a String, a Float, an Integer, a Formula, a Link or a Date or DateTime object - or nil if the cell is empty.
# 
# # More information about the formatting of a cell can be found in the format with the equivalent index:
# 
# row.format 2

book.write 'out.xls'
