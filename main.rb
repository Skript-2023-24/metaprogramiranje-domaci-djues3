# frozen_string_literal: true

require './worksheet'
key = '1E4s88FL79ff-hrAyCDrDHkdAN2ZJTvm-f59fNxguxYE'
session = GoogleDrive::Session.from_config('config.json')

# worksheet = Worksheet.new(session, key, 0)
#
# ## Displays memoization
# p 'call'
# a, b = worksheet.find_table_wrapper
#
# if a && b
#   rows_content = worksheet.rows
#   # cell_content = worksheet.at a, b
#
#   p rows_content
#   # p cell_content
# else
#   puts 'No non-empty cell found'
# end
# sleep 5
# p 'back'
# c, d = worksheet.find_table_wrapper
# p worksheet.row(c)[d]
#
# ## Displays table printing
#
# table_matrix = worksheet.table
#
# if table_matrix.nil?
#   puts 'No table found in the worksheet.'
# else
#   puts 'Table found!'
#   puts 'Table Matrix:'
#   table_matrix.each { |row| puts row.join("\t") }
# end
#
# ## Displays each
# worksheet.each do |cell|
#   p cell
# end
#
# ## Displays row
# table = worksheet
#
# # Accessing the third row (index 2) and printing its elements
# row = table.row(3)
# puts 'Elements in the third row:'
# puts row.inspect # Displays the entire row
#
# # Accessing a specific element in the third row (index 2) using array index operations
# element = row[1] # Accessing the second element in the row (index 1)
# puts "Element at index 1 in the third row: #{element}"

# ## Displays column access
# # Accessing a column by header name
# column = worksheet['a']
#
# # Accessing a cell in the column by row index
# cell = column[2] # Represents cell value at row index 2
#
# # Retrieving cell value
# puts "Cell value before update: #{cell}"
#
# # Updating cell value
# column[2] = 15 # Setting cell value at row index 2
# worksheet.save
# # Retrieving updated cell value
# puts "Cell value after update: #{column[2]}"

# ### METAPROGRAMMING
# ## Provides access to columns by header name
# p worksheet.a[0]
# p worksheet.c.sum
# p worksheet.b.average
# ## Provides access to rows by identifier
# row = worksheet.a.row_by_identifier('3')
# if row.nil?
#   puts "Row not found for identifier '2'"
# else
#   puts "Row found: #{row}"
# end

#
# # ## map, select, reduce
# column = worksheet.a
#
# # Using map to increment each cell value by 1
# mapped_values = column.map { |cell| cell.to_i + 1 }
#
# # Using select to filter cells meeting a condition (e.g., select cells greater than 10)
# selected_values = column.select { |cell| cell.to_i > 3 }
#
# # Using reduce to calculate the sum of the column cells
# sum = column.reduce(0) { |acc, cell| acc + cell.to_i }
#
# puts "Mapped values: #{mapped_values}"
# puts "Selected values: #{selected_values}"
# puts "Sum of column: #{sum}"
#
# mapped_values = column.map { |cell| cell + 1 }
#
# # Using select to filter cells meeting a condition (excluding subtotal rows)
# selected_values = column.select { |cell| cell.to_i > 10 }
#
# puts "Mapped values: #{mapped_values}"
# puts "Selected values: #{selected_values}"
#
# # Using reduce to calculate the sum of the column cells (excluding subtotal rows)
# sum = column.reduce(0) { |acc, cell| acc + cell }
#
# puts "Sum of column (excluding subtotal rows): #{sum}"

worksheet_operable_a = Worksheet.new(session, key, 924_246_715)
worksheet_operable_b = Worksheet.new(session, key, 864_058_606)
inoperable = Worksheet.new(session, key, 1_188_560_897)
# p worksheet_operable_a.to_s
# p worksheet_operable_b.to_s

p worksheet_operable_a.operable? worksheet_operable_b

p(worksheet_operable_b + worksheet_operable_a)
# Should raise an error
# inoperable + worksheet_operable_a

p(worksheet_operable_a - worksheet_operable_b)
