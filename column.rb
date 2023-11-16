# frozen_string_literal: true

require 'google_drive'

# comment
class Column
  def initialize(table, header_name, worksheet, start_row, start_col)
    @table = table
    @header_name = header_name
    @column_index = table[0].index(header_name)
    @worksheet = worksheet
    @start_row = start_row
    @start_col = start_col
  end

  def [](row_index)
    @table[row_index][@column_index]
  end

  def []=(row_index, value)
    @table[row_index][@column_index] = value
    update_worksheet(row_index, value)
  end

  def sum
    numeric_values.sum
  end

  def average
    sum.to_f / numeric_values.length
  end

  def row_by_identifier(identifier)
    row_index = @table[@start_row..].index { |row| row[@column_index] == identifier }
    return nil unless row_index

    @table[@start_row + row_index]
  end

  # Make sure to properly convert the values into numeric values before using this method
  def map
    filtered_values = @table[@start_row..].reject { |row| subtotal_row?(row) }
    filtered_values.map { |row| yield row[@column_index].to_i }
  end

  # Make sure to properly convert the values into numeric values before using this method
  def select(&block)
    filtered_values = @table[@start_row..].reject { |row| subtotal_row?(row) }
    filtered_values.map { |row| row[@column_index] }.select(&block)
  end
  # Make sure to properly convert the values into numeric values before using this method

  def reduce(initial = nil)
    filtered_values = @table[@start_row..].reject { |row| subtotal_row?(row) }
    result = initial
    filtered_values.each do |row|
      result = result.nil? ? row[@column_index].to_i : yield(result, row[@column_index].to_i)
    end
    result
  end

  def to_s
    @table[@start_row..].map { |row| row[@column_index] }.join("\n")
  end

  private

  def update_worksheet(row_index, value)
    worksheet_row = @start_row + row_index
    worksheet_col = @start_col + @column_index
    @worksheet[worksheet_row, worksheet_col] = value
    @worksheet.save
  end

  def numeric_values
    @table[@start_row..].map { |row| row[@column_index] }.compact.map(&:to_f)
  end

  def subtotal_row?(row)
    row.any? { |cell| cell.to_s.downcase.include?('subtotal') }
  end
end
