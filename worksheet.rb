# frozen_string_literal: true

require 'google_drive'
require 'column'

# A worksheet in a spreadsheet
class Worksheet
  include Enumerable
  attr_reader(:worksheet, :columns, :table)

  def initialize(session, key, index)
    @worksheet = session.spreadsheet_by_key(key).worksheets[index]
    @table = extract_table
    @columns = canonicalize_headers
  end

  def to_s
    @table
  end

  def save
    @worksheet.save
  end

  def each(&block)
    return enum_for(:each) unless block_given? # Return an enumerator if no block is given

    @table.each do |row|
      row.each(&block)
    end
  end

  def row(row_index)
    @table[row_index - 1] # Adjust row_index since indices start from 0
  end

  def rows
    @worksheet.rows
  end

  # Returns the first non-empty cell in the worksheet.
  # Exists just to show that the functionality exists.
  def find_table_wrapper
    find_table
  end

  def [](key)
    return unless key.is_a?(String) && @table[0].include?(key)

    # Start row and column of the table
    row, column = find_table
    Column.new(@table, key, @worksheet, row, column)
  end

  def method_missing(method_name, *_args)
    column_name = method_name.to_s
    #     "no such header name in: #{@table[0]} in worksheet: #{@worksheet.gid}"
    unless @table[0].include?(column_name)
      raise(NoMethodError,
            "undefined method `#{method_name}'! No such header name in #{@table[0]} in worksheet: #{@worksheet.gid}")
    end

    define_singleton_method(method_name) do
      Column.new(@table, column_name, @worksheet, 1, @table[0].index(column_name) + 1)
    end
    send(method_name)
  end

  def respond_to_missing?(method_name, include_private = false)
    column_name = method_name.to_s
    @table[0].include?(column_name) || super
  end

  private

  # Memoizes the first non-empty cell in the worksheet.
  # Column major order traversal.
  # returns [row, col] of the first non-empty cell
  def find_table
    @find_table ||= begin
      rows = @worksheet.num_rows
      cols = @worksheet.num_cols
      return if rows.zero? && cols.zero?

      # Multiple column table possible
      cols.times do |i|
        rows.times do |j|
          return [i + 1, j + 1] unless @worksheet[i + 1, j + 1].empty?
        end
      end

      # Only single column table possible
      return if @worksheet[rows, cols].empty?

      [rows, cols]
    end
  end

  # Returns the table found in the worksheet as a matrix
  def extract_table
    table_start = find_table

    return nil if table_start.nil?

    start_row, start_col = table_start
    rows = @worksheet.num_rows
    cols = @worksheet.num_cols

    table_matrix = []

    (start_row..rows).each do |i|
      row_data = []
      (start_col..cols).each do |j|
        cell_value = @worksheet[i, j]
        break if cell_value.empty? && j == start_col # Break if the first cell of a row is empty

        row_data << cell_value
      end
      table_matrix << row_data
    end

    table_matrix
  end

  def canonicalize_headers
    @table[0].map { |header| header.downcase.gsub(/\s+/, '_') }
  end

  def canonicalize_headers!
    @table[0].map! { |header| header.downcase.gsub(/\s+/, '_') }
  end

end

