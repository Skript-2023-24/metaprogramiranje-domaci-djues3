# frozen_string_literal: true

require 'google_drive'
require './column'

# A worksheet in a spreadsheet
class Worksheet
  include Enumerable
  attr_reader(:worksheet, :columns, :table)

  # Initializes a worksheet
  # Key is the key of the spreadsheet
  # Gid is the gid of the worksheet
  def initialize(session, key, gid, table = nil)
    @worksheet = session.spreadsheet_by_key(key).worksheet_by_gid(gid)
    @session = session
    @key = key
    @gid = gid
    @table = table || extract_table
    canonicalize_headers!
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

  def operable?(other)
    false unless other.is_a?(Worksheet)
    @table[0] == other.table[0]
  end

  def +(other)
    raise(ArgumentError, 'Tables are not operable.') unless operable?(other)

    @table + other.table[1..] # Concatenate rows excluding the header of other table
  end

  def -(other)
    raise(ArgumentError, 'Tables are not operable.') unless operable?(other)

    # Remove rows from self that are present in other
    @table.reject { |row| other.table.include?(row) }
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
      raise(ArgumentError, 'No table found') if rows.zero? && cols.zero?

      # Multiple column table possible
      rows.times do |i|
        cols.times do |j|
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
        # Break if the first cell of a row is empty
        break if cell_value.empty? && j == start_col

        row_data << cell_value
      end
      table_matrix << row_data
    end

    table_matrix
  end

  # Makes headers adhere to snake_case

  def canonicalize_headers
    @table[0].map { |header| header.downcase.gsub(/\s+/, '_') }
  end

  # Makes headers adhere to snake_case, overwrites the original headers
  def canonicalize_headers!
    @table[0].map! { |header| header.downcase.gsub(/\s+/, '_') }
  end

  def headers
    @table[0]
  end
end
