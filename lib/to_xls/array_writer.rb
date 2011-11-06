require 'rubygems'
require 'stringio'
require 'spreadsheet'

module ToXls

  class ArrayWriter
    def initialize(array, options = {})
      @array = array
      @options = options
    end

    def write_string(string = '')
      io = StringIO.new(string)
      write_io(io)
      io.string
    end

    def write_io(io)
      book = Spreadsheet::Workbook.new
      write_book(book)
      book.write(io)
    end

    def write_book(book)
      sheet = book.create_worksheet
      sheet.name = @options[:name] || 'Sheet 1'
      write_sheet(sheet)
      return book
    end

    def write_sheet(sheet)
      if columns.any?
        row_index = 0

        if headers_should_be_included?
          fill_row(sheet.row(0), headers)
          row_index = 1
        end

        @array.each do |model|
          row = sheet.row(row_index)
          fill_row(row, columns, model)
          row_index += 1
        end
      end
    end

    def columns
      return  @columns if @columns
      @columns = @options[:columns]
      raise ArgumentError.new(":columns (#{columns}) must be an array or nil") unless (@columns.nil? || @columns.is_a?(Array))
      @columns ||=  can_get_columns_from_first_element? ? get_columns_from_first_element : []
    end

    def can_get_columns_from_first_element?
      @array.first && 
      (@array.first.respond_to?(:attributes) &&
      @array.first.attributes.respond_to?(:keys) &&
      @array.first.attributes.keys.is_a?(Array)) or
      (@array.first.respond_to?(:keys) &&
       @array.first.keys.is_a?(Array))
    end

    def get_columns_from_first_element
      
      keys =
        if @array.first.respond_to?(:attributes)
          @array.first.attributes.keys
        elsif @array.first.respond_to?(:keys)  
          @array.first.keys
        end
      
      keys.sort_by {|sym| sym.to_s}.collect.to_a
    end

    def headers
      return  @headers if @headers
      @headers = @options[:headers] || columns
      raise ArgumentError, ":headers (#{@headers.inspect}) must be an array" unless @headers.is_a? Array
      if @options[:humanize_headers]
        @headers = @headers.map {|c| c.to_s.humanize}
      end
      @headers
    end

    def headers_should_be_included?
      @options[:headers] != false
    end

private

    def fill_row(row, column, row_data=nil)
      case column
      when String, Symbol
        if row_data and row_data.class == Hash
          row.push(row_data[column] )
        else  
          row.push(row_data ? row_data.send(column)  : column)
        end 
      when Hash
        column.each{|key, values| fill_row(row, values, row_data && (row_data.class == Hash ? row_data[key]  : row_data.send(key)) )}
      when Array
        column.each{|value| fill_row(row, value, row_data)}
      else
        raise ArgumentError, "column #{column} has an invalid class (#{ column.class })"
      end
    end

  end

end
