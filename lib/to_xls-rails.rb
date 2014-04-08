# Wooseok
# I forked this version from the following repository, and update source codes to support multiple Models
# https://github.com/liangwenke/to_xls-rails
#

require 'spreadsheet'

class Array

    # Example,
    # @items = [{:name => "woo", :first => "seo"}, {:addr => "aaa", :addr2 => "bbb"}]
    # @items.hash_array_to_xls
    #
    # if you want save all array into a sheet, you can use :one_sheet => true option like following
    # @items.hash_array_to_xls(:one_sheet => true)

    def hash_array_to_xls(options = {}, &block) 
        return '' if self.empty? && options[:prepend].blank?

        columns = []
        options.reverse_merge!(:header => true)

        xls_report = StringIO.new

        Spreadsheet.client_encoding = options[:client_encoding] || "UTF-8"

        book = Spreadsheet::Workbook.new

        sheet_created = false
        @sheet = nil
        @sheet_index = 0
        self.each do |item|
            if options[:one_sheet] == nil or options[:one_sheet] != true or sheet_created == false
                @sheet = book.create_worksheet
                @sheet_index = 0
                if options[:column_width]
                    options[:column_width].each_index {|index| sheet.column(index).width = options[:column_width][index]}
                end
            else
                @sheet.update_row @sheet_index, ' '
                @sheet_index += 1
            end

            sheet_created = true
                
            item.each_with_index do |(key,value), index|
                @sheet.update_row @sheet_index, key.to_s, value.to_s
                @sheet_index += 1
            end
        end

        book.write(xls_report)
        xls_report.string
    end 


    # Example,
    # @works = Work.all
    # @users = User.all
    # @items = []
    # @items.push(@works)
    # @items.push(@users)
    # @items.model_array_to_xls

    def model_array_to_xls(options = {}, &block)
        return '' if self.empty? && options[:prepend].blank?

        columns = []
        options.reverse_merge!(:header => true)

        xls_report = StringIO.new

        Spreadsheet.client_encoding = options[:client_encoding] || "UTF-8"

        book = Spreadsheet::Workbook.new
        self.each do |item|
            sheet = book.create_worksheet
            if options[:only]
              columns = Array(options[:only]).map(&:to_sym)
            elsif !item.empty?
              columns = item.first.class.column_names.map(&:to_sym) - Array(options[:except]).map(&:to_sym)
            end

            continue if columns.empty? && options[:prepend].blank?

            sheet_index = 0

            unless options[:prepend].blank?
              options[:prepend].each do |array|
                sheet.row(sheet_index).concat(array)
                sheet_index += 1
              end
            end

            if options[:header]
              sheet.row(sheet_index).concat(options[:header_columns].blank? ? columns.map(&:to_s).map(&:humanize) : options[:header_columns])
              sheet_index += 1
            end

            if options[:column_width]
              options[:column_width].each_index {|index| sheet.column(index).width = options[:column_width][index]}
            end

            item.each_with_index do |obj, index|
              if block
                sheet.row(sheet_index).replace(columns.map { |column| block.call(column, obj.send(column), index) })
              else
                sheet.row(sheet_index).replace(columns.map { |column| obj.send(column) })
              end
              sheet_index += 1
            end

            unless options[:append].blank?
              options[:append].each do |array|
                sheet.row(sheet_index).concat(array)
                sheet_index += 1
              end
            end
        end

        book.write(xls_report)
        xls_report.string
    end


    # Example,
    # @works = Work.all
    # @works.to_xls
 
    def to_xls(options = {}, &block)

        return '' if self.empty? && options[:prepend].blank?

        columns = []
        options.reverse_merge!(:header => true)

        xls_report = StringIO.new

        Spreadsheet.client_encoding = options[:client_encoding] || "UTF-8"

        book = Spreadsheet::Workbook.new
        sheet = book.create_worksheet

        if options[:only]
            columns = Array(options[:only]).map(&:to_sym)
        elsif !self.empty?
            columns = self.first.class.column_names.map(&:to_sym) - Array(options[:except]).map(&:to_sym)
        end

        return '' if columns.empty? && options[:prepend].blank?

        sheet_index = 0

        unless options[:prepend].blank?
            options[:prepend].each do |array|
                sheet.row(sheet_index).concat(array)
                sheet_index += 1
            end
        end

        if options[:header]
            sheet.row(sheet_index).concat(options[:header_columns].blank? ? columns.map(&:to_s).map(&:humanize) : options[:header_columns])
            sheet_index += 1
        end

        if options[:column_width]
            options[:column_width].each_index {|index| sheet.column(index).width = options[:column_width][index]}
        end

        self.each_with_index do |obj, index|
            if block
                sheet.row(sheet_index).replace(columns.map { |column| block.call(column, obj.send(column), index) })
            else
                sheet.row(sheet_index).replace(columns.map { |column| obj.send(column) })
            end

            sheet_index += 1
        end

        unless options[:append].blank?
            options[:append].each do |array|
                sheet.row(sheet_index).concat(array)
                sheet_index += 1
            end
        end

        book.write(xls_report)

        xls_report.string
    end

end

