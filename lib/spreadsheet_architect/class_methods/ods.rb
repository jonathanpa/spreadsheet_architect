require 'rodf'

module SpreadsheetArchitect
  module ClassMethods
    def to_ods(opts={})
      return to_rodf_spreadsheet(opts).bytes
    end

    def to_rodf_spreadsheet(opts={}, spreadsheet=nil)
      opts = SpreadsheetArchitect::Utils.get_cell_data(opts, self)
      options = SpreadsheetArchitect::Utils.get_options(opts, self)

      if !spreadsheet
        spreadsheet = RODF::Spreadsheet.new
      end

      spreadsheet.data_style :date_format, :date do
        section :year, style: "YYYY"
        section :text, textual: "-"
        section :month, style: "MM"
        section :text, textual: "-"
        section :day, style: "DD"
      end

      spreadsheet.data_style :datetime_format, :date do
        section :year, style: "YYYY"
        section :text, textual: "-"
        section :month, style: "MM"
        section :text, textual: "-"
        section :day, style: "DD"

        section :text, textual: " "

        section :hours, style: "HH"
        section :text, textual: ":"
        section :minutes, style: "MM"
        section :text, textual: ":"
        section :seconds, style: "SS"
      end

      spreadsheet.data_style :time_format, :date do
        section :hours, style: "HH"
        section :text, textual: ":"
        section :minutes, style: "MM"
        section :text, textual: ":"
        section :seconds, style: "SS"
      end

      spreadsheet.office_style :header_style, family: :cell do
        if options[:header_style]
          SpreadsheetArchitect::Utils.convert_styles_to_ods(options[:header_style]).each do |prop, styles|
            styles.each do |k,v|
              property prop.to_sym, k => v
            end
          end
        end
      end

      ods_row_styles = options[:row_style] ? SpreadsheetArchitect::Utils.convert_styles_to_ods(options[:row_style]) : {}

      spreadsheet.office_style :row_style, family: :cell do
        ods_row_styles.each do |prop, styles|
          styles.each do |k,v|
            property prop.to_sym, k => v
          end
        end
      end

      [:date, :datetime, :time].each do |wut|
        spreadsheet.office_style "row_style_with_#{wut}", family: :cell, data_style: "#{wut}_format" do
          ods_row_styles.each do |prop, styles|
            styles.each do |k,v|
              property prop.to_sym, k => v
            end
          end
        end
      end

      spreadsheet.table options[:sheet_name] do 
        if options[:headers]
          options[:headers].each do |header_row|
            row do
              header_row.each_with_index do |header, i|
                cell header, style: :header_style
              end
            end
          end
        end

        options[:data].each_with_index do |row_data, index|
          row do 
            row_data.each_with_index do |v,i|
              if options[:types]
                type = options[:types][i]

                if y.respond_to?(:strftime)
                  case type
                  when :date
                    v = v.strftime("%Y-%m-%d")
                    s = "date_row_style"
                  when :datetime
                    v = v.strftime("%Y-%m-%d %H:%M:%S")
                    s = "datetime_row_style"
                  when :time
                    v = v.strftime("%H:%M:%S")
                    s = "time_row_style"
                  end
                end
              end

              cell v, style: (s || :row_style), type: type
            end
          end
        end
      end

      return spreadsheet
    end

  end
end
