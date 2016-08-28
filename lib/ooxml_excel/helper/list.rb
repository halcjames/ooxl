module OOXML
  module Helper
    module List

      # excel.rb
      # fetch dropdown values based on given data validation formula
      # this will be depracated: use #named_range instead
      def list_values(formula)
        # "Lists!$J$2:$J$4"
        # transform into useful info

        # for list values explicitly stated
        if formula.include?(',')
          formula.gsub('"', '').split(',')
          # invalid format
        elsif !formula.include?('!') && formula[/$/]
          puts "Warning: This formula is not yet supported: #{formula} in your Data Validation's formula."
          []
        else
          # # required for fetching values
          sheet_name = formula.gsub(/[\$\']/, '').scan(/^[^!]*/).first
          cell_range_formula = formula.gsub(/\$/, '').scan(/(?<=!).+/).first

          # fetch the sheet of the cell reference
          working_sheet = sheet(sheet_name)

          # gather values
          list_values = working_sheet.list_values_from_formula(cell_range_formula)
        end
      end

      # Used in sheet.rb
      def list_value_formula(cell_ref)
        data_validation = data_validations.find { |data_validation| data_validation.sqref_range.include?(cell_ref)}
        if data_validation.respond_to?(:type) && data_validation.type == "list"
          if data_validation.formula[/[\s\$\,\:]/]
            (data_validation.formula[/\$/].present?) ? "#{name}!#{data_validation.formula}" : data_validation.formula
          else
            @defined_names.fetch(data_validation.formula)
          end
        end
      end

      def list_values_from_formula(formula)
        return [] if formula.blank?

        # Formula values separated by comma
        if formula.include?(":")
          cell_letters = formula.gsub(/[\d]/, '').split(':')
          start_index, end_index = formula.gsub(/[^\d:]/, '').split(':').map(&:to_i)

          cell_letter = cell_letters.uniq.first
          (start_index..end_index).to_a.map do |row_index|
            row = rows[row_index-1]
            next if row.blank?
            row["#{cell_letter}#{row_index}"].value
          end
        else
          # when only one value: B2
          row_index = formula.gsub(/[^\d:]/, '').split(':').map(&:to_i).first
          row = rows[row_index-1]
          return if row.blank?
          [row[formula].value]
        end
      end
    end
  end
end
