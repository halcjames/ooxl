class OOXL
  module ListHelper
    # fetch dropdown values based on given data validation formula
    def list_values(formula)
      # "Lists!$J$2:$J$4"

      # for list values explicitly stated
      if formula.include?(',')
        formula.gsub('"', '').split(',')
      elsif !formula.include?('!') && formula[/$/]
        puts "Warning: This formula is not yet supported: #{formula} in your Data Validation's formula."
        []
      else
        # # required for fetching values
        sheet_name = formula.gsub(/[\$\']/, '').scan(/^[^!]*/).first
        cell_range = formula.gsub(/\$/, '').scan(/(?<=!).+/).first

        # fetch the sheet of the cell reference
        working_sheet = sheet(sheet_name)

        # gather values
        list_values = working_sheet.load_cell_range(cell_range)
      end
    end
  end
end
