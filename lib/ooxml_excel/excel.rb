module OOXML
  class Excel
    include OOXML::Helper::List
    attr_reader :comments
    def initialize(spreadsheet, load_only_sheet:nil)
      @spreadsheet = spreadsheet
      @workbook = nil
      @sheets = {}
      @comments = {}
      # @themes = []
      @load_only_sheet = nil
      @styles = nil
      load_xml_contents
    end

    def sheets
      @workbook.sheets.map { |sheet| sheet[:name]}
    end

    def sheet(sheet_name)
      sheet_from_workbook = @workbook.sheets.find { |sheet| sheet[:name] == sheet_name}
      raise "No #{sheet_name} in workbook." if sheet_from_workbook.blank?
      sheet = @sheets.fetch(sheet_from_workbook[:relationship_id])

      # shared variables
      sheet.name = sheet_name
      sheet.comments = @comments[sheet_from_workbook[:relationship_id]]
      sheet.styles = @styles
      sheet.defined_names = @workbook.defined_names
      sheet
    end

    def [](text)
      # immediately treat as cell range if an exclamation point is detected
      # otherwise, normally load a sheet
      text.include?('!') ? load_list_values(text) : sheet(text)
    end

    def named_range(name)
      # "Hidden11390550_39"=>"Hidden!$B$734:$B$735"
      # ooxml.named_range('Hidden11390550_107')
      # a typical named range would be be
      # yes_no => 'Lists'!$A$1:$A$6
      defined_name = @workbook.defined_names[name]
      load_list_values(defined_name) if defined_name.present?
    end

    private

    def load_list_values(range_text)
      # get the sheet name => 'Lists'
      sheet_name = range_text.gsub(/[\$\']/, '').scan(/^[^!]*/).first
      # fetch the cell range => '$A$1:$A$6'
      cell_range = range_text.gsub(/\$/, '').scan(/(?<=!).+/).first
      # get the sheet object and fetch the cells in range
      sheet(sheet_name).list_values_from_formula(cell_range)
    end

    # Currently supports DataValidation (comment), columns (check if hidden)
    # TODO: list values, Font styles (if bold, colored in red etc..), background color
    def load_xml_contents
      shared_strings = []
      Zip::File.open(@spreadsheet) do |spreadsheet_zip|
        spreadsheet_zip.each do |entry|
          stream_xml = entry.get_input_stream.read
          if entry.name[/xl\/worksheets\/sheet(\d+)?\.xml/]
            sheet_id = entry.name.scan(/xl\/worksheets\/sheet(\d+)?\.xml/).flatten.first
            @sheets[sheet_id] =  Excel::Sheet.load_from_stream(stream_xml, shared_strings)
          elsif entry.name[/xl\/styles\.xml/]
            @styles = Excel::Styles.load_from_stream(stream_xml)
          # elsif entry.name[/xl\/theme(\d+)?\.xml/]
          #   @themes << Excel::Theme.load_from_stream(stream_xml)
          elsif entry.name[/xl\/comments(\d+)?\.xml/]
            comment_id = entry.name.scan(/xl\/comments(\d+)\.xml/).flatten.first
            @comments[comment_id] = Excel::Comments.load_from_stream(stream_xml)
          elsif entry.name == "xl/sharedStrings.xml"
            Nokogiri.XML(stream_xml).remove_namespaces!.xpath('sst/si').each do |shared_string_node|
              shared_strings << shared_string_node.at('t').text
            end
          elsif entry.name == "xl/workbook.xml"
            @workbook = Workbook.load_from_stream(stream_xml)
          end
        end
      end
    end
  end
end
