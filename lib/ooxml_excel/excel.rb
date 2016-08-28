module OOXML
  class Excel
    include OOXML::Helper::List
    attr_reader :comments
    def initialize(spreadsheet, load_only_sheet:nil)
      @spreadsheet = spreadsheet
      @sheets = {}
      @comments = {}
      # @themes = []
      @workbook = nil
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

    private

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
