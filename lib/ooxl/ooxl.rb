class OOXL
  include Enumerable
  include ListHelper

  def initialize(spreadsheet_filepath, options={})
    @workbook = nil
    @sheets = {}
    @styles = []
    @comments = {}
    @relationships = {}
    @options = options
    @tables = []
    parse_spreadsheet_contents(spreadsheet_filepath)
  end

  def self.open(spreadsheet_filepath, options={})
    new(spreadsheet_filepath, options)
  end

  def sheets(skip_hidden: false)
    @workbook.sheets.map do |sheet|
      next if sheet[:state] != 'visible' &&  (@options[:skip_hidden_sheets] || skip_hidden)
      sheet[:name]
    end.compact
  end

  def each
    sheets.each do |sheet_name|
      yield sheet(sheet_name)
    end
  end

  def sheet(sheet_name)
    sheet_index = @workbook.sheets.index { |sheet| sheet[:name] == sheet_name}
    raise "No #{sheet_name} in workbook." if sheet_index.nil?
    sheet = @sheets.fetch((sheet_index+1).to_s)

    # shared variables
    sheet.name = sheet_name
    sheet.comments = fetch_comments(sheet_index)
    sheet.styles = @styles
    sheet.defined_names = @workbook.defined_names
    sheet
  end

  def [](text)
    # immediately treat as cell range if an exclamation point is detected
    # otherwise, normally load a sheet
    text.include?('!') ? load_cell_range(text) : sheet(text)
  end

  def named_range(name, clean_range: false)
    # yes_no => 'Lists'!$A$1:$A$6
    defined_name = @workbook.defined_names[name]
    defined_name = defined_name.gsub(/\[.+\]/, '').squish if clean_range

    load_cell_range(defined_name) if defined_name.present?
  end

  def table(name)
    @tables.find { |tbl| tbl.name == name}
  end

  def load_cell_range(range_text)
    # get the sheet name => 'Lists'
    sheet_name = range_text.gsub(/[\$\']/, '').scan(/^[^!]*/).first
    # fetch the cell range => '$A$1:$A$6'
    cell_range = range_text.gsub(/\$/, '').scan(/(?<=!).+/).first
    # get the sheet object and fetch the cells in range
    sheet(sheet_name).list_values_from_cell_range(cell_range)
  end

  def fetch_comments(sheet_index)
    final_sheet_index = sheet_index+1
    relationship = @relationships[final_sheet_index.to_s]
    @comments[relationship.comment_id] if relationship.present?
  end

  def parse_spreadsheet_contents(spreadsheet)
    shared_strings = []
    Zip::File.open(spreadsheet) do |spreadsheet_zip|
      spreadsheet_zip.each do |entry|
        case entry.name
        when /xl\/worksheets\/sheet(\d+)?\.xml/
          sheet_id = entry.name.scan(/xl\/worksheets\/sheet(\d+)?\.xml/).flatten.first
          @sheets[sheet_id] = OOXL::Sheet.new(entry.get_input_stream.read, shared_strings, @options)
        when /xl\/styles\.xml/
          @styles = OOXL::Styles.load_from_stream(entry.get_input_stream.read)
        when /xl\/comments(\d+)?\.xml/
          comment_id = entry.name.scan(/xl\/comments(\d+)\.xml/).flatten.first
          @comments[comment_id] = OOXL::Comments.load_from_stream(entry.get_input_stream.read)
        when "xl/sharedStrings.xml"
          Nokogiri.XML(entry.get_input_stream.read).remove_namespaces!.xpath('sst/si').each do |shared_string_node|
            shared_strings << shared_string_node.xpath('r/t|t').map { |value_node| value_node.text}.join('')
          end
        when /xl\/tables\/.*?/i
          @tables << OOXL::Table.new(entry.get_input_stream.read)
        when "xl/workbook.xml"
          @workbook = OOXL::Workbook.load_from_stream(entry.get_input_stream.read)
        when /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels/
          sheet_id = entry.name.scan(/sheet(\d+)/).flatten.first
          @relationships[sheet_id] = Relationships.new(entry.get_input_stream.read)
        else
          # unsupported for now..
        end
      end
    end
  end
end
