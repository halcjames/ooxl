class OOXL
  include Enumerable
  include ListHelper
  attr_reader :filename
  
  def initialize(filepath = nil, contents: nil, **options)
    @workbook = nil
    @sheets = {}
    @styles = []
    @comments = {}
    @workbook_relationships = nil
    @sheet_relationships = {}
    @options = options
    @tables = []

    @filename = filepath && File.basename(filepath)
    if contents.present?
      parse_spreadsheet_contents(contents)
    elsif filepath.present?
      parse_spreadsheet_file(filepath)
    else
      raise 'no file path or contents were provided'
    end
  end

  def self.open(spreadsheet_filepath, options={})
    new(spreadsheet_filepath, **options)
  end

  def self.parse(spreadsheet_contents, options={})
    spreadsheet_contents.force_encoding('ASCII-8BIT') if spreadsheet_contents.respond_to?(:force_encoding)
    new(nil, contents: spreadsheet_contents, **options)
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
    sheet_meta = @workbook.sheets.find { |sheet| sheet[:name] == sheet_name }
    raise "No #{sheet_name} in workbook." if sheet_meta.nil?

    sheet_index = @workbook_relationships[sheet_meta[:relationship_id]].scan(/\d+/).first
    sheet = @sheets.fetch(sheet_index)

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
    relationship = @sheet_relationships[sheet_index]
    @comments[relationship.comment_id] if relationship.present?
  end

  def parse_spreadsheet_file(file_path)
    Zip::File.open(file_path) { |zip| parse_zip(zip) }
  end

  def parse_spreadsheet_contents(file_contents)
    # open_buffer works for strings and IO streams
    Zip::File.open_buffer(file_contents) { |zip| parse_zip(zip) }
  end

  def parse_zip(spreadsheet_zip)
    shared_strings = []
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
        @sheet_relationships[sheet_id] = Relationships.new(entry.get_input_stream.read)
      when /xl\/_rels\/workbook\.xml\.rels/
        @workbook_relationships = Relationships.new(entry.get_input_stream.read)
      else
        # unsupported for now..
      end
    end
  end
end
