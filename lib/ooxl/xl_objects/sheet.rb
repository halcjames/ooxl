require_relative 'sheet/data_validation'
class OOXL
  class Sheet
    include OOXL::Util
    include Enumerable
    attr_reader :columns, :data_validations, :shared_strings
    attr_accessor :comments, :styles, :defined_names, :name

    def initialize(xml, shared_strings)
      @xml = Nokogiri.XML(xml).remove_namespaces!
      @shared_strings = shared_strings
      @comments = {}
      @defined_names = {}
      @styles = []
    end

    def code_name
      @code_name ||= @xml.xpath('//sheetPr').attribute('codeName').try(:value)
    end

    def comment(cell_ref)
      @comments[cell_ref]
    end

    def data_validation(cell_ref)
      data_validations.find { |data_validation| data_validation.in_sqref_range?(cell_ref)}
    end

    def column(id)
      uniformed_reference = uniform_reference(id)
      columns.find { |column| column.id_range.include?(uniformed_reference)}
    end

    def columns
      @columns ||= begin
        @xml.xpath('//cols/col').map do |column_node|
          Column.load_from_node(column_node)
        end
      end
    end

    def [](id)
      if id.is_a?(String)
        rows.find { |row| row.id == id}
      else
        rows[id]
      end
    end

    def row(index)
      rows.find { |row| row.id == index.to_s}
    end

    def rows
      @rows ||= begin
        # TODO: get the value of merged cells
        # merged_cells = @xml.xpath('//mergeCells/mergeCell').map { |merged_cell| merged_cell.attributes["ref"].try(:value) }
        @xml.xpath('//sheetData/row').map do |row_node|
          Row.load_from_node(row_node, @shared_strings, styles)
        end
      end
    end

    def each
      rows.each { |row| yield row }
    end

    def font(cell_reference)
      style_id = fetch_style_style_id(cell_reference)
      if style_id.present?
        style = @styles.by_id(style_id.to_i)

        (style.present?) ? style[:font] : nil
      end
    end

    def fill(cell_reference)
      style_id = fetch_style_style_id(cell_reference)
      if style_id.present?
        style = @styles.by_id(style_id.to_i)
        (style.present?) ? style[:fill] : nil
      end
    end


    def data_validations
      @data_validations ||= begin

        # original validations
        dvalidations = @xml.xpath('//dataValidations/dataValidation').map do |data_validation_node|
          Sheet::DataValidation.load_from_node(data_validation_node)
        end

        # extended validations
        dvalidations_ext = @xml.xpath('//extLst//ext//dataValidations/dataValidation').map do |data_validation_node_ext|
          Sheet::DataValidation.load_from_node(data_validation_node_ext)
        end

        # merge validations
        [dvalidations, dvalidations_ext].flatten.compact
      end
    end

    # a shortcut for:
    # formula =  data_validation('A1').formula
    # ooxl.named_range(formula)
    def cell_range(cell_ref)
      data_validation = data_validations.find { |data_validation| data_validation.in_sqref_range?(cell_ref)}
      if data_validation.respond_to?(:type) && data_validation.type == "list"
        if data_validation.formula[/[\s\$\,\:]/]
          (data_validation.formula[/\$/].present?) ? "#{name}!#{data_validation.formula}" : data_validation.formula
        else
          @defined_names.fetch(data_validation.formula)
        end
      end
    end
    alias_method :list_value_formula, :cell_range

    def list_values_from_cell_range(cell_range)
      return [] if cell_range.blank?

      # cell_range values separated by comma
      if cell_range.include?(":")
        cell_letters = cell_range.gsub(/[\d]/, '').split(':')
        start_index, end_index = cell_range.gsub(/[^\d:]/, '').split(':').map(&:to_i)
        # This will allow values from this pattern
        # 'SheetName!A1:C3'
        # The number after the cell letter will be the index
        # 1 => start_index
        # 3 => end_index
        # Expected output would be: [['value', 'value', 'value'], ['value', 'value', 'value'], ['value', 'value', 'value']]
        if cell_letters.uniq.size > 1
          start_index.upto(end_index).map do  |row_index|
            (cell_letters.first..cell_letters.last).map do |cell_letter|
                row = fetch_row_by_id(row_index.to_s)
                next if row.blank?
                row["#{cell_letter}#{row_index}"].value
            end
          end
        else
          cell_letter = cell_letters.uniq.first
          (start_index..end_index).to_a.map do |row_index|
            row = fetch_row_by_id(row_index.to_s)
            next if row.blank?
            row["#{cell_letter}#{row_index}"].value
          end
        end
      else
        # when only one value: B2
        row_index = cell_range.gsub(/[^\d:]/, '').split(':').map(&:to_i).first
        row = fetch_row_by_id(row_index.to_s)
        return if row.blank?
        [row[cell_range].value]
      end
    end
    alias_method :list_values_from_formula, :list_values_from_cell_range

    def self.load_from_stream(xml_stream, shared_strings)
      self.new(Nokogiri.XML(xml_stream).remove_namespaces!, shared_strings)
    end

    private
    def fetch_row_by_id(row_id)
      rows.find { |row| row.id == row_id.to_s}
    end
    def fetch_style_style_id(cell_reference)
      raise 'Invalid Cell Reference!' if cell_reference[/[A-Z]{1,}\d+/].blank?
      row_index = cell_reference.scan(/[A-Z{1,}](\d+)/).flatten.first.to_i - 1
      return if rows[row_index].blank? || rows[row_index][cell_reference].blank?
      rows[row_index][cell_reference].style_id
    end

  end
end
