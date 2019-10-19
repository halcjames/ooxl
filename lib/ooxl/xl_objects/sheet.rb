require_relative 'sheet/data_validation'
require_relative 'row_cache'

class OOXL
  class Sheet
    include OOXL::Util
    include Enumerable

    attr_reader :columns, :data_validations, :shared_strings, :styles
    attr_accessor :comments, :defined_names, :name
    delegate :[], :each, :rows, :row, to: :@row_cache

    def initialize(xml, shared_strings, options={})
      @xml = Nokogiri.XML(xml).remove_namespaces!
      @shared_strings = shared_strings
      @comments = {}
      @defined_names = {}
      @styles = []
      @options = options
      @row_cache = RowCache.new(@xml, @shared_strings, options)
    end

    def code_name
      @code_name ||= @xml.xpath('//sheetPr').attribute('codeName').try(:value)
    end

    def comment(cell_ref)
      @comments[cell_ref] unless @comments.blank?
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


    # DEPRECATED: stream is no longer separate
    def stream_row(index)
      row(index)
    end

    # test mode
    def cells_by_column(column_letter)
      rows.map do |row|
        row.cells.find { |cell| to_column_letter(cell.id) == column_letter}
      end
    end

    def last_column(row_index=1)
      @last_column ||= {}
      @last_column[row_index] ||= begin
        cells = row(row_index).try(:cells) 
        cells.last.column if cells.present?
      end
    end

    def cell(cell_id)
      column_letter, row_index = cell_id.partition(/\d+/)
      current_row = row(row_index)
      current_row.cell(column_letter) unless current_row.nil?
    end

    def formula(cell_id)
      cell(cell_id).try(:formula)
    end

    def font(cell_reference)
      cell(cell_reference).try(:font)
    end

    def fill(cell_reference)
      cell(cell_reference).try(:fill)
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

    def styles=(styles)
      @styles = styles
      @row_cache.styles = styles
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
          @defined_names[data_validation.formula]
        end
      end
    end
    alias_method :list_value_formula, :cell_range

    def list_values_from_cell_range(cell_range)
      return [] if cell_range.blank?

      # cell_range values separated by comma
      if cell_range.include?(":")
        cell_letters = cell_range.gsub(/[\d]/, '').split(':')
        start_index, end_index = cell_range[/[A-Z]{1,}\d+/] ? cell_range.gsub(/[^\d:]/, '').split(':').map(&:to_i) : [1, @row_cache.max_row_index]
        if cell_letters.uniq.size > 1
          list_values_from_rectangle(cell_letters, start_index, end_index)
        else
          list_values_from_column(cell_letters.uniq.first, start_index, end_index)
        end
      else
        # when only one value: B2
        list_values_from_cell(cell_range)
      end
    end
    alias_method :list_values_from_formula, :list_values_from_cell_range

    # This will allow values from this pattern
    # 'SheetName!A1:C3'
    # The number after the cell letter will be the index
    # 1 => start_index
    # 3 => end_index
    # Expected output would be: [['value', 'value', 'value'], ['value', 'value', 'value'], ['value', 'value', 'value']]
    def list_values_from_rectangle(cell_letters, start_index, end_index)
      start_col = column_letter_to_number(cell_letters.first)
      end_col = column_letter_to_number(cell_letters.last)
      @row_cache.row_range(start_index, end_index).map do |row|
        (start_col..end_col).map do |col_index|
          col_letter = column_number_to_letter(col_index)
          row["#{col_letter}#{row.id}"].value
        end
      end
    end

    def list_values_from_column(column_letter, start_index, end_index)
      @row_cache.row_range(start_index, end_index).map do |row|
        row["#{column_letter}#{row.id}"].value
      end
    end

    def list_values_from_cell(cell_ref)
      row_index = cell_ref.gsub(/[^\d:]/, '').split(':').map(&:to_i).first
      row = row(row_index)
      return if row.blank?
      [row[cell_ref].value]
    end

    def in_merged_cells?(cell_id)
      column_letter, column_index = cell_id.partition(/\d+/)
      range = merged_cells.find { |column_range, index_range| column_range.cover?(column_letter) && index_range.cover?(column_index) }
      range.present?
    end

    def self.load_from_stream(xml_stream, shared_strings)
      self.new(xml_stream, shared_strings)
    end

    private

    def merged_cells
      @merged_cells ||= @xml.xpath('//mergeCells/mergeCell').map do |merged_cell|
        # <mergeCell ref="Q381:R381"/>
        start_reference, end_reference = merged_cell.attributes["ref"].try(:value).split(':')

        start_column_letter, start_index =  start_reference.partition(/\d+/)
        end_column_letter, end_index = end_reference.partition(/\d+/)
        [(start_column_letter..end_column_letter), (start_index..end_index)]
      end.to_h
    end
  end
end
