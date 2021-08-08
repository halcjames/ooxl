class OOXL
  class RowCache
    include Enumerable

    attr_accessor :styles

    # built on-demand -- use rows instead
    attr_reader :row_cache

    # built on-demand -- use fetch_row_by_id instead
    attr_reader :row_id_map

    delegate :size, to: :row_nodes

    def initialize(sheet_xml, shared_strings, options = {})
      @shared_strings = shared_strings
      @sheet_xml = sheet_xml
      @options = options
      @row_cache = []
      @row_id_map = {}
    end

    def [](id)
      fetch_row_by_id(id)
    end

    alias_method :row, :[]

    def each(&block)
      if @options[:padded_rows]
        padded_rows(&block)
      else
        rows(&block)
      end
    end

    def rows(&block)
      # track yield count to know if caller broke out of loop
      rows_yielded = 0
      row_cache.each do |r|
        yield r if block_given?
        rows_yielded += 1
      end

      if !all_rows_loaded? && rows_yielded == row_cache.count
        parse_more_rows(&block)
      end

      row_cache
    end

    def row_range(start_index, end_index)
      return enum_for(:row_range, start_index, end_index) unless block_given?

      rows do |row|
        row_id = row.id.to_i
        next if row_id < start_index
        break if row_id > end_index

        yield row
      end
    end

    def max_row_index
      return 0 if row_nodes.empty?

      if all_rows_loaded?
        row_cache.last.id.to_i
      else
        Row.extract_id(row_nodes.last).to_i
      end
    end

    private

    def parse_more_rows
      row_nodes.drop(row_cache.count).each do |row_node|
        row = parse_row(row_node)
        row_cache << row
        row_id_map[row.id] = row
        yield row if block_given?
      end
    end

    def all_rows_loaded?
      row_cache.size == row_nodes.size
    end

    def fetch_row_by_id(row_id)
      row_id = row_id.to_s
      return row_id_map[row_id] if all_rows_loaded? || row_id_map.key?(row_id)

      parse_more_rows do |row|
        return row if row.id == row_id
      end

      nil
    end

    def padded_rows
      real_rows_yielded = 0
      yielded_rows = []
      (1..Float::INFINITY).each do |row_index|
        row = row(row_index)
        if row.blank?
          row = Row.new(id: row_index.to_s, cells: [])
        else
          real_rows_yielded += 1
        end

        yielded_rows << row
        yield row if block_given?

        break if real_rows_yielded == row_cache.count && all_rows_loaded?
      end
      yielded_rows
    end

    def row_nodes
      @row_nodes ||= @sheet_xml.xpath('//sheetData/row')
    end

    def parse_row(row_node)
      Row.load_from_node(row_node, @shared_strings, @styles, @options)
    end
  end
end

