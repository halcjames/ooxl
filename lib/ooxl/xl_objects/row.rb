class OOXL
  class Row
    include Enumerable
    attr_accessor :id, :spans, :cells, :height
    DEFAULT_HEIGHT = '12.75'
    def initialize(**attrs)
      attrs.each { |property, value| property == :options ? instance_variable_set("@#{property}", value) : send("#{property}=", value)}
      @options ||= {}
      @height ||= DEFAULT_HEIGHT
    end

    def [](id)
      if id.is_a?(String)
        cell(id)
      else
        cells[id] || BlankCell.new(id)
      end
    end

    def cells
      if @options[:padded_cells]
        unless @cells.blank?
          'A'.upto(@cells.last.column).map do |column_letter|
            cell(column_letter)
          end
        end
      else
        @cells
      end
    end

    def cell(cell_id)
      cell_final_id = cell_id[/[A-Z]+\d+/] ? cell_id : "#{cell_id}#{id}"
      cell_id_map[cell_final_id]
    end

    def cell_id_map
      @cell_id_map ||= cells.each_with_object(Hash.new { |_, k| BlankCell.new(k) }) do |cell, result|
        result[cell.id] = cell
      end
    end

    def each
      cells.each { |cell| yield cell }
    end

    def self.load_from_node(row_node, shared_strings, styles, options)
      new(id: row_node.attributes["r"].try(:value),
          spans: row_node.attributes["spans"].try(:value),
          height: row_node.attributes["ht"].try(:value),
          cells: row_node.xpath('c').map {  |cell_node| OOXL::Cell.load_from_node(cell_node, shared_strings, styles)},
          options: options )
    end
  end
end

# <row r="1" spans="1:38" s="111" customFormat="1" ht="75" x14ac:dyDescent="0.2">
#   <c r="A1" s="108" t="s">
# </row>
