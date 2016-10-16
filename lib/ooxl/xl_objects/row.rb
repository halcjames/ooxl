class OOXL
  class Row
    include Enumerable
    attr_accessor :id, :spans, :cells

    def initialize(**attrs)
      attrs.each { |property, value| property == :options ? instance_variable_set("@#{property}", value) : send("#{property}=", value)}
      @options ||= {}
    end

    def [](id)
      cell = if id.is_a?(String)
        cells.find { |row| row.id == id}
      else
        cells[id]
      end
      (cell.present?) ? cell : BlankCell.new(id)
    end

    def cells
      if @options[:padded_cells]
        unless @cells.blank?
          'A'.upto(@cells.last.column).map do |column_letter|
            cell = @cells.find { |cell| cell.column == column_letter}
            (cell.blank?) ? BlankCell.new("#{column_letter}#{id}") : cell
          end
        end
      else
        @cells
      end
    end

    def cell(cell_id)
      cell_final_id = cell_id[/[A-Z]{1,}\d+/] ? cell_id : "#{cell_id}#{id}"
      cells.find { |cell| cell.id == cell_final_id}
    end

    def each
      cells.each { |cell| yield cell }
    end

    def self.load_from_node(row_node, shared_strings, styles, options)
      new(id: row_node.attributes["r"].try(:value),
          spans: row_node.attributes["spans"].try(:value),
          cells: row_node.xpath('c').map {  |cell_node| OOXL::Cell.load_from_node(cell_node, shared_strings, styles)},
          options: options )
    end
  end
end

# <row r="1" spans="1:38" s="111" customFormat="1" ht="75" x14ac:dyDescent="0.2">
#   <c r="A1" s="108" t="s">
# </row>
