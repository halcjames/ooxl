class OOXL
  class Row
    include Enumerable
    attr_accessor :id, :spans, :cells

    def initialize(**attrs)
      attrs.each { |property, value| send("#{property}=", value)}
    end

    def [](id)
      cell = if id.is_a?(String)
        cells.find { |row| row.id == id}
      else
        cells[id]
      end
      (cell.present?) ? cell : BlankCell.new(id)
    end

    def each
      cells.each { |cell| yield cell }
    end

    def self.load_from_node(row_node, shared_strings, styles)
      new(id: row_node.attributes["r"].try(:value),
          spans: row_node.attributes["spans"].try(:value),
          cells: row_node.xpath('c').map {  |cell_node| OOXL::Cell.load_from_node(cell_node, shared_strings, styles) } )
    end
  end
end
