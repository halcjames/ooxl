class OOXL
  class Table
    def initialize(stream)
      @table_node = Nokogiri.XML(stream).remove_namespaces!
    end

    def ref
      @table_node.at('/table').attributes['ref'].try(:value)
    end

    def name
      @table_node.at('/table').attributes['name'].try(:value)
    end
  end
end
