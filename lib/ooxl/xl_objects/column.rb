class OOXL
  class Column
    attr_accessor :id, :width, :custom_width, :id_range, :hidden, :best_fit
    alias_method :hidden?, :hidden

    def initialize(**attrs)
      attrs.each { |property, value| send("#{property}=", value)}
    end

    def self.load_from_node(column_node)
      hidden_attr = column_node.attributes["hidden"]
      new(id: column_node.attributes["min"].try(:value),
          width: column_node.attributes["width"].try(:value),
          best_fit: column_node.attributes["best_fit"].try(:value),
          custom_width: column_node.attributes["custom_width"].try(:value),
          id_range: (column_node.attributes["min"].value.to_i..column_node.attributes["max"].value.to_i).to_a,
          hidden: (hidden_attr.present?) ? hidden_attr.value == "1" : false)
    end
  end
end

#    <col min="3" max="3" width="6.42578125" bestFit="1" customWidth="1" />
#    <col min="4" max="5" width="12.7109375" customWidth="1" />
#    <col min="6" max="8" width="6.42578125" customWidth="1" />
