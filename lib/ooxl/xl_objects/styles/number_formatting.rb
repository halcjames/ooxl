class OOXL
  class Styles
    class NumberFormatting
      attr_accessor :id, :code
      def self.load_from_node(num_fmt_node)
        new_format = self.new.tap do |number_format|
          number_format.id = num_fmt_node.attributes["numFmtId"].try(:value)
          number_format.code = num_fmt_node.attributes["formatCode"].try(:value)
        end
      end
    end
  end
end
