class OOXL
  class Styles
    class CellStyleReference
      attr_accessor :id, :number_formatting_id, :fill_id, :font_id
      def initialize(**attrs)
        attrs.each { |property, value| send("#{property}=", value)}
      end
      def self.load_from_node(cell_style_xfs_node)
        attributes = cell_style_xfs_node.attributes
        self.new(
          id: attributes["xfId"].value.to_i,
          number_formatting_id: attributes["numFmtId"].value.to_i,
          fill_id: attributes["fillId"].value.to_i,
          font_id: attributes["fontId"].value.to_i
        )
      end
    end
  end
end

# <cellStyleXfs count="4">
#    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
#    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" />
#    <xf numFmtId="0" fontId="39" fillId="0" borderId="0" applyProtection="0" />
#    <xf numFmtId="9" fontId="1" fillId="0" borderId="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0" />
# </cellStyleXfs>
