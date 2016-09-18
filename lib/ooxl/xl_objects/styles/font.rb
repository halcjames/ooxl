class OOXL
  class Styles
    class Font
      attr_accessor :size, :name, :rgb_color, :theme, :bold
      alias_method :bold?, :bold
      def initialize(**attrs)
        attrs.each { |property, value| send("#{property}=", value)}
      end
      def self.load_from_node(font_node)
        font_size_node = font_node.at('sz')
        font_color_node = font_node.at('color')
        font_name_node = font_node.at('name')
        font_bold_node = font_node.at('b')
        self.new(
          size: font_size_node && font_size_node.attributes["val"].value,
          name: font_name_node && font_name_node.attributes["val"].value,
          rgb_color: font_color_node && font_color_node.attributes["rgb"].try(:value) ,
          theme: font_color_node && font_color_node.attributes["theme"].try(:value),
          bold: font_bold_node.present?
        )
      end
    end
  end
end

# <fonts count="3">
#    <font>
#       <sz val="10" />
#       <name val="Arial" />
#    </font>
#    <font>
#       <sz val="10" />
#       <name val="Arial" />
#       <family val="2" />
#    </font>
#    <font>
#       <b />
#       <sz val="9" />
#       <name val="Arial" />
#       <family val="2" />
#    </font>
# </fonts>
