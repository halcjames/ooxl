class OOXL
  class Styles
    class Fill
      attr_accessor :pattern_type, :fg_color, :fg_color_theme, :fg_color_tint, :bg_color_index, :bg_color, :fg_color_index

      def initialize(**attrs)
        attrs.each { |property, value| send("#{property}=", value)}
      end
      def self.load_from_node(fill_node)
        pattern_fill = fill_node.at('patternFill')

        pattern_type = pattern_fill.attributes["patternType"].value
        if pattern_type == "solid"
          fg_color = pattern_fill.at('fgColor')
          bg_color = pattern_fill.at('bgColor')
          fg_color_index = fg_color.class == Nokogiri::XML::Element ? fg_color.attributes["indexed"].try(:value) : nil

          self.new(pattern_type: pattern_type,
                  fg_color: (fg_color.present?) ? fg_color.attributes["rgb"].try(:value) : nil,
                  fg_color_theme: (fg_color.present?) ? fg_color.attributes["theme"].try(:value) : nil,
                  fg_color_tint:  (fg_color.present?) ? fg_color.attributes["tint"].try(:value) : nil,
                  bg_color: (bg_color.present?) ?  bg_color.attributes["rgb"].try(:value) : nil,
                  bg_color_index: (bg_color.present?) ?  bg_color.attributes["index"].try(:value) : nil,
                  fg_color_index: fg_color_index)
        else
          self.new(pattern_type: pattern_type)
        end
      end
    end
  end
end
