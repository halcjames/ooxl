class OOXL
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

# <fills count="9">
#    <fill>
#       <patternFill patternType="none" />
#    </fill>
#    <fill>
#       <patternFill patternType="gray125" />
#    </fill>
#    <fill>
#       <patternFill patternType="solid">
#          <fgColor theme="0" tint="-0.249977111117893" />
#          <bgColor indexed="64" />
#       </patternFill>
#    </fill>
#    <fill>
#       <patternFill patternType="solid">
#          <fgColor theme="0" tint="-0.249977111117893" />
#          <bgColor indexed="31" />
#       </patternFill>
#    </fill>
#    <fill>
#       <patternFill patternType="solid">
#          <fgColor theme="0" tint="-4.9989318521683403E-2" />
#          <bgColor indexed="64" />
#       </patternFill>
#    </fill>
#    <fill>
#       <patternFill patternType="solid">
#          <fgColor rgb="FFFFFF00" />
#          <bgColor indexed="64" />
#       </patternFill>
#    </fill>
#    <fill>
#       <patternFill patternType="solid">
#          <fgColor theme="0" tint="-0.24994659260841701" />
#          <bgColor indexed="64" />
#       </patternFill>
#    </fill>
#    <fill>
#       <patternFill patternType="solid">
#          <fgColor rgb="FFFFCC99" />
#          <bgColor indexed="64" />
#       </patternFill>
#    </fill>
# </fills>
