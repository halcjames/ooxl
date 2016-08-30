module OOXML
  class Excel
    class Styles
      attr_accessor :fonts, :fills, :number_formats, :cell_style_xfs
      def initialize(**attrs)
        attrs.each { |property, value| send("#{property}=", value)}
      end

      def by_id(id)
        cell_style = cell_style_xfs.fetch(id)
        {
          font: fonts_by_index(cell_style.font_id),
          fill: fills_by_index(cell_style.fill_id),
          number_format: number_formats_by_index(cell_style.number_formatting_id),
        }
      end

      def fonts_by_index(font_index)
        @fonts[font_index]
      end

      def fills_by_index(fill_index)
        @fills[fill_index]
      end

      def number_formats_by_index(number_format_index)
        @number_formats.find { |number_format| number_format.id == number_format_index.to_s}.try(:code)
      end

      def self.load_from_stream(xml_stream)
        style_doc = Nokogiri.XML(xml_stream).remove_namespaces!
        fonts = style_doc.xpath('//fonts/font')
        fills = style_doc.xpath('//fills/fill')
        number_formats = style_doc.xpath('//numFmts/numFmt')
        # This element contains the master formatting records (xf) which
        # define the formatting applied to cells in this workbook.
        # link: https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.cellformats(v=office.14).aspx
        cell_style_xfs =  style_doc.xpath('//cellXfs/xf')

        self.new(
          fonts: fonts.map { |font_node| Excel::Styles::Font.load_from_node(font_node)},
          fills: fills.map { |fill_node| Excel::Styles::Fill.load_from_node(fill_node)},
          number_formats: number_formats.map { |num_fmt_node| Excel::Styles::NumFmt.load_from_node(num_fmt_node) },
          cell_style_xfs: cell_style_xfs.map { |cell_style_xfs_node| Excel::Styles::CellStyleXfs.load_from_node(cell_style_xfs_node)}
        )
      end
    end
  end
end

module OOXML
  class Excel
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
            size: font_size_node.attributes["val"].value,
            name: font_name_node.attributes["val"].value,
            rgb_color: (font_color_node.present?) ? font_color_node.attributes["rgb"].try(:value) : nil,
            theme: (font_color_node.present?) ? font_color_node.attributes["theme"].try(:value) : nil,
            bold: font_bold_node.present?
          )
        end
      end
    end
  end
end
module OOXML
  class Excel
    class Styles
      class NumFmt
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
end
# <xf numFmtId="0" borderId="0" fillId="0" fontId="0" applyAlignment="1" applyFont="1" xfId="0"/>
module OOXML
  class Excel
    class Styles
      class CellStyleXfs
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
end


module OOXML
  class Excel
    class Styles
      class Fill
        attr_accessor :pattern_type, :fg_color, :fg_color_theme, :fg_color_tint, :bg_color_index, :bg_color

        def initialize(**attrs)
          attrs.each { |property, value| send("#{property}=", value)}
        end
        def self.load_from_node(fill_node)
          pattern_fill = fill_node.at('patternFill')

          pattern_type = pattern_fill.attributes["patternType"].value
          if pattern_type == "solid"
            fg_color = pattern_fill.at('fgColor')
            bg_color = pattern_fill.at('bgColor')
            self.new(pattern_type: pattern_type,
                    fg_color: (fg_color.present?) ? fg_color.attributes["rgb"].try(:value) : nil,
                    fg_color_theme: (fg_color.present?) ? fg_color.attributes["theme"].try(:value) : nil,
                    fg_color_tint:  (fg_color.present?) ? fg_color.attributes["tint"].try(:value) : nil,
                    bg_color: (bg_color.present?) ?  bg_color.attributes["rgb"].try(:value) : nil,
                    bg_color_index: (bg_color.present?) ?  bg_color.attributes["index"].try(:value) : nil)
          else
            self.new(pattern_type: pattern_type)
          end
        end
      end
    end
  end
end
