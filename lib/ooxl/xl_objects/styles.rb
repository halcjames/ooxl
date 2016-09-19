require_relative 'cell_style_reference'
require_relative 'fill'
require_relative 'font'
require_relative 'number_formatting'
class OOXL
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
        fonts: fonts.map { |font_node| Font.load_from_node(font_node)},
        fills: fills.map { |fill_node| Fill.load_from_node(fill_node) if fill_node.to_s.include?('patternFill')},
        number_formats: number_formats.map { |num_fmt_node| NumberFormatting.load_from_node(num_fmt_node) },
        cell_style_xfs: cell_style_xfs.map { |cell_style_xfs_node| CellStyleReference.load_from_node(cell_style_xfs_node)}
      )
    end
  end
end
