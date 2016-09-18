class OOXL
  class BlankCell
    attr_reader :id

    def initialize(id)
      @id = id
    end
    def value
      nil
    end
  end

  # where,
  # r = reference
  # s = style
  # t = type
  # v = value
  # <c r="A1" s="227" t="s">
  #  <v>113944</v>
  # </c>
  class Cell
    attr_accessor :id, :type_id, :style_id, :value_id, :shared_strings, :styles

    def initialize(**attrs)
      attrs.each { |property, value| send("#{property}=", value)}
    end

    def next_id(offset: 1, location: "bottom")
      _, column_letter, column_index = id.partition(/[A-Z]+/)

      # ensure that all are numbers
      column_index = column_index.to_i
      offset = offset.to_i if offset.is_a?(String)

      # increment based on specified location
      case location
      when "top"
        if column_index == 1 || column_index < offset
          column_index = 1
        else
          column_index -= offset
        end
      when "bottom"
        column_index += offset
      when "left"
        return id if column_letter == 'A'
        1.upto(offset) { |count| column_letter = (column_letter.ord-1).chr unless column_letter == 'A' }
      when "right"
        1.upto(offset) { |count| column_letter.next! }
      else
        id
      end

      "#{column_letter}#{column_index}"
    end

    def type
      @type ||= begin
        case type_id
        when 's' then :string
        when 'n' then :number
        when 'b' then :boolean
        when 'd' then :date
        when 'str' then :formula
        when 'inlineStr' then :inline_str
        else
          :error
        end
      end
    end

    def style
      @style ||= begin
        if style_id.present?
          style = styles.by_id(style_id.to_i)
        end
      end
    end

    def number_format
      if (style.present?)
        nf = style[:number_format]
        (nf.present?) ? nf.gsub("\\", "") : nil
      end
    end

    def font
      (style.present?) ? style[:font] : nil
    end

    def fill
      (style.present?) ? style[:fill]: nil
    end

    def value
      case type
      when :string
        (value_id.present?) ? shared_strings[value_id.to_i] : nil
      else
        # TODO: to support other types soon
        value_id
      end
    end

    def self.load_from_node(cell_node, shared_strings, styles)
      new(id: cell_node.attributes["r"].try(:value),
          type_id: cell_node.attributes["t"].try(:value),
          style_id: cell_node.attributes["s"].try(:value),
          value_id: cell_node.at('v').try(:text),
          shared_strings: shared_strings,
          styles: styles )
    end
  end
end
