class OOXL
  module Util
    COLUMN_LETTERS = [nil] + ('A'..'ZZZZ').to_a

    def letter_index(col_letter)
      column_letter_to_number(col_letter) - 1
    end

    def letter_equivalent(col_index)
      column_number_to_letter(col_index + 1)
    end

    def to_column_letter(reference)
      reference.gsub(/\d+/, '')
    end

    def uniform_reference(ref)
      ref.to_s[/[A-Z]/] ? column_letter_to_number(ref) : ref
    end

    def node_value_extractor(node)
      node.try(:value)
    end

    def column_number_to_letter(index)
      COLUMN_LETTERS.fetch(index)
    end

    def column_letter_to_number(column_letter)
      pow = column_letter.length - 1
      result = 0
      column_letter.each_byte do |b|
        result += 26**pow * (b - 64)
        pow -= 1
      end
      result
    end

    def node_attribute_value(node, attribute_name)
      unless node.blank?
        attribute = node.attributes.find { |key, attribute| key == attribute_name}
        attribute[1].value if attribute.present?
      end
    end
  end
end
