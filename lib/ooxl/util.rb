class OOXL
  module Util
    COLUMN_LETTERS = ('A'..'ZZZZ').to_a
    def letter_equivalent(index)
      COLUMN_LETTERS.fetch(index)
    end

    def letter_index(letter)
      COLUMN_LETTERS.index { |c_letter| c_letter == letter}
    end

    def to_column_letter(reference)
      reference.gsub(/\d+/, '')
    end

    def uniform_reference(ref)
      ref.to_s[/[A-Z]/] ? letter_index(ref) + 1 : ref
    end

    def node_value_extractor(node)
      node.try(:value)
    end

    def node_attribute_value(node, attribute_name)
      unless node.blank?
        attribute = node.attributes.find { |key, attribute| key == attribute_name}
        attribute[1].value if attribute.present?
      end
    end
  end
end
