class OOXL
  module Util
    COLUMN_LETTERS = ('A'..'ZZZZ').to_a
    def letter_equivalent(index)
      COLUMN_LETTERS.fetch(index)
    end

    def letter_index(letter)
      COLUMN_LETTERS.index { |c_letter| c_letter == letter}
    end

    def uniform_reference(ref)
      ref.to_s[/[A-Z]/] ? letter_index(ref) + 1 : ref
    end
  end
end
