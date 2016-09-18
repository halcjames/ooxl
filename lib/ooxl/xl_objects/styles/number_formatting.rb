class OOXL
  class Styles
    class NumberFormatting
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

# <numFmts count="16">
#    <numFmt numFmtId="172" formatCode="000\ 00000\ 00000\ 0" />
#    <numFmt numFmtId="173" formatCode="00000.0000" />
#    <numFmt numFmtId="174" formatCode="00.0" />
#    <numFmt numFmtId="175" formatCode="0000.0" />
#    <numFmt numFmtId="176" formatCode="000000000" />
#    <numFmt numFmtId="177" formatCode="0000" />
#    <numFmt numFmtId="178" formatCode="0000000" />
#    <numFmt numFmtId="179" formatCode="mm/dd/yy" />
#    <numFmt numFmtId="180" formatCode="000000" />
#    <numFmt numFmtId="181" formatCode="00" />
#    <numFmt numFmtId="182" formatCode="0.000" />
#    <numFmt numFmtId="183" formatCode="0.0" />
#    <numFmt numFmtId="184" formatCode="000\ 00000\ 00000" />
#    <numFmt numFmtId="185" formatCode="000000\ 000" />
#    <numFmt numFmtId="186" formatCode="00000" />
#    <numFmt numFmtId="187" formatCode="[$-F800]dddd\,\ mmmm\ dd\,\ yyyy" />
# </numFmts>
