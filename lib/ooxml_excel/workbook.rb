module OOXML
  class Excel
    class Workbook
      def initialize(xml)
        @xml = xml
      end

      def sheets
        @sheets ||= begin
          @xml.xpath('//sheets/sheet').map do |sheet_node|
            name = sheet_node.attribute('name').value
            rel_id = sheet_node.attribute('id').value.gsub(/[^\d+]/, '')
            sheet_id = sheet_node.attribute('sheetId').value
            { name: name, sheet_id: sheet_id, relationship_id: rel_id}
          end
        end
      end

      def defined_names
        @defined_names ||= begin
          @xml.xpath('//definedNames/definedName').map do |defined_names_node|
            name = defined_names_node.attribute('name').value
            reference = defined_names_node.text
            [name, reference]
          end.to_h
        end
      end


      def self.load_from_stream(xml_stream)
        self.new (Nokogiri.XML(xml_stream).remove_namespaces!)
      end
    end
  end
end
