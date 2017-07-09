class OOXL
  class Workbook
    def initialize(xml, relationship)
      @xml = xml
      @relationship = relationship
    end

    def sheets
      @sheets ||= begin
        # <sheet r:id="rId13" sheetId="1" name="Ceiling Fans" state="hidden"/>
        @xml.xpath('//sheets/sheet').map.with_index do |sheet_node, sheet_order_index|
          name = sheet_node.attribute('name').value
          rel_id = sheet_node.attribute('id').value.gsub(/[^\d+]/, '')
          sheet_id = sheet_node.attribute('sheetId').value
          state = sheet_node.attribute('state').try(:value)
          
          { 
            name: name, 
            sheet_id: sheet_id, 
            sheet_file_index: sheet_file_index(rel_id), 
            sheet_order_index: sheet_order_index, 
            relationship_id: rel_id,
            state: (state.blank?) ? 'visible' : state
          }
        end
      end
    end

    def sheet_file_index(relationship_id)
      @relationship.sheet_file_index(relationship_id)
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

    def self.load_from_stream(xml_stream, relationship)
      xml_node = Nokogiri.XML(xml_stream).remove_namespaces!
      new(xml_node, relationship)
    end
  end
end
