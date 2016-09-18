class OOXL
  class Relationships
    SUPPORTED_TYPES = ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']
    def initialize(relationships_node)
      @types = {}
      parse_relationships(relationships_node)
    end

    def comment_id
      @types['comments']
    end

    private
    def parse_relationships(relationships_node)
      relationships_node = Nokogiri.XML(relationships_node).remove_namespaces!
      relationships_node.xpath('//Relationship').each do |relationship_node|
        relationship_type = relationship_node.attributes["Type"].value
        target = relationship_node.attributes["Target"].value
        if supported_type?(relationship_type)
          @types[extract_type(relationship_type)] = extract_file_reference(target)
        end
      end
    end

    def supported_type?(type)
      SUPPORTED_TYPES.include?(type)
    end

    def extract_type(type)
      type.split('/').last
    end

    def extract_file_reference(file)
      file.scan(/(\d+)\.[\w]/).flatten.first
    end

  end
end

# Only supporting comments for now

# <Relationships>
#    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml" />
#    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml" />
#    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" Target="../customProperty1.bin" />
#    <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml" />
#    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp2.xml" />
#    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml" />
# </Relationships>
