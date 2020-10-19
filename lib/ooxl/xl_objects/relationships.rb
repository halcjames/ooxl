class OOXL
  class Relationships
    SUPPORTED_TYPES = ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']

    def initialize(relationships_node)
      @relationships = []
      parse_relationships(relationships_node)
    end

    def comment_id
      comment_target = by_type('comments').first
      comment_target && extract_file_reference(comment_target)
    end

    def [](id)
      @relationships.find { |rel| rel.id == id }&.target
    end

    def by_type(type)
      @relationships.select { |rel| rel.type == type }.map(&:target)
    end

    private

    def parse_relationships(relationships_node)
      relationships_node = Nokogiri.XML(relationships_node).remove_namespaces!
      relationships_node.xpath('//Relationship').each do |relationship_node|
        relationship_type = relationship_node.attributes["Type"].value
        target = relationship_node.attributes["Target"].value
        id = extract_number(relationship_node.attributes["Id"].value)
        type = extract_type(relationship_type)
        target = relationship_node.attributes["Target"].value
        @relationships << Relationship.new(id, type, target)
      end
    end

    def extract_number(str)
      str.scan(/(\d+)/).flatten.first
    end

    def extract_type(type)
      type.split('/').last
    end

    def extract_file_reference(file)
      file.scan(/(\d+)\.[\w]/).flatten.first
    end

    Relationship = Struct.new(:id, :type, :target)
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
