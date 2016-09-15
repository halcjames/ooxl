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
