module OOXML
  class Excel
    class Comments
      attr_reader :comments

      def initialize(comments)
        @comments = comments
      end

      def [](id)
        @comments[id]
      end

      def self.load_from_stream(comment_xml)
        comment_xml =Nokogiri.XML(comment_xml).remove_namespaces!

        comments = comment_xml.xpath("//comments/commentList/comment").map do |comment_node|
          value = (comment_node.xpath('./text/r/t').last || comment_node.at_xpath('./text/r/t') || comment_node.at_xpath('./text/t')).text
          id = comment_node.attributes["ref"].to_s
          [id, value]
        end.to_h
        new(comments)
      end
    end
  end
end
