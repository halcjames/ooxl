class OOXL
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
        comment_text_node = comment_node.xpath('./text/r/t') 
        comment_text_node = comment_node.xpath('text/t') if comment_text_node.nil? || comment_text_node.empty?

        value = if comment_text_node.is_a?(Array)
          comment_text_node.map { |comment_text_node| comment_text_node.text }.join('')
        else
          comment_text_node.text
        end

        id = comment_node.attributes["ref"].to_s
        [id, value]
      end.to_h
      new(comments)
    end
  end
end


# <comments>
#    <authors>
#       <author>Author1</author>
#    </authors>
#    <commentList>
#       <comment ref="J1" authorId="0" shapeId="0">
#          <text>
#             <r>
#                <rPr>
#                   <sz val="8" />
#                   <color indexed="81" />
#                   <rFont val="Tahoma" />
#                   <family val="2" />
#                </rPr>
#                <t>Is the product weight consistent? ( Yes or No)</t>
#             </r>
#          </text>
#       </comment>
#       <comment ref="L1" authorId="0" shapeId="0">
#          <text>
#             <r>
#                <rPr>
#                   <sz val="8" />
#                   <color indexed="81" />
#                   <rFont val="Tahoma" />
#                   <family val="2" />
#                </rPr>
#                <t xml:space="preserve">Enter weight of product without any packaging.
# </t>
#             </r>
#          </text>
#       </comment>
#       <comment ref="O1" authorId="0" shapeId="0">
#          <text>
#             <r>
#                <rPr>
#                   <sz val="8" />
#                   <color indexed="81" />
#                   <rFont val="Tahoma" />
#                   <family val="2" />
#                </rPr>
#                <t>Number of days product can remain on store shelf.</t>
#             </r>
#             <r>
#                <rPr>
#                   <sz val="8" />
#                   <color indexed="81" />
#                   <rFont val="Tahoma" />
#                   <family val="2" />
#                </rPr>
#                <t xml:space="preserve">
# </t>
#             </r>
#          </text>
#       </comment>
#       <comment ref="P1" authorId="0" shapeId="0">
#          <text>
#             <r>
#                <rPr>
#                   <sz val="8" />
#                   <color indexed="81" />
#                   <rFont val="Tahoma" />
#                   <family val="2" />
#                </rPr>
#                <t>Total life of product in number of days (including shelf life).</t>
#             </r>
#             <r>
#                <rPr>
#                   <sz val="8" />
#                   <color indexed="81" />
#                   <rFont val="Tahoma" />
#                   <family val="2" />
#                </rPr>
#                <t xml:space="preserve">
# </t>
#             </r>
#          </text>
#       </comment>
#       <comment ref="Q1" authorId="0" shapeId="0">
#          <text>
#             <r>
#                <rPr>
#                   <sz val="8" />
#                   <color indexed="81" />
#                   <rFont val="Tahoma" />
#                   <family val="2" />
#                </rPr>
#                <t>Product packaging Weight only!</t>
#             </r>
#          </text>
#       </comment>
#    </commentList>
# </comments>
