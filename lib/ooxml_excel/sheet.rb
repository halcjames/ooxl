module OOXML
  class Excel
    class Sheet
      include OOXML::Helper::List
      attr_reader :columns, :data_validations, :shared_strings
      attr_accessor :comments, :styles, :defined_names, :name

      def initialize(xml, shared_strings)
        @xml = xml
        @shared_strings = shared_strings
        @comments = {}
        @defined_names = {}
        @styles = []
      end

      def code_name
        @code_name ||= @xml.xpath('//sheetPr').attribute('codeName').try(:value)
      end

      def data_validation_for_cell(cell_ref)
        data_validations.find { |data_validation| data_validation.sqref_range.include?(cell_ref)}
      end


      def column(id)
        columns.select { |column| column.id_range.include?(id)}
      end

      def columns
        @columns ||= begin
          @xml.xpath('//cols/col').map do |column_node|
            Excel::Sheet::Column.load_from_node(column_node)
          end
        end
      end

      def [](id)
        if id.is_a?(String)
          rows.find { |row| row.id == id}
        else
          rows[id]
        end
      end

      def rows
        @rows ||= begin
          # TODO: get the value of merged cells
          # merged_cells = @xml.xpath('//mergeCells/mergeCell').map { |merged_cell| merged_cell.attributes["ref"].try(:value) }
          @xml.xpath('//sheetData/row').map do |row_node|
            Excel::Sheet::Row.load_from_node(row_node, shared_strings)
          end
        end
      end

      def font(cell_reference)
        style_id = fetch_style_style_id(cell_reference)
        if style_id.present?
          style = @styles.by_id(style_id.to_i)

          (style.present?) ? style[:font] : nil
        end
      end

      def fill(cell_reference)
        style_id = fetch_style_style_id(cell_reference)
        if style_id.present?
          style = @styles.by_id(style_id.to_i)
          (style.present?) ? style[:fill] : nil
        end
      end

      def each_row
        rows.each_with_index do |row, row_index|
          yield row.cells.map(&:value), row_index
        end
      end

      def each_row_as_object
        0.upto(rows.size).each do |row_index|
          yield rows[row_index]
        end
      end

      def data_validations
        @data_validations ||= begin
          @xml.xpath('//dataValidations/dataValidation').map do |data_validation_node|
            Excel::Sheet::DataValidation.load_from_node(data_validation_node)
          end
        end
      end

      def self.load_from_stream(xml_stream, shared_strings)
        self.new(Nokogiri.XML(xml_stream).remove_namespaces!, shared_strings)
      end

      private
      def fetch_style_style_id(cell_reference)
        raise 'Invalid Cell Reference!' if cell_reference[/[A-Z]{1,}\d+/].blank?
        row_index = cell_reference.scan(/[A-Z{1,}](\d+)/).flatten.first.to_i - 1
        return if rows[row_index].blank? || rows[row_index][cell_reference].blank?
        rows[row_index][cell_reference].s
      end

    end
  end
end

module OOXML
  class Excel
    class Sheet
      class Column
        attr_accessor :id, :width, :custom_width, :id_range, :hidden
        alias_method :hidden?, :hidden
        def initialize(**attrs)
          attrs.each { |property, value| send("#{property}=", value)}
        end

        def self.load_from_node(column_node)
          hidden_attr = column_node.attributes["hidden"]
          new(id: column_node.attributes["min"].try(:value),
              width: column_node.attributes["width"].try(:value),
              custom_width: column_node.attributes["custom_width"].try(:value),
              id_range: (column_node.attributes["min"].value.to_i..column_node.attributes["max"].value.to_i).to_a,
              hidden: (hidden_attr.present?) ? hidden_attr.value == "1" : false)
        end
      end
    end
  end
end

module OOXML
  class Excel
    class Sheet
      class Row
        attr_accessor :id, :spans, :cells

        def initialize(**attrs)
          attrs.each { |property, value| send("#{property}=", value)}
        end

        def [](id)
          if id.is_a?(String)
            cells.find { |row| row.id == id}
          else
            cells[id]
          end
        end

        def self.load_from_node(row_node, shared_strings)
          new(id: row_node.attributes["r"].try(:value),
              spans: row_node.attributes["spans"].try(:value),
              cells: row_node.xpath('c').map {  |cell_node| Row::Cell.load_from_node(cell_node, shared_strings) } )
        end
      end
    end
  end
end

module OOXML
  class Excel
    class Sheet
      class Row
        class Cell
          attr_accessor :id, :t, :s, :v, :shared_strings
          # t = type
          # v = value
          # s = ??
          def initialize(**attrs)
            attrs.each { |property, value| send("#{property}=", value)}
          end

          def value
            (v.present?) ? shared_strings[v.to_i] : nil
          end

          def self.load_from_node(cell_node, shared_strings)
            new(id: cell_node.attributes["r"].try(:value),
                t: cell_node.attributes["t"].try(:value),
                s: cell_node.attributes["s"].try(:value),
                v: cell_node.at('v').try(:text),
                shared_strings: shared_strings )
          end
        end
      end
    end
  end
end


module OOXML
  class Excel
    class Sheet
      class DataValidation
        attr_accessor :allow_blank, :prompt, :type, :sqref, :formula

        def sqref_range
          @sqref_range ||= begin
            # "BH5:BH271 BI5:BI271"
            sqref.split( ' ').map do |splitted_by_space_sqref|
              # ["BH5:BH271, "BI5:BI271"]
              if splitted_by_space_sqref.is_a?(Array)
                splitted_by_space_sqref.map do |sqref|
                  split_sqref(sqref)
                end
              else
                # "BH5:BH271"
                split_sqref(splitted_by_space_sqref)
              end
            end.flatten.uniq
          end
        end

        def self.load_from_node(data_validation_node)
          allow_blank = data_validation_node.attribute('allowBlank').try(:value)
          prompt = data_validation_node.attribute('prompt').try(:value)
          type = data_validation_node.attribute('type').try(:value)
          sqref = data_validation_node.attribute('sqref').try(:value)
          formula = data_validation_node.at('formula1').try(:content)

          self.new(allow_blank: allow_blank,
                   prompt: prompt,
                   type: type,
                   sqref: sqref,
                   formula: formula)
        end

        private
        def initialize(**attrs)
          attrs.each { |property, value| send("#{property}=", value)}
        end

        def split_sqref(sqref)
          # Example: "BH5:BH271"
          # starting_reference: BH5
          # ending_reference: BH271
          starting_reference, ending_reference = sqref.split(":")

          # if the starting_reference column letters are the same with ending_reference
          # use the first one otherwise use both
          if ending_reference.blank? || starting_reference[/A-Z{1,}/] == ending_reference[/A-Z{1,}/]
            starting_reference
          else
            [starting_reference, ending_reference]
          end
        end
      end
    end
  end
end
