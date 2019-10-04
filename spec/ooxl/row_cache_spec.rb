require 'spec_helper'
require 'pry'

describe OOXL::RowCache do
  let(:second_row_id) { 2 }
  let(:sheet_xml) do
    raw = <<~XML
      <worksheet mc:Ignorable="x14ac">
        <cols>
          <col min="1" max="2" width="8.5703125"/>
          <col min="3" max="3" width="0" hidden="1"/>
          <col min="4" max="1025" width="8.5703125"/>
        </cols>
        <sheetData>
        <row r="1" spans="1:2" x14ac:dyDescent="0.2" ht="102.33">
          <c r="A1" s="1" t="s"><f>VLOOKUP($C$5,TABLES!$A$904:$AG$910,AG5,0)</f></c>
          <c r="B1" s="2" t="s"><v>1</v></c>
          <c r="C1" s="2" t="s"><v>2</v></c>
        </row>
        <row r="#{second_row_id}" spans="1:2" x14ac:dyDescent="0.2">
          <c r="A2" s="3" t="s"><v>2</v></c>
          <c r="B2" s="2" t="s"><v>3</v></c>
          <c r="C3" s="2" t="s"><v>4</v></c>
          <c r="D3" s="2" t="s"><v>5</v></c>
        </row>
        <mergeCells count="1">
          <mergeCell ref="C1:C3"/>
        </mergeCells>
        <dataValidations count="1">
          <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B6">
            <formula1>named_range</formula1>
          </dataValidation>
        </dataValidations>
      </worksheet>
    XML

    Nokogiri.XML(raw).remove_namespaces!
  end

  let(:row_cache) { OOXL::RowCache.new(sheet_xml, []) }

  before do
    allow(OOXL::Row).to receive(:load_from_node).and_call_original
  end

  it "loads the rows" do
    expect(row_cache.rows.count).to eq(2)
    expect(OOXL::Row).to have_received(:load_from_node).twice
    expect(row_cache[2]['B2']).to be_a(OOXL::Cell)
  end

  describe "#row" do
    it "loads only as many rows as necessary" do
      row_cache.row(1)
      expect(OOXL::Row).to have_received(:load_from_node).once
    end

    it "loads more rows on subsequent calls" do
      row_cache.row(1)
      row_cache.row(2)
      expect(OOXL::Row).to have_received(:load_from_node).twice
    end
  end

  describe "#each" do
    it "loads only as many rows as necessary" do
      row_cache.each do |_row|
        break
      end
      expect(OOXL::Row).to have_received(:load_from_node).once
    end

    it "still loads only the first row when it's requested more times" do
      row_cache.each do |_row|
        break
      end

      row_cache.each do |_row|
        break
      end

      row_cache.row(1)

      expect(OOXL::Row).to have_received(:load_from_node).once
    end

    it "loads more rows on subsequent calls" do
      row_cache.each do |_row|
        break
      end

      row_cache.each do
        # nothing
      end

      expect(OOXL::Row).to have_received(:load_from_node).twice
    end

    it "loads more rows on subsequent calls" do
      row_cache.each do |_row|
        break
      end

      row_cache.each do
        # nothing
      end

      expect(OOXL::Row).to have_received(:load_from_node).twice
    end

    context "padded rows" do
      let(:second_row_id) { 3 }
      let(:row_cache) { OOXL::RowCache.new(sheet_xml, [], padded_rows: true) }

      it "fills blanks with empty rows" do
        ids = row_cache.map { |row| row.id }
        expect(ids).to eq(['1', '2', '3'])
      end
    end
  end
end
